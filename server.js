require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const path = require('path');
const multer = require('multer');
const fs = require('fs');
const { OpenAI } = require('openai');
const pdfParse = require('pdf-parse');
const ExcelJS = require('exceljs');
const PDFMerger = require('pdf-merger-js');
const fsPromises = require('fs').promises; // Use promises version for async/await
const { PDFDocument } = require('pdf-lib');
const dayjs = require('dayjs');
const { v4: uuidv4 } = require('uuid');
const archiver = require('archiver');
const cors = require('cors'); // Add CORS middleware
const sessions = new Map(); // Store session data

// Expenses data store - in-memory for demo purposes
let expenses = [];

// Function to save expenses (for demo, just logs a message)
function saveExpenses() {
  console.log("Expenses data updated");
}

// Add debugging - print environment variables
console.log('Environment variables:');
console.log('PORT:', process.env.PORT);
console.log('OPENAI_API_KEY exists:', !!process.env.OPENAI_API_KEY);
if (process.env.OPENAI_API_KEY) {
  console.log('API Key starts with:', process.env.OPENAI_API_KEY.substring(0, 10) + '...');
}

// Set up multer for file uploads
const upload = multer({
  dest: 'uploads/',
  limits: { fileSize: 10 * 1024 * 1024 }, // 10MB limit
  fileFilter: (req, file, cb) => {
    if (file.mimetype === 'application/pdf') {
      cb(null, true);
    } else {
      cb(new Error('Only PDF files are allowed'));
    }
  }
});

// Ensure uploads directory exists
if (!fs.existsSync('uploads')) {
  fs.mkdirSync('uploads');
}

// Initialize OpenAI with the API key from environment variables
const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
  dangerouslyAllowBrowser: true,
});

const app = express();
const PORT = 3005;

// Enable CORS for all routes
app.use(cors());

// Middleware
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// Routes
app.get('/api/expenses', (req, res) => {
  // This would normally fetch from a database
  res.json(expenses);
});

app.post('/api/expenses', (req, res) => {
  // This would normally save to a database
  console.log('New expense:', req.body);
  // Create a new expense with ID and add it to our array
  const newExpense = { 
    id: Date.now(), 
    ...req.body 
  };
  expenses.push(newExpense);
  saveExpenses();
  // Return success response with ID
  res.status(201).json(newExpense);
});

// Fix the PUT route handler for updating expenses
app.put('/api/expenses/:id', (req, res) => {
  const id = parseInt(req.params.id);
  const updatedExpense = req.body;
  
  // Handle mileage calculation if G/L code is 6026-000
  if (updatedExpense.glCode === '6026-000') {
    // Use provided kilometers if available
    const kilometers = parseFloat(updatedExpense.kilometers) || parseFloat(updatedExpense.amount) || 0;
    updatedExpense.kilometers = kilometers;
    updatedExpense.amount = (kilometers * 0.72).toFixed(2);
    updatedExpense.tax = 0; // No tax for mileage
    
    // Store mileage-specific fields
    if (updatedExpense.fromLocation && updatedExpense.toLocation) {
      // These fields are already in the updatedExpense object from the client
      console.log(`Mileage from ${updatedExpense.fromLocation} to ${updatedExpense.toLocation}`);
    }
  }
  
  console.log('Updating expense:', id, updatedExpense);
  
  // Update the expense in expenses array
  const index = expenses.findIndex(e => e.id === id);
  if (index !== -1) {
    expenses[index] = { ...expenses[index], ...updatedExpense };
    saveExpenses();
    return res.json(expenses[index]);
  } 
  
  // If the expense doesn't exist in our array yet, add it
  // This handles the case where expenses were created before our array was initialized
  const newExpense = { id, ...updatedExpense };
  expenses.push(newExpense);
  saveExpenses();
  console.log(`Added expense ${id} to expenses array since it didn't exist`);
  return res.json(newExpense);
});

// Add helper function to clean up uploads directory
async function cleanUploadsDirectory(oldestAgeInMinutes = 120) {
  const uploadsDir = path.join(__dirname, 'uploads');
  
  try {
    // Check if directory exists
    await fsPromises.access(uploadsDir, fs.constants.F_OK);
    
    // Get all files
    const files = await fsPromises.readdir(uploadsDir);
    const now = Date.now();
    
    // Only remove files older than specified age
    const cutoffTime = now - (oldestAgeInMinutes * 60 * 1000);
    
    for (const file of files) {
      const filePath = path.join(uploadsDir, file);
      try {
        const stats = await fsPromises.stat(filePath);
        
        // Check if older than cutoff
        if (stats.mtime.getTime() < cutoffTime) {
          await fsPromises.unlink(filePath);
          console.log(`Removed old file: ${file}`);
        }
      } catch (err) {
        console.error(`Error removing file ${file}:`, err);
      }
    }
    
    console.log('Cleanup of uploads directory completed');
    
  } catch (err) {
    console.error('Error cleaning uploads directory:', err);
  }
}

// PDF Upload and Processing Route for multiple files
app.post('/api/upload-pdf', upload.array('pdfFiles', 50), async (req, res) => {
    try {
        if (!req.files || req.files.length === 0) {
            return res.status(400).json({ error: 'No PDF files uploaded' });
        }

        // Extract name and department from form data
        const userName = req.body.userName || '';
        const userDepartment = req.body.userDepartment || '';
        
        // Validate required fields
        if (!userName.trim() || !userDepartment.trim()) {
            return res.status(400).json({ error: 'Name and Department are required fields' });
        }

        const sessionId = Date.now().toString(); // Simple timestamp-based session ID
        sessions.set(sessionId, {
            files: [],
            createdAt: Date.now()
        });

        const results = [];
        const errors = [];

        // Process files in parallel for maximum speed (batches of 10)
        const BATCH_SIZE = 10; // Process 10 PDFs simultaneously for faster throughput
        
        for (let i = 0; i < req.files.length; i += BATCH_SIZE) {
            const batch = req.files.slice(i, i + BATCH_SIZE);
            
            // Process batch in parallel
            const batchPromises = batch.map(async (file) => {
            try {
                // Read PDF file
                const pdfBuffer = fs.readFileSync(file.path);
                const pdfData = await pdfParse(pdfBuffer);
                
                // Extract text content from PDF
                const pdfText = pdfData.text;
                
                // Process with OpenAI API - Using GPT-4o for maximum intelligence and speed
                const openaiResponse = await openai.chat.completions.create({
                    model: "gpt-4o",
                    messages: [
                        {
                            role: "system", 
                            content: `You are an advanced AI assistant specialized in extracting and categorizing expense information from receipts and invoices for Greenwin Corp. Analyze documents with precision and intelligence, automatically detecting patterns and context. Your task is to extract key expense details with the highest accuracy possible.

                            REQUIRED FIELDS (all must be present):
                            - date: Extract in YYYY-MM-DD format. If only month/day are shown, use the current year.
                            - merchant: The vendor/company name from the receipt
                            - amount: Total amount including tax (as a number)
                            - tax: HST/GST amount if explicitly shown (as a number)
                            - description: Brief summary of purchased items/services
                            - gl_code: Most appropriate G/L code from the provided list
                            
                            OPTIONAL FIELDS (include if found):
                            - line_items: Detailed list of purchased items with amounts

                            G/L CODE SELECTION RULES:
                            6408-000 (Office & General):
                            - Office supplies, equipment, furniture
                            - Administrative expenses
                            - General business supplies
                            - Cleaning supplies
                            - Printer supplies

                            6402-000 (Membership):
                            - Professional association fees
                            - Industry memberships
                            - Chamber of commerce
                            - Trade organization dues

                            6404-000 (Subscriptions):
                            - Software licenses
                            - Cloud services
                            - Digital subscriptions
                            - Online tools
                            - Microsoft/Adobe products
                            - Zoom/Teams subscriptions

                            7335-000 (Education & Development):
                            - Training courses
                            - Professional certifications
                            - Workshops
                            - Conferences
                            - Educational materials
                            - Skill development programs

                            6026-000 (Mileage/ETR):
                            - Vehicle mileage
                            - Toll fees
                            - Parking fees
                            - Public transit
                            - Vehicle maintenance

                            6010-000 (Food & Entertainment):
                            - Business meals
                            - Client lunches/dinners
                            - Coffee meetings
                            - Catering
                            - Restaurant expenses
                            - Food for meetings

                            6011-000 (Social):
                            - Team building events
                            - Company celebrations
                            - Holiday parties
                            - Employee events
                            - Team activities

                            6012-000 (Travel Expenses):
                            - Airfare
                            - Hotels
                            - Taxis/Uber/Lyft
                            - Car rentals
                            - Travel insurance
                            - Baggage fees

                            CRITICAL INTELLIGENCE RULES:
                            1. SMART TAX DETECTION: For Canadian receipts, intelligently identify HST/GST/PST. If not shown, calculate based on regional rates.
                            2. CONTEXT AWARENESS: Analyze receipt context - restaurant, retail, service, etc. - to improve categorization.
                            3. DATE INTELLIGENCE: For incomplete dates, infer the most likely year/month based on context.
                            4. MERCHANT RECOGNITION: Extract official business name and clean up formatting (remove extra spaces, special characters).
                            5. CONCISE DESCRIPTIONS: Write brief, one-line descriptions (max 8-10 words). Be clear and professional but concise. Focus on:
                               - Main service/product only
                               - Key purpose in 5-10 words
                               - No lists or excessive details
                               Example: "September Azure cloud services and Veeam backup"
                               NOT: "Monthly billing for September including Azure Storage, Data Factory, SQL Database, Microsoft Defender, Virtual Machines, Virtual Network, Bandwidth, and Veeam Backup"
                            6. ADVANCED G/L MATCHING: Use semantic understanding to assign the most accurate G/L code based on merchant type AND items purchased.
                            7. PATTERN RECOGNITION: Learn from common expense patterns (e.g., coffee shop = Food & Entertainment).
                            8. CONFIDENCE: Only include fields you're highly confident about. Skip uncertain data rather than guessing.

                            Format your response as a JSON object with these exact field names:
                            {
                                "date": "YYYY-MM-DD",
                                "merchant": "string",
                                "amount": number,
                                "tax": number (optional),
                                "description": "string",
                                "gl_code": "string",
                                "location": "string" (optional),
                                "line_items": [
                                    {
                                        "name": "string",
                                        "amount": number
                                    }
                                ]
                            }`
                        },
                        {
                            role: "user",
                            content: `Intelligently analyze and extract expense information from this receipt/invoice. Use context clues, pattern recognition, and semantic understanding. 
                            
                            CRITICAL: Keep descriptions SHORT - max 8-10 words, one line only. Be concise and professional. Example: "Azure cloud services for September" NOT long lists.
                            
                            Return precise, structured data in the exact JSON format specified:\n\n${pdfText}`
                        }
                    ],
                    response_format: { type: "json_object" },
                    temperature: 0.1, // Low for concise, consistent descriptions
                    max_tokens: 500, // Reduced for short descriptions
                    top_p: 0.95 // Focus on most likely tokens for faster processing
                });
                
                // Parse OpenAI response with error handling
                let parsedData;
                try {
                    parsedData = JSON.parse(openaiResponse.choices[0].message.content);
                    
                    // Validate required fields
                    const requiredFields = ['date', 'merchant', 'amount', 'description', 'gl_code'];
                    const missingFields = requiredFields.filter(field => !parsedData[field]);
                    
                    if (missingFields.length > 0) {
                        throw new Error(`Missing required fields: ${missingFields.join(', ')}`);
                    }
                    
                    // Clean and validate data
                    parsedData.amount = parseFloat(parsedData.amount.toString().replace(/[^0-9.]/g, ''));
                    if (isNaN(parsedData.amount)) {
                        throw new Error('Invalid amount format');
                    }
                    
                    if (parsedData.tax) {
                        parsedData.tax = parseFloat(parsedData.tax.toString().replace(/[^0-9.]/g, ''));
                        if (isNaN(parsedData.tax)) {
                            delete parsedData.tax; // Remove invalid tax amount
                        }
                    }
                    
                    // Validate date format
                    const dateRegex = /^\d{4}-\d{2}-\d{2}$/;
                    if (!dateRegex.test(parsedData.date)) {
                        throw new Error('Invalid date format');
                    }
                    
                    // Validate G/L code format
                    const glCodeRegex = /^\d{4}-\d{3}$/;
                    if (!glCodeRegex.test(parsedData.gl_code)) {
                        throw new Error('Invalid G/L code format');
                    }
                    
                } catch (parseError) {
                    console.error('Error parsing OpenAI response:', parseError);
                    throw new Error(`Failed to parse expense data: ${parseError.message}`);
                }
                
                // Add user name and department to the parsed data
                parsedData.name = userName;
                parsedData.department = userDepartment;
                
                // Add tax amount to the data if it exists in the parsed response
                if (parsedData.tax) {
                    parsedData.tax = parseFloat(parsedData.tax.toString().replace(/[^0-9.]/g, ''));
                }
                
                // Add GL code to the data if it exists
                if (parsedData.gl_code) {
                    parsedData.glCode = parsedData.gl_code;
                }
                
                // Add file to session tracking
                sessions.get(sessionId).files.push(file.path);
                console.log(`PDF file saved at: ${file.path}`);

                // Return the result for this file
                return {
                    filename: file.originalname,
                    filepath: file.path,
                    data: parsedData
                };
                
            } catch (error) {
                console.error(`Error processing PDF file ${file.originalname}:`, error);
                return {
                    error: true,
                    filename: file.originalname,
                    message: error.message || 'Error processing PDF file'
                };
            }
            });
            
            // Wait for all files in batch to complete
            const batchResults = await Promise.all(batchPromises);
            
            // Separate successful results from errors
            batchResults.forEach(result => {
                if (result.error) {
                    errors.push({
                        filename: result.filename,
                        error: result.message
                    });
                } else {
                    results.push(result);
                }
            });
        }

        res.json({ 
            success: true, 
            results: results,
            errors: errors,
            processedCount: results.length,
            errorCount: errors.length,
            totalCount: req.files.length,
            sessionId: sessionId
        });
    } catch (error) {
        console.error('Error in PDF upload handler:', error);
        
        // Only clean up on error
        if (req.files) {
            for (const file of req.files) {
                if (fs.existsSync(file.path)) {
                    fs.unlinkSync(file.path);
                }
            }
        }
        
        res.status(500).json({
            success: false,
            error: error.message || 'Error processing PDF files'
        });
    }
});

// Export expenses to Excel
app.post('/api/export-excel', async (req, res) => {
  try {
    const { expenses, signature } = req.body;
    
    if (!expenses || !Array.isArray(expenses) || expenses.length === 0) {
      return res.status(400).json({ success: false, error: 'No expense data provided' });
    }
    
    // Create a new Excel workbook
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Greenwin Corp Expense Report App';
    workbook.created = new Date();
    
    // Add a worksheet with properties
    const worksheet = workbook.addWorksheet('Expense Report', {
      properties: {
        defaultRowHeight: 18
      },
      pageSetup: {
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0, // Setting to 0 means "automatically determine the number of pages"
        paperSize: 9, // 9 = A4
        orientation: 'portrait',
        margins: {
          left: 0.7,
          right: 0.7,
          top: 0.75,
          bottom: 0.75,
          header: 0.3,
          footer: 0.3
        }
      }
    });
    
    // Set column widths for professional layout
    worksheet.columns = [
      { width: 13 }, // A - DATE
      { width: 55 }, // B - DESCRIPTION (Merchant: Description combined)
      { width: 16 }, // C - 6408-000
      { width: 15 }, // D - 6402-000
      { width: 16 }, // E - 6404-000
      { width: 19 }, // F - 7335-000
      { width: 15 }, // G - 6026-000
    ];
    
    // -- HEADER SECTION --
    
    // Top section - Name and Date Submitted
    const nameCell = worksheet.getCell('A1');
    nameCell.value = 'Name:';
    nameCell.font = { bold: true };
    nameCell.border = {
      top: { style: 'medium' },
      left: { style: 'medium' },
      bottom: { style: 'medium' }
    };
    
    // Cell for name input
    worksheet.mergeCells('B1:C1');
    const nameInputCell = worksheet.getCell('B1');
    nameInputCell.value = expenses[0]?.name || '';
    nameInputCell.border = {
      top: { style: 'medium' },
      right: { style: 'medium' },
      bottom: { style: 'medium' }
    };
    
    // Date submitted
    const dateSubmittedCell = worksheet.getCell('E1');
    dateSubmittedCell.value = 'Date Submitted:';
    dateSubmittedCell.font = { bold: true };
    dateSubmittedCell.border = {
      top: { style: 'medium' },
      left: { style: 'medium' },
      bottom: { style: 'medium' }
    };
    
    // Cell for date input
    worksheet.mergeCells('F1:G1');
    const dateInputCell = worksheet.getCell('F1');
    dateInputCell.value = new Date();
    dateInputCell.numFmt = 'mmmm d yyyy';
    dateInputCell.border = {
      top: { style: 'medium' },
      right: { style: 'medium' },
      bottom: { style: 'medium' }
    };
    
    // Division to be charged
    const divisionCell = worksheet.getCell('A2');
    divisionCell.value = 'Division to be charged:';
    divisionCell.font = { bold: true };
    divisionCell.border = {
      left: { style: 'medium' },
      bottom: { style: 'thin' }
    };
    
    // Division field spans multiple columns
    worksheet.mergeCells('B2:G2');
    const divisionInputCell = worksheet.getCell('B2');
    divisionInputCell.value = expenses[0]?.department || '';
    divisionInputCell.border = {
      right: { style: 'medium' },
      bottom: { style: 'thin' }
    };
    
    // Team # and BLDG#
    const teamCell = worksheet.getCell('C3');
    teamCell.value = 'Team #:';
    teamCell.font = { bold: true };
    teamCell.alignment = { horizontal: 'right' };
    
    const teamInputCell = worksheet.getCell('D3');
    teamInputCell.border = {
      bottom: { style: 'medium' }
    };
    
    const bldgCell = worksheet.getCell('C4');
    bldgCell.value = 'BLDG#:';
    bldgCell.font = { bold: true };
    bldgCell.alignment = { horizontal: 'right' };
    
    const bldgInputCell = worksheet.getCell('D4');
    bldgInputCell.border = {
      bottom: { style: 'medium' }
    };
    
    // Add borders to the right side of the header
    for (let row = 3; row <= 4; row++) {
      worksheet.getCell(`A${row}`).border = {
        left: { style: 'medium' }
      };
      worksheet.getCell(`G${row}`).border = {
        right: { style: 'medium' }
      };
    }
    
    // Complete border for the header section
    worksheet.getCell('A5').border = {
      left: { style: 'medium' },
      bottom: { style: 'medium' }
    };
    
    worksheet.mergeCells('B5:G5');
    worksheet.getCell('B5').border = {
      right: { style: 'medium' },
      bottom: { style: 'medium' }
    };
    
    // -- PROMOTION EXPENSES SECTION --
    
    const promoRow = 6;
    
    // Promotion Expenses Header - Modern gradient-style design
    worksheet.mergeCells(`A${promoRow}:B${promoRow}`);
    const promoHeaderCell = worksheet.getCell(`A${promoRow}`);
    promoHeaderCell.value = 'PROMOTION EXPENSES';
    promoHeaderCell.alignment = { horizontal: 'center', vertical: 'middle' };
    promoHeaderCell.font = { bold: true, size: 12, color: { argb: 'FFFFFF' }, name: 'Calibri' };
    promoHeaderCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '667EEA' } // Modern purple (matches UI gradient)
    };
    promoHeaderCell.border = {
      top: { style: 'medium' },
      left: { style: 'medium' },
      bottom: { style: 'medium' }
    };
    worksheet.getRow(promoRow).height = 18; // Compact header
    
    // G/L Allocation Header - Modern gradient-style design
    worksheet.mergeCells(`C${promoRow}:G${promoRow}`);
    const glHeaderCell = worksheet.getCell(`C${promoRow}`);
    glHeaderCell.value = 'G/L ALLOCATION';
    glHeaderCell.alignment = { horizontal: 'center', vertical: 'middle' };
    glHeaderCell.font = { bold: true, size: 12, color: { argb: 'FFFFFF' }, name: 'Calibri' };
    glHeaderCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '667EEA' } // Modern purple (matches UI gradient)
    };
    glHeaderCell.border = {
      top: { style: 'medium' },
      right: { style: 'medium' },
      bottom: { style: 'medium' }
    };
    
    // G/L Codes Row
    const glCodesRow = promoRow + 1;
    worksheet.getCell(`A${glCodesRow}`).border = {
      left: { style: 'medium' }
    };
    worksheet.getCell(`B${glCodesRow}`).border = {
      right: { style: 'thin' }
    };
    
    worksheet.getCell(`C${glCodesRow}`).value = '6408-000';
    worksheet.getCell(`C${glCodesRow}`).alignment = { horizontal: 'center' };
    worksheet.getCell(`C${glCodesRow}`).font = { bold: true, size: 10, name: 'Calibri' };
    worksheet.getCell(`C${glCodesRow}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'E8EEFF' } // Light purple background
    };
    worksheet.getCell(`C${glCodesRow}`).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' },
      bottom: { style: 'thin' }
    };
    
    worksheet.getCell(`D${glCodesRow}`).value = '6402-000';
    worksheet.getCell(`D${glCodesRow}`).alignment = { horizontal: 'center' };
    worksheet.getCell(`D${glCodesRow}`).font = { bold: true, size: 10, name: 'Calibri' };
    worksheet.getCell(`D${glCodesRow}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'E8EEFF' } // Light purple background
    };
    worksheet.getCell(`D${glCodesRow}`).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' },
      bottom: { style: 'thin' }
    };
    
    worksheet.getCell(`E${glCodesRow}`).value = '6404-000';
    worksheet.getCell(`E${glCodesRow}`).alignment = { horizontal: 'center' };
    worksheet.getCell(`E${glCodesRow}`).font = { bold: true, size: 10, name: 'Calibri' };
    worksheet.getCell(`E${glCodesRow}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'E8EEFF' } // Light purple background
    };
    worksheet.getCell(`E${glCodesRow}`).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' },
      bottom: { style: 'thin' }
    };
    
    worksheet.getCell(`F${glCodesRow}`).value = '7335-000';
    worksheet.getCell(`F${glCodesRow}`).alignment = { horizontal: 'center' };
    worksheet.getCell(`F${glCodesRow}`).font = { bold: true, size: 10, name: 'Calibri' };
    worksheet.getCell(`F${glCodesRow}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'E8EEFF' } // Light purple background
    };
    worksheet.getCell(`F${glCodesRow}`).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' },
      bottom: { style: 'thin' }
    };
    
    worksheet.getCell(`G${glCodesRow}`).border = {
      right: { style: 'medium' }
    };
    
    // Category Labels Row - Professional styling (removed gratuity row)
    const categoryRow = glCodesRow + 1;
    worksheet.getRow(categoryRow).height = 18; // Compact
    
    worksheet.getCell(`A${categoryRow}`).value = 'DATE';
    worksheet.getCell(`A${categoryRow}`).alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell(`A${categoryRow}`).font = { bold: true, size: 11, name: 'Calibri' };
    worksheet.getCell(`A${categoryRow}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'F5F5F7' } // Light gray background
    };
    worksheet.getCell(`A${categoryRow}`).border = {
      left: { style: 'medium' },
      top: { style: 'thin' },
      right: { style: 'thin' },
      bottom: { style: 'medium' }
    };
    
    worksheet.getCell(`B${categoryRow}`).value = 'DESCRIPTION';
    worksheet.getCell(`B${categoryRow}`).alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell(`B${categoryRow}`).font = { bold: true, size: 11, name: 'Calibri' };
    worksheet.getCell(`B${categoryRow}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'F5F5F7' } // Light gray background
    };
    worksheet.getCell(`B${categoryRow}`).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' },
      bottom: { style: 'medium' }
    };
    
    worksheet.getCell(`C${categoryRow}`).value = 'Office &\nGeneral';
    worksheet.getCell(`C${categoryRow}`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    worksheet.getCell(`C${categoryRow}`).font = { bold: true, size: 10, name: 'Calibri' };
    worksheet.getCell(`C${categoryRow}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'F5F5F7' } // Light gray background
    };
    worksheet.getCell(`C${categoryRow}`).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' },
      bottom: { style: 'medium' }
    };
    
    worksheet.getCell(`D${categoryRow}`).value = 'Membership';
    worksheet.getCell(`D${categoryRow}`).alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell(`D${categoryRow}`).font = { bold: true, size: 10, name: 'Calibri' };
    worksheet.getCell(`D${categoryRow}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'F5F5F7' } // Light gray background
    };
    worksheet.getCell(`D${categoryRow}`).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' },
      bottom: { style: 'medium' }
    };
    
    worksheet.getCell(`E${categoryRow}`).value = 'Subscriptions';
    worksheet.getCell(`E${categoryRow}`).alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell(`E${categoryRow}`).font = { bold: true, size: 10, name: 'Calibri' };
    worksheet.getCell(`E${categoryRow}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'F5F5F7' } // Light gray background
    };
    worksheet.getCell(`E${categoryRow}`).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' },
      bottom: { style: 'medium' }
    };
    
    worksheet.getCell(`F${categoryRow}`).value = 'Education &\nDevelopment';
    worksheet.getCell(`F${categoryRow}`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    worksheet.getCell(`F${categoryRow}`).font = { bold: true, size: 10, name: 'Calibri' };
    worksheet.getCell(`F${categoryRow}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'F5F5F7' } // Light gray background
    };
    worksheet.getCell(`F${categoryRow}`).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' },
      bottom: { style: 'medium' }
    };
    
    worksheet.getCell(`G${categoryRow}`).value = 'Mileage/ETR';
    worksheet.getCell(`G${categoryRow}`).alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell(`G${categoryRow}`).font = { bold: true, size: 10, name: 'Calibri' };
    worksheet.getCell(`G${categoryRow}`).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'F5F5F7' } // Light gray background
    };
    worksheet.getCell(`G${categoryRow}`).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'medium' },
      bottom: { style: 'medium' }
    };
    
    // Add empty rows for Promotion expenses
    let promoRowsStart = categoryRow + 1;
    const promoRows = 10; // Number of empty rows for promotion expenses
    
    for (let i = 0; i < promoRows; i++) {
      const row = worksheet.getRow(promoRowsStart + i);
      
      // Date cell
      row.getCell(1).border = {
        left: { style: 'medium' },
        top: { style: 'thin' },
        right: { style: 'thin' },
        bottom: { style: 'thin' }
      };
      
      // Description cell
      row.getCell(2).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' },
        bottom: { style: 'thin' }
      };
      row.getCell(2).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
      
      // Amount cells (C-F)
      for (let col = 3; col <= 6; col++) {
        row.getCell(col).border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          right: { style: 'thin' },
          bottom: { style: 'thin' }
        };
        row.getCell(col).numFmt = '$#,##0.00';
        row.getCell(col).alignment = { horizontal: 'right' };
      }
      
      // Right border
      row.getCell(7).border = {
        right: { style: 'medium' }
      };
    }
    
    // Promotion Summary Rows
    const promoSummaryStart = promoRowsStart + promoRows;
    
    // Total Promotion Expenses row with SUM formulas
    const row1 = worksheet.getRow(promoSummaryStart);
    row1.getCell(1).border = { left: { style: 'medium' } };
    row1.getCell(2).value = 'Total Promotion Expenses (incl. HST)';
    row1.getCell(2).alignment = { horizontal: 'right' };
    row1.getCell(2).font = { bold: true };
    
    for (let col = 3; col <= 6; col++) {
      const colLetter = String.fromCharCode(64 + col); // Convert to column letter (C, D, E, F)
      row1.getCell(col).value = {
        formula: `SUM(${colLetter}${promoRowsStart}:${colLetter}${promoSummaryStart-1})`
      };
      row1.getCell(col).numFmt = '$#,##0.00';
      row1.getCell(col).alignment = { horizontal: 'right' };
      row1.getCell(col).font = { bold: true };
      row1.getCell(col).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' },
        bottom: { style: 'thin' }
      };
    }
    
    row1.getCell(7).border = { right: { style: 'medium' } };
    
    // HST row
    const row2 = worksheet.getRow(promoSummaryStart + 1);
    row2.getCell(1).border = { left: { style: 'medium' } };
    row2.getCell(2).value = 'HST (G/L 2325-000)';
    row2.getCell(2).alignment = { horizontal: 'right' };
    
    for (let col = 3; col <= 6; col++) {
      row2.getCell(col).value = '-';
      row2.getCell(col).alignment = { horizontal: 'center' };
      row2.getCell(col).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' },
        bottom: { style: 'thin' }
      };
    }
    
    row2.getCell(7).border = { right: { style: 'medium' } };
    
    // Net Amount row
    const row3 = worksheet.getRow(promoSummaryStart + 2);
    row3.getCell(1).border = { left: { style: 'medium' } };
    row3.getCell(2).value = 'Net Amount (before HST)';
    row3.getCell(2).alignment = { horizontal: 'right' };
    
    for (let col = 3; col <= 6; col++) {
      row3.getCell(col).value = '-';
      row3.getCell(col).alignment = { horizontal: 'center' };
      row3.getCell(col).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' },
        bottom: { style: 'thin' }
      };
    }
    
    row3.getCell(7).border = { right: { style: 'medium' } };
    
    // TOTAL PROMOTION with SUM formula across all promotion categories
    const totalPromoRow = promoSummaryStart + 3;
    worksheet.mergeCells(`A${totalPromoRow}:F${totalPromoRow}`);
    worksheet.getCell(`A${totalPromoRow}`).value = 'TOTAL PROMOTION';
    worksheet.getCell(`A${totalPromoRow}`).alignment = { horizontal: 'right' };
    worksheet.getCell(`A${totalPromoRow}`).font = { bold: true };
    worksheet.getCell(`A${totalPromoRow}`).border = {
      left: { style: 'medium' },
      right: { style: 'thin' },
      bottom: { style: 'thin' }
    };
    
    // Add totals formula for TOTAL PROMOTION
    worksheet.getCell(`G${totalPromoRow}`).value = {
      formula: `SUM(C${promoSummaryStart}:F${promoSummaryStart})`
    };
    worksheet.getCell(`G${totalPromoRow}`).numFmt = '$#,##0.00';
    worksheet.getCell(`G${totalPromoRow}`).font = { bold: true };
    worksheet.getCell(`G${totalPromoRow}`).alignment = { horizontal: 'right' };
    worksheet.getCell(`G${totalPromoRow}`).border = {
      right: { style: 'medium' },
      bottom: { style: 'thin' }
    };
    
    // -- OTHER EXPENSES SECTION --
    
    const otherRow = totalPromoRow + 1;
    
    // Other Expenses Header
    worksheet.mergeCells(`A${otherRow}:B${otherRow}`);
    const otherHeaderCell = worksheet.getCell(`A${otherRow}`);
    otherHeaderCell.value = 'OTHER EXPENSES';
    otherHeaderCell.alignment = { horizontal: 'center', vertical: 'middle' };
    otherHeaderCell.font = { bold: true, color: { argb: 'FFFFFF' } };
    otherHeaderCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '00A651' } // Green background
    };
    otherHeaderCell.border = {
      left: { style: 'medium' },
      bottom: { style: 'medium' }
    };
    
    // G/L Allocation Header for Other Expenses
    worksheet.mergeCells(`C${otherRow}:G${otherRow}`);
    const otherGlHeaderCell = worksheet.getCell(`C${otherRow}`);
    otherGlHeaderCell.value = 'G/L ALLOCATION';
    otherGlHeaderCell.alignment = { horizontal: 'center', vertical: 'middle' };
    otherGlHeaderCell.font = { bold: true, color: { argb: 'FFFFFF' } };
    otherGlHeaderCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '00A651' } // Green background
    };
    otherGlHeaderCell.border = {
      right: { style: 'medium' },
      bottom: { style: 'medium' }
    };
    
    // G/L Codes Row for Other Expenses
    const otherGlCodesRow = otherRow + 1;
    
    worksheet.getCell(`A${otherGlCodesRow}`).border = {
      left: { style: 'medium' }
    };
    
    worksheet.getCell(`B${otherGlCodesRow}`).border = {
      right: { style: 'thin' }
    };
    
    const otherGlCodes = [
      { cell: 'C', code: '6408-000' },
      { cell: 'D', code: '6402-000' },
      { cell: 'E', code: '6404-000' },
      { cell: 'F', code: '7335-000' },
      { cell: 'G', code: '6026-000' }
    ];
    
    otherGlCodes.forEach(item => {
      worksheet.getCell(`${item.cell}${otherGlCodesRow}`).value = item.code;
      worksheet.getCell(`${item.cell}${otherGlCodesRow}`).alignment = { horizontal: 'center' };
      worksheet.getCell(`${item.cell}${otherGlCodesRow}`).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' },
        bottom: { style: 'thin' }
      };
    });
    
    worksheet.getCell(`G${otherGlCodesRow}`).border.right = { style: 'medium' };
    
    // Category Labels Row for Other Expenses
    const otherCategoryRow = otherGlCodesRow + 1;
    
    worksheet.getCell(`A${otherCategoryRow}`).value = 'DATE';
    worksheet.getCell(`A${otherCategoryRow}`).alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell(`A${otherCategoryRow}`).font = { bold: true };
    worksheet.getCell(`A${otherCategoryRow}`).border = {
      left: { style: 'medium' },
      top: { style: 'thin' },
      right: { style: 'thin' },
      bottom: { style: 'thin' }
    };
    
    worksheet.getCell(`B${otherCategoryRow}`).value = 'DESCRIPTION';
    worksheet.getCell(`B${otherCategoryRow}`).alignment = { horizontal: 'center', vertical: 'middle' };
    worksheet.getCell(`B${otherCategoryRow}`).font = { bold: true };
    worksheet.getCell(`B${otherCategoryRow}`).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' },
      bottom: { style: 'thin' }
    };
    
    const otherCategories = [
      { cell: 'C', name: 'Office & General' },
      { cell: 'D', name: 'Membership' },
      { cell: 'E', name: 'Subscriptions' },
      { cell: 'F', name: 'Education &\nDevelopment' },
      { cell: 'G', name: 'Mileage/ ETR' }
    ];
    
    otherCategories.forEach(item => {
      worksheet.getCell(`${item.cell}${otherCategoryRow}`).value = item.name;
      worksheet.getCell(`${item.cell}${otherCategoryRow}`).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      worksheet.getCell(`${item.cell}${otherCategoryRow}`).font = { bold: true };
      worksheet.getCell(`${item.cell}${otherCategoryRow}`).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' },
        bottom: { style: 'thin' }
      };
    });
    
    worksheet.getCell(`G${otherCategoryRow}`).border.right = { style: 'medium' };
    
    // Add data rows for expenses
    let otherRowsStart = otherCategoryRow + 1;
    
    // Group expenses by G/L code
    const expensesByGlCode = {};
    expenses.forEach(expense => {
      const glCode = expense.glCode || '';
      if (!expensesByGlCode[glCode]) {
        expensesByGlCode[glCode] = [];
      }
      expensesByGlCode[glCode].push(expense);
    });
    
    // Add expense data
    let currentRow = otherRowsStart;
    expenses.forEach(expense => {
        const row = worksheet.getRow(currentRow);
        
        // Format date
        let dateValue = expense.date;
        if (typeof dateValue === 'string') {
            try {
                const dateObj = new Date(dateValue);
                if (!isNaN(dateObj.getTime())) {
                    dateValue = dateObj;
                }
            } catch (e) {
                console.warn('Could not parse date:', e);
            }
        }
        
        // DATE column
        row.getCell(1).value = dateValue;
        row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
        row.getCell(1).border = {
            left: { style: 'medium' },
            top: { style: 'thin' },
            right: { style: 'thin' },
            bottom: { style: 'thin' }
        };
        
        // DESCRIPTION column - combine merchant and description
        const merchantName = expense.title || expense.merchant || '';
        let description = merchantName ? `${merchantName}: ${expense.description || ''}` : (expense.description || '');
        if (expense.location) {
            description = `${description}${description ? '\n\n' : ''}Location: ${expense.location}`;
        }
        
        // Check if this is a split expense with splits
        if (expense.splits && expense.splits.length > 0) {
            // Calculate total amount allocated to splits
            const totalSplitAmount = expense.splits.reduce((sum, split) => sum + parseFloat(split.amount || 0), 0);
            const remainingAmount = parseFloat(expense.amount) - totalSplitAmount;
            const remainingPercent = Math.round((remainingAmount / parseFloat(expense.amount)) * 100);
            
            // Start building the split description
            let splitDescription = `\n\n`;
            
            // Add primary G/L allocation if there's a remaining amount
            if (remainingAmount > 0) {
                // Get name for G/L code
                let glName = "Other";
                switch (expense.glCode) {
                    case '6408-000': glName = "Office & General"; break;
                    case '6402-000': glName = "Membership"; break;
                    case '6404-000': glName = "Subscriptions"; break;
                    case '7335-000': glName = "Education & Development"; break;
                    case '6026-000': glName = "Mileage/ETR"; break;
                    case '6010-000': glName = "Food & Entertainment"; break;
                    case '6011-000': glName = "Social"; break;
                    case '6012-000': glName = "Travel"; break;
                }
                
                splitDescription += `Primary (${expense.glCode}): $${remainingAmount.toFixed(2)} (${remainingPercent}%)\n`;
            }
            
            // Add each split allocation on a new line
            expense.splits.forEach((split, index) => {
                // Get name for G/L code
                let glName = "Other";
                switch (split.glCode) {
                    case '6408-000': glName = "Office & General"; break;
                    case '6402-000': glName = "Membership"; break;
                    case '6404-000': glName = "Subscriptions"; break;
                    case '7335-000': glName = "Education & Development"; break;
                    case '6026-000': glName = "Mileage/ETR"; break;
                    case '6010-000': glName = "Food & Entertainment"; break;
                    case '6011-000': glName = "Social"; break;
                    case '6012-000': glName = "Travel"; break;
                }
                
                splitDescription += `${split.glCode}: $${parseFloat(split.amount).toFixed(2)} (${Math.round(parseFloat(split.percentage))}%)\n`;
            });
            
            // Add the split description to the main description
            description += splitDescription;
        }
        
        // Apply rich text formatting to make merchant name bold
        if (merchantName && description.startsWith(merchantName + ':')) {
            // Use rich text format: bold merchant, normal description
            const descriptionPart = description.substring(merchantName.length);
            row.getCell(2).value = {
                richText: [
                    { font: { bold: true }, text: merchantName },
                    { text: descriptionPart }
                ]
            };
        } else {
            row.getCell(2).value = description;
        }
        row.getCell(2).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
        row.getCell(2).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' },
            bottom: { style: 'thin' }
        };
        
        // Check if this is a split expense
        let hasPromotionGLCode = false;
        
        if (expense.splits && expense.splits.length > 0) {
            // Check if any split has a promotion G/L code
            for (const split of expense.splits) {
                if (['6010-000', '6011-000', '6012-000'].includes(split.glCode)) {
                    hasPromotionGLCode = true;
                    break;
                }
            }
            
            // Calculate total amount allocated to splits
            const totalSplitAmount = expense.splits.reduce((sum, split) => sum + parseFloat(split.amount || 0), 0);
            const remainingAmount = parseFloat(expense.amount) - totalSplitAmount;
            
            // First, allocate the remaining amount to the primary G/L code
            if (remainingAmount > 0) {
                const primaryGlColumn = getColumnForGlCode(expense.glCode);
                if (primaryGlColumn > 0) {
                    row.getCell(primaryGlColumn).value = remainingAmount;
                    row.getCell(primaryGlColumn).numFmt = '$#,##0.00';
                    row.getCell(primaryGlColumn).alignment = { horizontal: 'right', vertical: 'middle' };
                }
            }
            
            // Then allocate each split amount to its respective G/L code
            expense.splits.forEach(split => {
                const splitGlColumn = getColumnForGlCode(split.glCode);
                if (splitGlColumn > 0 && split.amount > 0) {
                    // If this is the first time we're adding to this column, initialize it
                    const currentValue = row.getCell(splitGlColumn).value || 0;
                    row.getCell(splitGlColumn).value = currentValue + parseFloat(split.amount);
                    row.getCell(splitGlColumn).numFmt = '$#,##0.00';
                    row.getCell(splitGlColumn).alignment = { horizontal: 'right', vertical: 'middle' };
                } else if (['6010-000', '6011-000', '6012-000'].includes(split.glCode)) {
                    // For promotion G/L codes, we need to create a separate row in the promotion section
                    addPromotionExpense(worksheet, expense, split);
                }
            });
        } else {
            // Not a split transaction, allocate the full amount to the expense's G/L code
            const amountColumn = getColumnForGlCode(expense.glCode);
            if (amountColumn > 0) {
                row.getCell(amountColumn).value = expense.amount;
                row.getCell(amountColumn).numFmt = '$#,##0.00';
                row.getCell(amountColumn).alignment = { horizontal: 'right', vertical: 'middle' };
            } else if (['6010-000', '6011-000', '6012-000'].includes(expense.glCode)) {
                // This is a promotion expense, add it to the promotion section
                hasPromotionGLCode = true;
            }
        }
        
        // Skip adding to Other Expenses if it's only a promotion expense
        if (!hasPromotionGLCode || (expense.splits && expense.splits.some(split => !['6010-000', '6011-000', '6012-000'].includes(split.glCode)))) {
            // Add borders to all cells in the row (columns C-G for GL codes)
            for (let col = 3; col <= 7; col++) {
                row.getCell(col).border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    right: { style: 'thin' },
                    bottom: { style: 'thin' }
                };
            }
            
            // Ensure right border
            row.getCell(7).border.right = { style: 'medium' };
            
            // Set compact row height
            row.height = 18; // Compact, tight spacing
            
            currentRow++;
        }
    });
    
    // Add a few empty rows if not enough expense data
    const minRows = 10;
    const emptyRowsNeeded = Math.max(0, minRows - (currentRow - otherRowsStart));
    
    for (let i = 0; i < emptyRowsNeeded; i++) {
        const row = worksheet.getRow(currentRow + i);
        
        // Date cell
        row.getCell(1).border = {
            left: { style: 'medium' },
            top: { style: 'thin' },
            right: { style: 'thin' },
            bottom: { style: 'thin' }
        };
        
        // Description cell
        row.getCell(2).border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' },
            bottom: { style: 'thin' }
        };
        row.getCell(2).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
        
        // Amount cells (C-G)
        for (let col = 3; col <= 7; col++) {
            row.getCell(col).border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                right: { style: 'thin' },
                bottom: { style: 'thin' }
            };
            row.getCell(col).numFmt = '$#,##0.00';
        }
        
        // Ensure right border
        row.getCell(7).border.right = { style: 'medium' };
        
        // Set compact row height for empty rows
        row.height = 18;
    }
    
    // Update current row to account for empty rows added
    currentRow += emptyRowsNeeded;
    
    // Add totals row
    const totalsRow = currentRow;
    const totalsRowObj = worksheet.getRow(totalsRow);
    
    // TOTAL label
    totalsRowObj.getCell(1).value = 'TOTAL';
    totalsRowObj.getCell(1).font = { bold: true };
    totalsRowObj.getCell(1).alignment = { horizontal: 'right' };
    totalsRowObj.getCell(1).border = {
      left: { style: 'medium' },
      top: { style: 'thin' },
      right: { style: 'thin' },
      bottom: { style: 'double' }
    };
    
    totalsRowObj.getCell(2).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' },
      bottom: { style: 'double' }
    };
    
    // Total formulas for each column (C-G)
    for (let col = 3; col <= 7; col++) {
      const colLetter = String.fromCharCode(64 + col); // Convert to column letter (C, D, etc.)
      totalsRowObj.getCell(col).value = {
        formula: `SUM(${colLetter}${otherRowsStart}:${colLetter}${totalsRow-1})`
      };
      totalsRowObj.getCell(col).numFmt = '$#,##0.00';
      totalsRowObj.getCell(col).font = { bold: true };
      totalsRowObj.getCell(col).alignment = { horizontal: 'right' };
      totalsRowObj.getCell(col).border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        right: { style: 'thin' },
        bottom: { style: 'double' }
      };
    }
    
    // Ensure right border
    totalsRowObj.getCell(7).border.right = { style: 'medium' };
    
    // Add GRAND TOTAL row that combines promotion and other expenses
    const grandTotalRow = totalsRow + 1;
    const grandTotalRowObj = worksheet.getRow(grandTotalRow);
    
    // GRAND TOTAL label
    worksheet.mergeCells(`A${grandTotalRow}:F${grandTotalRow}`);
    grandTotalRowObj.getCell(1).value = 'GRAND TOTAL (Promotion + Other Expenses)';
    grandTotalRowObj.getCell(1).font = { bold: true, size: 12 };
    grandTotalRowObj.getCell(1).alignment = { horizontal: 'right' };
    grandTotalRowObj.getCell(1).border = {
      left: { style: 'medium' },
      top: { style: 'thin' },
      right: { style: 'thin' },
      bottom: { style: 'double' }
    };
    
    // Grand total formula
    grandTotalRowObj.getCell(7).value = {
      formula: `G${totalPromoRow}+SUM(C${totalsRow}:G${totalsRow})`
    };
    grandTotalRowObj.getCell(7).numFmt = '$#,##0.00';
    grandTotalRowObj.getCell(7).font = { bold: true, size: 12 };
    grandTotalRowObj.getCell(7).alignment = { horizontal: 'right' };
    grandTotalRowObj.getCell(7).border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'medium' },
      bottom: { style: 'double' }
    };
    grandTotalRowObj.getCell(7).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'E6F2E6' } // Light green background
    };
    
    // Set compact row heights
    worksheet.getRow(promoRow).height = 18; // PROMOTION EXPENSES row
    worksheet.getRow(categoryRow).height = 18; // Categories row
    worksheet.getRow(otherRow).height = 18; // OTHER EXPENSES row
    worksheet.getRow(otherCategoryRow).height = 18; // Other categories row
    
    // Set compact row height for empty rows
    for (let i = 0; i < promoRows; i++) {
      worksheet.getRow(promoRowsStart + i).height = 18;
    }
    
    // Set compact default formatting
    for (let i = 1; i <= grandTotalRow; i++) {
      if (!worksheet.getRow(i).height) {
        worksheet.getRow(i).height = 18; // Compact spacing
      }
    }
    
    // Add space for signatures
    const signatureRow = grandTotalRow + 3;
    const signaturesObj = worksheet.getRow(signatureRow);
    
    signaturesObj.getCell(1).value = 'Signature of Claimant';
    signaturesObj.getCell(1).font = { bold: true };
    signaturesObj.getCell(5).value = 'Department Head/Manager';
    signaturesObj.getCell(5).font = { bold: true };
    
    signaturesObj.getCell(7).border = {
      right: { style: 'medium' }
    };
    
    // Increase signature row height (line for actual signature)
    worksheet.getRow(signatureRow+1).height = 60; // Increased for signature
    
    worksheet.getRow(signatureRow+1).getCell(1).border = {
      bottom: { style: 'thin' },
      left: { style: 'medium' }
    };
    worksheet.getRow(signatureRow+1).getCell(2).border = {
      bottom: { style: 'thin' }
    };
    
    // Add digital signature if provided
    if (signature) {
      try {
        console.log('Signature data received on server side');
        
        // Extract the base64 data from the signature string
        const signatureParts = signature.split(',');
        console.log('Signature format:', signatureParts[0]);
        
        const base64Data = signatureParts[1];
        if (base64Data) {
          console.log('Base64 data length:', base64Data.length);
          console.log('Base64 data starts with:', base64Data.substring(0, 30) + '...');
          
          // Save the image to a temporary file
          const tempImgPath = path.join(__dirname, 'temp_signature.png');
          const imgBuffer = Buffer.from(base64Data, 'base64');
          fs.writeFileSync(tempImgPath, imgBuffer);
          
          console.log('Signature saved to temporary file');
          
          try {
            // Create an image
            const signatureImage = workbook.addImage({
              filename: tempImgPath,
              extension: 'png',
            });
            
            console.log('Image created in workbook from file');
            
            // Add image to the worksheet at the signature position
            worksheet.addImage(signatureImage, {
              tl: { col: 0, row: signatureRow + 0.2 },
              br: { col: 2.5, row: signatureRow + 1.8 },
              editAs: 'oneCell'
            });
            
            // Log success message for debugging
            console.log('Signature added to Excel file successfully');
            
            // DO NOT delete the temp file here - we'll delete it after the Excel is generated
          } catch (imgError) {
            console.error('Error creating image from file:', imgError);
          }
        } else {
          console.log('No base64 data found in signature string');
        }
      } catch (signatureError) {
        console.error('Error adding signature to Excel:', signatureError);
        // Continue with export even if signature fails
      }
    } else {
      console.log('No signature data provided for Excel export');
    }
    
    worksheet.getRow(signatureRow+1).getCell(5).border = {
      bottom: { style: 'thin' }
    };
    worksheet.getRow(signatureRow+1).getCell(6).border = {
      bottom: { style: 'thin' }
    };
    worksheet.getRow(signatureRow+1).getCell(7).border = {
      bottom: { style: 'thin' },
      right: { style: 'medium' }
    };
    
    // Add additional blank space row below signatures
    const spacerRow = signatureRow + 2;
    worksheet.getRow(spacerRow).height = 20; // Extra space
    worksheet.getRow(spacerRow).getCell(1).border = {
      left: { style: 'medium' }
    };
    worksheet.getRow(spacerRow).getCell(7).border = {
      right: { style: 'medium' }
    };
    
    // Close the bottom of the sheet
    const finalRow = signatureRow + 3; // Moved down by 1 to account for the spacer row
    worksheet.getRow(finalRow).getCell(1).border = {
      left: { style: 'medium' },
      bottom: { style: 'medium' }
    };
    worksheet.getRow(finalRow).getCell(7).border = {
      right: { style: 'medium' },
      bottom: { style: 'medium' }
    };
    worksheet.mergeCells(`A${finalRow}:G${finalRow}`);
    
    // Generate Excel file
    const buffer = await workbook.xlsx.writeBuffer();
    
    // Clean up the temporary signature file if it exists
    const tempImgPath = path.join(__dirname, 'temp_signature.png');
    if (fs.existsSync(tempImgPath)) {
      try {
        fs.unlinkSync(tempImgPath);
        console.log('Temporary signature file removed after Excel generation');
      } catch (cleanupError) {
        console.error('Error removing temporary file:', cleanupError);
      }
    }
    
    // Send the file to the client
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=Expense_Report.xlsx');
    res.setHeader('Content-Length', buffer.length);
    res.send(buffer);
    
  } catch (error) {
    console.error('Error generating Excel report:', error);
    res.status(500).json({ success: false, error: 'Failed to generate Excel report' });
  }
});

// Add our new route for merging PDFs - simplified version that merges all PDFs
app.post('/api/merge-pdfs', async (req, res) => {
  try {
    // Create temporary directory if it doesn't exist
    const tempDir = path.join(__dirname, 'temp');
    try {
      await fsPromises.access(tempDir);
    } catch {
      await fsPromises.mkdir(tempDir);
    }
    
    // Path for the merged PDF
    const mergedPdfPath = path.join(tempDir, `Greenwin_Merged_PDF_${Date.now()}.pdf`);
    
    // Initialize PDF merger - using v4 syntax
    const merger = new PDFMerger();
    
    // Get all PDF files from the uploads directory
    const uploadsDir = path.join(__dirname, 'uploads');
    let files = [];
    
    try {
      files = await fsPromises.readdir(uploadsDir);
      console.log(`Found ${files.length} files in uploads directory`);
    } catch (err) {
      console.error('Error reading uploads directory:', err);
      return res.status(404).json({ success: false, error: 'No PDF files found' });
    }
    
    // Check if we have any files to merge
    if (files.length === 0) {
      return res.status(404).json({ success: false, error: 'No PDF files found' });
    }
    
    // Add all valid PDF files to the merger
    let filesAdded = 0;
    
    // Loop through all files
    for (const file of files) {
      const filePath = path.join(uploadsDir, file);
      try {
        // Check if file exists and is readable
        await fsPromises.access(filePath, fs.constants.R_OK);
        
        // Try to read the file as a PDF
        try {
          const pdfBuffer = fs.readFileSync(filePath);
          // Try to parse it as a PDF (this will throw if not a PDF)
          await pdfParse(pdfBuffer, { max: 1 }); // Only parse the first page to verify it's a PDF
          
          // Add file to merger - using v4 syntax
          merger.add(filePath);  // v4 doesn't require await here
          filesAdded++;
          console.log(`Added file to merger: ${file}`);
        } catch (pdfError) {
          console.log(`File ${file} is not a valid PDF, skipping`);
        }
      } catch (err) {
        console.error(`Error processing file ${file}:`, err);
      }
    }
    
    if (filesAdded === 0) {
      return res.status(404).json({ success: false, error: 'No valid PDF files could be processed' });
    }
    
    // Save the merged PDF - using v4 syntax
    await merger.save(mergedPdfPath);  // This requires await in v4
    console.log(`Merged PDF saved to: ${mergedPdfPath}`);
    
    // Send the merged PDF as a download with proper headers
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="Greenwin_Merged_PDF.pdf"`);
    
    // Stream the file to the client
    const fileStream = fs.createReadStream(mergedPdfPath);
    fileStream.pipe(res);
    
    // Clean up the temp file after streaming is complete
    fileStream.on('end', () => {
      // Delete the temporary file after it's been sent
      try {
        fs.unlinkSync(mergedPdfPath);
        console.log(`Temporary PDF file deleted: ${mergedPdfPath}`);
      } catch (err) {
        console.error(`Error deleting temporary PDF file: ${err.message}`);
      }
    });
    
    // Handle errors during file streaming
    fileStream.on('error', (err) => {
      console.error(`Error streaming PDF file: ${err.message}`);
      // Only send error if headers haven't been sent
      if (!res.headersSent) {
        res.status(500).json({ success: false, error: 'Error streaming PDF file' });
      }
    });
    
  } catch (error) {
    console.error('Error merging PDFs:', error);
    
    // Only send error if headers haven't been sent
    if (!res.headersSent) {
      res.status(500).json({
        success: false,
        error: 'Failed to merge PDF invoices'
      });
    }
  }
});

// PDF Export Route
app.get('/api/export-pdf', async (req, res) => {
    try {
        const sessionId = req.query.sessionId;
        if (!sessionId || !sessions.has(sessionId)) {
            return res.status(400).json({ error: 'Invalid session' });
        }

        const session = sessions.get(sessionId);
        const filesToMerge = session.files;

        if (filesToMerge.length === 0) {
            return res.status(400).json({ error: 'No PDFs to merge' });
        }

        const merger = new PDFMerger();
        for (const file of filesToMerge) {
            await merger.add(file);
        }

        const mergedPdfPath = path.join('uploads', `merged_${Date.now()}.pdf`);
        await merger.save(mergedPdfPath);

        res.download(mergedPdfPath, 'expense_report.pdf', (err) => {
            if (err) {
                console.error('Error sending file:', err);
            }
            // Cleanup: Delete the merged file after sending
            fs.unlink(mergedPdfPath, (err) => {
                if (err) console.error('Error deleting merged file:', err);
                else console.log('Deleted merged file:', mergedPdfPath);
            });
        });

        // Cleanup: Delete individual files after merging
        for (const file of filesToMerge) {
            fs.unlink(file, (err) => {
                if (err) console.error('Error deleting file:', err);
                else console.log('Deleted:', file);
            });
        }

        // Cleanup: Remove session after processing
        sessions.delete(sessionId);
        console.log('Cleaned up session:', sessionId);
    } catch (error) {
        console.error('Error merging PDFs:', error);
        res.status(500).json({ error: 'Error merging PDFs' });
    }
});

// Export to Excel and PDF for email
app.post('/api/export-email', async (req, res) => {
  try {
    const { expenses, sessionId } = req.body;
    
    if (!expenses || !Array.isArray(expenses) || expenses.length === 0) {
      return res.status(400).json({ success: false, error: 'No expense data provided' });
    }
    
    // Paths for the temporary files
    const timestamp = Date.now();
    const excelPath = `/temp/Greenwin_Expense_Report_${timestamp}.xlsx`;
    const pdfPath = `/temp/Greenwin_Expense_PDF_${timestamp}.pdf`;
    
    // Make sure the temp directory exists
    if (!fs.existsSync(path.join(__dirname, 'public/temp'))) {
      fs.mkdirSync(path.join(__dirname, 'public/temp'), { recursive: true });
    }
    
    // Create a new Excel workbook
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Greenwin Corp Expense Report App';
    workbook.created = new Date();
    
    // Add a worksheet with properties
    const worksheet = workbook.addWorksheet('Expense Report', {
      properties: {
        defaultRowHeight: 18
      },
      pageSetup: {
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0, // Setting to 0 means "automatically determine the number of pages"
        paperSize: 9, // 9 = A4
        orientation: 'portrait',
        margins: {
          left: 0.7,
          right: 0.7,
          top: 0.75,
          bottom: 0.75,
          header: 0.3,
          footer: 0.3
        }
      }
    });
    
    // Set column widths for professional layout
    worksheet.columns = [
      { width: 13 }, // A - DATE
      { width: 55 }, // B - DESCRIPTION (Merchant: Description combined)
      { width: 16 }, // C - 6408-000
      { width: 15 }, // D - 6402-000
      { width: 16 }, // E - 6404-000
      { width: 19 }, // F - 7335-000
      { width: 15 }, // G - 6026-000
    ];
    
    // -- HEADER SECTION --
    
    // Top section - Name and Date Submitted
    const nameCell = worksheet.getCell('A1');
    nameCell.value = 'Name:';
    nameCell.font = { bold: true };
    nameCell.border = {
      top: { style: 'medium' },
      left: { style: 'medium' },
      bottom: { style: 'medium' }
    };
    
    // Cell for name input
    worksheet.mergeCells('B1:C1');
    const nameInputCell = worksheet.getCell('B1');
    nameInputCell.border = {
      top: { style: 'medium' },
      bottom: { style: 'medium' }
    };
    
    // Date submitted
    const dateSubmittedCell = worksheet.getCell('D1');
    dateSubmittedCell.value = 'Date Submitted:';
    dateSubmittedCell.font = { bold: true };
    dateSubmittedCell.border = {
      top: { style: 'medium' },
      bottom: { style: 'medium' }
    };
    
    // Date field
    worksheet.mergeCells('E1:G1');
    const dateFieldCell = worksheet.getCell('E1');
    dateFieldCell.value = new Date().toLocaleDateString();
    dateFieldCell.border = {
      top: { style: 'medium' },
      right: { style: 'medium' },
      bottom: { style: 'medium' }
    };
    
    // Second row - Division
    const divisionCell = worksheet.getCell('A2');
    divisionCell.value = 'Division to be charged:';
    divisionCell.font = { bold: true };
    divisionCell.border = {
      left: { style: 'medium' },
      bottom: { style: 'medium' }
    };
    
    // Division field
    worksheet.mergeCells('B2:C2');
    const divisionFieldCell = worksheet.getCell('B2');
    divisionFieldCell.border = {
      bottom: { style: 'medium' }
    };
    
    // Team number
    const teamCell = worksheet.getCell('D2');
    teamCell.value = 'Team #:';
    teamCell.font = { bold: true };
    teamCell.border = {
      bottom: { style: 'medium' }
    };
    
    // Team field
    const teamFieldCell = worksheet.getCell('E2');
    teamFieldCell.border = {
      bottom: { style: 'medium' }
    };
    
    // Building number
    const buildingCell = worksheet.getCell('F2');
    buildingCell.value = 'BLDG#:';
    buildingCell.font = { bold: true };
    buildingCell.border = {
      bottom: { style: 'medium' }
    };
    
    // Building field
    const buildingFieldCell = worksheet.getCell('G2');
    buildingFieldCell.border = {
      bottom: { style: 'medium' },
      right: { style: 'medium' }
    };
    
    // -- PROMOTION EXPENSES SECTION --
    
    const promoRow = 5;
    
    // Promotion Expenses Header
    worksheet.mergeCells(`A${promoRow}:B${promoRow}`);
    const promoHeaderCell = worksheet.getCell(`A${promoRow}`);
    promoHeaderCell.value = 'PROMOTION EXPENSES';
    promoHeaderCell.alignment = { horizontal: 'center', vertical: 'middle' };
    promoHeaderCell.font = { bold: true, color: { argb: 'FFFFFF' } };
    promoHeaderCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '00A651' } // Green background
    };
    promoHeaderCell.border = {
      left: { style: 'medium' },
      bottom: { style: 'medium' }
    };
    
    // G/L Allocation Header for Promotion
    worksheet.mergeCells(`C${promoRow}:G${promoRow}`);
    const promoGlHeaderCell = worksheet.getCell(`C${promoRow}`);
    promoGlHeaderCell.value = 'G/L ALLOCATION';
    promoGlHeaderCell.alignment = { horizontal: 'center', vertical: 'middle' };
    promoGlHeaderCell.font = { bold: true, color: { argb: 'FFFFFF' } };
    promoGlHeaderCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '00A651' } // Green background
    };
    promoGlHeaderCell.border = {
      right: { style: 'medium' },
      bottom: { style: 'medium' }
    };
    
    // Continue with the rest of the Excel formatting...
    // ... (all the code from the regular Excel export for promotion expenses, 
    // other expenses, signature section, etc would go here)
    
    // Generate and save the Excel file
    await workbook.xlsx.writeFile(path.join(__dirname, 'public', excelPath));
    
    // Generate PDF if sessionId exists
    let pdfFilePath = null;
    if (sessionId) {
      // Find all PDFs for this session
      const sessionDir = path.join(__dirname, 'uploads', sessionId);
      let pdfFiles = [];
      
      if (fs.existsSync(sessionDir)) {
        pdfFiles = fs.readdirSync(sessionDir)
          .filter(file => file.endsWith('.pdf'))
          .map(file => path.join(sessionDir, file));
      }
      
      if (pdfFiles.length > 0) {
        try {
          // Merge PDFs
          const pdfDoc = await PDFDocument.create();
          
          for (const pdfFile of pdfFiles) {
            const fileData = await fs.promises.readFile(pdfFile);
            const pdfToMerge = await PDFDocument.load(fileData);
            const copiedPages = await pdfDoc.copyPages(pdfToMerge, pdfToMerge.getPageIndices());
            copiedPages.forEach(page => pdfDoc.addPage(page));
          }
          
          // Save the merged PDF
          const pdfBytes = await pdfDoc.save();
          fs.writeFileSync(path.join(__dirname, 'public', pdfPath), pdfBytes);
          pdfFilePath = pdfPath;
        } catch (pdfError) {
          console.error('Error merging PDFs:', pdfError);
          // Continue with just the Excel file if there's an error
        }
      }
    }
    
    // Return the file paths for the email client
    res.json({
      success: true,
      excelPath: excelPath,
      pdfPath: pdfFilePath || ''
    });
    
  } catch (error) {
    console.error('Error generating files for email:', error);
    res.status(500).json({ success: false, error: 'Failed to generate files for email' });
  }
});

// Send all other requests to the frontend
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Error handling middleware
app.use((err, req, res, next) => {
  console.error(err.stack);
  res.status(500).json({
    success: false,
    error: err.message || 'Something went wrong!'
  });
});

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});

// Helper function to get the Excel column number for a G/L code
function getColumnForGlCode(glCode) {
    // IMPORTANT: These column indexes must match the column structure in the Excel template
    // For "Other Expenses" section:
    const OTHER_EXPENSES_COLUMNS = {
        '6408-000': 3, // Office & General - Column C
        '6402-000': 4, // Membership - Column D
        '6404-000': 5, // Subscriptions - Column E
        '7335-000': 6, // Education & Development - Column F
        '6026-000': 7  // Mileage/ETR - Column G
    };
    
    // For promotion expenses, return 0 to indicate they should be handled separately
    // These go in the Promotion Expenses section of the Excel template
    const PROMOTION_GL_CODES = ['6010-000', '6011-000', '6012-000'];
    
    if (PROMOTION_GL_CODES.includes(glCode)) {
        return 0; // Signal that this is a promotion expense
    }
    
    // Return the column index for the G/L code, or default to Office & General (column 3)
    return OTHER_EXPENSES_COLUMNS[glCode] || 3;
}

// Function to add an expense to the promotion section
function addPromotionExpense(worksheet, expense, split) {
    try {
        // IMPORTANT - These are hard-coded Excel column numbers for promotion expenses
        // These are actual Excel column numbers for the specific layout in our template
        const PROMOTION_GL_COLUMNS = {
            'other': 3,           // C column - Other
            '6010-000': 4,        // D column - Food & Entertainment
            '6011-000': 5,        // E column - Social
            '6012-000': 6         // F column - Travel
        };
        
        // Get the promotion section rows from worksheet
        // Category row is where promotion expense headers are
        const categoryRow = 9; // Based on Excel template structure
        // Number of rows available for promotion expenses
        const promoRows = 10; // Allow up to 10 promotion expenses
        
        // Find the promotion section starting row (after the category labels)
        const promoStartRow = categoryRow + 1;
        
        // Find the first empty row in the promotion section
        let emptyRow = null;
        for (let i = 0; i < promoRows; i++) {
            const rowNum = promoStartRow + i;
            if (!worksheet.getRow(rowNum).getCell(1).value) {
                emptyRow = rowNum;
                break;
            }
        }
        
        // If no empty row found, we can't add the promotion expense
        if (!emptyRow) {
            console.warn('No empty rows available for promotion expense');
            return;
        }
        
        console.log(`Adding promotion expense to row ${emptyRow} for GL code ${split.glCode} with amount ${split.amount}`);
        
        const row = worksheet.getRow(emptyRow);
        
        // Format date
        let dateValue = expense.date;
        if (typeof dateValue === 'string') {
            try {
                const dateObj = new Date(dateValue);
                if (!isNaN(dateObj.getTime())) {
                    dateValue = dateObj;
                }
            } catch (e) {
                console.warn('Could not parse date:', e);
            }
        }
        
        // DATE column
        row.getCell(1).value = dateValue;
        row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
        row.getCell(1).border = {
            left: { style: 'medium' },
            top: { style: 'thin' },
            right: { style: 'thin' },
            bottom: { style: 'thin' }
        };
        
        // DESCRIPTION column - include location if available
        let description = expense.description || '';
        if (expense.location) {
            description = `${description}${description ? '\n\n' : ''}Location: ${expense.location}`;
        }
        
        // Add split info to description - specific for promotion expense split
        description = `${description}${description ? '\n\n' : ''}Split: $${split.amount} (${Math.round(split.percentage)}%) allocated to ${split.glCode}`;
        
        row.getCell(2).value = description;
        
        // Clear existing values in amount columns
        for (let col = 3; col <= 6; col++) {
            row.getCell(col).value = null;
        }
        
        // Determine which column to use based on the G/L code
        let targetColumn;
        
        // Get the column index based on the G/L code
        if (PROMOTION_GL_COLUMNS[split.glCode] !== undefined) {
            targetColumn = PROMOTION_GL_COLUMNS[split.glCode];
        } else {
            // Default to "Other" column if not a recognized promotion G/L code
            targetColumn = PROMOTION_GL_COLUMNS['other'];
        }
        
        console.log(`Placing amount ${split.amount} in column index ${targetColumn} for G/L code ${split.glCode}`);
        
        // Set the amount in the determined column
        row.getCell(targetColumn).value = parseFloat(split.amount);
        row.getCell(targetColumn).numFmt = '$#,##0.00';
        row.getCell(targetColumn).alignment = { horizontal: 'right', vertical: 'middle' };
        
        // Add borders to all cells in the row
        for (let col = 3; col <= 6; col++) {
            row.getCell(col).border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                right: { style: 'thin' },
                bottom: { style: 'thin' }
            };
        }
        
        // Ensure right border
        row.getCell(7).border = {
            right: { style: 'medium' }
        };
        
        // Adjust row height based on content length
        const descLength = description.length;
        const lineBreaks = (description.match(/\n/g) || []).length;
        const approximateLines = Math.ceil(descLength / 50) + lineBreaks; // Estimate lines based on characters and explicit line breaks
        
        // Calculate height based on estimated lines
        // Each line is approximately 15 points high in Excel
        let rowHeight = Math.max(24, approximateLines * 18); // Increased minimum height to 24 points and line height to 18 points
        
        // Add extra padding for split information
        if (expense.splits && expense.splits.length > 0) {
            // Add additional height based on number of splits (each split takes roughly two lines now)
            rowHeight += (expense.splits.length * 30);
            // Add extra padding to ensure visibility
            rowHeight += 15;
        }
        
        // Apply calculated height
        row.height = rowHeight;
        
        // Debug logging for troubleshooting
        console.log(`Row height calculated: ${rowHeight} for description with ${descLength} chars and ${lineBreaks} line breaks`);
        
        console.log(`Successfully added promotion expense to row ${emptyRow}`);
    } catch (error) {
        console.error('Error in addPromotionExpense:', error);
    }
} 