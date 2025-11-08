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
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');
const { supabase } = require('./supabaseClient');
// In-memory store no longer used; persisting to Supabase DB

// Avoid logging sensitive environment details

// Configure multer with per-session directories and 50MB limit
const storage = multer.diskStorage({
  destination: async (req, file, cb) => {
    try {
      // sessionId will be set by an init middleware
      if (!req.sessionId) {
        req.sessionId = uuidv4();
      }
      const sessionDir = path.join(__dirname, 'uploads', req.sessionId);
      await fsPromises.mkdir(sessionDir, { recursive: true });
      cb(null, sessionDir);
    } catch (err) {
      cb(err);
    }
  },
  filename: (req, file, cb) => {
    const safeBase = (file.originalname || `upload_${Date.now()}.pdf`).replace(/[^A-Za-z0-9._-]/g, '_');
    const name = `${uuidv4()}_${safeBase}`;
    cb(null, name);
  }
});

const upload = multer({
  storage,
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    if (file.mimetype === 'application/pdf' || (file.originalname && file.originalname.toLowerCase().endsWith('.pdf'))) {
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

// Ensure temp directory exists for downloads/merges
const TEMP_DIR = path.join(__dirname, 'temp');
if (!fs.existsSync(TEMP_DIR)) {
  fs.mkdirSync(TEMP_DIR, { recursive: true });
}

// Attempt to ensure Supabase storage bucket exists (best-effort)
async function ensureReceiptsBucket() {
  if (!supabase) return;
  try {
    const { data: buckets } = await supabase.storage.listBuckets();
    const exists = (buckets || []).some(b => b.name === 'receipts');
    if (!exists) {
      await supabase.storage.createBucket('receipts', { public: false, fileSizeLimit: 50 * 1024 * 1024 });
      console.log('Created Supabase bucket: receipts');
    }
  } catch (e) {
    // Ignore if lacking permissions or already exists
    console.warn('Bucket ensure warning:', e.message);
  }
}
ensureReceiptsBucket();

// Initialize OpenAI with the API key from environment variables
const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

const app = express();
const PORT = process.env.PORT || 3005;

// Security middleware (allow cross-origin images like the Greenwin logo)
app.use(helmet({
  contentSecurityPolicy: false, // we are not setting an explicit CSP here
  crossOriginEmbedderPolicy: false, // allow embedding cross-origin resources without CORP/CORS
  crossOriginResourcePolicy: { policy: 'cross-origin' } // permit loading images from external CDNs
}));

// Enable CORS with optional restriction via env
const allowedOrigin = process.env.ALLOWED_ORIGIN;
if (allowedOrigin) {
  app.use(cors({ origin: allowedOrigin }));
} else {
  app.use(cors());
}

// Basic rate limiting for API
const apiLimiter = rateLimit({ windowMs: 60 * 1000, max: 60 });
app.use('/api/', apiLimiter);

// Middleware
app.use(bodyParser.json({ limit: '5mb' }));
app.use(bodyParser.urlencoded({ extended: true, limit: '5mb' }));
app.use(express.static(path.join(__dirname, 'public')));

// Routes
// Fetch recent expenses (optional: filter by sessionId)
app.get('/api/expenses', async (req, res) => {
  try {
    if (!supabase) return res.json([]);
    const q = supabase
      .from('expenses')
      .select('*')
      .order('created_at', { ascending: false })
      .limit(100);
    const { data, error } = await q;
    if (error) throw error;
    res.json(data || []);
  } catch (e) {
    console.error('Error fetching expenses:', e);
    res.status(500).json([]);
  }
});

// Create an expense (and optional splits)
app.post('/api/expenses', async (req, res) => {
  try {
    const body = req.body || {};
    if (!supabase) {
      // Fallback: echo back with generated id for local/dev without DB
      const fallback = { id: Date.now(), ...body };
      return res.status(201).json(fallback);
    }

    const expenseRow = {
      session_id: body.sessionId || null,
      date: body.date || null,
      merchant: body.title || body.merchant || '',
      amount: body.amount != null ? Number(body.amount) : null,
      tax: body.tax != null ? Number(body.tax) : null,
      gl_code: body.glCode || body.gl_code || null,
      description: body.description || '',
      name: body.name || '',
      department: body.department || '',
      location: body.location || '',
      property_code: body.propertyCode || ''
    };
    let inserted;
    try {
      const resp = await supabase
        .from('expenses')
        .insert(expenseRow)
        .select('*')
        .single();
      if (resp.error) throw resp.error;
      inserted = resp.data;
    } catch (dbErr) {
      console.warn('Supabase insert failed, using local fallback:', dbErr.message);
      const fallback = { id: Date.now(), ...expenseRow };
      return res.status(201).json(fallback);
    }

    // Insert splits if provided
    if (Array.isArray(body.splits) && body.splits.length > 0) {
      const splitsRows = body.splits.map(s => ({
        expense_id: inserted.id,
        gl_code: s.glCode || s.gl_code,
        amount: Number(s.amount),
        percentage: s.percentage != null ? Number(s.percentage) : null
      }));
      const { error: splitErr } = await supabase.from('expense_splits').insert(splitsRows);
      if (splitErr) console.error('Error inserting splits:', splitErr.message);
    }

    res.status(201).json(inserted);
  } catch (e) {
    console.error('Error creating expense:', e);
    res.status(500).json({ error: 'Failed to create expense' });
  }
});

// Fix the PUT route handler for updating expenses
app.put('/api/expenses/:id', async (req, res) => {
  try {
    const id = parseInt(req.params.id);
    const body = req.body || {};

    // Mileage handling
    if (body.glCode === '6026-000') {
      const kilometers = parseFloat(body.kilometers) || parseFloat(body.amount) || 0;
      body.kilometers = kilometers;
      body.amount = Number((kilometers * 0.72).toFixed(2));
      body.tax = 0;
      if (body.fromLocation && body.toLocation) {
        console.log(`Mileage from ${body.fromLocation} to ${body.toLocation}`);
      }
    }

    const updateRow = {
      session_id: body.sessionId || null,
      date: body.date || null,
      merchant: body.title || body.merchant || '',
      amount: body.amount != null ? Number(body.amount) : null,
      tax: body.tax != null ? Number(body.tax) : null,
      gl_code: body.glCode || body.gl_code || null,
      description: body.description || '',
      name: body.name || '',
      department: body.department || '',
      location: body.location || '',
      property_code: body.propertyCode || '',
      updated_at: new Date()
    };

    if (!supabase) {
      // Fallback: echo back update for local/dev without DB
      return res.json({ id, ...body });
    }

    let updated;
    try {
      const resp = await supabase
        .from('expenses')
        .update(updateRow)
        .eq('id', id)
        .select('*')
        .single();
      if (resp.error) throw resp.error;
      updated = resp.data;
    } catch (dbErr) {
      console.warn('Supabase update failed, using local fallback:', dbErr.message);
      return res.json({ id, ...updateRow });
    }

    // Replace splits
    const { error: delErr } = await supabase.from('expense_splits').delete().eq('expense_id', id);
    if (delErr) console.error('Error deleting splits:', delErr.message);
    if (Array.isArray(body.splits) && body.splits.length > 0) {
      const rows = body.splits.map(s => ({
        expense_id: id,
        gl_code: s.glCode || s.gl_code,
        amount: Number(s.amount),
        percentage: s.percentage != null ? Number(s.percentage) : null
      }));
      const { error: insErr } = await supabase.from('expense_splits').insert(rows);
      if (insErr) console.error('Error inserting splits:', insErr.message);
    }

    res.json(updated);
  } catch (e) {
    console.error('Error updating expense:', e);
    res.status(500).json({ error: 'Failed to update expense' });
  }
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
        const isOld = stats.mtime.getTime() < cutoffTime;
        if (!isOld) continue;
        if (stats.isDirectory()) {
          await fsPromises.rm(filePath, { recursive: true, force: true });
          console.log(`Removed old directory: ${file}`);
        } else {
          await fsPromises.unlink(filePath);
          console.log(`Removed old file: ${file}`);
        }
      } catch (err) {
        console.error(`Error removing path ${file}:`, err);
      }
    }
    
    console.log('Cleanup of uploads directory completed');
    
  } catch (err) {
    console.error('Error cleaning uploads directory:', err);
  }
}

// Initialize a new upload session
async function initUploadSession(req, res, next) {
  try {
    if (!req.sessionId) {
      req.sessionId = uuidv4();
    }
    if (supabase) {
      // Create session row if not exists
      const { error } = await supabase.from('sessions').insert({ id: req.sessionId, status: 'active' });
      if (error && !/duplicate key/i.test(error.message)) {
        console.warn('Unable to create session in DB:', error.message);
      }
    }
  } catch (e) {
    console.warn('initUploadSession warning:', e.message);
  }
  next();
}

// Simple retry helper for transient failures
async function withRetry(fn, { retries = 3, baseMs = 500 } = {}) {
  let lastErr;
  for (let i = 0; i < retries; i++) {
    try {
      return await fn();
    } catch (err) {
      lastErr = err;
      const wait = baseMs * Math.pow(2, i);
      await new Promise(r => setTimeout(r, wait));
    }
  }
  throw lastErr;
}

// PDF Upload and Processing Route for multiple files
app.post('/api/upload-pdf', initUploadSession, upload.array('pdfFiles', 50), async (req, res) => {
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

        const sessionId = req.sessionId;

        const results = [];
        const errors = [];

        // Process files in parallel with limited concurrency
        const BATCH_SIZE = 5; // Reduce concurrency to limit API pressure
        
        for (let i = 0; i < req.files.length; i += BATCH_SIZE) {
            const batch = req.files.slice(i, i + BATCH_SIZE);
            
            // Process batch in parallel
            const batchPromises = batch.map(async (file) => {
            try {
                // Read PDF file (async) and quick header validation
                const pdfBuffer = await fsPromises.readFile(file.path);
                if (pdfBuffer.slice(0, 5).toString() !== '%PDF-') {
                  throw new Error('Invalid PDF file content');
                }
                const pdfData = await pdfParse(pdfBuffer);
                
                // Extract text content from PDF
                let pdfText = pdfData.text || '';

                // OCR fallback for scanned/image-only PDFs
                if (pdfText.trim().length < 20 && process.env.OCR_SPACE_API_KEY) {
                  try {
                    const form = new FormData();
                    form.append('apikey', process.env.OCR_SPACE_API_KEY);
                    form.append('language', 'eng');
                    form.append('isOverlayRequired', 'false');
                    form.append('OCREngine', '2');
                    form.append('file', new Blob([pdfBuffer], { type: 'application/pdf' }), file.originalname || 'upload.pdf');
                    const ocrResp = await fetch('https://api.ocr.space/parse/image', { method: 'POST', body: form });
                    const ocrJson = await ocrResp.json();
                    if (ocrJson && Array.isArray(ocrJson.ParsedResults)) {
                      pdfText = ocrJson.ParsedResults.map(r => r.ParsedText || '').join('\n');
                    }
                  } catch (ocrErr) {
                    console.warn('OCR fallback failed:', ocrErr.message);
                  }
                }
                
                // Process with OpenAI API - Using GPT-4o for maximum intelligence and speed
                const openaiResponse = await withRetry(() => openai.chat.completions.create({
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
                }), { retries: 3, baseMs: 600 });
                
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
                
                // Upload PDF to Supabase Storage and track storage path
                let storagePath = null;
                if (supabase) {
                  const safeOrig = (file.originalname || 'upload.pdf').replace(/[^A-Za-z0-9._-]/g, '_');
                  const key = `${sessionId}/${uuidv4()}_${safeOrig}`;
                  const { error: upErr } = await supabase.storage.from('receipts').upload(key, pdfBuffer, {
                    contentType: 'application/pdf',
                    upsert: true,
                  });
                  if (upErr) {
                    console.error('Supabase upload error:', upErr);
                  } else {
                    storagePath = key;
                  }
                }

                // Persist upload record (prefer storage path)
                if (supabase && (storagePath || file.path)) {
                  const { error: upRecErr } = await supabase.from('uploads').insert({
                    session_id: sessionId,
                    storage_path: storagePath || file.path,
                    original_name: file.originalname || null,
                    size: file.size || null
                  });
                  if (upRecErr) console.error('Error writing upload record:', upRecErr.message);
                }
                console.log(`PDF file processed: ${storagePath || file.path}`);

                // Remove local file if uploaded to storage
                if (storagePath) {
                  try { await fsPromises.unlink(file.path); } catch (_) {}
                }

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
        // Extract the base64 data from the signature string
        const signatureParts = signature.split(',');
        const base64Data = signatureParts[1];
        if (base64Data) {
          const imgBuffer = Buffer.from(base64Data, 'base64');
          try {
            // Embed image directly from buffer
            const signatureImage = workbook.addImage({ buffer: imgBuffer, extension: 'png' });
            worksheet.addImage(signatureImage, {
              tl: { col: 0, row: signatureRow + 0.2 },
              br: { col: 2.5, row: signatureRow + 1.8 },
              editAs: 'oneCell'
            });
          } catch (imgError) {
            console.error('Error embedding signature image:', imgError);
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
// (Removed deprecated global merge endpoint to prevent mixing sessions)

// PDF Export Route
app.get('/api/export-pdf', async (req, res) => {
    try {
        const sessionId = req.query.sessionId;
        if (!sessionId) {
            return res.status(400).json({ error: 'Invalid session' });
        }

        // Fetch upload records for this session
        let filesToMerge = [];
        if (supabase) {
            const { data: rows, error } = await supabase
                .from('uploads')
                .select('storage_path')
                .eq('session_id', sessionId);
            if (error) throw error;
            filesToMerge = (rows || []).map(r => r.storage_path);
        }

        if (filesToMerge.length === 0) {
            return res.status(400).json({ error: 'No PDFs to merge' });
        }

        const merger = new PDFMerger();
        const tempFiles = [];
        for (const entry of filesToMerge) {
            const isStorage = typeof entry === 'string' && !entry.startsWith('/') && !entry.startsWith('uploads');
            if (isStorage && supabase) {
                const { data, error } = await supabase.storage.from('receipts').download(entry);
                if (error) {
                    console.error('Supabase download error:', error);
                    continue;
                }
                const arrBuf = await data.arrayBuffer();
                const buf = Buffer.from(arrBuf);
                const tempPath = path.join(TEMP_DIR, `${sessionId}_${path.basename(entry)}`);
                await fsPromises.writeFile(tempPath, buf);
                tempFiles.push(tempPath);
                await merger.add(tempPath);
            } else {
                // Local fallback
                await merger.add(entry);
            }
        }

        const mergedPdfPath = path.join(TEMP_DIR, `merged_${Date.now()}.pdf`);
        await merger.save(mergedPdfPath);

        res.download(mergedPdfPath, 'expense_report.pdf', async (err) => {
            if (err) {
                console.error('Error sending file:', err);
            }
            // Cleanup temp merged and downloaded files
            try { await fsPromises.unlink(mergedPdfPath); } catch (_) {}
            for (const f of tempFiles) { try { await fsPromises.unlink(f); } catch (_) {} }
        });

        // Cleanup: remove files and rows for this session
        if (supabase) {
            const storagePaths = filesToMerge.filter(e => typeof e === 'string' && !e.startsWith('/') && !e.startsWith('uploads'));
            if (storagePaths.length > 0) {
                const { error: remErr } = await supabase.storage.from('receipts').remove(storagePaths);
                if (remErr) console.error('Supabase remove error:', remErr);
            }
            const { error: delRowsErr } = await supabase.from('uploads').delete().eq('session_id', sessionId);
            if (delRowsErr) console.error('Error deleting upload rows:', delRowsErr.message);
            const { error: updSessErr } = await supabase.from('sessions').update({ status: 'completed' }).eq('id', sessionId);
            if (updSessErr) console.error('Error updating session status:', updSessErr.message);
        }

        // Cleanup: Remove session after processing
        console.log('Completed session:', sessionId);
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
    const excelPath = `temp/Greenwin_Expense_Report_${timestamp}.xlsx`;
    const pdfPath = `temp/Greenwin_Expense_PDF_${timestamp}.pdf`;
    
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

// Healthcheck endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'ok', uptime: process.uptime() });
});

// Periodic cleanup of uploads directory (old files)
setTimeout(() => cleanUploadsDirectory(120), 10 * 1000); // initial delayed run
setInterval(() => cleanUploadsDirectory(120), 60 * 60 * 1000); // hourly

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
