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
const sessions = new Map(); // Store session data

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
});

const app = express();
const PORT = 3005;

// Middleware
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// Routes
app.get('/api/expenses', (req, res) => {
  // This would normally fetch from a database
  const sampleExpenses = [];
  
  res.json(sampleExpenses);
});

app.post('/api/expenses', (req, res) => {
  // This would normally save to a database
  console.log('New expense:', req.body);
  // Return success response with mock ID
  res.status(201).json({ id: Date.now(), ...req.body });
});

// Add the missing PUT route handler for updating expenses
app.put('/api/expenses/:id', (req, res) => {
  // This would normally update in a database
  console.log('Updating expense:', req.params.id, req.body);
  
  // Return success response
  res.json({ id: req.params.id, ...req.body });
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

        // Process each file
        for (const file of req.files) {
            try {
                // Read PDF file
                const pdfBuffer = fs.readFileSync(file.path);
                const pdfData = await pdfParse(pdfBuffer);
                
                // Extract text content from PDF
                const pdfText = pdfData.text;
                
                // Process with OpenAI API
                const openaiResponse = await openai.chat.completions.create({
                    model: "gpt-4-turbo",
                    messages: [
                        {
                            role: "system", 
                            content: `You are an AI assistant trained to extract and categorize expense information from receipts and invoices for Greenwin Corp. Extract the following details in a structured format: 
                            - date (YYYY-MM-DD)
                            - merchant/vendor name
                            - total amount
                            - tax amount (specifically look for HST, GST, or other Canadian taxes listed on the receipt)
                            - description of purchase (brief summary of what was purchased)
                            - any itemized expenses with their individual amounts (if available)

                            IMPORTANT: For Canadian receipts and invoices, carefully extract the actual HST/tax amount shown on the receipt. If the HST (13%) is explicitly listed, extract this exact value. Do not calculate or estimate this value - only provide the exact tax amount shown on the receipt.

                            Assign the most appropriate G/L allocation code from this list:
                            - 6408-000 (Office & General): For office supplies, general equipment, administrative expenses
                            - 6402-000 (Membership): For professional memberships, dues, associations
                            - 6404-000 (Subscriptions): For software subscriptions, publications, online services
                            - 7335-000 (Education & Development): For training, courses, professional development
                            - 6026-000 (Mileage/ETR): For mileage reimbursements and toll expenses
                            - 6010-000 (Food & Entertainment): For meals, catering, food-related expenses
                            - 6011-000 (Social): For company events, team activities, celebrations
                            - 6012-000 (Travel Expenses): For flights, hotels, taxis, and other travel costs (excluding mileage)

                            Be very precise in your G/L code selections based on the merchant and items purchased. For example:
                            - Restaurant receipts → G/L Code: 6010-000 (Food & Entertainment)
                            - Zoom subscription → G/L Code: 6404-000 (Subscriptions)
                            - Conference registration → G/L Code: 7335-000 (Education & Development)
                            - Office supplies → G/L Code: 6408-000 (Office & General)

                            Format the response as JSON with these exact field names: date, merchant, amount, tax, description, gl_code, line_items (array of objects with name and amount).
                            
                            If the tax amount is not explicitly stated on the receipt, do not include a tax field in your response.`
                        },
                        {
                            role: "user",
                            content: `Extract detailed expense information from this document in the exact JSON format specified: ${pdfText}`
                        }
                    ],
                    response_format: { type: "json_object" }
                });
                
                // Parse OpenAI response
                const parsedData = JSON.parse(openaiResponse.choices[0].message.content);
                
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

                // Add filename to the results for reference
                results.push({
                    filename: file.originalname,
                    filepath: file.path,
                    data: parsedData
                });
                
            } catch (error) {
                console.error(`Error processing PDF file ${file.originalname}:`, error);
                errors.push({
                    filename: file.originalname,
                    error: error.message || 'Error processing PDF file'
                });
            }
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
    const { expenses } = req.body;
    
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
        tabColor: { argb: '00A651' }, // Green tab color
        defaultRowHeight: 18
      }
    });
    
    // Set column widths for better display - removed Name/Department columns as per image
    worksheet.columns = [
      { width: 15, style: { numFmt: 'mm/dd/yyyy' } }, // A - Date
      { width: 30 }, // B - Merchant
      { width: 45 }, // C - Description
      { width: 15, style: { numFmt: '$#,##0.00', alignment: { horizontal: 'right' } } }, // D - Amount
      { width: 15, style: { numFmt: '$#,##0.00', alignment: { horizontal: 'right' } } }, // E - Tax
      { width: 15 }, // F - G/L Code
    ];
    
    // Add header/title section (rows 1-3)
    worksheet.mergeCells('A1:C1');
    const titleCell = worksheet.getCell('A1');
    titleCell.value = 'GREENWIN';
    titleCell.font = { bold: true, size: 16, color: { argb: '000000' } };
    titleCell.alignment = { horizontal: 'left', vertical: 'middle' };
    
    // Create report title
    worksheet.mergeCells('A2:C2');
    const subtitleCell = worksheet.getCell('A2');
    subtitleCell.value = 'EXPENSE REPORT';
    subtitleCell.font = { bold: true, size: 14, color: { argb: '000000' } };
    subtitleCell.alignment = { horizontal: 'left', vertical: 'middle' };
    
    // Add date and employee info on right side
    worksheet.mergeCells('D1:E1');
    const dateLabel = worksheet.getCell('D1');
    dateLabel.value = 'Date:';
    dateLabel.alignment = { horizontal: 'right', vertical: 'middle' };
    dateLabel.font = { bold: true, size: 11 };
    
    // Current date
    const currentDate = new Date();
    const dateCell = worksheet.getCell('F1');
    dateCell.value = currentDate;
    dateCell.numFmt = 'mmmm d yyyy';
    dateCell.alignment = { horizontal: 'left', vertical: 'middle' };
    
    // Employee name from the first expense
    worksheet.mergeCells('D2:E2');
    const employeeLabel = worksheet.getCell('D2');
    employeeLabel.value = 'Employee:';
    employeeLabel.alignment = { horizontal: 'right', vertical: 'middle' };
    employeeLabel.font = { bold: true, size: 11 };
    
    // Get employee name from the first expense
    const employeeName = expenses[0]?.name || '';
    const employeeCell = worksheet.getCell('F2');
    employeeCell.value = employeeName;
    employeeCell.alignment = { horizontal: 'left', vertical: 'middle' };
    
    // Add G/L ALLOCATION header in row 4 with green background as in the image
    worksheet.mergeCells('A4:F4');
    const glHeaderCell = worksheet.getCell('A4');
    glHeaderCell.value = 'G/L ALLOCATION';
    glHeaderCell.alignment = { horizontal: 'center', vertical: 'middle' };
    glHeaderCell.font = { bold: true, size: 12, color: { argb: 'FFFFFF' } };
    glHeaderCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: '00A651' } // Green background
    };
    worksheet.getRow(4).height = 24;
    
    // Add gratuity note in yellow
    worksheet.mergeCells('A5:F5');
    const gratuityCell = worksheet.getCell('A5');
    gratuityCell.value = 'include gratuity where applicable';
    gratuityCell.alignment = { horizontal: 'center', vertical: 'middle' };
    gratuityCell.font = { italic: true, size: 11 };
    gratuityCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFFF99' } // Light yellow background
    };
    worksheet.getRow(5).height = 20;
    
    // Add the table headers at row 6
    const headerRow = worksheet.getRow(6);
    headerRow.values = ['Date', 'Merchant', 'Description', 'Amount', 'Tax (13% HST)', 'G/L Code'];
    headerRow.height = 22;
    
    // Style the header row
    headerRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFCC' } // Light yellow background
      };
      cell.font = { bold: true, size: 11 };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    });
    
    // Add rows for expense data starting at row 7
    let rowIndex = 7;
    let rowColor = false; // For alternating row colors
    
    expenses.forEach((expense) => {
      const row = worksheet.getRow(rowIndex);
      
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
      
      // Calculate tax if it doesn't exist
      let taxAmount = expense.tax;
      if (taxAmount === undefined && expense.amount) {
        taxAmount = parseFloat((expense.amount * 0.13).toFixed(2));
      }
      
      // Populate data - aligned to the image layout
      row.getCell(1).value = dateValue; // Date
      row.getCell(2).value = expense.merchant || expense.title; // Merchant
      row.getCell(3).value = expense.description || ''; // Description
      row.getCell(4).value = expense.amount; // Amount
      row.getCell(5).value = taxAmount || 0; // Tax
      row.getCell(6).value = expense.glCode || ''; // G/L Code
      
      // Format the cells
      row.height = 22;
      
      // Add alternating row colors for better readability
      if (rowColor) {
        row.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'F5F5F5' } // Light gray background
          };
        });
      }
      rowColor = !rowColor;
      
      row.eachCell((cell) => {
        cell.border = {
          top: { style: 'thin', color: { argb: 'D0D0D0' } },
          left: { style: 'thin', color: { argb: 'D0D0D0' } },
          bottom: { style: 'thin', color: { argb: 'D0D0D0' } },
          right: { style: 'thin', color: { argb: 'D0D0D0' } }
        };
        
        // Set alignment based on column
        if (cell.col === 1) { // Date
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        } else if (cell.col === 2 || cell.col === 3) { // Merchant and Description
          cell.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };
        } else if (cell.col === 4 || cell.col === 5) { // Amount and Tax
          cell.numFmt = '$#,##0.00';
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
        } else if (cell.col === 6) { // G/L Code
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }
      });
      
      rowIndex++;
    });
    
    // Add empty rows for manual entry (10 extra rows)
    for (let i = 0; i < 10; i++) {
      const row = worksheet.getRow(rowIndex + i);
      row.height = 22;
      
      // Add alternating row colors for better readability
      if (rowColor) {
        row.eachCell((cell) => {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'F5F5F5' } // Light gray background
          };
        });
      }
      rowColor = !rowColor;
      
      // Add empty cells with borders
      for (let col = 1; col <= 6; col++) {
        const cell = row.getCell(col);
        cell.border = {
          top: { style: 'thin', color: { argb: 'D0D0D0' } },
          left: { style: 'thin', color: { argb: 'D0D0D0' } },
          bottom: { style: 'thin', color: { argb: 'D0D0D0' } },
          right: { style: 'thin', color: { argb: 'D0D0D0' } }
        };
        
        // Set alignment based on column
        if (col === 1) { // Date
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        } else if (col === 2 || col === 3) { // Merchant and Description
          cell.alignment = { horizontal: 'left', vertical: 'middle' };
        } else if (col === 4 || col === 5) { // Amount and Tax
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
        } else if (col === 6) { // G/L Code
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        }
      }
    }
    
    // Add total row
    const totalRowIndex = rowIndex + 10;
    const totalRow = worksheet.getRow(totalRowIndex);
    totalRow.height = 25;
    
    // Merge cells for "Total" label
    worksheet.mergeCells(`A${totalRowIndex}:C${totalRowIndex}`);
    const totalLabelCell = totalRow.getCell(1);
    totalLabelCell.value = 'TOTAL';
    totalLabelCell.alignment = { horizontal: 'right', vertical: 'middle' };
    totalLabelCell.font = { bold: true, size: 12 };
    
    // Add formula for total amount
    const totalAmountCell = totalRow.getCell(4);
    totalAmountCell.value = { formula: `SUM(D7:D${rowIndex-1})` };
    totalAmountCell.numFmt = '$#,##0.00';
    totalAmountCell.font = { bold: true };
    
    // Add formula for total tax
    const totalTaxCell = totalRow.getCell(5);
    totalTaxCell.value = { formula: `SUM(E7:E${rowIndex-1})` };
    totalTaxCell.numFmt = '$#,##0.00';
    totalTaxCell.font = { bold: true };
    
    // Style the total row
    totalRow.eachCell((cell) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'E6E6E6' } // Light gray background
      };
      cell.border = {
        top: { style: 'medium' },
        left: { style: 'thin' },
        bottom: { style: 'double' },
        right: { style: 'thin' }
      };
    });
    
    // Add signature section
    const signatureRow = totalRowIndex + 3;
    worksheet.mergeCells(`A${signatureRow}:C${signatureRow}`);
    const signatureCell = worksheet.getCell(`A${signatureRow}`);
    signatureCell.value = 'Signature of Claimant';
    signatureCell.font = { bold: true, size: 11 };
    signatureCell.alignment = { horizontal: 'left' };
    
    // Add a line for the signature
    ['A', 'B', 'C'].forEach(col => {
      worksheet.getCell(`${col}${signatureRow+1}`).border = {
        top: { style: 'thin' }
      };
    });
    
    worksheet.mergeCells(`D${signatureRow}:F${signatureRow}`);
    const managerCell = worksheet.getCell(`D${signatureRow}`);
    managerCell.value = 'Department Head/Manager';
    managerCell.font = { bold: true, size: 11 };
    managerCell.alignment = { horizontal: 'left' };
    
    // Add a line for the manager signature
    ['D', 'E', 'F'].forEach(col => {
      worksheet.getCell(`${col}${signatureRow+1}`).border = {
        top: { style: 'thin' }
      };
    });
    
    // Generate a temporary file path
    const tempFilePath = path.join(__dirname, 'uploads', `expense_report_${Date.now()}.xlsx`);
    
    // Write to file
    await workbook.xlsx.writeFile(tempFilePath);
    
    // Send file as download
    res.download(tempFilePath, 'Greenwin_Expense_Report.xlsx', (err) => {
      // Delete the temp file after download
      if (fs.existsSync(tempFilePath)) {
        fs.unlinkSync(tempFilePath);
      }
      
      if (err) {
        console.error('Error downloading file:', err);
      }
    });
  } catch (error) {
    console.error('Error generating Excel export:', error);
    res.status(500).json({
      success: false,
      error: error.message || 'Error generating Excel export'
    });
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