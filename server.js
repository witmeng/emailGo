const express = require('express');
const dotenv = require('dotenv');
const multer = require('multer');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const path = require('path');
const fs = require('fs');
const { v4: uuidv4 } = require('uuid'); // For generating unique job IDs
const cheerio = require('cheerio'); // <<<--- 添加 cheerio 引入

// Determine base directory first, as it's needed for external .env loading
const isPackaged = typeof process.pkg !== 'undefined';
const baseDir = isPackaged ? path.dirname(process.execPath) : __dirname;

// Initial dotenv.config() to load bundled .env or project .env in dev
dotenv.config(); 

// If packaged, try to load an external .env file from the .exe's directory to override
if (isPackaged) {
    const externalEnvPath = path.join(baseDir, '.env');
    if (fs.existsSync(externalEnvPath)) {
        console.log(`Loading external .env file from: ${externalEnvPath}`);
        try {
            const externalConfig = dotenv.parse(fs.readFileSync(externalEnvPath));
            for (const k in externalConfig) {
                process.env[k] = externalConfig[k];
            }
            console.log("External .env file loaded and applied.");
        } catch (e) {
            console.error(`Error parsing external .env file at ${externalEnvPath}:`, e);
        }
    } else {
        console.log("No external .env file found at application root. Using packaged configuration.");
    }
}

const app = express();
const port = process.env.PORT || 3000;

// Configuration check for email user and from email
if (!process.env.EMAIL_USER) {
    console.error("ERROR: EMAIL_USER is not configured in the .env file. This is required for SMTP authentication. Please check your .env file.");
    process.exit(1); // Exit if critical config is missing
}
if (!process.env.EMAIL_FROM_EMAIL) {
    console.error("ERROR: EMAIL_FROM_EMAIL is not configured in the .env file. As per your requirements, this must be configured and match EMAIL_USER. Please check your .env file.");
    process.exit(1); // Exit if critical config is missing
}
if (process.env.EMAIL_USER !== process.env.EMAIL_FROM_EMAIL) {
    console.error(`ERROR: EMAIL_USER ("${process.env.EMAIL_USER}") and EMAIL_FROM_EMAIL ("${process.env.EMAIL_FROM_EMAIL}") in the .env file do not match. As per your requirements, these two values must be the same. Please check your .env file.`);
    process.exit(1); // Exit if they are not consistent as per requirement
}

// Middlewares
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Adjust public and uploads paths based on baseDir
const publicDir = path.join(baseDir, 'public');
const uploadsDir = path.join(baseDir, 'uploads');

app.use(express.static(publicDir)); 
app.use('/uploads', express.static(uploadsDir));

// Ensure uploads directory exists (relative to baseDir now)
if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
}

// Multer setup for initial XLS file upload (to /upload-and-preview)
const xlsStorage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, uploadsDir);
    },
    filename: function (req, file, cb) {
        // 嘗試修正文件名編碼：如果UTF-8字節被錯誤地解析為latin1，則進行還原
        const originalNameBuffer = Buffer.from(file.originalname, 'latin1');
        const correctedOriginalName = originalNameBuffer.toString('utf-8');
        // Keep original name (corrected) with a timestamp for uniqueness
        cb(null, `${Date.now()}-${correctedOriginalName}`);
    }
});

const xlsFileFilter = (req, file, cb) => {
    if (file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || file.originalname.toLowerCase().endsWith('.xlsx')) {
        cb(null, true);
    } else {
        cb(new Error('Invalid file type, only .xlsx files are allowed!'), false);
    }
};

const uploadXls = multer({ storage: xlsStorage, fileFilter: xlsFileFilter });

// Multer setup for email attachments (to /send-emails)
// Attachments will be temporarily stored and then deleted after sending.
const attachmentStorage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, uploadsDir); // Store attachments in the same uploads folder temporarily
    },
    filename: function (req, file, cb) {
        // Use a unique name for temp storage to avoid conflicts, but keep original extension
        cb(null, `attachment-${Date.now()}-${Math.round(Math.random() * 1E9)}${path.extname(file.originalname)}`);
    }
});
const uploadAttachments = multer({ storage: attachmentStorage });

// Nodemailer transporter setup
let transporter;
try {
    transporter = nodemailer.createTransport({
        host: process.env.EMAIL_HOST,
        port: parseInt(process.env.EMAIL_PORT, 10),
        secure: process.env.EMAIL_SECURE === 'true', // true for 465, false for other ports
        auth: {
            user: process.env.EMAIL_USER,
            pass: process.env.EMAIL_PASS,
        },
        // 徹底修正 self signed certificate 問題 (如果需要)
        // tls: {
        //     rejectUnauthorized: false // 不建議在 production environment 使用!
        // }
    });
    console.log("Nodemailer transporter configured successfully.");
    transporter.verify((error, success) => {
        if (error) {
            console.error("Nodemailer transporter verification failed:", error);
        } else {
            console.log("Nodemailer transporter is ready to send emails.");
        }
    });
} catch (error) {
    console.error("Failed to create Nodemailer transporter:", error);
    // May exit process or set a flag that mailing is disabled
}

// In-memory store for active sending jobs
const activeJobs = {};

// Helper function to send SSE messages
function sendSseMessage(res, eventName, data) {
    res.write(`event: ${eventName}\n`);
    res.write(`data: ${JSON.stringify(data)}\n\n`);
    // Flush the data to the client if possible (may not be supported by all Node versions/setups for res.write)
    if (res.flushHeaders) {
        res.flushHeaders();
    }
}

// --- Helper function to convert Quill HTML to email-safe HTML ---
function convertQuillHtmlToEmailSafeHtml(htmlBody) {
    if (!htmlBody) return '';
    const $ = cheerio.load(htmlBody);

    // Define mappings from Quill classes to pixel sizes
    const sizeMap = {
        'ql-size-small': '12px',
        'ql-size-large': '20px',
        'ql-size-huge': '28px'
        // Default size for normal text will be handled below or can be set here
    };

    // Define default sizes for headings
    const headingSizeMap = {
        'h1': '32px',
        'h2': '26px',
        'h3': '22px'
        // h4, h5, h6 can be added if Quill supports them and you want specific sizes
    };
    
    const defaultParagraphFontSize = '16px'; // Default for p, span if no other size is applied

    // Apply sizes based on Quill classes
    for (const qlClass in sizeMap) {
        $(`.${qlClass}`).each((i, el) => {
            const currentStyle = $(el).attr('style') || '';
            $(el).attr('style', `${currentStyle}font-size: ${sizeMap[qlClass]} !important;`);
            $(el).removeClass(qlClass); // Remove class after applying style
        });
    }

    // Apply sizes for headings
    for (const headingTag in headingSizeMap) {
        $(headingTag).each((i, el) => {
            const currentStyle = $(el).attr('style') || '';
            // Only apply if no more specific font-size is already set by a ql-size class
            if (!currentStyle.includes('font-size:')) {
                 $(el).attr('style', `${currentStyle}font-size: ${headingSizeMap[headingTag]} !important; margin: 0.5em 0;`); // Added basic margin for headings
            }
        });
    }
    
    // Apply default font size to all p and span elements if they don't have one
    $('p, span').each((i, el) => {
        const currentStyle = $(el).attr('style') || '';
        if (!currentStyle.includes('font-size:')) {
            $(el).attr('style', `${currentStyle}font-size: ${defaultParagraphFontSize};`);
        }
        // Ensure paragraphs have some basic margin for readability in emails
        if (el.name === 'p' && !currentStyle.includes('margin:')) {
             $(el).attr('style', `${$(el).attr('style')}margin: 10px 0;`);
        }
    });
    
    // Ensure all text elements (p, span, li, blockquote, etc.) have a base color if not specified
    // This helps if the email client has a weird default.
    $('p, span, li, blockquote, h1, h2, h3, h4, h5, h6, div, td, th').each((i, el) => {
        const currentStyle = $(el).attr('style') || '';
        if (!currentStyle.includes('color:')) {
            // $(el).attr('style', `${currentStyle}color: #333333;`); // A common default text color
        }
        // Ensure common block elements have some line-height for readability
        if (['p', 'li', 'div'].includes(el.name) && !currentStyle.includes('line-height:')) {
            $(el).attr('style', `${currentStyle}line-height: 1.6;`);
        }
    });


    // Handle text alignment classes (Quill uses ql-align-*)
    $('.ql-align-center').each((i, el) => {
        const currentStyle = $(el).attr('style') || '';
        $(el).attr('style', `${currentStyle}text-align: center !important;`);
        $(el).removeClass('ql-align-center');
    });
    $('.ql-align-right').each((i, el) => {
        const currentStyle = $(el).attr('style') || '';
        $(el).attr('style', `${currentStyle}text-align: right !important;`);
        $(el).removeClass('ql-align-right');
    });
    $('.ql-align-justify').each((i, el) => {
        const currentStyle = $(el).attr('style') || '';
        $(el).attr('style', `${currentStyle}text-align: justify !important;`);
        $(el).removeClass('ql-align-justify');
    });
    // .ql-align-left is default, but we can remove the class if it exists
    $('.ql-align-left').removeClass('ql-align-left');


    // Convert font classes if any (e.g., ql-font-serif)
    // This is an example, you might need to adjust based on Quill's actual output
    $('.ql-font-serif').each((i, el) => {
        const currentStyle = $(el).attr('style') || '';
        $(el).attr('style', `${currentStyle}font-family: serif !important;`);
        $(el).removeClass('ql-font-serif');
    });
    $('.ql-font-monospace').each((i, el) => {
        const currentStyle = $(el).attr('style') || '';
        $(el).attr('style', `${currentStyle}font-family: monospace !important;`);
        $(el).removeClass('ql-font-monospace');
    });
    // Default .ql-font-sans-serif can be removed
    $('.ql-font-sans-serif').removeClass('ql-font-sans-serif');


    // Color and background color (Quill might use classes or inline styles)
    // If Quill uses classes like ql-color-red, ql-bg-yellow:
    // $('[class*="ql-color-"]').each((i, el) => { ... map to actual color ... });
    // $('[class*="ql-bg-"]').each((i, el) => { ... map to actual background-color ... });
    // However, Quill often outputs inline styles for colors directly, which is good.
    // We just need to ensure `!important` is not overused by Quill itself if it causes issues.

    return $('body').html(); // Return the content of the body
}

// --- API Endpoints ---

// POST /upload-and-preview
app.post('/upload-and-preview', (req, res) => {
    uploadXls.single('xlsfile')(req, res, function (err) {
        if (err instanceof multer.MulterError) {
            // A Multer error occurred when uploading.
            return res.status(500).json({ message: `Multer error: ${err.message}` });
        } else if (err) {
            // An unknown error occurred when uploading, or fileFilter error.
            return res.status(400).json({ message: err.message || 'File upload error.' });
        }

        // Everything went fine, proceed with file processing.
        if (!req.file) {
            return res.status(400).json({ message: 'No file uploaded. Please select an .xlsx file.' });
        }

        const filePath = req.file.path;
        try {
            const workbook = xlsx.readFile(filePath);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            // Get headers (first row)
            const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            if (jsonData.length === 0) {
                return res.status(400).json({ message: 'XLS file is empty or has incorrect format.' });
            }
            const headers = jsonData[0].map(h => String(h).trim()); // Ensure headers are strings and trimmed

            // Validate standard headers (optional, but good practice)
            const requiredHeaders = ['email', 'title']; // status and send_time are for output
            for (const reqHeader of requiredHeaders) {
                if (!headers.includes(reqHeader)) {
                     // fs.unlinkSync(filePath); // Clean up uploaded file
                     console.warn(`XLS file might be missing a recommended header: ${reqHeader} (this will not affect operation, but it is recommended to include it)`);
                }
            }
            
            // Get preview data (e.g., first 5 data rows)
            const dataRows = xlsx.utils.sheet_to_json(worksheet, { defval: '' }); // Convert all to objects

            const previewData = dataRows.slice(0, 5); // Get first 5 data rows for preview

            // Correct the original filename encoding before sending to client
            const originalNameBuffer = Buffer.from(req.file.originalname, 'latin1');
            const correctedOriginalName = originalNameBuffer.toString('utf-8');

            res.json({
                message: 'File uploaded and parsed successfully',
                fileName: req.file.filename, // This is timestamped-correctedOriginalName from xlsStorage
                filePath: req.file.filename, // This is timestamped-correctedOriginalName from xlsStorage
                originalXlsName: correctedOriginalName, // Use the corrected name here
                headers: headers,
                rowCount: dataRows.length, // Number of data rows (excluding header)
                previewData: previewData 
            });

        } catch (error) {
            console.error('Failed to parse XLS file:', error);
            if (fs.existsSync(filePath)) {
               // fs.unlinkSync(filePath); // Clean up uploaded file in case of error
            }
            res.status(500).json({ message: `Failed to parse .xlsx file: ${error.message}` });
        }
    });
});

// POST /initiate-sending-job
app.post('/initiate-sending-job', uploadAttachments.array('attachments', 10), async (req, res) => {
    const { filePath: relativeXlsFilePath, subjectTemplate, bodyTemplate, headers: headersString, sendInterval, originalXlsName } = req.body;
    const receivedHeaders = JSON.parse(headersString || '[]');
    const uploadedAttachmentFiles = req.files || [];

    // Basic validation (can be expanded)
    if (!relativeXlsFilePath || !subjectTemplate || !bodyTemplate || !receivedHeaders || !originalXlsName) {
        uploadedAttachmentFiles.forEach(file => fs.unlink(file.path, err => { if(err) console.error("Error deleting temp attachment during init validation:", err); }));
        return res.status(400).json({ message: 'Missing required parameters (XLS file, templates, headers, original XLS name)' });
    }

    const actualXlsFilePath = path.join(uploadsDir, relativeXlsFilePath);
    if (!fs.existsSync(actualXlsFilePath)) {
        uploadedAttachmentFiles.forEach(file => fs.unlink(file.path, err => { if(err) console.error("Error deleting temp attachment, XLS not found:", err); }));
        return res.status(404).json({ message: `XLS file not found: ${actualXlsFilePath}` });
    }

    const jobId = uuidv4();
    activeJobs[jobId] = {
        relativeXlsFilePath,
        actualXlsFilePath,
        originalXlsName,
        subjectTemplate,
        bodyTemplate,
        receivedHeaders,
        sendIntervalMs: (parseInt(sendInterval, 10) || 0) * 1000,
        nodemailerAttachments: uploadedAttachmentFiles.map(file => {
            // 嘗試修正文件名編碼：如果UTF-8字節被錯誤地解析為latin1，則進行還原
            const originalNameBuffer = Buffer.from(file.originalname, 'latin1');
            const correctedOriginalName = originalNameBuffer.toString('utf-8');
            return {
                filename: correctedOriginalName, // 使用修正後的UTF-8文件名
                path: file.path
            };
        }),
        status: 'pending',
        creationTime: Date.now()
    };

    console.log(`Job created: ${jobId} for XLS: ${relativeXlsFilePath}`);
    res.json({ jobId });
});

// GET /send-emails-stream
app.get('/send-emails-stream', async (req, res) => {
    const { jobId } = req.query;
    const job = activeJobs[jobId];

    if (!job || job.status !== 'pending') {
        return res.status(404).json({ message: 'Invalid job ID or job already processed/processing' });
    }

    job.status = 'processing'; // Mark job as processing

    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    res.flushHeaders(); // Flush the headers to establish the connection

    sendSseMessage(res, 'job_started', { jobId, message: 'Email sending task has started...' });

    // --- Start of the actual email sending logic (moved from old /send-emails) ---
    const SAVE_BATCH_SIZE = 20; 
    let successCount = 0;
    let failCount = 0;
    const errors = [];
    let emailsAttemptedThisRun = 0;
    let workbook;
    let worksheet;
    const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

    try {
        workbook = xlsx.readFile(job.actualXlsFilePath);
        const sheetName = workbook.SheetNames[0];
        worksheet = workbook.Sheets[sheetName];
        const dataObjects = xlsx.utils.sheet_to_json(worksheet, { defval: '' });

        sendSseMessage(res, 'data_loaded', { totalRows: dataObjects.length });
        
        if (dataObjects.length === 0) {
            sendSseMessage(res, 'error', { message: 'No data to process in XLS file.' });
            throw new Error('No data to process in XLS file.'); // Will be caught by outer catch
        }

        const statusColIndex = job.receivedHeaders.indexOf('status');
        const sendTimeColIndex = job.receivedHeaders.indexOf('send_time');
        const emailColHeader = job.receivedHeaders.find(h => String(h).trim().toLowerCase() === 'email');
        const titleColHeader = job.receivedHeaders.find(h => String(h).trim().toLowerCase() === 'title');

        if (!emailColHeader) {
            sendSseMessage(res, 'error', { message: 'Could not find "email" header in XLS file.' });
            throw new Error('Could not find "email" header in XLS file.');
        }

        for (let i = 0; i < dataObjects.length; i++) {
            const rowData = dataObjects[i];
            const rowIndexForDisplay = i + 1; 
            const dataRowIndexInSheet = i + 1; 
            let currentRowStatus = 'pending';
            let currentRowError = null;

            try {
                let shouldSkip = false;
                if (statusColIndex !== -1 && job.receivedHeaders[statusColIndex]) {
                    const currentStatusInFile = String(rowData[job.receivedHeaders[statusColIndex]] || '').trim().toLowerCase();
                    if (currentStatusInFile === 'success') {
                        shouldSkip = true;
                        currentRowStatus = 'skipped_previously_success';
                    }
                }
                if (shouldSkip) {
                    sendSseMessage(res, 'progress', { jobId, rowIndex: rowIndexForDisplay, email: rowData[emailColHeader], status: currentRowStatus });
                    continue; 
                }

                const recipientEmail = String(rowData[emailColHeader] || '').trim();
                if (!recipientEmail || !recipientEmail.includes('@')) {
                    failCount++;
                    currentRowStatus = 'validation_failed';
                    currentRowError = `Row ${rowIndexForDisplay} (${recipientEmail || 'empty email'}): invalid or empty email address.`;
                    errors.push({ rowIndex: rowIndexForDisplay, email: recipientEmail, error: currentRowError });
                    if (statusColIndex !== -1) xlsx.utils.sheet_add_aoa(worksheet, [['Format Error']], { origin: xlsx.utils.encode_cell({c:statusColIndex, r:dataRowIndexInSheet})});
                    if (sendTimeColIndex !== -1) xlsx.utils.sheet_add_aoa(worksheet, [[new Date()]], { origin: xlsx.utils.encode_cell({c:sendTimeColIndex, r:dataRowIndexInSheet})});
                } else {
                    if (emailsAttemptedThisRun > 0 && job.sendIntervalMs > 0) await sleep(job.sendIntervalMs);
                    emailsAttemptedThisRun++;

                    let mailSubject = job.subjectTemplate;
                    let mailBody = job.bodyTemplate; // Original HTML from Quill
                    
                    // Replace placeholders
                    for (const header of job.receivedHeaders) {
                        if (header) { 
                            const placeholder = `{{${String(header)}}}`;
                            const value = String(rowData[header] || ''); 
                            mailSubject = mailSubject.replace(new RegExp(placeholder.replace(/[.*+?^${}()|[\\]\\]/g, '\\$&'), 'g'), value);
                            mailBody = mailBody.replace(new RegExp(placeholder.replace(/[.*+?^${}()|[\\]\\]/g, '\\$&'), 'g'), value);
                        }
                    }
                    const specificTitle = (titleColHeader && rowData[titleColHeader]) ? String(rowData[titleColHeader]).trim() : '';
                    if (specificTitle && !mailSubject.includes(`{{${String(titleColHeader)}}}`)) mailSubject = specificTitle;

                    // Convert Quill HTML for email clients (font sizes, etc.)
                    let emailSafeHtmlBody = convertQuillHtmlToEmailSafeHtml(mailBody); // <<<--- 轉換HTML

                    // Handle embedded Base64 images (after initial placeholder replacement and HTML conversion)
                    const newNodemailerAttachments = [...job.nodemailerAttachments]; 
                    let imageCounterForCid = 0;
                    const imgRegex = /<img[^>]+src="data:(image\/(png|jpeg|gif|webp))(?:;charset=[^;]+)?(?:;name=([^;]+))?;base64,([^\"\s]+)"([^>]*)>/gi;
                    
                    emailSafeHtmlBody = emailSafeHtmlBody.replace(imgRegex, (fullMatch, fullMimeType, imageType, embeddedFilename, base64Data, otherAttributes) => {
                        imageCounterForCid++;
                        const cid = `emb_${jobId}_${i}_${imageCounterForCid}`;
                        newNodemailerAttachments.push({
                            content: base64Data, 
                            encoding: 'base64',
                            cid: cid, 
                            contentType: fullMimeType,
                        });
                        return `<img src="cid:${cid}"${otherAttributes}>`;
                    });

                    const mailOptions = {
                        from: `"${process.env.EMAIL_FROM_NAME}" <${process.env.EMAIL_FROM_EMAIL}>`,
                        to: recipientEmail,
                        subject: mailSubject,
                        html: emailSafeHtmlBody, // <<<--- 使用轉換後的 HTML
                        attachments: newNodemailerAttachments 
                    };
                    
                    if (!transporter) throw new Error("Nodemailer transporter is not configured.");
                    await transporter.sendMail(mailOptions);
                    successCount++;
                    currentRowStatus = 'success';
                    if (statusColIndex !== -1) xlsx.utils.sheet_add_aoa(worksheet, [['Success']], { origin: xlsx.utils.encode_cell({c:statusColIndex, r:dataRowIndexInSheet})});
                }
            } catch (emailError) {
                failCount++;
                currentRowStatus = 'send_failed';
                currentRowError = emailError.message;
                errors.push({ rowIndex: rowIndexForDisplay, email: rowData[emailColHeader] || 'N/A', error: currentRowError });
                if (statusColIndex !== -1) xlsx.utils.sheet_add_aoa(worksheet, [[`Failed: ${currentRowError.substring(0,100)}`]], { origin: xlsx.utils.encode_cell({c:statusColIndex, r:dataRowIndexInSheet})});
            } finally {
                if (currentRowStatus !== 'skipped_previously_success' && currentRowStatus !== 'validation_failed') { // only update time if an attempt was made
                     if (sendTimeColIndex !== -1) xlsx.utils.sheet_add_aoa(worksheet, [[new Date()]], { origin: xlsx.utils.encode_cell({c:sendTimeColIndex, r:dataRowIndexInSheet})});
                }
                sendSseMessage(res, 'progress', { jobId, rowIndex: rowIndexForDisplay, email: rowData[emailColHeader], status: currentRowStatus, error: currentRowError });
            }

            if ((i + 1) % SAVE_BATCH_SIZE === 0 || (i + 1) === dataObjects.length) {
                xlsx.writeFile(workbook, job.actualXlsFilePath);
                sendSseMessage(res, 'batch_save', { jobId, message: `Processed ${i + 1} rows, XLS progress saved.` });
            }
        } // End of for loop

        // Extract original filename for download attribute
        let suggestedDownloadName = job.originalXlsName; // Use the stored original name
        // Fallback if originalXlsName wasn't passed or stored correctly (though validation should catch it)
        if (!suggestedDownloadName) {
            const parts = job.relativeXlsFilePath.match(/^(\d+)-(.*)$/);
            if (parts && parts[2]) {
                suggestedDownloadName = parts[2]; 
            } else {
                suggestedDownloadName = job.relativeXlsFilePath; // Full path with timestamp as last resort
            }
        }
        
        sendSseMessage(res, 'complete', {
            jobId,
            message: 'Email processing completed',
            totalRows: dataObjects.length,
            successCount,
            failCount,
            errors,
            outputFilePath: `/uploads/${job.relativeXlsFilePath}`,
            suggestedDownloadName: suggestedDownloadName // Provide the original name
        });

    } catch (error) {
        console.error(`Job ${jobId} failed:`, error);
        sendSseMessage(res, 'error', { jobId, message: `Internal server error: ${error.message}` });
        job.status = 'failed'; 
    } finally {
        // --- Clean up job and attachments --- 
        job.nodemailerAttachments.forEach(attachment => {
            fs.unlink(attachment.path, (err) => {
                if (err) console.error(`Failed to clean up attachment ${attachment.path} for job ${jobId}:`, err);
                else console.log(`Cleaned up attachment for job ${jobId}: ${attachment.path}`);
            });
        });
        delete activeJobs[jobId];
        console.log(`Job ${jobId} processing finished and cleaned up.`);
        res.end();
    }
});

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
    console.log("Please ensure your .env file is correctly configured with SMTP server information.");
    console.log("XLS file requirements: First row as header. Column A=email, Column B=title, Column C=status (auto-filled), Column D=send_time (auto-filled).");
    if (isPackaged) {
        console.log("EmailGo is running as a packaged application.");
        console.log("Press Ctrl+C in this window to stop the server.");
    }
});

// Keep the process alive and window open when double-clicked as an EXE
if (isPackaged) {
    process.stdin.resume(); 
    process.on('SIGINT', () => {
        console.log('SIGINT received, shutting down...');
        process.exit(0);
    });
    process.on('SIGTERM', () => {
        console.log('SIGTERM received, shutting down...');
        process.exit(0);
    });
    // Optional: A more explicit way to keep it open, e.g., wait for a key press
    // console.log("Press any key to exit.");
    // process.stdin.setRawMode(true);
    // process.stdin.resume();
    // process.stdin.on('data', process.exit.bind(process, 0));
}
