const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');
const path = require('path');
const fs = require('fs');
const os = require('os'); // Added os module
const { exec } = require('child_process'); // Added exec from child_process
const { v4: uuidv4 } = require('uuid');
const cheerio = require('cheerio');
const net = require('net'); // Required for port checking

const isPackaged = typeof process.pkg !== 'undefined';
const appDataDir = path.join(os.homedir(), '.emailGo');
const configFilePath = path.join(appDataDir, 'config.json');
const projectBaseDir = isPackaged ? path.dirname(process.execPath) : __dirname; // For public and uploads relative to app

// --- Configuration Management ---
let appConfig = {
    EMAIL_HOST: '',
    EMAIL_PORT: 587,
    EMAIL_SECURE: false,
    EMAIL_USER: '',
    EMAIL_PASS: '',
    EMAIL_FROM_NAME: 'EmailGo App',
    EMAIL_FROM_EMAIL: ''
};

function ensureConfigDirExists() {
    if (!fs.existsSync(appDataDir)) {
        try {
            fs.mkdirSync(appDataDir, { recursive: true });
            console.log(`Configuration directory created at: ${appDataDir}`);
        } catch (error) {
            console.error(`Error creating configuration directory at ${appDataDir}:`, error);
            // If we can't create the dir, we might have to fall back or exit
            // For now, we'll proceed and saving/loading might fail, which will be logged.
        }
    }
}

function loadConfig() {
    ensureConfigDirExists(); // Make sure directory exists before trying to read/write
    try {
        if (fs.existsSync(configFilePath)) {
            console.log(`Loading configuration from: ${configFilePath}`);
            const rawData = fs.readFileSync(configFilePath);
            let jsonData = JSON.parse(rawData);
            delete jsonData.PORT; // Remove PORT if it exists from old config
            appConfig = { ...appConfig, ...jsonData };
            console.log("Configuration loaded successfully.");
        } else {
            console.log(`Configuration file not found at ${configFilePath}. Creating a new one with default values.`);
            // Save defaults (appConfig doesn't have PORT, so it won't be saved)
            const configToSave = { ...appConfig }; // Create a fresh copy from current appConfig
            // delete configToSave.PORT; // Not strictly necessary as appConfig doesn't have it, but for safety.
            fs.writeFileSync(configFilePath, JSON.stringify(configToSave, null, 2));
            console.log(`Default configuration saved to ${configFilePath}`);
        }
    } catch (error) {
        console.error('Error loading or parsing config.json:', error);
        console.log('Using in-memory default configuration due to error.');
    }
}

function saveConfig(newConfig) {
    ensureConfigDirExists(); // Ensure directory exists
    try {
        if (newConfig) {
            const newConfigFiltered = { ...newConfig };
            delete newConfigFiltered.PORT; // Ensure PORT is not processed from incoming newConfig

            if (typeof newConfigFiltered.EMAIL_PORT === 'string') newConfigFiltered.EMAIL_PORT = parseInt(newConfigFiltered.EMAIL_PORT, 10);
            if (typeof newConfigFiltered.EMAIL_SECURE === 'string') newConfigFiltered.EMAIL_SECURE = newConfigFiltered.EMAIL_SECURE.toLowerCase() === 'true';
            
            appConfig = { ...appConfig, ...newConfigFiltered };
        }
        
        const configToSave = { ...appConfig };
        delete configToSave.PORT; // Explicitly ensure PORT is not saved

        fs.writeFileSync(configFilePath, JSON.stringify(configToSave, null, 2));
        console.log(`Configuration saved to ${configFilePath}`);
        
        if (newConfig && (newConfig.hasOwnProperty('EMAIL_HOST') || newConfig.hasOwnProperty('EMAIL_USER') || newConfig.hasOwnProperty('EMAIL_PASS') || newConfig.hasOwnProperty('EMAIL_PORT') || newConfig.hasOwnProperty('EMAIL_SECURE'))) {
            console.log("SMTP configuration changed, re-initializing Nodemailer.");
            initializeNodemailer();
        }
    } catch (error) {
        console.error('Error saving config.json:', error);
    }
}

const app = express();

// Configuration validation function
function isSmtpConfigComplete() {
    return appConfig.EMAIL_HOST && appConfig.EMAIL_USER && appConfig.EMAIL_FROM_EMAIL && appConfig.EMAIL_USER === appConfig.EMAIL_FROM_EMAIL;
}

// Initial Nodemailer setup
let transporter;
function initializeNodemailer() {
    if (!appConfig.EMAIL_HOST || !appConfig.EMAIL_USER || !appConfig.EMAIL_PASS || !appConfig.EMAIL_FROM_EMAIL) {
        console.warn("Nodemailer transporter cannot be initialized: Critical SMTP settings (Host, User, Pass, From Email) are missing in config.json.");
        transporter = null;
        return;
    }
    if (appConfig.EMAIL_USER !== appConfig.EMAIL_FROM_EMAIL) {
        console.warn(`Nodemailer transporter cannot be initialized: EMAIL_USER ("${appConfig.EMAIL_USER}") and EMAIL_FROM_EMAIL ("${appConfig.EMAIL_FROM_EMAIL}") must match. Please update config.json.`);
        transporter = null;
        return;
    }

    try {
        transporter = nodemailer.createTransport({
            host: appConfig.EMAIL_HOST,
            port: parseInt(appConfig.EMAIL_PORT, 10),
            secure: appConfig.EMAIL_SECURE === true || String(appConfig.EMAIL_SECURE).toLowerCase() === 'true',
            auth: {
                user: appConfig.EMAIL_USER,
                pass: appConfig.EMAIL_PASS,
            },
        });
        console.log("Nodemailer transporter configured/re-configured.");
        transporter.verify((error, success) => {
            if (error) {
                console.error("Nodemailer transporter verification failed:", error.message);
            } else {
                console.log("Nodemailer transporter is ready to send emails.");
            }
        });
    } catch (error) {
        console.error("Failed to create/re-create Nodemailer transporter:", error);
        transporter = null;
    }
}

// Middlewares
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const publicDir = path.join(projectBaseDir, 'public'); // Use projectBaseDir
const uploadsDir = path.join(projectBaseDir, 'uploads'); // Use projectBaseDir

app.use(express.static(publicDir));
app.use('/uploads', express.static(uploadsDir));

if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
}

// --- API Endpoints for Configuration ---
app.get('/api/config', (req, res) => {
    // Return a version of the config suitable for the client, omitting sensitive details like passwords
    const clientSafeConfig = { ...appConfig };
    delete clientSafeConfig.EMAIL_PASS; // Never send password to client
    res.json(clientSafeConfig);
});

app.post('/api/config', (req, res) => {
    const newConfigData = req.body;
    // Add validation for received data here (e.g., check types, required fields)
    if (!newConfigData.EMAIL_HOST || !newConfigData.EMAIL_USER || !newConfigData.EMAIL_FROM_EMAIL) {
        return res.status(400).json({ message: "Missing required SMTP fields: Host, User, From Email." });
    }
    if (newConfigData.EMAIL_USER !== newConfigData.EMAIL_FROM_EMAIL) {
        return res.status(400).json({ message: "EMAIL_USER and EMAIL_FROM_EMAIL must match." });
    }
    // Securely handle EMAIL_PASS - if it's empty or not provided, do not clear an existing password
    // unless explicitly intended. If it's provided, update it.
    // For simplicity now, we just take what's given.
    
    saveConfig(newConfigData); // Save the received config
    res.json({ message: 'Configuration updated successfully. Nodemailer re-initialized if necessary.' });
});

app.get('/api/config/status', (req, res) => {
    const complete = isSmtpConfigComplete();
    const message = complete ? "SMTP configuration is complete." : "SMTP configuration is incomplete. Please configure SMTP settings.";
    res.json({
        isConfigured: complete,
        message: message,
        // Optionally include which fields are missing for more detailed feedback
        // missingFields: [...] 
    });
});

// Multer setup for initial XLS file upload (to /upload-and-preview)
const xlsStorage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, uploadsDir);
    },
    filename: function (req, file, cb) {
        const originalNameBuffer = Buffer.from(file.originalname, 'latin1');
        const correctedOriginalName = originalNameBuffer.toString('utf-8');
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
            const dataRows = xlsx.utils.sheet_to_json(worksheet, { defval: '' });

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
                if (statusColIndex !== -1 && job.receivedHeaders[statusIndex]) {
                    const currentStatusInFile = String(rowData[job.receivedHeaders[statusIndex]] || '').trim().toLowerCase();
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
                        from: `"${appConfig.EMAIL_FROM_NAME}" <${appConfig.EMAIL_FROM_EMAIL}>`,
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
    res.sendFile(path.join(publicDir, 'index.html'));
});

// Keep track of whether the browser has been opened
let browserOpened = false;

// --- Dynamic Port Finding Logic ---
async function tryPort(portToTry) {
    return new Promise((resolve, reject) => {
        const server = net.createServer();
        server.unref(); // Allow program to exit if this is the only active server.
        server.once('error', (err) => {
            if (err.code === 'EADDRINUSE') {
                resolve(false); // Port is in use
            } else {
                reject(err); // Other error
            }
        });
        server.once('listening', () => {
            server.close(() => {
                resolve(true); // Port is available
            });
        });
        server.listen(portToTry, '127.0.0.1');
    });
}

async function findAvailablePort(initialPort, maxTotalAttempts = 10) {
    const triedPorts = new Set();

    // Attempt 1: Initial port
    console.log(`Attempting to use initial port ${initialPort}...`);
    triedPorts.add(initialPort);
    try {
        const isAvailable = await tryPort(initialPort);
        if (isAvailable) {
            console.log(`Port ${initialPort} is available.`);
            return initialPort;
        }
        console.log(`Port ${initialPort} is in use.`);
    } catch (error) {
        // Log error but continue to random attempts, as initial port might be an expected conflict
        console.error(`Error checking initial port ${initialPort}: ${error.message}. Will proceed to try random ports.`);
    }

    // Attempts 2 to maxTotalAttempts: Random ports
    const randomAttemptsToMake = maxTotalAttempts - 1;
    if (randomAttemptsToMake > 0) {
        console.log(`Initial port ${initialPort} is unavailable or errored. Attempting to find a random available port (${randomAttemptsToMake} attempts left)...`);
    }

    for (let i = 0; i < randomAttemptsToMake; i++) {
        let randomPort;
        let safetyBreak = 0; // To prevent infinite loop if all random ports in range are tried
        const minRandom = 7000;
        const maxRandom = 9999;

        do {
            randomPort = Math.floor(Math.random() * (maxRandom - minRandom + 1)) + minRandom;
            safetyBreak++;
            if (safetyBreak > (maxRandom - minRandom + 1) * 2 && safetyBreak > 50) { // Heuristic safety break
                // This condition means we've tried to find a new unique random port many times
                console.warn("Having trouble finding a new unique random port to test. There might be few available ports in the chosen random range.");
                break; // Break from do-while, will effectively skip this random attempt
            }
        } while (triedPorts.has(randomPort));

        if (triedPorts.has(randomPort) && safetyBreak > 50) { // If do-while was broken by safetyBreak and port is still a duplicate (shouldn't happen with proper break)
            continue; // Skip this attempt
        }

        console.log(`Attempting random port ${randomPort} (attempt ${i + 2}/${maxTotalAttempts})...`);
        triedPorts.add(randomPort);
        try {
            const isAvailable = await tryPort(randomPort);
            if (isAvailable) {
                console.log(`Port ${randomPort} is available.`);
                return randomPort;
            }
            console.log(`Port ${randomPort} is in use.`);
        } catch (error) {
            console.error(`Error checking random port ${randomPort}: ${error.message}`);
            // Continue to the next random attempt
        }
    }

    throw new Error(`Could not find an available port after ${maxTotalAttempts} attempts (initial port ${initialPort} and ${randomAttemptsToMake} random ports in range ${7000}-${9999}).`);
}

// --- Main Application Start Function ---
async function startApp() {
    loadConfig(); // Load SMTP config
    initializeNodemailer(); // Initialize Nodemailer

    const app = express();

    // Middlewares
    app.use(express.json());
    app.use(express.urlencoded({ extended: true }));

    // For packaged app (pkg), __dirname is the root. For dev, it's the current file's dir.
    // This ensures that 'public' is correctly found whether running from source or packaged.
    const publicDir = path.join(__dirname, 'public'); 
    const uploadsDir = path.join(projectBaseDir, 'uploads'); // uploads are external to the package

    app.use(express.static(publicDir)); 
    app.use('/uploads', express.static(uploadsDir));

    if (!fs.existsSync(uploadsDir)) {
        fs.mkdirSync(uploadsDir, { recursive: true });
    }

    // --- API Endpoints for Configuration ---
    app.get('/api/config', (req, res) => {
        const clientSafeConfig = { ...appConfig };
        // PORT is not in appConfig, so no need to delete explicitly
        delete clientSafeConfig.EMAIL_PASS; // Ensure password is not sent
        res.json(clientSafeConfig);
    });

    app.post('/api/config', (req, res) => {
        const newConfigData = req.body;
        if (!newConfigData.EMAIL_HOST || !newConfigData.EMAIL_USER || !newConfigData.EMAIL_FROM_EMAIL) {
            return res.status(400).json({ message: "Missing required SMTP fields: Host, User, From Email." });
        }
        if (newConfigData.EMAIL_USER !== newConfigData.EMAIL_FROM_EMAIL) {
            return res.status(400).json({ message: "EMAIL_USER and EMAIL_FROM_EMAIL must match." });
        }
        saveConfig(newConfigData);
        res.json({ message: 'Configuration updated successfully.' });
    });

    app.get('/api/config/status', (req, res) => {
        const complete = isSmtpConfigComplete();
        const message = complete ? "SMTP configuration is complete." : "SMTP configuration is incomplete. Please configure SMTP settings.";
        res.json({ isConfigured: complete, message: message });
    });

    // Multer setup
    const xlsStorage = multer.diskStorage({
        destination: (req, file, cb) => cb(null, uploadsDir),
        filename: (req, file, cb) => {
            const originalNameBuffer = Buffer.from(file.originalname, 'latin1');
            const correctedOriginalName = originalNameBuffer.toString('utf-8');
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
    const attachmentStorage = multer.diskStorage({
        destination: (req, file, cb) => cb(null, uploadsDir),
        filename: (req, file, cb) => cb(null, `attachment-${Date.now()}-${Math.round(Math.random() * 1E9)}${path.extname(file.originalname)}`)
    });
    const uploadAttachments = multer({ storage: attachmentStorage });

    // --- Other API Endpoints (now defined within startApp) ---
    app.post('/upload-and-preview', uploadXls.single('xlsfile'), (req, res) => {
        // ... (Implementation from existing global scope, ensuring no multer errors are missed)
        if (req.fileValidationError) { // Example if fileFilter passes error via req
            return res.status(400).json({ message: req.fileValidationError });
        }
        if (err instanceof multer.MulterError) { // From original global scope error handling
            return res.status(500).json({ message: `Multer error: ${err.message}` });
        }
        if (err) { // From original global scope error handling
            return res.status(400).json({ message: err.message || 'File upload error.' });
        }
        if (!req.file) {
            return res.status(400).json({ message: 'No file uploaded or invalid file type.' });
        }
        const filePath = req.file.path;
        try {
            const workbook = xlsx.readFile(filePath);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = xlsx.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            if (jsonData.length === 0) return res.status(400).json({ message: 'XLS file is empty.' });
            const headers = jsonData[0].map(h => String(h).trim());
            const dataRows = xlsx.utils.sheet_to_json(worksheet, { defval: '' });
            const previewData = dataRows.slice(0, 5);
            const originalNameBuffer = Buffer.from(req.file.originalname, 'latin1');
            const correctedOriginalName = originalNameBuffer.toString('utf-8');
            res.json({
                message: 'File uploaded and parsed successfully',
                filePath: req.file.filename,
                originalXlsName: correctedOriginalName,
                headers: headers,
                rowCount: dataRows.length,
                previewData: previewData
            });
        } catch (error) {
            console.error('Failed to parse XLS file:', error);
            // if (fs.existsSync(filePath)) { // Consider if unlinking is needed here on error
            //    fs.unlinkSync(filePath);
            // }
            res.status(500).json({ message: `Failed to parse .xlsx file: ${error.message}` });
        }
    });
    
    app.post('/initiate-sending-job', uploadAttachments.array('attachments', 10), async (req, res) => {
        // ... (Implementation from existing global scope)
        const { filePath: relativeXlsFilePath, subjectTemplate, bodyTemplate, headers: headersString, sendInterval, originalXlsName } = req.body;
        const receivedHeaders = JSON.parse(headersString || '[]');
        const uploadedAttachmentFiles = req.files || [];

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
                const originalNameBuffer = Buffer.from(file.originalname, 'latin1');
                const correctedOriginalName = originalNameBuffer.toString('utf-8');
                return {
                    filename: correctedOriginalName,
                    path: file.path
                };
            }),
            status: 'pending',
            creationTime: Date.now()
        };
        console.log(`Job created: ${jobId} for XLS: ${relativeXlsFilePath}`);
        res.json({ jobId });
    });

    app.get('/send-emails-stream', async (req, res) => {
        // ... (Implementation from existing global scope, ensure `err` from multer in uploadXls.single is not used here)
        const { jobId } = req.query;
        const job = activeJobs[jobId];

        if (!job || job.status !== 'pending') {
            return res.status(404).json({ message: 'Invalid job ID or job already processed/processing' });
        }

        job.status = 'processing'; 
        res.setHeader('Content-Type', 'text/event-stream');
        res.setHeader('Cache-Control', 'no-cache');
        res.setHeader('Connection', 'keep-alive');
        res.flushHeaders(); 
        sendSseMessage(res, 'job_started', { jobId, message: 'Email sending task has started...' });
        
        const SAVE_BATCH_SIZE=20;let sC=0;let fC=0;const errs=[];let eATR=0;let wB;let wS;
        const sleep=(ms)=>new Promise(r=>setTimeout(r,ms));
        try{
            wB=xlsx.readFile(job.actualXlsFilePath);wS=wB.Sheets[wB.SheetNames[0]];
            const dO=xlsx.utils.sheet_to_json(wS,{defval:''});sendSseMessage(res,'data_loaded',{totalRows:dO.length});
            if(dO.length===0)throw new Error('No data in XLS.');
            const sCI=job.receivedHeaders.indexOf('status');const sTI=job.receivedHeaders.indexOf('send_time');
            const eCH=job.receivedHeaders.find(h=>String(h).trim().toLowerCase()==='email');
            const tCH=job.receivedHeaders.find(h=>String(h).trim().toLowerCase()==='title');
            if(!eCH)throw new Error('No "email" header.');
            for(let i=0;i<dO.length;i++){
                const rD=dO[i];const rDI=i+1;const dRIS=i+1;let cRS='pending';let cRE=null;
                try{
                    let skip=false;
                    if(sCI!==-1&&job.receivedHeaders[sCI]){if(String(rD[job.receivedHeaders[sCI]]||'').trim().toLowerCase()==='success'){skip=true;cRS='skipped_previously_success';}}
                    if(skip){sendSseMessage(res,'progress',{jobId,rowIndex:rDI,email:rD[eCH],status:cRS});continue;}
                    const rE=String(rD[eCH]||'').trim();
                    if(!rE||!rE.includes('@')){fC++;cRS='validation_failed';cRE=`Invalid email: ${rE||'empty'}`;errs.push({rowIndex:rDI,email:rE,error:cRE});
                        if(sCI!==-1)xlsx.utils.sheet_add_aoa(wS,[['Format Error']],{origin:xlsx.utils.encode_cell({c:sCI,r:dRIS})});
                        if(sTI!==-1)xlsx.utils.sheet_add_aoa(wS,[[new Date()]],{origin:xlsx.utils.encode_cell({c:sTI,r:dRIS})});
                    }else{
                        if(eATR>0&&job.sendIntervalMs>0)await sleep(job.sendIntervalMs);eATR++;
                        let mS=job.subjectTemplate;let mB=job.bodyTemplate;
                        for(const h of job.receivedHeaders){if(h){const p=`{{${String(h)}}}`;const v=String(rD[h]||'');mS=mS.replace(new RegExp(p.replace(/[.*+?^${}()|[\]\\]/g,'\$&'),'g'),v);mB=mB.replace(new RegExp(p.replace(/[.*+?^${}()|[\]\\]/g,'\$&'),'g'),v);}}
                        const sT=(tCH&&rD[tCH])?String(rD[tCH]).trim():'';if(sT&&!mS.includes(`{{${String(tCH)}}}`))mS=sT;
                        let eSHB=convertQuillHtmlToEmailSafeHtml(mB);
                        const nNA=[...job.nodemailerAttachments];let iCFC=0;
                        const iR=/<img[^>]+src="data:(image\/(png|jpeg|gif|webp))(?:;charset=[^;]+)?(?:;name=([^;]+))?;base64,([^\"\s]+)"([^>]*)>/gi;
                        eSHB=eSHB.replace(iR,(fM,fMT,iT,eFN,bD,oA)=>{iCFC++;const cid=`emb_${jobId}_${i}_${iCFC}`;nNA.push({content:bD,encoding:'base64',cid:cid,contentType:fMT});return `<img src="cid:${cid}"${oA}>`;});
                        const mO={from:`"${appConfig.EMAIL_FROM_NAME}" <${appConfig.EMAIL_FROM_EMAIL}>`,to:rE,subject:mS,html:eSHB,attachments:nNA};
                        if(!transporter)throw new Error("Nodemailer not configured.");
                        await transporter.sendMail(mO);sC++;cRS='success';
                        if(sCI!==-1)xlsx.utils.sheet_add_aoa(wS,[['Success']],{origin:xlsx.utils.encode_cell({c:sCI,r:dRIS})});
                    }
                }catch(eErr){fC++;cRS='send_failed';cRE=eErr.message;errs.push({rowIndex:rDI,email:rD[eCH]||'N/A',error:cRE});
                    if(sCI!==-1)xlsx.utils.sheet_add_aoa(wS,[[`Failed: ${cRE.substring(0,100)}`]],{origin:xlsx.utils.encode_cell({c:sCI,r:dRIS})});
                }finally{
                    if(cRS!=='skipped_previously_success'&&cRS!=='validation_failed'){if(sTI!==-1)xlsx.utils.sheet_add_aoa(wS,[[new Date()]],{origin:xlsx.utils.encode_cell({c:sTI,r:dRIS})});}
                    sendSseMessage(res,'progress',{jobId,rowIndex:rDI,email:rD[eCH],status:cRS,error:cRE});
                }
                if((i+1)%SAVE_BATCH_SIZE===0||(i+1)===dO.length){xlsx.writeFile(wB,job.actualXlsFilePath);sendSseMessage(res,'batch_save',{jobId,message:`Saved ${i+1} rows.`});}
            }
            let sDN=job.originalXlsName;
            if(!sDN){const pts=job.relativeXlsFilePath.match(/^(\d+)-(.*)$/);sDN=(pts&&pts[2])?pts[2]:job.relativeXlsFilePath;}
            sendSseMessage(res,'complete',{jobId,totalRows:dO.length,successCount:sC,failCount:fC,errors:errs,outputFilePath:`/uploads/${job.relativeXlsFilePath}`,suggestedDownloadName:sDN});
        }catch(err){console.error(`Job ${jobId} failed:`,err);sendSseMessage(res,'error',{jobId,message:`Error: ${err.message}`});job.status='failed';
        }finally{
            job.nodemailerAttachments.forEach(att=>fs.unlink(att.path,e=>{if(e)console.error(`Failed to cleanup ${att.path}:`,e);}));
            delete activeJobs[jobId];console.log(`Job ${jobId} finished.`);res.end();
        }
    });

    app.get('/', (req, res) => {
        res.sendFile(path.join(publicDir, 'index.html'));
    });

    // Start listening on a dynamically found port
    try {
        const initialPortToTry = 8600;
        const actualPort = await findAvailablePort(initialPortToTry, 10);
        
        let browserOpened = false;

        app.listen(actualPort, '127.0.0.1', () => {
            const serverUrl = `http://127.0.0.1:${actualPort}`;
            console.log(`Server running at ${serverUrl}`);
            console.log(`Application base directory: ${projectBaseDir}`);
            console.log(`Is packaged: ${isPackaged}`);

            if (isPackaged && !browserOpened) {
                let command;
                const platform = os.platform();
                if (platform === 'win32') command = `start ${serverUrl}`;
                else if (platform === 'darwin') command = `open ${serverUrl}`;
                else command = `xdg-open ${serverUrl}`;

                if (command) {
                    exec(command, (error) => {
                        if (error) {
                            console.error(`Failed to open browser: ${error.message}`);
                            console.error(`Please manually open your browser and navigate to ${serverUrl}`);
                            return;
                        }
                        console.log(`Attempted to open browser with command: ${command}`);
                        browserOpened = true;
                    });
                } else {
                    console.log(`Unsupported platform for auto browser opening: ${platform}. Please open ${serverUrl} manually.`);
                }
            } else if (!isPackaged) {
                console.log(`Development mode: Please open ${serverUrl} in your browser.`);
            }

            if (isPackaged) {
                console.log("Application started. This window will remain open. Press Ctrl+C to exit.");
                process.stdin.resume();
                process.on('SIGINT', () => { console.log('SIGINT signal received. Shutting down...'); process.exit(0); });
                process.on('SIGTERM', () => { console.log('SIGTERM signal received. Shutting down...'); process.exit(0); });
            }
        });
    } catch (error) {
        console.error("Failed to start the application:", error);
        process.exit(1);
    }
}

// --- Start the Application ---
startApp();
