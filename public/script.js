document.addEventListener('DOMContentLoaded', () => {
    // --- Language Switcher Elements and Config ---
    const langSwitcher = document.getElementById('langSwitcher');
    let currentLang = localStorage.getItem('emailGoLang') || 'zh-CN'; // Default to Simplified Chinese
    let i18nData = {};

    // Quill instance (initialized later)
    let quill;

    // DOM Elements that need i18n
    const DOMElements = {
        // Meta
        pageTitle: document.querySelector('title'),
        // Headings & General Labels
        mainHeading: document.querySelector('.container > h1'),
        langSwitchLabel: document.getElementById('langSwitchLabel'),
        
        // Steps UI (New)
        stepUploadTitle: document.querySelector('#stepUpload .step-title'),
        stepUploadDesc: document.querySelector('#stepUpload .step-description'),
        stepTemplateTitle: document.querySelector('#stepTemplate .step-title'),
        stepTemplateDesc: document.querySelector('#stepTemplate .step-description'),
        stepSendTitle: document.querySelector('#stepSend .step-title'),
        stepSendDesc: document.querySelector('#stepSend .step-description'),
        stepCompleteTitle: document.querySelector('#stepComplete .step-title'),
        stepCompleteDesc: document.querySelector('#stepComplete .step-description'),
        
        // Section 1: Upload
        uploadSectionTitle: document.querySelector('#configSection h2'),
        uploadInstructions: document.querySelector('#configSection p'),
        selectXlsFileLabel: document.querySelector('label[for="xlsFile"]'), // More specific
        previewButton: document.getElementById('previewButton'),
        previewAreaTitle: document.querySelector('#previewArea h3'),
        // XLS Preview Table Title is usually static in HTML, if dynamic, add ID

        // Section 2: Template
        templateSectionTitle: document.querySelector('#templateSection h2'),
        subjectTemplateLabel: document.querySelector('label[for="emailSubjectTemplate"]'),
        emailSubjectTemplate: document.getElementById('emailSubjectTemplate'), // for placeholder
        subjectTemplateHelp: document.querySelector('label[for="emailSubjectTemplate"] + small'),
        bodyTemplateLabel: document.querySelector('label[for="quillEditorContainer"]'),
        attachmentsLabel: document.querySelector('label[for="emailAttachments"]'),
        emailAttachmentsHelp: document.querySelector('label[for="emailAttachments"] + small'),
        currentAttachmentsLabel: document.querySelector('#attachmentList').previousElementSibling, // Assuming it's a label or h4 before #attachmentList

        // Section 3: Send Config
        sendConfigSectionTitle: document.querySelector('#sendSection h2'),
        sendIntervalLabel: document.querySelector('label[for="sendInterval"]'),
        sendIntervalHelp: document.querySelector('label[for="sendInterval"] + small'),
        sendButton: document.getElementById('sendButton'),
        sendingStatusSectionTitle: document.querySelector('#progressArea h3'), // "發送進度:"
        progressLabel: document.querySelector('#progressBarContainer').previousElementSibling, // Assuming this is the label like "Overall Progress:"
        errorLogTitle: document.querySelector('#errorLog h4'),

        // Section 4: Results
        resultsSectionTitle: document.querySelector('#resultsSection h2'),
        resultsInfo: document.querySelector('#resultsSection p'),

        // Section 5: Anti-Spam Advice
        antiSpamAdviceSectionTitle: document.querySelector('#antiSpamAdvice h2'),
        // For list items, it's better to give them IDs or data-i18n-key if their content is purely static
        // For now, I'll assume they might have mixed content (bold tags) and skip direct mapping here,
        // but you should ideally add IDs like antiSpamPoint1, antiSpamPoint2 etc. if they are in json.
        // If the <li> content is complex and includes HTML, it's better to rebuild them in JS.
        // For simplicity, I will attempt to map them if they have IDs.
        antiSpamPoint1: document.getElementById('antiSpamPoint1'),
        antiSpamPoint2: document.getElementById('antiSpamPoint2'),
        antiSpamPoint3: document.getElementById('antiSpamPoint3'),
        antiSpamPoint4: document.getElementById('antiSpamPoint4'),
        antiSpamPoint5: document.getElementById('antiSpamPoint5'),
        antiSpamPoint6: document.getElementById('antiSpamPoint6'),
        antiSpamPoint7: document.getElementById('antiSpamPoint7'),
        antiSpamPoint8: document.getElementById('antiSpamPoint8'),


        // Modal (title and ok button are handled in showModal)
        modalOkButton: document.getElementById('modalOkButton'), // For OK button text

        // Navigation Buttons (New)
        nextToTemplateBtn: document.getElementById('nextToTemplateBtn'),
        prevToUploadBtn: document.getElementById('prevToUploadBtn'),
        nextToSendBtn: document.getElementById('nextToSendBtn'),
        // prevToTemplateBtn is sendButton, no explicit prev button from send section during sending
        // No explicit nav from complete section, only download
    };

    // Step elements for UI manipulation
    const stepUploadEl = document.getElementById('stepUpload');
    const stepTemplateEl = document.getElementById('stepTemplate');
    const stepSendEl = document.getElementById('stepSend');
    const stepCompleteEl = document.getElementById('stepComplete');

    // Section elements
    const configSection = document.getElementById('configSection');
    const templateSection = document.getElementById('templateSection');
    const sendSection = document.getElementById('sendSection');
    const resultsSection = document.getElementById('resultsSection');
    const antiSpamAdviceSection = document.getElementById('antiSpamAdvice');

    const allSections = [configSection, templateSection, sendSection, resultsSection];

    async function fetchAndApplyLanguage(lang) {
        try {
            const response = await fetch(`./locales/${lang}.json`);
            if (!response.ok) {
                console.error(`Could not load ${lang}.json. Status: ${response.status}`);
                if (lang !== 'zh-CN') {
                    console.warn('Fallback to zh-CN');
                    fetchAndApplyLanguage('zh-CN');
                }
                return;
            }
            i18nData = await response.json();
            applyTranslations();
            currentLang = lang;
            localStorage.setItem('emailGoLang', lang);
            if (langSwitcher) langSwitcher.value = lang;
            // Update Quill placeholder explicitly after language load
            if (quill && i18nData.quillPlaceholder) {
                 quill.root.setAttribute('data-placeholder', i18nData.quillPlaceholder);
            }
            // Update "Remove" button for any existing attachments
            document.querySelectorAll('.remove-attachment-btn').forEach(btn => {
                if(i18nData.btnRemove) btn.textContent = i18nData.btnRemove;
            });

        } catch (error) {
            console.error('Error fetching or parsing language file:', error);
            if (lang !== 'zh-CN') {
                console.warn('Fallback to zh-CN on error');
                fetchAndApplyLanguage('zh-CN');
            }
        }
    }

    function applyTranslations() {
        if (!i18nData || Object.keys(i18nData).length === 0) {
            console.warn("i18nData is not loaded or empty. Translations not applied.");
            return;
        }
        for (const key in DOMElements) {
            const element = DOMElements[key];
            if (element && i18nData[key]) {
                if (key === 'pageTitle') {
                    element.textContent = i18nData[key];
                } else if (element.tagName === 'INPUT' && element.type === 'button' || element.tagName === 'BUTTON') {
                    element.textContent = i18nData[key];
                } else if (element.tagName === 'INPUT' && element.placeholder !== undefined) {
                    element.placeholder = i18nData[key];
                } else if (element.tagName === 'SMALL' || element.tagName === 'P' || element.tagName === 'LABEL' || element.tagName === 'H1' || element.tagName === 'H2' || element.tagName === 'H3' || element.tagName === 'H4' || element.tagName === 'LI' || element.tagName === 'SPAN') {
                     // For elements where complex HTML might be in the translation (e.g. <strong>)
                    element.innerHTML = i18nData[key];
                } else {
                    element.textContent = i18nData[key];
                }
            }
        }
    }

    quill = new Quill('#quillEditorContainer', {
        theme: 'snow',
        modules: {
            toolbar: [
                [{ 'header': [1, 2, 3, false] }],
                ['bold', 'italic', 'underline', 'strike'],
                [{ 'list': 'ordered'}, { 'list': 'bullet' }],
                [{ 'script': 'sub'}, { 'script': 'super' }],
                [{ 'indent': '-1'}, { 'indent': '+1' }],
                [{ 'direction': 'rtl' }],
                [{ 'size': ['small', false, 'large', 'huge'] }],
                [{ 'color': [] }, { 'background': [] }],
                [{ 'font': [] }],
                [{ 'align': [] }],
                ['link', 'image', 'video'],
                ['clean']
            ]
        },
        // Placeholder is now set dynamically after language load
    });

    if (langSwitcher) {
        langSwitcher.addEventListener('change', (event) => {
            fetchAndApplyLanguage(event.target.value);
        });
    }
    
    // Initial Load - Must be called after Quill is initialized if Quill needs i18n placeholder
    fetchAndApplyLanguage(currentLang).then(() => {
        resetAllStepsToPending(); // Set initial state after language loaded
        navigateToStep(0, true); // Show first step/section, force activation
    });

    const xlsFileElement = document.getElementById('xlsFile');
    const previewButton = document.getElementById('previewButton'); // Already in DOMElements
    const fileInfoElement = document.getElementById('fileInfo');
    const placeholdersElement = document.getElementById('placeholders');
    const xlsPreviewTableElement = document.getElementById('xlsPreviewTable');
    const emailSubjectTemplateElement = document.getElementById('emailSubjectTemplate'); // Already in DOMElements
    const sendButton = document.getElementById('sendButton'); // Already in DOMElements
    const sendIntervalElement = document.getElementById('sendInterval');
    const emailAttachmentsElement = document.getElementById('emailAttachments');
    const attachmentListElement = document.getElementById('attachmentList');
    const progressBarElement = document.getElementById('progressBar');
    const progressTextElement = document.getElementById('progressText');
    const errorListElement = document.getElementById('errorList');
    const downloadLinkElement = document.getElementById('downloadLink');
    const customModal = document.getElementById('customModal');
    const modalCloseButton = document.querySelector('.modal-close-button');
    const modalTitleElement = document.getElementById('modalTitle');
    const modalMessageElement = document.getElementById('modalMessage');
    const modalOkButton = document.getElementById('modalOkButton'); // Already in DOMElements

    let uploadedFileName = null;
    let originalXlsNameFromServer = null;
    let headers = [];
    let currentEventSource = null;
    let currentAttachmentFiles = [];

    try {
        const savedSubject = localStorage.getItem('emailGoSubjectTemplate');
        const savedBodyHTML = localStorage.getItem('emailGoBodyTemplateHTML');
        if (savedSubject) {
            emailSubjectTemplateElement.value = savedSubject;
        }
        if (savedBodyHTML && quill) {
            quill.clipboard.dangerouslyPasteHTML(0, savedBodyHTML);
        }
    } catch (e) {
        console.warn('Failed to load templates from localStorage:', e);
    }

    function showModal(titleKey, messageKey, messageParams = {}) {
        modalTitleElement.textContent = i18nData[titleKey] || titleKey; // Fallback to key if no translation
        let message = i18nData[messageKey] || messageKey; // Fallback to key
        for (const param in messageParams) {
            message = message.replace(`{${param}}`, messageParams[param]);
        }
        modalMessageElement.textContent = message;
        if (i18nData.modalOkButton) modalOkButton.textContent = i18nData.modalOkButton;
        customModal.style.display = 'block';
    }

    function closeModal() {
        customModal.style.display = 'none';
    }

    modalCloseButton.addEventListener('click', closeModal);
    modalOkButton.addEventListener('click', closeModal);
    window.addEventListener('click', (event) => {
        if (event.target == customModal) {
            closeModal();
        }
    });

    function renderAttachmentList() {
        attachmentListElement.innerHTML = '';
        currentAttachmentFiles.forEach((file, index) => {
            const fileDiv = document.createElement('div');
            fileDiv.classList.add('attachment-item');
            const fileNameSpan = document.createElement('span');
            fileNameSpan.textContent = file.name;
            fileDiv.appendChild(fileNameSpan);
            const removeButton = document.createElement('button');
            removeButton.textContent = i18nData.btnRemove || 'Remove';
            removeButton.classList.add('remove-attachment-btn');
            removeButton.dataset.fileIndex = index;
            removeButton.addEventListener('click', (event) => {
                const indexToRemove = parseInt(event.target.dataset.fileIndex, 10);
                currentAttachmentFiles.splice(indexToRemove, 1);
                renderAttachmentList();
                updateFileInput();
            });
            fileDiv.appendChild(removeButton);
            attachmentListElement.appendChild(fileDiv);
        });
    }

    function updateFileInput() {
        const dataTransfer = new DataTransfer();
        currentAttachmentFiles.forEach(file => dataTransfer.items.add(file));
        emailAttachmentsElement.files = dataTransfer.files;
    }

    emailAttachmentsElement.addEventListener('change', () => {
        currentAttachmentFiles = Array.from(emailAttachmentsElement.files);
        renderAttachmentList();
    });

    previewButton.addEventListener('click', async () => {
        if (!xlsFileElement.files.length) {
            showModal('modalFileNotFoundTitle', 'modalFileNotFoundMessageXlsx');
            return;
        }
        const file = xlsFileElement.files[0];
        if (!file.name.toLowerCase().endsWith('.xlsx')) {
            showModal('modalInvalidFileTypeTitle', 'modalInvalidFileTypeMessageXlsx');
            return;
        }
        navigateToStep(0, true); // Ensure we are on step 0 and it's active

        const formData = new FormData();
        formData.append('xlsfile', file);
        const originalButtonText = previewButton.textContent;
        previewButton.disabled = true;
        previewButton.textContent = i18nData.previewButtonLoading || 'Uploading and Parsing...';
        try {
            const response = await fetch('/upload-and-preview', { method: 'POST', body: formData });
            if (!response.ok) {
                const errorData = await response.json();
                throw new Error(errorData.message || `Server error: ${response.status}`);
            }
            const result = await response.json();
            fileInfoElement.textContent = (i18nData.fileInfo || 'Uploaded: {fileName}. Contains {rowCount} data rows.')
                .replace('{fileName}', file.name)
                .replace('{rowCount}', result.rowCount);
            headers = result.headers;
            placeholdersElement.innerHTML = `<h3>${i18nData.placeholdersAvailable || 'Available Placeholders:'}</h3>` + headers.map(h => `<span>{{${h}}}</span>`).join('');
            xlsPreviewTableElement.innerHTML = '';
            const headerRow = xlsPreviewTableElement.insertRow();
            headers.forEach(headerText => {
                const th = document.createElement('th');
                th.textContent = headerText;
                headerRow.appendChild(th);
            });
            result.previewData.forEach(rowData => {
                const row = xlsPreviewTableElement.insertRow();
                headers.forEach(header => {
                    const cell = row.insertCell();
                    cell.textContent = rowData[header] || '';
                });
            });
            uploadedFileName = result.filePath;
            originalXlsNameFromServer = result.originalXlsName;
            sendButton.disabled = false;
            stepsOrder[0].classList.add('completed'); // Mark upload as completed
            updateStepStatus(stepUploadEl, 'completed'); // Explicitly update the current step's visual state
            showSection(configSection); // Ensure the current section (configSection) remains visible
        } catch (error) {
            console.error('Preview failed:', error);
            showModal('modalPreviewErrorTitle', 'modalPreviewErrorMessage', { errorMessage: error.message });
            fileInfoElement.textContent = '';
            placeholdersElement.innerHTML = '';
            xlsPreviewTableElement.innerHTML = '';
            sendButton.disabled = true;
            navigateToStep(0, true); // Go back to step 0 on error, force activation
        } finally {
            previewButton.disabled = false;
            previewButton.textContent = i18nData.previewButton || originalButtonText; // Restore or use translated
        }
    });

    sendButton.addEventListener('click', async () => {
        if (!uploadedFileName) {
            showModal('modalSendErrorTitle', 'modalSendErrorMessageFile');
            return;
        }
        const subjectTemplate = emailSubjectTemplateElement.value.trim();
        const bodyTemplateHTML = quill.root.innerHTML;
        const sendInterval = parseInt(sendIntervalElement.value, 10);
        if (!subjectTemplate) {
            showModal('modalInputIncompleteTitle', 'modalInputIncompleteSubject');
            return;
        }
        if (!bodyTemplateHTML || bodyTemplateHTML === '<p><br></p>') {
            showModal('modalInputIncompleteTitle', 'modalInputIncompleteBody');
            return;
        }
        if (isNaN(sendInterval) || sendInterval < 0) {
            showModal('modalInputInvalidTitle', 'modalInputInvalidInterval');
            return;
        }

        // Force navigation to Step 3 (Send) to ensure UI is in correct state.
        // This will set step bar correctly (Upload & Template as completed, Send as active) and show sendSection.
        navigateToStep(2, true);

        try {
            localStorage.setItem('emailGoSubjectTemplate', subjectTemplate);
            localStorage.setItem('emailGoBodyTemplateHTML', bodyTemplateHTML);
        } catch (e) {
            console.warn('Failed to save templates to localStorage:', e);
        }

        const originalButtonText = sendButton.textContent;
        sendButton.disabled = true;
        sendButton.textContent = i18nData.sendButtonInitializing || 'Initializing...';
        progressTextElement.textContent = i18nData.statusPending || 'Preparing email task...';
        progressBarElement.style.width = '0%';
        progressBarElement.textContent = '';
        progressBarElement.style.backgroundColor = '#5cb85c';
        errorListElement.innerHTML = '';
        downloadLinkElement.style.display = 'none';

        const formData = new FormData();
        formData.append('filePath', uploadedFileName);
        formData.append('subjectTemplate', subjectTemplate);
        formData.append('bodyTemplate', bodyTemplateHTML);
        formData.append('sendInterval', sendInterval.toString());
        formData.append('headers', JSON.stringify(headers));
        formData.append('originalXlsName', originalXlsNameFromServer);
        if (currentAttachmentFiles.length > 0) {
            for (let i = 0; i < currentAttachmentFiles.length; i++) {
                formData.append('attachments', currentAttachmentFiles[i]);
            }
        }

        try {
            const initResponse = await fetch('/initiate-sending-job', { method: 'POST', body: formData });
            if (!initResponse.ok) {
                const errorData = await initResponse.json();
                throw new Error(errorData.message || 'Initialization failed');
            }
            const initResult = await initResponse.json();
            const jobId = initResult.jobId;
            if (!jobId) {
                 throw new Error('Could not retrieve job ID.');
                 navigateToStep(0, true); // Reset to first step on critical error
            }

            progressTextElement.textContent = (i18nData.statusJobStarted || 'Email task started (ID: {jobId}), connecting to event stream...')
                .replace('{jobId}', jobId);
            sendButton.textContent = i18nData.sendButtonSending || 'Sending...';

            if (currentEventSource) currentEventSource.close();
            currentEventSource = new EventSource(`/send-emails-stream?jobId=${jobId}`);

            currentEventSource.addEventListener('job_started', (event) => {
                // This message is now mainly for console, progressText already updated
                console.log('SSE: Job started', JSON.parse(event.data));
            });

            currentEventSource.addEventListener('data_loaded', (event) => {
                const data = JSON.parse(event.data);
                progressTextElement.textContent = (i18nData.statusDataLoaded || 'Loaded {totalRows} data rows, preparing to process...')
                    .replace('{totalRows}', data.totalRows);
            });

            currentEventSource.addEventListener('progress', (event) => {
                const data = JSON.parse(event.data);
                const percentage = data.totalRows > 0 ? (data.rowIndex / data.totalRows) * 100 : 0;
                progressBarElement.style.width = `${Math.min(percentage, 100)}%`;
                progressBarElement.textContent = percentage > 0 ? `${Math.ceil(Math.min(percentage, 100))}%` : '0%';
                
                let statusText = data.status; // Default to English status from server if not found
                if (data.status === 'success' && i18nData.statusSuccess) statusText = i18nData.statusSuccess;
                else if (data.status === 'send_failed' && i18nData.statusSendFailed) statusText = i18nData.statusSendFailed;
                else if (data.status === 'validation_failed' && i18nData.statusValidationFailed) statusText = i18nData.statusValidationFailed;
                else if (data.status === 'skipped_previously_success' && i18nData.statusSkipped) statusText = i18nData.statusSkipped;

                progressTextElement.textContent = (i18nData.statusProcessing || 'Processing row {rowIndex}: {email} - Status: {status}')
                    .replace('{rowIndex}', data.rowIndex)
                    .replace('{email}', data.email)
                    .replace('{status}', statusText);

                if (data.status === 'send_failed' || data.status === 'validation_failed') {
                    const li = document.createElement('li');
                    li.textContent = `Row ${data.rowIndex}: ${data.email} - ${data.error || 'Unknown error'}`; // Error details remain in original lang from server
                    errorListElement.appendChild(li);
                    progressBarElement.style.backgroundColor = '#d9534f';
                }
            });

            currentEventSource.addEventListener('batch_save', (event) => {
                console.log('SSE: Batch save', JSON.parse(event.data)); // Mostly for console
            });

            currentEventSource.addEventListener('complete', (event) => {
                const data = JSON.parse(event.data);
                progressTextElement.textContent = (i18nData.statusComplete || 'Processing complete! Total {totalRows} rows, {successCount} successful, {failCount} failed.')
                    .replace('{totalRows}', data.totalRows)
                    .replace('{successCount}', data.successCount)
                    .replace('{failCount}', data.failCount);
                progressBarElement.style.width = '100%';
                progressBarElement.textContent = '100%';
                progressBarElement.style.backgroundColor = data.failCount > 0 ? '#d9534f' : '#5cb85c';

                if (data.errors && data.errors.length > 0 && errorListElement.children.length === 0) {
                    data.errors.forEach(err => {
                        const li = document.createElement('li');
                        li.textContent = `Row ${err.rowIndex}: ${err.email} - ${err.error}`; // Error details remain in original lang
                        errorListElement.appendChild(li);
                    });
                }
                if (data.outputFilePath && data.suggestedDownloadName) {
                    downloadLinkElement.href = data.outputFilePath;
                    downloadLinkElement.download = data.suggestedDownloadName;
                    downloadLinkElement.style.display = 'block';
                    downloadLinkElement.textContent = (i18nData.downloadUpdatedXls || 'Download Updated XLS File ({fileName})')
                        .replace('{fileName}', data.suggestedDownloadName);
                }
                sendButton.disabled = false;
                sendButton.textContent = i18nData.sendButton || originalButtonText; // Restore or use translated
                if (currentEventSource) currentEventSource.close();
                currentEventSource = null;
                navigateToStep(3); // Move to complete step
                
                if(data.failCount > 0){
                    updateStepStatus(stepCompleteEl, 'active'); 
                } else {
                    updateStepStatus(stepCompleteEl, 'completed');
                }
            });

            currentEventSource.addEventListener('error', (event) => {
                console.error('SSE Error event:', event);
                let errorMessageFromServer = "Unknown SSE error";
                 if (event.data) {
                    try {
                        const parsedError = JSON.parse(event.data);
                        if(parsedError.message) errorMessageFromServer = parsedError.message;
                    } catch(e) { /* ignore if data is not json */ }
                } else if (event.target && event.target.readyState === EventSource.CLOSED) {
                    errorMessageFromServer = "Connection was closed.";
                }

                progressTextElement.textContent = (i18nData.statusErrorSSE || 'Error: {errorMessage}')
                    .replace('{errorMessage}', errorMessageFromServer); // Show server error message (likely English)
                progressBarElement.style.backgroundColor = '#d9534f';
                showModal('modalJobFailedTitle', 'modalJobFailedMessage', { errorMessage: errorMessageFromServer });
                sendButton.disabled = false;
                sendButton.textContent = i18nData.sendButton || originalButtonText;
                if (currentEventSource) currentEventSource.close();
                currentEventSource = null;
                // On SSE error, the current active step (Send) should reflect an error,
                updateStepStatus(stepSendEl, 'active'); // Keep send step active, maybe add an error class
                stepSendEl.classList.add('error-state'); // Add a generic error state for styling if needed
                updateStepStatus(stepCompleteEl, 'pending');
                showSection(sendSection); // Stay on send section to show error
            });

        } catch (error) {
            console.error('Main send process error:', error);
            progressTextElement.textContent = (i18nData.modalOperationFailedMessage || 'Operation failed: {errorMessage}')
                .replace('{errorMessage}', error.message);
            progressBarElement.style.backgroundColor = '#d9534f';
            showModal('modalOperationFailedTitle', 'modalOperationFailedMessage', { errorMessage: error.message });
            sendButton.disabled = false;
            sendButton.textContent = i18nData.sendButton || originalButtonText;
             if (currentEventSource) currentEventSource.close();
             currentEventSource = null;
             // On general error during send initiation
             updateStepStatus(stepSendEl, 'active'); 
             stepSendEl.classList.add('error-state');
             updateStepStatus(stepCompleteEl, 'pending');
             showSection(sendSection); // Stay on send section
        }
    });

    // --- Step UI Update Function ---
    function updateStepStatus(step, status) { // status can be 'pending', 'active', 'completed'
        step.classList.remove('pending', 'active', 'completed');
        if (status) step.classList.add(status);
    }

    // --- Section Display Function ---
    function showSection(sectionToShow) {
        allSections.forEach(sec => {
            if (sec === sectionToShow) {
                sec.classList.add('active-section');
                // Move and append the antiSpamAdvice section to the currently shown section
                // unless it's the resultsSection, where it might not be desired or fits differently.
                // Or if antiSpamAdviceSection itself is the one to show (though it's not in allSections array)
                if (antiSpamAdviceSection && sectionToShow !== antiSpamAdviceSection) {
                    sectionToShow.appendChild(antiSpamAdviceSection);
                    antiSpamAdviceSection.style.display = 'block'; // Ensure it's visible
                }
            } else {
                sec.classList.remove('active-section');
            }
        });
        // If no specific section is active (e.g. during initial load before a step is chosen)
        // or if the results section is shown, we might want to hide or place anti-spam advice elsewhere.
        // For now, if it's not appended to an active section, it will remain where it is in the HTML
        // or be hidden if its parent becomes hidden. 
        // If configSection is the initial active section, it will be appended to it.
    }

    // --- Navigate Function ---
    let currentStepIndex = 0; // 0: Upload, 1: Template, 2: Send, 3: Complete
    const stepsOrder = [stepUploadEl, stepTemplateEl, stepSendEl, stepCompleteEl];
    const sectionsOrder = [configSection, templateSection, sendSection, resultsSection];

    function navigateToStep(targetStepIndex, forceActivation = false) {
        // Validate targetStepIndex
        if (targetStepIndex < 0 || targetStepIndex >= stepsOrder.length) return;

        const targetStep = stepsOrder[targetStepIndex];
        const currentStep = stepsOrder[currentStepIndex];

        // Allow navigation only if target step is already completed, or is the next logical step, or forced
        if (!forceActivation && !stepsOrder[targetStepIndex].classList.contains('completed') && targetStepIndex > currentStepIndex +1 ){
            if(!(currentStep.classList.contains('completed') && targetStepIndex === currentStepIndex + 1)){
                 console.warn(`Navigation to step ${targetStepIndex} blocked. Current: ${currentStepIndex}, Target not completed or not next.`);
                 return;
            }
        }
        if (!forceActivation && targetStepIndex > currentStepIndex && !currentStep.classList.contains('completed')){
            // If trying to move forward but current step is not completed (e.g. XLS not previewed yet)
            // This logic might need refinement based on exact conditions for allowing next step (e.g. preview success)
            if(currentStepIndex === 0 && !uploadedFileName){ // Specific check for step 0
                showModal('modalPreviewNeededTitle', 'modalPreviewNeededMessage');
                return;
            }
            if(currentStepIndex === 1){ // Specific check for step 1 (template)
                const subjectTemplate = emailSubjectTemplateElement.value.trim();
                const bodyTemplateHTML = quill.root.innerHTML;
                if (!subjectTemplate) {
                    showModal('modalInputIncompleteTitle', 'modalInputIncompleteSubject');
                    return;
                }
                if (!bodyTemplateHTML || bodyTemplateHTML === '<p><br></p>') {
                    showModal('modalInputIncompleteTitle', 'modalInputIncompleteBody');
                    return;
                } 
            }
            // For other steps, completion is usually triggered by an action like 'send'.
        }


        currentStepIndex = targetStepIndex;

        stepsOrder.forEach((step, index) => {
            if (index < currentStepIndex) {
                updateStepStatus(step, 'completed');
            } else if (index === currentStepIndex) {
                updateStepStatus(step, 'active');
            } else {
                updateStepStatus(step, 'pending');
            }
        });
        showSection(sectionsOrder[currentStepIndex]);

        // Special handling for the last step (Complete) which might not have a next button
        // and its state depends on the sending process result.
        if (currentStepIndex === stepsOrder.length - 1) { // If on 'Complete' step
             // The 'Complete' step's status (completed or active with errors) is set by SSE 'complete' event.
             // Here we just ensure it's at least 'active'.
             if(!stepsOrder[currentStepIndex].classList.contains('completed')){
                updateStepStatus(stepsOrder[currentStepIndex], 'active');
             }
        }
    }


    // Initial step states
    function resetAllStepsToPending() {
        updateStepStatus(stepUploadEl, 'pending');
        updateStepStatus(stepTemplateEl, 'pending');
        updateStepStatus(stepSendEl, 'pending');
        updateStepStatus(stepCompleteEl, 'pending');
    }

    // --- Navigation Event Listeners ---
    // Step click listeners
    stepsOrder.forEach((stepElement, index) => {
        stepElement.addEventListener('click', () => {
            // Allow clicking to preview section content
            // This does NOT change the actual currentStepIndex or step completion status
            showSection(sectionsOrder[index]);

            // Optional: If you want to visually highlight the clicked step for preview temporarily
            // without changing its actual status (pending, active, completed), you might add a temporary class.
            // For now, just showing the section is the primary goal.
            
            // The original navigation logic via navigateToStep() is now primarily for
            // prev/next buttons and process-driven progression.
            /* Original logic that changed actual step:
            if (sendButton.disabled && index < 2) { 
                 if (!(index === 2 && currentEventSource)) { 
                    console.warn("Navigation disabled during active sending process.");
                    return;
                 }
            }
            navigateToStep(index);
            */
        });
    });

    // Button listeners
    if (DOMElements.nextToTemplateBtn) {
        DOMElements.nextToTemplateBtn.addEventListener('click', () => {
            if (uploadedFileName) {
                stepsOrder[0].classList.add('completed'); 
                navigateToStep(1);
            } else {
                showModal('modalPreviewNeededTitle', 'modalPreviewNeededMessage');
            }
        });
    }

    if (DOMElements.prevToUploadBtn) {
        DOMElements.prevToUploadBtn.addEventListener('click', () => navigateToStep(0));
    }

    if (DOMElements.nextToSendBtn) {
        DOMElements.nextToSendBtn.addEventListener('click', () => {
            const subjectTemplate = emailSubjectTemplateElement.value.trim();
            const bodyTemplateHTML = quill.root.innerHTML;
            if (!subjectTemplate || bodyTemplateHTML === '<p><br></p>' || !bodyTemplateHTML) { // Check for empty or placeholder-only quill
                showModal('modalInputIncompleteTitle', 'modalInputIncompleteBodyAndSubject'); 
                return;
            }
            stepsOrder[1].classList.add('completed'); 
            // sendButton.click(); // DO NOT trigger send here
            navigateToStep(2); // Just navigate to the send section
        });
    }

}); 