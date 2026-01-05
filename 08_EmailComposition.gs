/**
 * FILE: Email Composition and Template Processing
 * 
 * PURPOSE:
 * Handles individual email composition, template placeholder replacement,
 * signature integration, and dark mode styling for email content.
 * 
 * KEY FUNCTIONS:
 * - composeEmailWithTemplate() - Compose/send email using template
 * - replacePlaceholders() - Replace template variables with contact data
 * - getSignatureHtml() - Fetch signature from Google Doc
 * - getDarkModeStyles() - Generate email dark mode CSS
 * - cleanAndWrapSignature() - Clean and wrap signature HTML
 * 
 * DEPENDENCIES:
 * - 01_Config.gs: CONFIG, CONTACT_COLS
 * - 04_ContactData.gs: getContactByEmail
 * - 06_SequenceData.gs: getSequenceTemplateForStep, getSequenceStepCount
 * - 09_EmailProcessing.gs: findSentStep1Message (for reply logic)
 * 
 * @version 2.3
 */

/* ======================== EMAIL COMPOSITION ======================== */

/**
 * Shows a preview of the email before sending/drafting (with collapsible sections)
 */
function previewEmailBeforeSend(e) {
  const email = e.parameters.email;
  const contact = getContactByEmail(email);
  
  if (!contact) {
    return createNotification("Contact not found.");
  }

  const template = getTemplateForStep(contact.currentStep, contact.sequence);
  if (!template) {
    return createNotification("No template found for Step " + contact.currentStep);
  }

  // Render the email with placeholders replaced
  const renderedSubject = replacePlaceholders(template.subject, contact);
  const renderedBody = replacePlaceholders(template.body, contact);
  const signatureHtml = getSignatureHtml() || "";

  const card = CardService.newCardBuilder();
  
  card.setHeader(CardService.newCardHeader()
      .setTitle("📧 Email Preview")
      .setSubtitle(`Step ${contact.currentStep} → ${contact.firstName} ${contact.lastName}`));

  // Email details section (collapsible)
  const detailsSection = CardService.newCardSection()
      .setHeader("Email Details")
      .setCollapsible(true)
      .setNumUncollapsibleWidgets(0); // Collapsed by default
  
  detailsSection.addWidget(CardService.newKeyValue()
      .setTopLabel("To")
      .setContent(contact.email));
  detailsSection.addWidget(CardService.newKeyValue()
      .setTopLabel("Subject")
      .setContent(renderedSubject)
      .setMultiline(true));
  card.addSection(detailsSection);

  // Email body section (collapsible to save space)
  const bodySection = CardService.newCardSection()
      .setHeader("Email Body (Click to expand)")
      .setCollapsible(true)
      .setNumUncollapsibleWidgets(0); // Collapsed by default
  
  // Strip HTML tags from signature for preview display
  const signaturePreview = signatureHtml ? "\n\n[Signature will be appended]" : "";
  const fullBody = renderedBody + signaturePreview;
  // Truncate if too long (keep under card limits)
  const truncatedBody = fullBody.length > 1500 ? fullBody.substring(0, 1500) + "...\n\n[Body truncated for preview]" : fullBody;
  bodySection.addWidget(CardService.newTextParagraph().setText(truncatedBody));
  card.addSection(bodySection);

  // Check if auto-send is enabled for Step 1
  const userProps = PropertiesService.getUserProperties();
  const autoSendStep1Enabled = userProps.getProperty("AUTO_SEND_STEP_1_ENABLED") === 'true';
  const willSendImmediately = (contact.currentStep === 1 && autoSendStep1Enabled);

  // Action buttons section (always visible)
  const actionSection = CardService.newCardSection().setHeader("Ready to Send?");
  
  if (willSendImmediately) {
    actionSection.addWidget(CardService.newTextParagraph()
        .setText("⚠️ <b>Auto-send enabled:</b> This email will be sent immediately."));
  } else {
    actionSection.addWidget(CardService.newTextParagraph()
        .setText("This will be created as a draft for you to review and send manually."));
  }

  const buttonSet = CardService.newButtonSet();
  
  // Main action button
  buttonSet.addButton(CardService.newTextButton()
      .setText(willSendImmediately ? "✓ Send Now" : "✓ Create Draft")
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setOnClickAction(CardService.newAction()
          .setFunctionName("composeEmailWithTemplate")
          .setParameters({ email: contact.email })));

  // Test email button
  buttonSet.addButton(CardService.newTextButton()
      .setText("📬 Test Email")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("sendTestEmailToSelf")
          .setParameters({ email: contact.email })));

  actionSection.addWidget(buttonSet);
  card.addSection(actionSection);

  // Navigation
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
      .setText("← Back to Contact")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("viewContactCard")
          .setParameters({ email: contact.email })));
  card.addSection(navSection);

  return card.build();
}

/**
 * Sends a test email to the current user
 */
function sendTestEmailToSelf(e) {
  const email = e.parameters.email;
  const contact = getContactByEmail(email);
  
  if (!contact) {
    return createNotification("Contact not found.");
  }
  
  const userEmail = Session.getActiveUser().getEmail();
  const template = getTemplateForStep(contact.currentStep, contact.sequence);
  
  if (!template) {
    return createNotification("No template found for this step.");
  }
  
  const renderedSubject = "[TEST] " + replacePlaceholders(template.subject, contact);
  const renderedBody = replacePlaceholders(template.body, contact);
  const signatureHtml = getSignatureHtml();
  
  // Build email body with signature (matching production email format)
  const darkModeStyles = getDarkModeStyles();
  let finalEmailBodyHtml = darkModeStyles + renderedBody.replace(/\n/g, "<br>");
  if (signatureHtml) {
    finalEmailBodyHtml += cleanAndWrapSignature(signatureHtml);
  }
  
  try {
    GmailApp.sendEmail(userEmail, renderedSubject, "", {
      htmlBody: finalEmailBodyHtml
    });
    
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("✓ Test email sent to " + userEmail))
      .build();
  } catch (error) {
    console.error("Error sending test email: " + error);
    return createNotification("Error sending test email: " + error.message);
  }
}

/**
 * Composes an email to a contact using the default template for their current step.
 * Creates a draft or sends immediately based on settings for Step 1, then updates sheet data.
 * CORRECTED: Handles Step 1 as new email, Steps 2-5 as REPLIES to Step 1 thread.
 * ADDED: Checks for replies before composing follow-ups (Steps > 1).
 * NEW: Integrates signature from Google Doc.
 * NEW: Checks AUTO_SEND_STEP_1_ENABLED setting.
 */
function composeEmailWithTemplate(e) {
  const email = e.parameters.email;
  const contact = getContactByEmail(email); 

  if (!contact) {
    logAction("Error", `Compose Template Error: Contact not found ${email}`);
    return createNotification("Contact not found: " + email);
  }
  if (contact.status !== 'Active') {
    return createNotification(`Cannot compose email: Contact status is "${contact.status}". Activate the contact first.`);
  }

  const currentStep = contact.currentStep;
  const template = getSequenceTemplateForStep(contact.sequence, currentStep);

  if (!template) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(`Missing template for Step ${currentStep} in sequence "${contact.sequence}"`))
      .build();
  }

  let subjectForStep1Storage = ""; 
  let emailBody = "";
  let originalStep1Message = null;

  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) { return createNotification("Database not configured."); }
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);
  if (!contactsSheet) { return createNotification("Contacts sheet not found."); }
  const currentUserEmail = Session.getActiveUser() ? Session.getActiveUser().getEmail().toLowerCase() : null;
  // NEW: Check auto-send setting for Step 1
  const userProps = PropertiesService.getUserProperties();
  const autoSendStep1Enabled = userProps.getProperty("AUTO_SEND_STEP_1_ENABLED") === 'true';
  let actionVerbPast = "Draft created"; // Default action verb

  let draftComposeActionBuilder; // Define here to be accessible later

  try {
    if (currentStep > 1 && currentUserEmail) { /* ... keep your reply check logic here ... */ }
    
    emailBody = replacePlaceholders(template.body, contact);

    // --- SIGNATURE INTEGRATION ---
    const signatureHtml = getSignatureHtml();
    const darkModeStyles = getDarkModeStyles();
    let finalEmailBodyHtml = darkModeStyles + emailBody.replace(/\n/g, "<br>");

    if (signatureHtml) {
        finalEmailBodyHtml += cleanAndWrapSignature(signatureHtml);
    }
    // --- END SIGNATURE INTEGRATION ---

    // --- MODIFIED: Conditional Logic Block for Step 1 ---
    if (currentStep === 1) {
        const actualSubjectForStep1Email = replacePlaceholders(template.subject, contact); 
        subjectForStep1Storage = actualSubjectForStep1Email; 
        const emailOptions = { 
          htmlBody: finalEmailBodyHtml, 
          cc: userProps.getProperty("CC_EMAILS") || "",
          name: userProps.getProperty("SENDER_NAME") || (Session.getActiveUser() ? Session.getActiveUser().getEmail().split('@')[0] : "")
        };

        if (autoSendStep1Enabled) {
            // AUTO-SEND LOGIC
            GmailApp.sendEmail(contact.email, actualSubjectForStep1Email, '', emailOptions);
            actionVerbPast = "Sent";
        } else {
            // DRAFT LOGIC
            const draft = GmailApp.createDraft(contact.email, actualSubjectForStep1Email, '', emailOptions);
            draftComposeActionBuilder = CardService.newComposeActionResponseBuilder().setGmailDraft(draft);
        }
    } else { // Steps 2+ ALWAYS DRAFT
        if (!originalStep1Message) { 
            originalStep1Message = findSentStep1Message(contact, contactsSheet); 
            if (!originalStep1Message) {
                return createNotification(`Could not find the original Step 1 email for ${contact.email}.`);
            }
        }
        contact.originalSubject = originalStep1Message.getSubject(); 
        const existingCcEmails = userProps.getProperty("CC_EMAILS") || "";
        const ccList = [contact.email];
        if (existingCcEmails.trim()) ccList.push(existingCcEmails.trim());
        const finalCcString = ccList.join(',');
        const replyOptions = { htmlBody: finalEmailBodyHtml, cc: finalCcString };
        
        // Attachment logic remains for drafts
        const pdfFileId = userProps.getProperty("PDF_ATTACHMENT_FILE_ID");
        const sequencesForPdfString = userProps.getProperty("SEQUENCES_FOR_PDF") || "";
        const sequencesForPdf = sequencesForPdfString ? sequencesForPdfString.split(',') : [];
        if (currentStep === 2 && pdfFileId && sequencesForPdf.includes(contact.sequence)) {
             try {
                replyOptions.attachments = [DriveApp.getFileById(pdfFileId).getBlob()];
            } catch(e) { console.error("Could not attach PDF for single compose: " + e); }
        }

        const draft = originalStep1Message.createDraftReply("", replyOptions);
        draftComposeActionBuilder = CardService.newComposeActionResponseBuilder().setGmailDraft(draft);
    }
    // --- END CONDITIONAL ---

  } catch (emailError) {
    console.error(`Error processing email for ${email}, Step ${currentStep}: ${emailError}\n${emailError.stack}`);
    return createNotification(`Error processing email: ${emailError.message}`);
  }

  // This part runs for both send and draft
  try {
    const now = new Date();
    const delayDays = parseInt(userProps.getProperty("DELAY_DAYS")) || CONFIG.DEFAULT_DELAY_DAYS;
    const nextDate = new Date(now.getTime() + (delayDays * 24 * 60 * 60 * 1000));
    let nextStepVal = contact.currentStep + 1;
    let newStatus = contact.status; 
    
    // Use per-sequence step count instead of global CONFIG.SEQUENCE_STEPS
    const sequenceStepCount = getSequenceStepCount(contact.sequence);
    if (nextStepVal > sequenceStepCount) {
      nextStepVal = sequenceStepCount; 
      newStatus = "Completed"; 
    }
    // Update sheet data...
    const lastCol = Math.max(contactsSheet.getLastColumn(), CONTACT_COLS.STEP1_SENT_MESSAGE_ID + 1); 
    const rowRange = contactsSheet.getRange(contact.rowIndex, 1, 1, lastCol);
    const rowData = rowRange.getValues()[0];
    rowData[CONTACT_COLS.LAST_EMAIL_DATE] = now;
    rowData[CONTACT_COLS.NEXT_STEP_DATE] = (newStatus === "Completed") ? "" : nextDate;
    rowData[CONTACT_COLS.CURRENT_STEP] = nextStepVal;
    rowData[CONTACT_COLS.STATUS] = newStatus;
    rowData[CONTACT_COLS.PERSONAL_CALLED] = "No"; 
    rowData[CONTACT_COLS.WORK_CALLED] = "No";   
    if (currentStep === 1 && subjectForStep1Storage) {
        rowData[CONTACT_COLS.STEP1_SUBJECT] = subjectForStep1Storage; 
    }
    rowRange.setValues([rowData]);
    SpreadsheetApp.flush();
    
    // Invalidate cache since contact step/status changed
    removeContactFromCache(contact.email);
    
    logAction("Email Processed", `${actionVerbPast} for Step ${currentStep} for ${contact.email}. Moved to Step ${nextStepVal}.`);
  } catch (sheetError) {
    logAction("Error", `Sheet Update Failed for ${email} after email action: ${sheetError}`);
    // Handle error differently for send vs draft
    if (draftComposeActionBuilder) {
        return CardService.newActionResponseBuilder()
          .setNotification(CardService.newNotification().setText("Email draft created, BUT failed to update contact record."))
          .setComposeAction(draftComposeActionBuilder.build().getComposeAction()) 
          .build();
    }
    return createNotification("Email sent, BUT failed to update the contact's record. Please update manually.");
  }
  
  // Final response builder
  const responseBuilder = CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(`${actionVerbPast} for Step ${currentStep}.`));

  // Only set compose action if a draft was actually created
  if (draftComposeActionBuilder) {
      responseBuilder.setComposeAction(draftComposeActionBuilder.build().getComposeAction());
  } else {
      // If auto-sent, we rebuild the contact card to show the updated status
      responseBuilder.setNavigation(CardService.newNavigation().updateCard(viewContactCard(e)));
  }

  return responseBuilder.build();
}

/* ======================== TEMPLATE PROCESSING ======================== */

/**
 * Replaces placeholders in template text with contact and sender data.
 */
function replacePlaceholders(text, contact, senderInfoOverride) {
  if (!text || typeof text !== 'string') {
    return "";
  }
  if (!contact) {
    console.warn("replacePlaceholders called without contact data.");
    return text; // Return original text if no contact
   }

  // --- Get Sender Info ---
  let senderInfo;
   if (senderInfoOverride) {
       senderInfo = senderInfoOverride; // Use provided override if available
   } else {
       const userProps = PropertiesService.getUserProperties();
       // Provide sensible defaults if properties aren't set
       senderInfo = {
           name: userProps.getProperty("SENDER_NAME") || (Session.getActiveUser() ? Session.getActiveUser().getEmail().split('@')[0] : "Sender"),
           company: userProps.getProperty("SENDER_COMPANY") || "",
           title: userProps.getProperty("SENDER_TITLE") || ""
       };
   }

  // --- Extract Custom Fields ---
  // Use dedicated Industry column (prioritize over notes)
  let industry = contact.industry || "";
  
  // Extract department from Notes (still using notes as no dedicated column exists)
  let department = "";
  if (contact.notes && typeof contact.notes === 'string') {
      // Look for "key: value" pairs on separate lines or simple "key:value"
      const noteLines = contact.notes.split('\n');
      noteLines.forEach(line => {
          // Fallback: still check notes for industry if dedicated field is empty
          if (!industry) {
              const industryMatch = line.match(/^industry:\s*(.*)/i);
              if (industryMatch && industryMatch[1]) {
                  industry = industryMatch[1].trim();
              }
          }
          const departmentMatch = line.match(/^department:\s*(.*)/i);
           if (departmentMatch && departmentMatch[1]) {
              department = departmentMatch[1].trim();
          }
          // Add more custom field extractions here if needed
      });
  }

  // --- Perform Replacements ---
  // Use function replacement for case-insensitivity and efficiency
  const replacements = {
      '{{firstname}}': contact.firstName || "",
      '{{lastname}}': contact.lastName || "",
      '{{email}}': contact.email || "",
      '{{company}}': contact.company || "",
      '{{title}}': contact.title || "",
      '{{sendername}}': senderInfo.name || "",
      '{{sendercompany}}': senderInfo.company || "",
      '{{sendertitle}}': senderInfo.title || "",
      '{{industry}}': industry || "", // Replace with empty string if not found
      '{{department}}': department || "", // Replace with empty string if not found
      '{{originalsubject}}': contact.originalSubject || "" // Added for reply context
  };

  // Regex to find {{placeholder}} case-insensitively
  const placeholderRegex = /\{\{([\w\s]+?)\}\}/gi;

  let replacedText = text.replace(placeholderRegex, (match, placeholderName) => {
      const lowerPlaceholderName = placeholderName.toLowerCase().trim();
      const key = `{{${lowerPlaceholderName}}}`; // Construct the key used in replacements map
      return replacements.hasOwnProperty(key) ? replacements[key] : match; // Return replacement or original match if not found
  });


  return replacedText;
}

/* ======================== SIGNATURE INTEGRATION ======================== */

/**
 * Fetches the processed signature HTML from cache.
 * Returns cached signature instantly - no auto-processing.
 * User must click Refresh button after editing their signature.
 * @returns {string|null} The processed signature HTML, or null if not configured/enabled.
 */
function getSignatureHtml() {
  const userProps = PropertiesService.getUserProperties();
  const signatureEnabled = userProps.getProperty("SIGNATURE_ENABLED") === 'true';
  const docId = userProps.getProperty("SIGNATURE_DOC_ID");
  
  // Check if signature is enabled and configured
  if (!signatureEnabled || !docId) {
    return null;
  }

  // Return cached processed signature (stored in UserProperties for persistence)
  const cachedSignature = userProps.getProperty("SIGNATURE_PROCESSED_CACHE");
  if (cachedSignature) {
    return cachedSignature;
  }
  
  // No cache - signature was enabled but not processed yet
  console.warn("Signature enabled but not processed. User should click Refresh.");
  return null;
}

/**
 * Processes signature by converting embedded base64 images to Drive-hosted URLs.
 * Creates a folder "SoS Signature Images" to store the images.
 * Caches the processed signature for fast retrieval.
 * @returns {object} {success: boolean, html: string|null, error: string|null, sizeKB: number}
 */
function processSignatureImages() {
  const userProps = PropertiesService.getUserProperties();
  const docId = userProps.getProperty("SIGNATURE_DOC_ID");
  const SIGNATURE_FOLDER_NAME = "SoS Signature Images";
  
  if (!docId) {
    return { success: false, html: null, error: "No signature document configured", sizeKB: 0 };
  }

  try {
    // Fetch fresh signature HTML from Google Doc
    const url = "https://www.googleapis.com/drive/v2/files/" + docId + "/export?mimeType=text/html";
    const params = {
      method: "get",
      headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, params);
    
    if (response.getResponseCode() !== 200) {
      return { success: false, html: null, error: "Failed to fetch signature document", sizeKB: 0 };
    }
    
    let signatureHtml = response.getContentText();
    const originalSizeKB = Utilities.newBlob(signatureHtml).getBytes().length / 1024;
    console.log(`Original signature size: ${originalSizeKB.toFixed(2)} KB`);
    
    // Get or create the signature images folder
    let folder = getOrCreateSignatureFolder(SIGNATURE_FOLDER_NAME);
    
    // Clear old images from the folder
    const oldFiles = folder.getFiles();
    while (oldFiles.hasNext()) {
      oldFiles.next().setTrashed(true);
    }
    
    // Find all base64 encoded images in the HTML (capture the full tag to extract attributes)
    const imgRegex = /<img([^>]*)src="data:image\/([^;]+);base64,([^"]+)"([^>]*)>/gi;
    let processedHtml = signatureHtml;
    let match;
    let imageCount = 0;
    const replacements = [];
    
    // Collect all image matches first (regex exec with replace can be tricky)
    const tempHtml = signatureHtml;
    while ((match = imgRegex.exec(tempHtml)) !== null) {
      const fullMatch = match[0];
      const beforeSrc = match[1] || '';
      const imageType = match[2].toLowerCase();
      const base64Data = match[3];
      const afterSrc = match[4] || '';
      
      // Extract width and height from attributes
      const widthMatch = fullMatch.match(/width="(\d+)"/i) || fullMatch.match(/width:(\d+)px/i);
      const heightMatch = fullMatch.match(/height="(\d+)"/i) || fullMatch.match(/height:(\d+)px/i);
      const styleMatch = fullMatch.match(/style="([^"]*)"/i);
      
      replacements.push({
        fullMatch: fullMatch,
        imageType: imageType,
        base64Data: base64Data,
        width: widthMatch ? widthMatch[1] : null,
        height: heightMatch ? heightMatch[1] : null,
        style: styleMatch ? styleMatch[1] : null
      });
    }
    
    // Process each image
    for (const img of replacements) {
      try {
        // Decode base64 to blob
        const imageBytes = Utilities.base64Decode(img.base64Data);
        const imageSizeKB = imageBytes.length / 1024;
        console.log(`Processing image: ${imageSizeKB.toFixed(2)} KB, width=${img.width}, height=${img.height}`);
        
        // Create blob and upload to Drive
        const mimeType = img.imageType === 'png' ? 'image/png' : 'image/jpeg';
        const blob = Utilities.newBlob(imageBytes, mimeType, `signature_image_${imageCount + 1}.${img.imageType}`);
        const file = folder.createFile(blob);
        
        // Set sharing to anyone with link can view
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        // Get the direct image URL
        const fileId = file.getId();
        const imageUrl = `https://lh3.googleusercontent.com/d/${fileId}`;
        
        // Build new img tag preserving original dimensions
        let styleAttr = '';
        if (img.width && img.height) {
          styleAttr = `width:${img.width}px; height:${img.height}px;`;
        } else if (img.width) {
          styleAttr = `width:${img.width}px;`;
        } else if (img.height) {
          styleAttr = `height:${img.height}px;`;
        } else if (img.style) {
          // Preserve original style but remove any background-related properties
          styleAttr = img.style.replace(/background[^;]*;?/gi, '');
        }
        
        const newImgTag = `<img src="${imageUrl}"${styleAttr ? ` style="${styleAttr}"` : ''}>`;
        processedHtml = processedHtml.replace(img.fullMatch, newImgTag);
        
        imageCount++;
        console.log(`Uploaded image ${imageCount} to Drive: ${fileId}`);
      } catch (imgError) {
        console.error("Error processing image: " + imgError);
        // Keep original image tag if upload fails
      }
    }
    
    // Strip Google redirect wrappers from hyperlinks
    // Google Docs exports links as: https://www.google.com/url?q=REAL_URL&sa=D&source=editors&ust=...&usg=...
    // We extract just the real URL from the q= parameter
    processedHtml = processedHtml.replace(
      /https?:\/\/www\.google\.com\/url\?q=([^&"]+)(?:&[^"]*)?/gi,
      function(match, encodedUrl) {
        try {
          // Decode the URL (Google encodes special chars like %3A for :)
          return decodeURIComponent(encodedUrl);
        } catch (e) {
          // If decoding fails, return the encoded URL as-is
          return encodedUrl;
        }
      }
    );
    
    // Strip Google Docs HTML bloat
    processedHtml = processedHtml
      .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
      .replace(/ class="[^"]*"/gi, '')
      .replace(/<span[^>]*>\s*<\/span>/gi, '')
      .replace(/<p[^>]*>\s*<\/p>/gi, '')
      .replace(/\s+/g, ' ')
      .trim();
    
    // Calculate final size
    const finalSizeKB = Utilities.newBlob(processedHtml).getBytes().length / 1024;
    console.log(`Processed signature size: ${finalSizeKB.toFixed(2)} KB (was ${originalSizeKB.toFixed(2)} KB)`);
    
    // Check if processed signature fits in UserProperties (9KB limit)
    if (finalSizeKB > 9) {
      logAction("Error", `Processed signature (${finalSizeKB.toFixed(0)} KB) exceeds storage limit of 9KB`);
      return { success: false, html: null, error: `Signature still too large (${finalSizeKB.toFixed(0)} KB). Max is 9KB.`, sizeKB: finalSizeKB };
    }
    
    // Store processed signature in UserProperties (persistent, no expiry)
    userProps.setProperty("SIGNATURE_PROCESSED_CACHE", processedHtml);
    
    if (imageCount > 0) {
      logAction("Signature Processing", `Converted ${imageCount} image(s) to Drive URLs. Size: ${originalSizeKB.toFixed(0)} KB → ${finalSizeKB.toFixed(0)} KB`);
    }
    
    return { success: true, html: processedHtml, error: null, sizeKB: finalSizeKB };
    
  } catch (e) {
    console.error("Exception in processSignatureImages: " + e.toString());
    logAction("Error", "Exception processing signature: " + e.toString());
    return { success: false, html: null, error: e.toString(), sizeKB: 0 };
  }
}

/**
 * Gets or creates the folder for signature images.
 */
function getOrCreateSignatureFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

/**
 * Generates a powerful CSS style block to force correct dark mode rendering in email clients.
 * Includes more aggressive selectors and meta tags for compatibility.
 * @returns {string} The HTML <style> and <meta> block.
 */
function getDarkModeStyles() {
  // This block is designed to be as compatible as possible with clients like Gmail on iOS.
  const styles = `
    <meta name="color-scheme" content="light dark">
    <meta name="supported-color-schemes" content="light dark">
    <style>
      /* Tells email clients that this email supports light and dark schemes */
      :root {
        color-scheme: light dark;
        supported-color-schemes: light dark;
      }

      /* Dark Mode Media Query */
      @media (prefers-color-scheme: dark) {
        /*
         * --- HYPER-AGGRESSIVE BACKGROUND FIX ---
         * The ".signature-container *" selector targets EVERY SINGLE ELEMENT
         * inside your signature (spans, divs, tables, cells, etc.) and forces
         * its background to be transparent. This is much stronger.
        */
        .signature-container * {
          background-color: transparent !important;
          background: transparent !important;
        }

        /*
         * --- TEXT AND LINK COLOR FIX ---
         * This ensures all text and links inside the signature become a readable
         * light grey color in dark mode.
        */
        .signature-container span,
        .signature-container p,
        .signature-container a,
        .signature-container td,
        .signature-container div,
        .signature-container strong {
          color: #E1E1E1 !important;
        }
      }
    </style>
  `;
  return styles;
}

/**
 * Takes the raw signature HTML, cleans it of problematic background styles,
 * wraps it in the dark-mode-ready container, and returns the final HTML.
 * @param {string} signatureHtml The raw HTML content from the Google Doc.
 * @returns {string} The cleaned and wrapped HTML for the signature.
 */
function cleanAndWrapSignature(signatureHtml) {
  if (!signatureHtml) {
    return "";
  }

  // Use regular expressions to find and remove bgcolor attributes and background-color styles.
  // This prevents the email client from finding any color to invert.
  let cleanedHtml = signatureHtml
    .replace(/bgcolor="[^"]*"/gi, '') // Removes bgcolor="#FFFFFF"
    .replace(/background-color:[^;"]*/gi, '') // Removes style="... background-color: #ffffff; ..."
    // NEW LINE: Aggressively removes any CSS border properties from inline styles.
    .replace(/border:[^;"]*;/gi, 'border: none;'); // Finds "border: 1px solid #ccc;" and replaces it.

  // Wrap the now-cleaned HTML in our targetable container div.
  return '<br><br><div class="signature-container">' + cleanedHtml + '</div>';
}

