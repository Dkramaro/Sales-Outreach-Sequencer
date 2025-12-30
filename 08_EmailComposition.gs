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
 * Fetches the content of a Google Doc as HTML using the Drive API.
 * This is used to retrieve a fully-formatted email signature.
 * @returns {string|null} The HTML content of the signature doc, or null if not configured/enabled or an error occurs.
 */
function getSignatureHtml() {
  const userProps = PropertiesService.getUserProperties();
  const signatureEnabled = userProps.getProperty("SIGNATURE_ENABLED") === 'true';
  const docId = userProps.getProperty("SIGNATURE_DOC_ID");
  
  // Check if signature is enabled and configured
  if (!signatureEnabled || !docId) {
    return null; // Signature is disabled or not configured, return null gracefully.
  }

  try {
    // The Drive API v2 provides a direct export link.
    // Ensure "Drive API" advanced service is enabled in the Apps Script editor.
    const url = "https://www.googleapis.com/drive/v2/files/" + docId + "/export?mimeType=text/html";
    
    const params = {
      method: "get",
      headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true // Important to check the response code manually
    };
    
    const response = UrlFetchApp.fetch(url, params);

    if (response.getResponseCode() == 200) {
      return response.getContentText();
    } else {
      // Log the error for debugging but don't break the add-on for the user.
      const errorDetails = `Code: ${response.getResponseCode()}, Response: ${response.getContentText()}`;
      console.error("Error fetching signature doc as HTML. " + errorDetails);
      logAction("Error", "Failed to fetch signature doc HTML. " + errorDetails);
      return null;
    }
  } catch (e) {
    console.error("Exception occurred while fetching signature doc: " + e.toString());
    logAction("Error", "Exception fetching signature doc: " + e.toString());
    return null;
  }
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

