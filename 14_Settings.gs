/**
 * FILE: Settings and Configuration
 * 
 * PURPOSE:
 * Handles user settings UI and persistence including email delays, sender info,
 * CC recipients, PDF attachments, priority filters, and email signature configuration.
 * 
 * KEY FUNCTIONS:
 * - buildSettingsCard() - Settings UI with all configuration options
 * - saveSettings() - Persist user settings
 * 
 * NOTE: Today's Sent Emails reporting has been moved to 18_Analytics.gs
 * 
 * DEPENDENCIES:
 * - 01_Config.gs: CONFIG
 * - 03_Database.gs: logAction
 * - 04_ContactData.gs: getAllContactsData
 * - 06_SequenceData.gs: getAvailableSequences
 * - 17_Utilities.gs: extractIdFromUrl
 * 
 * @version 2.4
 */

/* ======================== SETTINGS UI ======================== */

/**
* Builds the complete Settings card, including the button to trigger the batch labeling process.
*/
function buildSettingsCard() {
 const card = CardService.newCardBuilder();

 card.setHeader(CardService.newCardHeader()
     .setTitle("Settings")
     .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/settings_black_48dp.png"));

 const userProps = PropertiesService.getUserProperties();

 // --- Email Sending Behavior Section ---
 const sendingBehaviorSection = CardService.newCardSection()
     .setHeader("Email Sending Behavior");
 
 const isAutoSendStep1Enabled = userProps.getProperty("AUTO_SEND_STEP_1_ENABLED") === 'true';
 sendingBehaviorSection.addWidget(CardService.newSelectionInput()
     .setType(CardService.SelectionInputType.SWITCH)
     .setTitle("Enable Auto-Sending for Step 1")
     .setFieldName("autoSendStep1Emails")
     .addItem("When enabled, ONLY Step 1 emails are sent immediately. All other steps (2-5) remain as drafts.", "true", isAutoSendStep1Enabled));
 card.addSection(sendingBehaviorSection);

 // --- Sequence Settings Section ---
 const sequenceSection = CardService.newCardSection()
     .setHeader("Sequence Settings");
 const delayDays = userProps.getProperty("DELAY_DAYS") || CONFIG.DEFAULT_DELAY_DAYS.toString();
 sequenceSection.addWidget(CardService.newTextInput()
     .setFieldName("delayDays")
     .setTitle("Days Between Sequence Emails")
     .setHint(`Default: ${CONFIG.DEFAULT_DELAY_DAYS}`)
     .setValue(delayDays));
 card.addSection(sequenceSection);

 // --- Sender Information Section ---
 const senderSection = CardService.newCardSection()
     .setHeader("Sender Info (for Template Placeholders)");
 const senderName = userProps.getProperty("SENDER_NAME") || (Session.getActiveUser() ? Session.getActiveUser().getEmail().split('@')[0] : "");
 const senderCompany = userProps.getProperty("SENDER_COMPANY") || "";
 const senderTitle = userProps.getProperty("SENDER_TITLE") || "";
 senderSection.addWidget(CardService.newTextInput()
     .setFieldName("senderName").setTitle("Your Name (for {{senderName}})")
     .setValue(senderName).setHint("e.g., John Doe"));
 senderSection.addWidget(CardService.newTextInput()
     .setFieldName("senderCompany").setTitle("Your Company (for {{senderCompany}})")
     .setValue(senderCompany).setHint("e.g., ACME Corp"));
 senderSection.addWidget(CardService.newTextInput()
     .setFieldName("senderTitle").setTitle("Your Title (for {{senderTitle}})")
     .setValue(senderTitle).setHint("e.g., Sales Manager"));
 card.addSection(senderSection);

 // --- Email CC Section ---
 const ccSection = CardService.newCardSection()
     .setHeader("Email CC Settings");
 const ccEmails = userProps.getProperty("CC_EMAILS") || "";
 ccSection.addWidget(CardService.newTextInput()
     .setFieldName("ccEmails")
     .setTitle("Emails to CC")
     .setHint("Comma-separated email addresses")
     .setValue(ccEmails));
 card.addSection(ccSection);
 
 // --- Signature From Google Doc Section ---
 const signatureSection = CardService.newCardSection()
     .setHeader("Email Signature from Google Doc");
 const signatureDocId = userProps.getProperty("SIGNATURE_DOC_ID") || "";
 signatureSection.addWidget(CardService.newTextParagraph()
     .setText("Automatically append a signature from a Google Doc. The Doc must be shared so 'anyone with the link can view'."));
 signatureSection.addWidget(CardService.newTextInput()
     .setFieldName("signatureDocUrlOrId")
     .setTitle("Google Doc URL or File ID for Signature")
     .setHint("e.g., https://docs.google.com/document/d/1a2b3c.../edit")
     .setValue(signatureDocId));
 card.addSection(signatureSection);

 // --- Step 2 PDF Attachment Section ---
 const attachmentSection = CardService.newCardSection()
    .setHeader("Step 2 PDF Attachment");

 const pdfFileId = userProps.getProperty("PDF_ATTACHMENT_FILE_ID") || "";
 const sequencesForPdfString = userProps.getProperty("SEQUENCES_FOR_PDF") || "";
 const sequencesForPdf = sequencesForPdfString ? sequencesForPdfString.split(',') : [];

 attachmentSection.addWidget(CardService.newTextParagraph()
    .setText("To automatically attach a PDF to Step 2 emails, paste its Google Drive share link below."));

 attachmentSection.addWidget(CardService.newTextInput()
    .setFieldName("pdfAttachmentUrlOrId")
    .setTitle("Google Drive Link or File ID for PDF")
    .setHint("e.g., 1mSlH2WJUr0jSZKyfySk44oG8nphZYs3g")
    .setValue(pdfFileId));

 if (pdfFileId) {
    try {
        const fileName = DriveApp.getFileById(pdfFileId).getName();
        attachmentSection.addWidget(CardService.newKeyValue()
            .setTopLabel("✅ Currently Saved File")
            .setContent(fileName)
            .setBottomLabel("To remove, clear the field above and save."));
    } catch(e) {
        attachmentSection.addWidget(CardService.newKeyValue()
            .setTopLabel("⚠️ Currently Saved File")
            .setContent("Error: Could not access file.")
            .setBottomLabel("ID: " + pdfFileId));
    }
 }

 attachmentSection.addWidget(CardService.newTextParagraph()
     .setText("Select which sequences should get this attachment:"));
 
 const availableSequences = getAvailableSequences();
 if (availableSequences.length > 0) {
     for (const sequenceName of availableSequences) {
         const fieldName = "sequence_for_pdf_" + sequenceName.replace(/[^a-zA-Z0-9]/g, "_");
         attachmentSection.addWidget(CardService.newSelectionInput()
             .setType(CardService.SelectionInputType.CHECK_BOX)
             .setFieldName(fieldName)
             .addItem(sequenceName, sequenceName, sequencesForPdf.includes(sequenceName))
         );
     }
 } else {
     attachmentSection.addWidget(CardService.newTextParagraph().setText("No sequences found. Create a sequence first."));
 }
 card.addSection(attachmentSection);

 // --- Priority Filter Section ---
 const priorityFilterSection = CardService.newCardSection()
    .setHeader("Contact Priority Filters (Select all that apply)");
 const priorityFiltersString = userProps.getProperty("PRIORITY_FILTERS");
 let initiallySelectedPriorities = [];
 if (priorityFiltersString === null || priorityFiltersString === undefined || priorityFiltersString === "") {
    initiallySelectedPriorities = ["High", "Medium", "Low"];
 } else {
    initiallySelectedPriorities = priorityFiltersString.split(',').map(p => p.trim());
 }
 priorityFilterSection.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setTitle("High Priority 🔥")
    .setFieldName("priorityFilter_High")
    .addItem("Include High Priority", "High", initiallySelectedPriorities.includes("High")));
 priorityFilterSection.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setTitle("Medium Priority 🟠")
    .setFieldName("priorityFilter_Medium")
    .addItem("Include Medium Priority", "Medium", initiallySelectedPriorities.includes("Medium")));
 priorityFilterSection.addWidget(CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.CHECK_BOX)
    .setTitle("Low Priority ⚪")
    .setFieldName("priorityFilter_Low")
    .addItem("Include Low Priority", "Low", initiallySelectedPriorities.includes("Low")));
 card.addSection(priorityFilterSection);

 // --- Save Button Section ---
 const buttonSection = CardService.newCardSection();
 buttonSection.addWidget(CardService.newButtonSet()
     .addButton(CardService.newTextButton()
         .setText("Save Settings")
         .setOnClickAction(CardService.newAction()
             .setFunctionName("saveSettings"))));
 card.addSection(buttonSection);

 // --- Manual Actions Section ---
 const actionsSection = CardService.newCardSection()
     .setHeader("Manual Actions");
 actionsSection.addWidget(CardService.newTextParagraph()
     .setText("Run background tasks on your entire database. This may take some time to complete."));
 actionsSection.addWidget(CardService.newButtonSet()
     .addButton(CardService.newTextButton()
         .setText("Find & Label Sent Emails")
         .setOnClickAction(CardService.newAction()
             .setFunctionName("startLabelingProcess"))));
 card.addSection(actionsSection);

// --- Database Section ---
const databaseSection = CardService.newCardSection()
    .setHeader("Database Connection");
const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
 if (spreadsheetId) {
     try {
         const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
         databaseSection.addWidget(CardService.newKeyValue()
              .setTopLabel("Connected Spreadsheet")
              .setContent(spreadsheet.getName())
              .setBottomLabel(`ID: ${spreadsheetId}`)
              .setOpenLink(CardService.newOpenLink()
                   .setUrl("https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/edit")));
         
         // Disconnect button
         databaseSection.addWidget(CardService.newButtonSet()
             .addButton(CardService.newTextButton()
                 .setText("🔌 Disconnect Database")
                 .setOnClickAction(CardService.newAction()
                     .setFunctionName("showDisconnectConfirmation"))));
     } catch (error) {
         databaseSection.addWidget(CardService.newTextParagraph()
              .setText(`⚠️ Error accessing connected spreadsheet (ID: ${spreadsheetId}).`));
         databaseSection.addWidget(CardService.newButtonSet()
             .addButton(CardService.newTextButton()
                 .setText("🔌 Disconnect & Reconnect")
                 .setOnClickAction(CardService.newAction()
                     .setFunctionName("disconnectDatabase"))));
     }
 } else {
     databaseSection.addWidget(CardService.newTextParagraph()
         .setText("No database currently connected."));
     databaseSection.addWidget(CardService.newButtonSet()
         .addButton(CardService.newTextButton()
             .setText("🔗 Connect Database")
             .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
             .setOnClickAction(CardService.newAction()
                 .setFunctionName("showConnectDatabaseForm"))));
 }
card.addSection(databaseSection);

 // --- Navigation Section ---
 const navSection = CardService.newCardSection();
 navSection.addWidget(CardService.newTextButton()
     .setText("Back to Main Menu")
     .setOnClickAction(CardService.newAction()
         .setFunctionName("buildAddOn")));
 card.addSection(navSection);

 return card.build();
}

/* ======================== SAVE SETTINGS ======================== */

/**
* Saves settings from the Settings card.
*/
function saveSettings(e) {
 const delayDaysInput = e.formInput.delayDays;
 const senderName = e.formInput.senderName || "";
 const senderCompany = e.formInput.senderCompany || "";
 const senderTitle = e.formInput.senderTitle || "";
 const ccEmails = e.formInput.ccEmails || "";
 const pdfUrlOrId = e.formInput.pdfAttachmentUrlOrId || "";
 const signatureUrlOrId = e.formInput.signatureDocUrlOrId || "";
 const autoSendStep1Enabled = e.formInput.autoSendStep1Emails === "true";

 const selectedSequencesForPdf = [];
 for (const key in e.formInput) {
    if (key.startsWith("sequence_for_pdf_")) {
        selectedSequencesForPdf.push(e.formInput[key]);
    }
 }

 const collectedPriorities = [];
 if (e.formInput.priorityFilter_High) { collectedPriorities.push("High"); }
 if (e.formInput.priorityFilter_Medium) { collectedPriorities.push("Medium"); }
 if (e.formInput.priorityFilter_Low) { collectedPriorities.push("Low"); }
 
 const delayDays = parseInt(delayDaysInput);
 if (isNaN(delayDays) || delayDays < 0) {
   return createNotification("Please enter a valid, non-negative number of days.");
 }

 const pdfFileId = extractIdFromUrl(pdfUrlOrId);
 if (pdfUrlOrId && !pdfFileId) {
     return createNotification("The provided PDF attachment link or ID is invalid.");
 }
 if (pdfFileId) {
    try {
        DriveApp.getFileById(pdfFileId).getName();
    } catch(err) {
        return createNotification("Error accessing the PDF file. Please check permissions.");
    }
 }

 let signatureDocId = "";
 if (signatureUrlOrId) {
     signatureDocId = extractIdFromUrl(signatureUrlOrId);
     if (!signatureDocId) {
         return createNotification("The provided Signature Google Doc link or ID is invalid.");
     }
     try {
         DriveApp.getFileById(signatureDocId).getName(); 
     } catch (err) {
         return createNotification("Error accessing the Signature Google Doc. Please check permissions.");
     }
 }

 const userProps = PropertiesService.getUserProperties();
 userProps.setProperty("DELAY_DAYS", delayDays.toString());
 userProps.setProperty("SENDER_NAME", senderName);
 userProps.setProperty("SENDER_COMPANY", senderCompany);
 userProps.setProperty("SENDER_TITLE", senderTitle);
 userProps.setProperty("CC_EMAILS", ccEmails);
 userProps.setProperty("PDF_ATTACHMENT_FILE_ID", pdfFileId || "");
 userProps.setProperty("SEQUENCES_FOR_PDF", selectedSequencesForPdf.join(','));
 userProps.setProperty("SIGNATURE_DOC_ID", signatureDocId);
 userProps.setProperty("AUTO_SEND_STEP_1_ENABLED", autoSendStep1Enabled.toString());

 if (collectedPriorities.length > 0) {
  userProps.setProperty("PRIORITY_FILTERS", collectedPriorities.join(','));
 } else {
  userProps.setProperty("PRIORITY_FILTERS", ""); 
 }

 const priorityText = collectedPriorities.length > 0 ? collectedPriorities.join(',') : 'ALL';
 const ccText = ccEmails ? ccEmails : 'None';
 const pdfText = pdfFileId ? `PDF ID: ${pdfFileId}` : 'None';
 const pdfSeqText = selectedSequencesForPdf.length > 0 ? selectedSequencesForPdf.join(', ') : 'None';
 const sigText = signatureDocId ? `Sig Doc ID: ${signatureDocId}` : 'None';
 const autoSendText = autoSendStep1Enabled ? "Auto-Send Step 1: ON" : "Auto-Send Step 1: OFF";
 
 logAction("Update Settings", `Saved: Delay=${delayDays}, Sender=${senderName}, Priorities=${priorityText}, CC=${ccText}, Attachment=${pdfText}, PDF Sequences=[${pdfSeqText}], ${sigText}, ${autoSendText}`);
 
 return CardService.newActionResponseBuilder()
   .setNotification(CardService.newNotification().setText("Settings saved successfully!"))
   .setNavigation(CardService.newNavigation()
     .updateCard(buildAddOn()))
   .build();
}

/* ======================== TODAY'S SENT EMAILS ======================== */
// NOTE: getTodaysSentCompanies() and buildTodaysSentEmailsCard() have been moved to 18_Analytics.gs

