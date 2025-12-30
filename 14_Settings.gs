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
const signatureEnabled = userProps.getProperty("SIGNATURE_ENABLED") === 'true';

if (!signatureDocId) {
    // First-time setup: Show setup button
    signatureSection.addWidget(CardService.newTextParagraph()
        .setText("Create a Google Doc for your email signature. The doc will have 0 margins for perfect email formatting."));
    signatureSection.addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
            .setText("📝 Setup Signature")
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(CardService.newAction()
                .setFunctionName("createSignatureDoc"))));
} else {
    // Signature exists: Show view button and toggle
    try {
        const sigDoc = DriveApp.getFileById(signatureDocId);
        signatureSection.addWidget(CardService.newKeyValue()
            .setTopLabel("Signature Document")
            .setContent(sigDoc.getName())
            .setOpenLink(CardService.newOpenLink()
                .setUrl("https://docs.google.com/document/d/" + signatureDocId + "/edit")));
    } catch(e) {
        signatureSection.addWidget(CardService.newTextParagraph()
            .setText("⚠️ Error accessing signature document. It may have been deleted."));
    }
    
    signatureSection.addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
            .setText("👁️ View Signature")
            .setOpenLink(CardService.newOpenLink()
                .setUrl("https://docs.google.com/document/d/" + signatureDocId + "/edit")))
        .addButton(CardService.newTextButton()
            .setText("🔄 Create New Signature")
            .setOnClickAction(CardService.newAction()
                .setFunctionName("createSignatureDoc"))));
    
    // Toggle to enable/disable signature
    signatureSection.addWidget(CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.SWITCH)
        .setTitle("Include Signature in Emails")
        .setFieldName("signatureEnabled")
        .addItem("Automatically append signature to all emails", "true", signatureEnabled));
}
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
         // Get the Contacts sheet gid to open directly to that tab
         const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);
         const contactsSheetGid = contactsSheet ? contactsSheet.getSheetId() : 0;
         databaseSection.addWidget(CardService.newKeyValue()
              .setTopLabel("Connected Spreadsheet")
              .setContent(spreadsheet.getName())
              .setBottomLabel(`ID: ${spreadsheetId}`)
              .setOpenLink(CardService.newOpenLink()
                   .setUrl("https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/edit#gid=" + contactsSheetGid)));
         
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

/* ======================== SIGNATURE DOCUMENT SETUP ======================== */

/**
* Creates a new Google Doc for email signature with 0 margins and a pre-formatted table.
* IMPORTANT: Uses table format because pulling plain text from Docs breaks formatting.
*/
function createSignatureDoc() {
  try {
    // Create a new Google Doc
    const doc = DocumentApp.create("Email Signature - SoS Outreach");
    const docId = doc.getId();
    
    // Set all margins to 0 (in points)
    const body = doc.getBody();
    body.setMarginTop(0);
    body.setMarginBottom(0);
    body.setMarginLeft(0);
    body.setMarginRight(0);
    
    // Add instructional text at the top
    const instructionText = body.appendParagraph("⚠️ IMPORTANT: Your signature MUST use the table format below.");
    instructionText.setForegroundColor("#cc0000");
    instructionText.editAsText().setBold(true).setFontSize(11);
    
    const tipsPara = body.appendParagraph(
      "Delete this instruction text when done. Replace the placeholder text in the table with your info. " +
      "The first column is for your logo (merged cells). Keep it relatively small!"
    );
    tipsPara.setForegroundColor("#666666");
    tipsPara.editAsText().setFontSize(9);
    
    body.appendParagraph(" "); // Spacing
    
    // Create a 2-column, 4-row table for the signature
    const table = body.appendTable();
    
    // Row 1: Logo cell (will be merged) + Name
    const row1 = table.appendTableRow();
    const logoCell1 = row1.appendTableCell("🖼️ YOUR LOGO HERE");
    const nameCell = row1.appendTableCell("Your Full Name");
    
    // Row 2: Logo cell (to be merged) + Title
    const row2 = table.appendTableRow();
    const logoCell2 = row2.appendTableCell("(Merge cells in first column)");
    const titleCell = row2.appendTableCell("Your Title");
    
    // Row 3: Logo cell (to be merged) + Email
    const row3 = table.appendTableRow();
    const logoCell3 = row3.appendTableCell("Use Insert > Image");
    const emailCell = row3.appendTableCell("E: your.email@company.com");
    
    // Row 4: Logo cell (to be merged) + Phone
    const row4 = table.appendTableRow();
    const logoCell4 = row4.appendTableCell("to add your logo");
    const phoneCell = row4.appendTableCell("M: 555-123-4567");
    
    // Style the table - set borders to transparent/white with 0 width
    table.setBorderWidth(0);
    table.setBorderColor("#FFFFFF");
    
    // Style logo column cells (first column)
    [logoCell1, logoCell2, logoCell3, logoCell4].forEach(cell => {
      cell.setBackgroundColor("#F5F5F5"); // Light grey to show it's placeholder
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      const cellText = cell.editAsText();
      cellText.setFontSize(8);
      cellText.setForegroundColor("#999999");
      cellText.setItalic(true);
      cell.setPaddingLeft(5);
      cell.setPaddingRight(5);
      cell.setPaddingTop(2);
      cell.setPaddingBottom(2);
      cell.setWidth(80); // Narrow column for logo
    });
    
    // Style info column cells (second column)
    // Name cell - larger, bold
    nameCell.editAsText().setBold(true).setFontSize(12).setForegroundColor("#000000");
    nameCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    nameCell.setPaddingLeft(10);
    
    // Title cell - medium
    titleCell.editAsText().setFontSize(10).setForegroundColor("#333333");
    titleCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    titleCell.setPaddingLeft(10);
    
    // Email cell - medium
    emailCell.editAsText().setFontSize(10).setForegroundColor("#0066cc");
    emailCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    emailCell.setPaddingLeft(10);
    
    // Phone cell - medium
    phoneCell.editAsText().setFontSize(10).setForegroundColor("#333333");
    phoneCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    phoneCell.setPaddingLeft(10);
    
    // Add bottom instructions
    body.appendParagraph(" "); // Spacing
    const bottomTips = body.appendParagraph(
      "💡 TO MERGE LOGO CELLS: Select all 4 cells in the first column → Right-click → Merge cells\n" +
      "💡 TO ADD LOGO: Click in merged cell → Insert → Image → Upload your logo\n" +
      "💡 KEEP IT SMALL: Signature should be compact (see example above)\n" +
      "💡 DELETE THESE INSTRUCTIONS when done!"
    );
    bottomTips.setForegroundColor("#0066cc");
    bottomTips.editAsText().setFontSize(9);
    
    // Save the document
    doc.saveAndClose();
    
    // Store the doc ID and enable signature by default
    const userProps = PropertiesService.getUserProperties();
    userProps.setProperty("SIGNATURE_DOC_ID", docId);
    userProps.setProperty("SIGNATURE_ENABLED", "true");
    
    // Log the action
    logAction("Signature Setup", `Created signature document with table format: ${docId}`);
    
    // Return success notification and rebuild settings card
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("✓ Signature doc created with table template! Click 'View Signature' to customize."))
      .setNavigation(CardService.newNavigation()
        .updateCard(buildSettingsCard()))
      .build();
      
  } catch (error) {
    console.error("Error creating signature doc: " + error);
    logAction("Error", "Failed to create signature document: " + error.toString());
    return createNotification("Error creating signature document: " + error.message);
  }
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
const signatureEnabled = e.formInput.signatureEnabled === "true";
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

const userProps = PropertiesService.getUserProperties();
userProps.setProperty("DELAY_DAYS", delayDays.toString());
userProps.setProperty("SENDER_NAME", senderName);
userProps.setProperty("SENDER_COMPANY", senderCompany);
userProps.setProperty("SENDER_TITLE", senderTitle);
userProps.setProperty("CC_EMAILS", ccEmails);
userProps.setProperty("PDF_ATTACHMENT_FILE_ID", pdfFileId || "");
userProps.setProperty("SEQUENCES_FOR_PDF", selectedSequencesForPdf.join(','));
userProps.setProperty("SIGNATURE_ENABLED", signatureEnabled.toString());
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
const sigText = signatureEnabled ? "Signature: ON" : "Signature: OFF";
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

