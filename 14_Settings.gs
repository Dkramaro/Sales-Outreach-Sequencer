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
 * - autoSaveSignatureToggle() - Auto-save when signature toggle changes
 * - autoSaveAutoSendToggle() - Auto-save when auto-send toggle changes
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
 * @version 2.7 - Consolidated UI, auto-save toggles, sticky save footer
 */

/* ======================== SETTINGS UI ======================== */

/**
* Builds the complete Settings card with consolidated sections for better UX.
* Sections: Identity, Email Behavior, Attachments & Filters, System
* Fixed Footer: Sticky Save button always visible at bottom
*/
function buildSettingsCard() {
  const card = CardService.newCardBuilder();
  const userProps = PropertiesService.getUserProperties();

  card.setHeader(CardService.newCardHeader()
      .setTitle("Settings")
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/settings_black_48dp.png"));

  // ═══════════════════════════════════════════════════════════════════
  // SECTION 1: YOUR EMAIL IDENTITY
  // ═══════════════════════════════════════════════════════════════════
  const identitySection = CardService.newCardSection()
      .setHeader("📧 Your Email Identity");

  // Sender info fields
  const senderName = userProps.getProperty("SENDER_NAME") || 
      (Session.getActiveUser() ? Session.getActiveUser().getEmail().split('@')[0] : "");
  const senderCompany = userProps.getProperty("SENDER_COMPANY") || "";
  const senderTitle = userProps.getProperty("SENDER_TITLE") || "";

  identitySection.addWidget(CardService.newTextInput()
      .setFieldName("senderName")
      .setTitle("Your Name")
      .setHint("Used for {{senderName}} in templates")
      .setValue(senderName));
  identitySection.addWidget(CardService.newTextInput()
      .setFieldName("senderCompany")
      .setTitle("Your Company")
      .setHint("Used for {{senderCompany}} in templates")
      .setValue(senderCompany));
  identitySection.addWidget(CardService.newTextInput()
      .setFieldName("senderTitle")
      .setTitle("Your Title")
      .setHint("Used for {{senderTitle}} in templates")
      .setValue(senderTitle));

  // Divider before signature
  identitySection.addWidget(CardService.newDivider());

  // Signature section (conditional)
  const signatureDocId = userProps.getProperty("SIGNATURE_DOC_ID") || "";
  const signatureEnabled = userProps.getProperty("SIGNATURE_ENABLED") === 'true';

  if (!signatureDocId) {
      // No signature set up yet
      identitySection.addWidget(CardService.newDecoratedText()
          .setText("No email signature configured")
          .setBottomLabel("Create a Google Doc with your signature"));
      identitySection.addWidget(CardService.newButtonSet()
          .addButton(CardService.newTextButton()
              .setText("📝 Setup Signature")
              .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("createSignatureDoc"))));
  } else {
      // Signature exists
      try {
          const sigDoc = DriveApp.getFileById(signatureDocId);
          identitySection.addWidget(CardService.newDecoratedText()
              .setTopLabel("Email Signature")
              .setText(sigDoc.getName())
              .setWrapText(true)
              .setOpenLink(CardService.newOpenLink()
                  .setUrl("https://docs.google.com/document/d/" + signatureDocId + "/edit")));
      } catch(e) {
          identitySection.addWidget(CardService.newDecoratedText()
              .setText("⚠️ Signature document not accessible")
              .setBottomLabel("It may have been deleted"));
      }
      
      identitySection.addWidget(CardService.newButtonSet()
          .addButton(CardService.newTextButton()
              .setText("👁️ View")
              .setOpenLink(CardService.newOpenLink()
                  .setUrl("https://docs.google.com/document/d/" + signatureDocId + "/edit")))
          .addButton(CardService.newTextButton()
              .setText("🔄 New")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("createSignatureDoc"))));
      
      identitySection.addWidget(CardService.newSelectionInput()
          .setType(CardService.SelectionInputType.SWITCH)
          .setTitle("Append Signature to Emails")
          .setFieldName("signatureEnabled")
          .addItem("Automatically add signature to all outgoing emails", "true", signatureEnabled)
          .setOnChangeAction(CardService.newAction()
              .setFunctionName("autoSaveSignatureToggle")));
  }
  card.addSection(identitySection);

  // ═══════════════════════════════════════════════════════════════════
  // SECTION 2: EMAIL BEHAVIOR
  // ═══════════════════════════════════════════════════════════════════
  const behaviorSection = CardService.newCardSection()
      .setHeader("⚙️ Email Behavior");

  // Auto-send toggle (auto-saves on change)
  const isAutoSendStep1Enabled = userProps.getProperty("AUTO_SEND_STEP_1_ENABLED") === 'true';
  behaviorSection.addWidget(CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.SWITCH)
      .setTitle("Auto-Send Step 1 Emails")
      .setFieldName("autoSendStep1Emails")
      .addItem("Send Step 1 immediately; Steps 2-5 remain as drafts", "true", isAutoSendStep1Enabled)
      .setOnChangeAction(CardService.newAction()
          .setFunctionName("autoSaveAutoSendToggle")));

  // Delay days
  const delayDays = userProps.getProperty("DELAY_DAYS") || CONFIG.DEFAULT_DELAY_DAYS.toString();
  behaviorSection.addWidget(CardService.newTextInput()
      .setFieldName("delayDays")
      .setTitle("Days Between Sequence Steps")
      .setHint("Default: " + CONFIG.DEFAULT_DELAY_DAYS)
      .setValue(delayDays));

  // CC emails
  const ccEmails = userProps.getProperty("CC_EMAILS") || "";
  behaviorSection.addWidget(CardService.newTextInput()
      .setFieldName("ccEmails")
      .setTitle("CC Recipients")
      .setHint("Comma-separated email addresses")
      .setValue(ccEmails));

  card.addSection(behaviorSection);

  // ═══════════════════════════════════════════════════════════════════
  // SECTION 3: ATTACHMENTS & FILTERS
  // ═══════════════════════════════════════════════════════════════════
  const filtersSection = CardService.newCardSection()
      .setHeader("📎 Attachments & Filters");

  // PDF Attachment
  const pdfFileId = userProps.getProperty("PDF_ATTACHMENT_FILE_ID") || "";
  const sequencesForPdfString = userProps.getProperty("SEQUENCES_FOR_PDF") || "";
  const sequencesForPdf = sequencesForPdfString ? sequencesForPdfString.split(',') : [];

  filtersSection.addWidget(CardService.newTextInput()
      .setFieldName("pdfAttachmentUrlOrId")
      .setTitle("Step 2 PDF Attachment")
      .setHint("Paste Google Drive link or file ID")
      .setValue(pdfFileId));

  // Show current file status if PDF is configured
  if (pdfFileId) {
      try {
          const fileName = DriveApp.getFileById(pdfFileId).getName();
          filtersSection.addWidget(CardService.newDecoratedText()
              .setTopLabel("✅ Attached File")
              .setText(fileName)
              .setBottomLabel("Clear field above to remove"));
      } catch(e) {
          filtersSection.addWidget(CardService.newDecoratedText()
              .setTopLabel("⚠️ File Error")
              .setText("Could not access file")
              .setBottomLabel("ID: " + pdfFileId));
      }
  }

  // Sequence selection for PDF - SINGLE widget with multiple items
  const availableSequences = getAvailableSequences();
  if (availableSequences.length > 0) {
      const sequenceSelector = CardService.newSelectionInput()
          .setType(CardService.SelectionInputType.CHECK_BOX)
          .setTitle("Attach PDF to Sequences")
          .setFieldName("sequencesForPdf");
      
      for (const seq of availableSequences) {
          sequenceSelector.addItem(seq, seq, sequencesForPdf.includes(seq));
      }
      filtersSection.addWidget(sequenceSelector);
  } else {
      filtersSection.addWidget(CardService.newDecoratedText()
          .setText("No sequences available")
          .setBottomLabel("Create a sequence first to attach PDFs"));
  }

  // Divider before priority filters
  filtersSection.addWidget(CardService.newDivider());

  // Priority filters - SINGLE widget with 3 items
  const priorityFiltersString = userProps.getProperty("PRIORITY_FILTERS");
  let initiallySelectedPriorities = [];
  if (priorityFiltersString === null || priorityFiltersString === undefined || priorityFiltersString === "") {
      initiallySelectedPriorities = ["High", "Medium", "Low"];
  } else {
      initiallySelectedPriorities = priorityFiltersString.split(',').map(p => p.trim());
  }

  filtersSection.addWidget(CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setTitle("Show Contacts by Priority")
      .setFieldName("priorityFilters")
      .addItem("🔥 High Priority", "High", initiallySelectedPriorities.includes("High"))
      .addItem("🟠 Medium Priority", "Medium", initiallySelectedPriorities.includes("Medium"))
      .addItem("⚪ Low Priority", "Low", initiallySelectedPriorities.includes("Low")));

  card.addSection(filtersSection);

  // ═══════════════════════════════════════════════════════════════════
  // SECTION 4: SYSTEM & ACTIONS
  // ═══════════════════════════════════════════════════════════════════
  const systemSection = CardService.newCardSection()
      .setHeader("🔧 System");

  // Database connection status
  const spreadsheetId = userProps.getProperty("SPREADSHEET_ID");
  if (spreadsheetId) {
      try {
          const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
          const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);
          const contactsSheetGid = contactsSheet ? contactsSheet.getSheetId() : 0;
          
          systemSection.addWidget(CardService.newDecoratedText()
              .setTopLabel("✅ Database Connected")
              .setText(spreadsheet.getName())
              .setWrapText(true)
              .setOpenLink(CardService.newOpenLink()
                  .setUrl("https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/edit#gid=" + contactsSheetGid)));
          
          systemSection.addWidget(CardService.newButtonSet()
              .addButton(CardService.newTextButton()
                  .setText("🔌 Disconnect")
                  .setOnClickAction(CardService.newAction()
                      .setFunctionName("showDisconnectConfirmation"))));
      } catch (error) {
          systemSection.addWidget(CardService.newDecoratedText()
              .setTopLabel("⚠️ Database Error")
              .setText("Cannot access spreadsheet")
              .setBottomLabel("ID: " + spreadsheetId));
          systemSection.addWidget(CardService.newButtonSet()
              .addButton(CardService.newTextButton()
                  .setText("🔌 Disconnect & Reconnect")
                  .setOnClickAction(CardService.newAction()
                      .setFunctionName("disconnectDatabase"))));
      }
  } else {
      systemSection.addWidget(CardService.newDecoratedText()
          .setText("No database connected")
          .setBottomLabel("Connect a Google Sheet to get started"));
      systemSection.addWidget(CardService.newButtonSet()
          .addButton(CardService.newTextButton()
              .setText("🔗 Connect Database")
              .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("showConnectDatabaseForm"))));
  }

  // Divider before manual actions
  systemSection.addWidget(CardService.newDivider());

  // Manual labeling action
  systemSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("🏷️ Find & Label Sent Emails")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("startLabelingProcess"))));

  // Back navigation
  systemSection.addWidget(CardService.newDivider());
  systemSection.addWidget(CardService.newTextButton()
      .setText("← Back to Main Menu")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("buildAddOn")));

  card.addSection(systemSection);

  // ═══════════════════════════════════════════════════════════════════
  // FIXED FOOTER: Save button (always visible, sticky at bottom)
  // ═══════════════════════════════════════════════════════════════════
  const fixedFooter = CardService.newFixedFooter()
      .setPrimaryButton(CardService.newTextButton()
          .setText("💾 Save")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction()
              .setFunctionName("saveSettings")));
  card.setFixedFooter(fixedFooter);

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
* Auto-saves the Signature toggle when changed.
* Provides instant feedback without requiring manual save.
*/
function autoSaveSignatureToggle(e) {
  const formInput = e.formInput || {};
  const signatureEnabled = formInput.signatureEnabled === "true";
  
  const userProps = PropertiesService.getUserProperties();
  userProps.setProperty("SIGNATURE_ENABLED", signatureEnabled.toString());
  
  const statusText = signatureEnabled ? "ON" : "OFF";
  logAction("Toggle Setting", `Signature: ${statusText}`);
  
  return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
          .setText(`✓ Signature ${statusText}`))
      .build();
}

/**
* Auto-saves the Auto-Send Step 1 toggle when changed.
* Provides instant feedback without requiring manual save.
*/
function autoSaveAutoSendToggle(e) {
  const formInput = e.formInput || {};
  const autoSendEnabled = formInput.autoSendStep1Emails === "true";
  
  const userProps = PropertiesService.getUserProperties();
  userProps.setProperty("AUTO_SEND_STEP_1_ENABLED", autoSendEnabled.toString());
  
  const statusText = autoSendEnabled ? "ON" : "OFF";
  logAction("Toggle Setting", `Auto-Send Step 1: ${statusText}`);
  
  return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
          .setText(`✓ Auto-Send Step 1 ${statusText}`))
      .build();
}

/**
* Saves settings from the Settings card.
* Handles consolidated field names from the redesigned UI.
*/
function saveSettings(e) {
  const formInput = e.formInput || {};
  
  // Extract form values with defaults
  const delayDaysInput = formInput.delayDays;
  const senderName = formInput.senderName || "";
  const senderCompany = formInput.senderCompany || "";
  const senderTitle = formInput.senderTitle || "";
  const ccEmails = formInput.ccEmails || "";
  const pdfUrlOrId = formInput.pdfAttachmentUrlOrId || "";
  const signatureEnabled = formInput.signatureEnabled === "true";
  const autoSendStep1Enabled = formInput.autoSendStep1Emails === "true";

  // Handle sequences for PDF - now a single field that returns array or string
  let selectedSequencesForPdf = [];
  if (formInput.sequencesForPdf) {
      // Can be array (multiple selected) or string (single selected)
      if (Array.isArray(formInput.sequencesForPdf)) {
          selectedSequencesForPdf = formInput.sequencesForPdf;
      } else {
          selectedSequencesForPdf = [formInput.sequencesForPdf];
      }
  }

  // Handle priority filters - now a single field that returns array or string
  let collectedPriorities = [];
  if (formInput.priorityFilters) {
      // Can be array (multiple selected) or string (single selected)
      if (Array.isArray(formInput.priorityFilters)) {
          collectedPriorities = formInput.priorityFilters;
      } else {
          collectedPriorities = [formInput.priorityFilters];
      }
  }
 
  // Validate delay days
  const delayDays = parseInt(delayDaysInput);
  if (isNaN(delayDays) || delayDays < 0) {
      return createNotification("Please enter a valid, non-negative number of days.");
  }

  // Validate PDF file if provided
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

  // Persist all settings
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

  // Log the action with summary
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

