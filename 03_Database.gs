/**
 * FILE: Database Setup and Management
 * 
 * PURPOSE:
 * Handles all database (Google Sheets) setup, initialization, schema management,
 * and connection management. Creates required sheets with proper headers and validation.
 * 
 * KEY FUNCTIONS:
 * - showConnectDatabaseForm() - UI for connecting to database
 * - connectToDatabase() - Connect to existing spreadsheet
 * - createNewDatabase() - Create new spreadsheet
 * - setupDatabase() - Initialize database structure
 * - setupContactsSheet() - Configure Contacts sheet
 * - setupTemplatesSheet() - Configure Templates sheet
 * - setupLogsSheet() - Configure Logs sheet
 * - logAction() - Write action logs
 * 
 * DEPENDENCIES:
 * - 01_Config.gs: CONFIG, CONTACT_COLS, INDUSTRY_OPTIONS
 * - 06_SequenceData.gs: getAvailableSequences, createSequenceSheet
 * 
 * @version 2.3
 */

/* ======================== DATABASE CONNECTION UI ======================== */

/**
 * Shows the form to connect to a database
 */
function showConnectDatabaseForm() {
  const card = CardService.newCardBuilder();

  // Add header
  card.setHeader(CardService.newCardHeader()
      .setTitle("Connect to Database")
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/storage_black_48dp.png"));

  const formSection = CardService.newCardSection();

  // NEW: Prominent "Create New Database" as primary action
  formSection.addWidget(CardService.newTextParagraph()
      .setText("👋 <b>First time here?</b> Let us create your database automatically."));

  formSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("🚀 Create My Database (Recommended)")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction()
              .setFunctionName("createNewDatabase"))));

  formSection.addWidget(CardService.newTextParagraph()
      .setText("\n<i>Or connect to an existing database:</i>"));

  formSection.addWidget(CardService.newTextInput()
      .setFieldName("spreadsheetIdOrUrl")
      .setTitle("Spreadsheet ID or Full URL")
      .setHint("Paste the full URL or just the ID"));

  formSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("Connect to Existing")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("connectToDatabase"))));

  card.addSection(formSection);

  // --- Recent Databases Section ---
  const recentDatabases = getRecentDatabases();
  if (recentDatabases.length > 0) {
    const recentSection = CardService.newCardSection()
        .setHeader("📋 Recent Databases")
        .setCollapsible(true)
        .setNumUncollapsibleWidgets(1);
    
    recentSection.addWidget(CardService.newTextParagraph()
        .setText("Quick reconnect to a previously used database:"));
    
    for (const db of recentDatabases) {
      recentSection.addWidget(CardService.newDecoratedText()
          .setTopLabel(db.name)
          .setText(db.id.substring(0, 20) + "...")
          .setBottomLabel("Last used: " + db.lastUsed)
          .setOnClickAction(CardService.newAction()
              .setFunctionName("quickConnectToDatabase")
              .setParameters({ spreadsheetId: db.id })));
    }
    
    card.addSection(recentSection);
  }

  // Navigation section
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
      .setText("Back")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("buildAddOn")));

  card.addSection(navSection);

  return card.build();
}

/* ======================== DATABASE DISCONNECT ======================== */

/**
 * Shows confirmation before disconnecting database
 */
function showDisconnectConfirmation() {
  const card = CardService.newCardBuilder();
  
  card.setHeader(CardService.newCardHeader()
      .setTitle("⚠️ Disconnect Database?")
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/warning_amber_48dp.png"));
  
  const section = CardService.newCardSection();
  
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  let dbName = "Unknown";
  try {
    if (spreadsheetId) {
      dbName = SpreadsheetApp.openById(spreadsheetId).getName();
    }
  } catch (e) {
    dbName = "Error accessing database";
  }
  
  section.addWidget(CardService.newTextParagraph()
      .setText(`<b>Current Database:</b> ${dbName}\n\n` +
               `This will disconnect the add-on from your database.\n\n` +
               `<b>Your data will NOT be deleted</b> - you can reconnect anytime using the same spreadsheet ID.`));
  
  section.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("🔌 Disconnect")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setBackgroundColor("#ea4335")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("disconnectDatabase")))
      .addButton(CardService.newTextButton()
          .setText("Cancel")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("buildSettingsCard"))));
  
  card.addSection(section);
  
  return card.build();
}

/**
 * Disconnects from the current database and saves it to history
 */
function disconnectDatabase() {
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  
  // Save to history before disconnecting
  if (spreadsheetId) {
    saveToRecentDatabases(spreadsheetId);
  }
  
  // Clear connection
  PropertiesService.getUserProperties().deleteProperty("SPREADSHEET_ID");
  PropertiesService.getUserProperties().deleteProperty("SKIP_WIZARD");
  PropertiesService.getUserProperties().deleteProperty("DEMO_CONTACTS_ADDED");
  
  logAction("Database Disconnect", "Disconnected from database: " + spreadsheetId);
  
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText("Database disconnected. You can reconnect anytime."))
    .setNavigation(CardService.newNavigation()
      .popToRoot()
      .updateCard(buildAddOn()))
    .build();
}

/* ======================== DATABASE HISTORY ======================== */

/**
 * Gets the list of recently used databases
 * @returns {Array<{id: string, name: string, lastUsed: string}>}
 */
function getRecentDatabases() {
  const historyJson = PropertiesService.getUserProperties().getProperty("DATABASE_HISTORY");
  if (!historyJson) return [];
  
  try {
    const history = JSON.parse(historyJson);
    // Filter out current database and verify access
    const currentId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
    return history.filter(db => db.id !== currentId);
  } catch (e) {
    return [];
  }
}

/**
 * Saves a database to the recent history
 * @param {string} spreadsheetId The spreadsheet ID to save
 */
function saveToRecentDatabases(spreadsheetId) {
  if (!spreadsheetId) return;
  
  let dbName = "Unknown Database";
  try {
    dbName = SpreadsheetApp.openById(spreadsheetId).getName();
  } catch (e) {
    // Can't access, but still save for reference
  }
  
  const historyJson = PropertiesService.getUserProperties().getProperty("DATABASE_HISTORY");
  let history = [];
  try {
    history = historyJson ? JSON.parse(historyJson) : [];
  } catch (e) {
    history = [];
  }
  
  // Remove if already exists
  history = history.filter(db => db.id !== spreadsheetId);
  
  // Add to front
  history.unshift({
    id: spreadsheetId,
    name: dbName,
    lastUsed: new Date().toLocaleDateString()
  });
  
  // Keep only last 5
  history = history.slice(0, 5);
  
  PropertiesService.getUserProperties().setProperty("DATABASE_HISTORY", JSON.stringify(history));
}

/**
 * Quick connect to a database from history
 */
function quickConnectToDatabase(e) {
  const spreadsheetId = e.parameters.spreadsheetId;
  
  if (!spreadsheetId) {
    return createNotification("Invalid database ID.");
  }
  
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const spreadsheetName = spreadsheet.getName();
    
    // Save current to history before switching
    const currentId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
    if (currentId && currentId !== spreadsheetId) {
      saveToRecentDatabases(currentId);
    }
    
    // Set new connection
    PropertiesService.getUserProperties().setProperty("SPREADSHEET_ID", spreadsheetId);
    
    logAction("Database Connection", "Quick-connected to: " + spreadsheetName);
    
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("Connected to " + spreadsheetName + "!"))
      .setNavigation(CardService.newNavigation()
        .popToRoot()
        .updateCard(buildAddOn()))
      .build();
    
  } catch (error) {
    return createNotification("Error connecting: " + error.message + ". The database may have been deleted or you've lost access.");
  }
}

/* ======================== DATABASE CONNECTION ======================== */

/**
 * Connects to an existing database
 * Accepts either a spreadsheet ID or a full Google Sheets URL
 */
function connectToDatabase(e) {
  const input = e.formInput.spreadsheetIdOrUrl || e.formInput.spreadsheetId || "";

  if (!input || !input.trim()) {
    return createNotification("Please enter a spreadsheet ID or URL.");
  }

  // Extract the ID from URL or use as-is if it's already an ID
  const spreadsheetId = extractSpreadsheetIdFromUrl(input.trim());
  
  if (!spreadsheetId) {
    return createNotification("Could not extract a valid spreadsheet ID. Please check the URL or ID.");
  }

  try {
    // Verify spreadsheet and access
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const spreadsheetName = spreadsheet.getName(); // Get name for logging

    // Save current to history before switching
    const currentId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
    if (currentId && currentId !== spreadsheetId) {
      saveToRecentDatabases(currentId);
    }

    // Check for required sheets
    let sheetsToCreate = [];

    if (!spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME)) {
      sheetsToCreate.push(CONFIG.CONTACTS_SHEET_NAME);
    }

    if (!spreadsheet.getSheetByName(CONFIG.TEMPLATES_SHEET_NAME)) {
      sheetsToCreate.push(CONFIG.TEMPLATES_SHEET_NAME);
    }

    if (!spreadsheet.getSheetByName(CONFIG.LOGS_SHEET_NAME)) {
      sheetsToCreate.push(CONFIG.LOGS_SHEET_NAME);
    }

    // Save the spreadsheet ID
    PropertiesService.getUserProperties().setProperty("SPREADSHEET_ID", spreadsheetId);

    // Create missing sheets if needed
    if (sheetsToCreate.length > 0) {
      logAction("Database Setup", `Connecting to ${spreadsheetName}. Creating missing sheets: ${sheetsToCreate.join(', ')}`);
      // Call setupDatabase which handles sheet creation and returns an ActionResponse
      return setupDatabase(); // setupDatabase handles navigation/notification
    }

    logAction("Database Connection", "Connected to existing database: " + spreadsheetName + " (ID: " + spreadsheetId + ")");

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("Connected to database successfully!"))
      .setNavigation(CardService.newNavigation()
        .updateCard(buildAddOn())) // Rebuild the main card now that we're connected
      .build();

  } catch (error) {
     console.error("Error connecting to database ID " + spreadsheetId + ": " + error);
     PropertiesService.getUserProperties().deleteProperty("SPREADSHEET_ID"); // Clear invalid ID
     logAction("Error", "Error connecting to database: " + error.message);
    return createNotification("Error connecting to database: Spreadsheet ID might be incorrect or you lack permission. Details: " + error.message);
  }
}

/**
 * Extracts spreadsheet ID from a Google Sheets URL or returns the ID if already provided
 * @param {string} urlOrId - The full URL or just the ID
 * @returns {string|null} - The extracted spreadsheet ID or null if invalid
 */
function extractSpreadsheetIdFromUrl(urlOrId) {
  if (!urlOrId) return null;
  
  // If it's already just the ID (alphanumeric, -, _), return it
  if (/^[a-zA-Z0-9_-]+$/.test(urlOrId) && urlOrId.length > 20) {
    return urlOrId;
  }
  
  // Try to extract from various Google Sheets URL formats
  // Format 1: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
  // Format 2: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit#gid=0
  // Format 3: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID
  const patterns = [
    /\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/,  // Standard format
    /\/d\/([a-zA-Z0-9_-]+)/,                  // Shorter format
    /[-\w]{25,}/                               // Just look for a long alphanumeric string
  ];
  
  for (const pattern of patterns) {
    const match = urlOrId.match(pattern);
    if (match) {
      return match[1] || match[0];
    }
  }
  
  return null;
}

/**
 * Creates a new database
 */
function createNewDatabase() {
  try {
    // Create a new spreadsheet
    const spreadsheet = SpreadsheetApp.create(CONFIG.SPREADSHEET_NAME);
    const spreadsheetId = spreadsheet.getId();
    const spreadsheetName = spreadsheet.getName();

    // Save the spreadsheet ID
    PropertiesService.getUserProperties().setProperty("SPREADSHEET_ID", spreadsheetId);

    // AUTO-POPULATE user info from Google account
    const userEmail = Session.getActiveUser().getEmail();
    const userName = userEmail.split('@')[0].replace(/[._]/g, ' ');
    // Capitalize first letter of each word
    const capitalizedName = userName.split(' ')
      .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
      .join(' ');
    PropertiesService.getUserProperties().setProperty("SENDER_NAME", capitalizedName);

    logAction("Database Setup", `Created new database: ${spreadsheetName} (ID: ${spreadsheetId})`);

    // Set up the database structure (sheets and headers)
    // setupDatabase handles the navigation/notification on completion
    return setupDatabase();

  } catch (error) {
    console.error("Error creating new database: " + error);
    logAction("Error", "Error creating new database: " + error.message);
    PropertiesService.getUserProperties().deleteProperty("SPREADSHEET_ID"); // Clean up if creation failed
    return createNotification("Error creating database: " + error.message);
  }
}

/* ======================== DATABASE SETUP ======================== */

/**
 * Sets up the database structure with required sheets
 */
function setupDatabase() {
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
   if (!spreadsheetId) {
       logAction("Error", "setupDatabase called without SPREADSHEET_ID set.");
       return createNotification("Database setup failed: No spreadsheet ID found.");
   }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let setupPerformed = false;

    // Set up Contacts sheet if it doesn't exist
    let contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);
    if (!contactsSheet) {
      contactsSheet = spreadsheet.insertSheet(CONFIG.CONTACTS_SHEET_NAME);
      setupContactsSheet(contactsSheet);
      logAction("Database Setup", `Created and set up sheet: ${CONFIG.CONTACTS_SHEET_NAME}`);
      setupPerformed = true;
    } else {
        // Optionally: Verify headers if sheet exists? For now, assume it's okay.
        console.log(`Sheet ${CONFIG.CONTACTS_SHEET_NAME} already exists.`);
    }

    // Set up default sequence sheets instead of single templates sheet
    const defaultSequences = ["SaaS / B2B Tech", "Ecommerce / Retail", "Local Services"];
    for (const sequenceName of defaultSequences) {
      const success = createSequenceSheet(sequenceName);
      if (success) {
        setupPerformed = true;
      }
    }

    // Set up Logs sheet if it doesn't exist
    let logsSheet = spreadsheet.getSheetByName(CONFIG.LOGS_SHEET_NAME);
    if (!logsSheet) {
      logsSheet = spreadsheet.insertSheet(CONFIG.LOGS_SHEET_NAME);
      setupLogsSheet(logsSheet);
      logAction("Database Setup", `Created and set up sheet: ${CONFIG.LOGS_SHEET_NAME}`);
      setupPerformed = true;
    } else {
        console.log(`Sheet ${CONFIG.LOGS_SHEET_NAME} already exists.`);
    }

    if (setupPerformed) {
        logAction("Database Setup", "Database structure verified/updated successfully.");
        SpreadsheetApp.flush(); // Ensure sheet creation is committed
        
        // Add demo contacts for first-time users
        addDemoContactsIfNeeded();
    }

    // Show simple success notification and go to dashboard
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("✅ Database ready! Demo contact added."))
      .setNavigation(CardService.newNavigation()
        .updateCard(buildAddOn()))
      .build();

  } catch (error) {
     console.error("Error setting up database: " + error);
     logAction("Error", "Error setting up database structure: " + error.message);
    return createNotification("Error setting up database: " + error.message);
  }
}

/* ======================== SHEET SETUP FUNCTIONS ======================== */

/**
 * Sets up the Contacts sheet with all required headers, formatting, and data validation.
 * This is the complete version including THREAD_ID, LABELED, and REPLY tracking columns.
 */
function setupContactsSheet(sheet) {
  // Set all headers in the correct order
  const headers = [
    "First Name", "Last Name", "Email", "Company", "Title",
    "Current Step", "Last Email Date", "Next Step Date", "Status", "Notes",
    "Personal Phone", "Work Phone", "Personal Called", "Work Called",
    "Priority", "Tags", "Sequence", "Industry", "Step 1 Subject", 
    "Step 1 Sent Message ID", "Connect Sales Link", "Thread ID", "Labeled",
    "Reply Received", "Reply Date"
  ];

  // Write headers to the first row and make them bold
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  sheet.setFrozenRows(1);

  // Set specific column widths for better readability
  sheet.autoResizeColumn(1); // First Name
  sheet.autoResizeColumn(2); // Last Name
  sheet.setColumnWidth(CONTACT_COLS.EMAIL + 1, 250);
  sheet.setColumnWidth(CONTACT_COLS.COMPANY + 1, 150);
  sheet.setColumnWidth(CONTACT_COLS.TITLE + 1, 150);
  sheet.setColumnWidth(CONTACT_COLS.NOTES + 1, 300); 
  sheet.setColumnWidth(CONTACT_COLS.TAGS + 1, 200);
  sheet.setColumnWidth(CONTACT_COLS.SEQUENCE + 1, 200);
  sheet.setColumnWidth(CONTACT_COLS.INDUSTRY + 1, 150);
  sheet.setColumnWidth(CONTACT_COLS.STEP1_SUBJECT + 1, 250);
  sheet.setColumnWidth(CONTACT_COLS.STEP1_SENT_MESSAGE_ID + 1, 250);
  sheet.setColumnWidth(CONTACT_COLS.CONNECT_SALES_LINK + 1, 250);
  sheet.setColumnWidth(CONTACT_COLS.THREAD_ID + 1, 250); // Set width for Thread ID
  sheet.setColumnWidth(CONTACT_COLS.LABELED + 1, 70);    // Set width for Labeled
  sheet.setColumnWidth(CONTACT_COLS.REPLY_RECEIVED + 1, 100); // Set width for Reply Received
  sheet.setColumnWidth(CONTACT_COLS.REPLY_DATE + 1, 150);     // Set width for Reply Date

  // --- Add Data Validation Rules ---

  // Status column validation
  const statusRange = sheet.getRange(2, CONTACT_COLS.STATUS + 1, sheet.getMaxRows() - 1, 1);
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Active", "Paused", "Completed", "Unsubscribed"], true)
    .setAllowInvalid(false)
    .build();
  statusRange.setDataValidation(statusRule);

  // Call status columns validation (Yes/No)
  const personalCalledRange = sheet.getRange(2, CONTACT_COLS.PERSONAL_CALLED + 1, sheet.getMaxRows() - 1, 1);
  const workCalledRange = sheet.getRange(2, CONTACT_COLS.WORK_CALLED + 1, sheet.getMaxRows() - 1, 1);
  const callRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Yes", "No"], true)
    .setAllowInvalid(false)
    .build();
  personalCalledRange.setDataValidation(callRule);
  workCalledRange.setDataValidation(callRule);

  // Priority column validation (High/Medium/Low)
  const priorityRange = sheet.getRange(2, CONTACT_COLS.PRIORITY + 1, sheet.getMaxRows() - 1, 1);
  const priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["High", "Medium", "Low"], true)
    .setAllowInvalid(false)
    .build();
  priorityRange.setDataValidation(priorityRule);

  // Sequence column validation (dynamic)
  const sequenceRange = sheet.getRange(2, CONTACT_COLS.SEQUENCE + 1, sheet.getMaxRows() - 1, 1);
  const availableSequences = getAvailableSequences();
  if (availableSequences.length > 0) {
    const sequenceRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(availableSequences, true)
      .setAllowInvalid(false) // Sequence is required
      .build();
    sequenceRange.setDataValidation(sequenceRule);
  }

  // Industry column validation
  const industryRange = sheet.getRange(2, CONTACT_COLS.INDUSTRY + 1, sheet.getMaxRows() - 1, 1);
  const industryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(INDUSTRY_OPTIONS, true)
    .setAllowInvalid(true) // Allow empty values
    .build();
  industryRange.setDataValidation(industryRule);

  // Labeled column validation (Yes/No)
  const labeledRange = sheet.getRange(2, CONTACT_COLS.LABELED + 1, sheet.getMaxRows() - 1, 1);
  const labeledRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Yes", "No"], true)
    .setAllowInvalid(false)
    .build();
  labeledRange.setDataValidation(labeledRule);

  // Reply Received column validation (Yes/No)
  const replyReceivedRange = sheet.getRange(2, CONTACT_COLS.REPLY_RECEIVED + 1, sheet.getMaxRows() - 1, 1);
  const replyReceivedRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Yes", "No"], true)
    .setAllowInvalid(false)
    .build();
  replyReceivedRange.setDataValidation(replyReceivedRule);

  // --- Set Number Formats ---

  // Format specific columns as Plain Text to prevent auto-formatting issues
  sheet.getRange(2, CONTACT_COLS.PERSONAL_PHONE + 1, sheet.getMaxRows() - 1, 1).setNumberFormat("@");
  sheet.getRange(2, CONTACT_COLS.WORK_PHONE + 1, sheet.getMaxRows() - 1, 1).setNumberFormat("@");
  sheet.getRange(2, CONTACT_COLS.TAGS + 1, sheet.getMaxRows() - 1, 1).setNumberFormat("@");
  sheet.getRange(2, CONTACT_COLS.SEQUENCE + 1, sheet.getMaxRows() - 1, 1).setNumberFormat("@");
  sheet.getRange(2, CONTACT_COLS.INDUSTRY + 1, sheet.getMaxRows() - 1, 1).setNumberFormat("@");
  sheet.getRange(2, CONTACT_COLS.STEP1_SUBJECT + 1, sheet.getMaxRows() - 1, 1).setNumberFormat("@");
  sheet.getRange(2, CONTACT_COLS.STEP1_SENT_MESSAGE_ID + 1, sheet.getMaxRows() - 1, 1).setNumberFormat("@");
  sheet.getRange(2, CONTACT_COLS.CONNECT_SALES_LINK + 1, sheet.getMaxRows() - 1, 1).setNumberFormat("@");
  sheet.getRange(2, CONTACT_COLS.THREAD_ID + 1, sheet.getMaxRows() - 1, 1).setNumberFormat("@");
  sheet.getRange(2, CONTACT_COLS.LABELED + 1, sheet.getMaxRows() - 1, 1).setNumberFormat("@");
  sheet.getRange(2, CONTACT_COLS.REPLY_RECEIVED + 1, sheet.getMaxRows() - 1, 1).setNumberFormat("@");

  // Format date columns for correct date handling
  sheet.getRange(2, CONTACT_COLS.LAST_EMAIL_DATE + 1, sheet.getMaxRows() - 1, 1).setNumberFormat("yyyy-mm-dd h:mm:ss");
  sheet.getRange(2, CONTACT_COLS.NEXT_STEP_DATE + 1, sheet.getMaxRows() - 1, 1).setNumberFormat("yyyy-mm-dd");
  sheet.getRange(2, CONTACT_COLS.REPLY_DATE + 1, sheet.getMaxRows() - 1, 1).setNumberFormat("yyyy-mm-dd h:mm:ss");
}

/**
 * Sets up the Templates sheet with default templates
 */
function setupTemplatesSheet(sheet) {
  // Set headers - Sequence, Template Name, Subject, Body
  const headers = ["Sequence", "Template Name", "Subject", "Body"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  sheet.setFrozenRows(1);

  // Add default templates for fallback purposes
  const defaultTemplates = [
    [
      "Step 1 - Introduction",
      "{{company}} and Google Ads?",
      "Hi {{firstName}},\n\nI'm part of Google's New Business team here in Canada.\n\nWe help a small group of high-potential companies get started on Google Ads with dedicated strategist support — including campaign setup, optimization, tagging, and competitive insights.\n\nThe program is fully covered by Google, and {{company}} looked like it could be a strong fit based on your growth and category.\n\nOpen to a quick chat to explore?",
    ],
    [
      "Step 2 - Quick Reply",
      "Re: {{originalSubject}}",
      "Hi {{firstName}},\n\nJust a quick follow-up on my previous email. Any thoughts on this?\n\nIf it's easier, feel free to book some time directly on my calendar: [Link to your calendar if applicable]\n\nBest,\n{{senderName}}"
    ],
    [
      "Step 3 - Follow Up",
      "Following up: Google Ads for {{company}}",
      "Hi {{firstName}},\n\nJust following up on my earlier note in case this slipped through.\n\nWe work with a small number of Canadian companies each quarter that are new to Google Ads, helping them launch with full support — no agency required, no cost for the service.\n\nIf {{company}} is still planning to scale this year, happy to connect and share what's working across your industry.\n\nWould a 15-minute chat make sense?",
    ],
    [
      "Step 4 - Value Proposition",
      "What we've done for similar companies",
      "Hi {{firstName}},\n\nTotally understand things can get busy — here's what our Google onboarding team has recently helped other companies accomplish:\n\n1. Activated high-intent traffic on Google Search\n2. Drove new customer growth through YouTube & Display\n3. Launched campaigns with full tagging and conversion tracking in under a week\n\nWould it be helpful to hear how this might apply to {{company}} specifically?\n\nHappy to connect when you're free.",
    ],
    [
      "Step 5 - Final Outreach",
      "Final check-in: Google Ads opportunity for {{company}}",
      "Hi {{firstName}},\n\nThis will be my final note unless I hear back.\n\nGoogle offers a fully-supported onboarding experience for new advertisers — no fees, no commitments, just a focused strategy to help brands like {{company}} scale with confidence.\n\nIf now isn't the right time, happy to revisit later. Otherwise, feel free to point me to the right contact if that's easier.\n\nThanks again for considering!",
    ]
  ];

  // Add default templates only if the sheet has only the header row
  if (sheet.getLastRow() < 2) {
     sheet.getRange(2, 1, defaultTemplates.length, headers.length).setValues(defaultTemplates);
  }

  // Format the sheet
  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);
  sheet.setColumnWidth(3, 400); // Make body column wider
}

/**
 * Sets up the Logs sheet for tracking actions
 */
function setupLogsSheet(sheet) {
  // Set headers
  const headers = ["Timestamp", "Action", "Details", "User"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
  sheet.setFrozenRows(1);

  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
   // Format timestamp column
   sheet.getRange(2, 1, sheet.getMaxRows() -1, 1).setNumberFormat("yyyy-mm-dd h:mm:ss");
}

/* ======================== LOGGING ======================== */

/**
 * Adds ONE demo contact on first database setup - ready to send at Step 1
 * This lets users immediately try the email flow
 */
function addDemoContactsIfNeeded() {
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) return;
  
  // Check if we've already added demo contacts
  if (PropertiesService.getUserProperties().getProperty("DEMO_CONTACTS_ADDED") === "true") {
    return;
  }
  
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);
    
    if (!contactsSheet || contactsSheet.getLastRow() > 1) {
      // Sheet doesn't exist or already has contacts
      return;
    }
    
    // Single demo contact at Step 1, ready to send immediately
    // Uses SaaS / B2B Tech sequence which is created by default
    const demoContact = [
      "Demo", "Contact", "demo@example.com", "Example Company", "Manager",
      1, "", "", "Active", "Try sending an email to this demo contact! Delete when done.",
      "", "", "No", "No", "Medium", "demo",
      "SaaS / B2B Tech", "SaaS / B2B Tech", "", "", "", "", "No",
      "No", ""  // Reply Received, Reply Date
    ];
    
    contactsSheet.getRange(2, 1, 1, demoContact.length).setValues([demoContact]);
    SpreadsheetApp.flush();
    
    PropertiesService.getUserProperties().setProperty("DEMO_CONTACTS_ADDED", "true");
    logAction("Demo Setup", "Added demo contact for onboarding");
    
  } catch (error) {
    console.error("Error adding demo contact: " + error);
  }
}

/**
 * Deletes the demo contact from the database after user sends their first email
 */
function deleteDemoContact(demoEmail) {
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) return;
  
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);
    
    if (!contactsSheet) return;
    
    const data = contactsSheet.getDataRange().getValues();
    
    // Find and delete the demo contact row
    for (let i = data.length - 1; i >= 1; i--) {
      const email = data[i][CONTACT_COLS.EMAIL];
      if (email && email.toString().toLowerCase().includes("example.com")) {
        contactsSheet.deleteRow(i + 1); // +1 for 1-based row index
        SpreadsheetApp.flush();
        logAction("Demo Cleanup", "Auto-deleted demo contact after first email: " + email);
        break;
      }
    }
  } catch (error) {
    console.error("Error deleting demo contact: " + error);
  }
}

/**
 * Logs an action to the Logs sheet
 */
function logAction(action, details) {
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    console.log("Log Action Skipped: No database connected.");
    return;
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const logsSheet = spreadsheet.getSheetByName(CONFIG.LOGS_SHEET_NAME);

    if (!logsSheet) {
       console.error("Log Action Failed: Logs sheet not found.");
      return; // Logs sheet doesn't exist
    }

    const user = Session.getActiveUser() ? Session.getActiveUser().getEmail() : "Unknown User"; // Handle cases where user might not be available (e.g., triggers)
    const timestamp = new Date();

    // Append row at the first empty row
    logsSheet.appendRow([timestamp, action, details, user]);

  } catch (error) {
    console.error("Error logging action '" + action + "': " + error.toString());
    // Avoid creating infinite loops by trying to log the logging error
  }
}

