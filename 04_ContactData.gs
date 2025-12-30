/**
 * FILE: Contact Data Operations
 * 
 * PURPOSE:
 * Handles all contact data retrieval, querying, filtering, and statistics.
 * Provides data layer for contact management without UI concerns.
 * 
 * KEY FUNCTIONS:
 * - getAllContactsData() - Retrieves all contacts with filtering
 * - getContactByEmail() - Find specific contact by email
 * - getContactByEmailFast() - Fast single-contact lookup using TextFinder (optimized for contextual triggers)
 * - getContactsReadyForEmail() - Get contacts ready for next step
 * - getContactsInStep() - Get contacts in a specific sequence step
 * - getContactStats() - Calculate contact statistics
 * - isContactReadyForEmail() - Check if contact is ready for email
 * - migrateExistingContactsToSequences() - Data migration utility
 * - validateSequenceConsistency() - Data validation utility
 * 
 * CACHE FUNCTIONS:
 * - getContactFromCache() - Retrieve contact from memory cache
 * - updateContactInCache() - Update/add contact in cache
 * - removeContactFromCache() - Remove contact from cache
 * - clearContactCache() - Clear all cached contacts
 * - populateContactCache() - Pre-load recent contacts into cache
 * 
 * DEPENDENCIES:
 * - 01_Config.gs: CONFIG, CONTACT_COLS
 * - 06_SequenceData.gs: getAvailableSequences
 * 
 * @version 2.4 - Added contact caching for faster contextual trigger loading
 */

/* ======================== CONTACT CACHE ======================== */

// Cache configuration
const CONTACT_CACHE_KEY = "RECENT_CONTACTS_CACHE";
const CONTACT_CACHE_EXPIRY = 7200; // 2 hours in seconds

/**
 * Gets a contact from cache by email.
 * @param {string} email - The email to look up
 * @returns {Object|null} Contact object or null if not in cache
 */
function getContactFromCache(email) {
  if (!email) return null;
  
  try {
    const cache = CacheService.getUserCache();
    const cachedData = cache.get(CONTACT_CACHE_KEY);
    
    if (!cachedData) return null;
    
    const contacts = JSON.parse(cachedData);
    const normalizedEmail = email.toLowerCase().trim();
    
    const contact = contacts.find(c => 
      c.email && c.email.toLowerCase().trim() === normalizedEmail
    );
    
    if (contact) {
      console.log("Cache HIT for: " + email);
      // Restore Date objects (JSON serialization converts them to strings)
      if (contact.lastEmailDate) contact.lastEmailDate = new Date(contact.lastEmailDate);
      if (contact.nextStepDate) contact.nextStepDate = new Date(contact.nextStepDate);
      if (contact.replyDate) contact.replyDate = new Date(contact.replyDate);
      // Recalculate isReady since it depends on current date
      contact.isReady = isContactReadyForEmail(contact.nextStepDate);
    }
    
    return contact || null;
  } catch (error) {
    console.error("Error reading from cache: " + error);
    return null;
  }
}

/**
 * Updates a single contact in the cache.
 * @param {Object} contact - The contact object to update
 */
function updateContactInCache(contact) {
  if (!contact || !contact.email) return;
  
  try {
    const cache = CacheService.getUserCache();
    const cachedData = cache.get(CONTACT_CACHE_KEY);
    
    let contacts = [];
    if (cachedData) {
      contacts = JSON.parse(cachedData);
    }
    
    const normalizedEmail = contact.email.toLowerCase().trim();
    const existingIndex = contacts.findIndex(c => 
      c.email && c.email.toLowerCase().trim() === normalizedEmail
    );
    
    if (existingIndex >= 0) {
      // Update existing
      contacts[existingIndex] = contact;
    } else {
      // Add to beginning (most recent)
      contacts.unshift(contact);
      // Keep only last 50
      if (contacts.length > 50) {
        contacts = contacts.slice(0, 50);
      }
    }
    
    cache.put(CONTACT_CACHE_KEY, JSON.stringify(contacts), CONTACT_CACHE_EXPIRY);
    console.log("Cache UPDATED for: " + contact.email);
  } catch (error) {
    console.error("Error updating cache: " + error);
  }
}

/**
 * Adds a new contact to the cache (alias for updateContactInCache).
 * @param {Object} contact - The contact object to add
 */
function addContactToCache(contact) {
  updateContactInCache(contact);
}

/**
 * Removes a contact from the cache by email.
 * @param {string} email - The email of the contact to remove
 */
function removeContactFromCache(email) {
  if (!email) return;
  
  try {
    const cache = CacheService.getUserCache();
    const cachedData = cache.get(CONTACT_CACHE_KEY);
    
    if (!cachedData) return;
    
    let contacts = JSON.parse(cachedData);
    const normalizedEmail = email.toLowerCase().trim();
    
    contacts = contacts.filter(c => 
      !c.email || c.email.toLowerCase().trim() !== normalizedEmail
    );
    
    cache.put(CONTACT_CACHE_KEY, JSON.stringify(contacts), CONTACT_CACHE_EXPIRY);
    console.log("Cache REMOVED: " + email);
  } catch (error) {
    console.error("Error removing from cache: " + error);
  }
}

/**
 * Clears the entire contact cache.
 */
function clearContactCache() {
  try {
    const cache = CacheService.getUserCache();
    cache.remove(CONTACT_CACHE_KEY);
    console.log("Contact cache cleared.");
  } catch (error) {
    console.error("Error clearing cache: " + error);
  }
}

/**
 * Pre-populates the cache with recent contacts.
 * @param {number} limit - Maximum number of contacts to cache (default 50)
 * @returns {number} Number of contacts cached
 */
function populateContactCache(limit = 50) {
  try {
    const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
    if (!spreadsheetId) return 0;
    
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);
    
    if (!contactsSheet) return 0;
    
    const lastRow = contactsSheet.getLastRow();
    if (lastRow <= 1) return 0;
    
    // Get the most recent contacts (last N rows)
    const startRow = Math.max(2, lastRow - limit + 1);
    const numRows = lastRow - startRow + 1;
    const numCols = Object.keys(CONTACT_COLS).length;
    
    const dataRange = contactsSheet.getRange(startRow, 1, numRows, numCols);
    const dataValues = dataRange.getValues();
    
    const contacts = [];
    for (let i = 0; i < dataValues.length; i++) {
      const row = dataValues[i];
      if (!row[CONTACT_COLS.EMAIL] && !row[CONTACT_COLS.FIRST_NAME]) {
        continue;
      }
      
      const isReady = isContactReadyForEmail(row[CONTACT_COLS.NEXT_STEP_DATE]);
      
      contacts.push({
        firstName: row[CONTACT_COLS.FIRST_NAME] || "",
        lastName: row[CONTACT_COLS.LAST_NAME] || "",
        email: row[CONTACT_COLS.EMAIL] || "",
        company: row[CONTACT_COLS.COMPANY] || "",
        title: row[CONTACT_COLS.TITLE] || "",
        currentStep: parseInt(row[CONTACT_COLS.CURRENT_STEP]) || 1,
        lastEmailDate: row[CONTACT_COLS.LAST_EMAIL_DATE],
        nextStepDate: row[CONTACT_COLS.NEXT_STEP_DATE],
        status: row[CONTACT_COLS.STATUS] || "Active",
        notes: row[CONTACT_COLS.NOTES] || "",
        personalPhone: row[CONTACT_COLS.PERSONAL_PHONE] || "",
        workPhone: row[CONTACT_COLS.WORK_PHONE] || "",
        personalCalled: row[CONTACT_COLS.PERSONAL_CALLED] || "No",
        workCalled: row[CONTACT_COLS.WORK_CALLED] || "No",
        priority: row[CONTACT_COLS.PRIORITY] || "Medium",
        tags: row[CONTACT_COLS.TAGS] || "",
        sequence: row[CONTACT_COLS.SEQUENCE] || "",
        industry: row[CONTACT_COLS.INDUSTRY] || "",
        step1Subject: row[CONTACT_COLS.STEP1_SUBJECT] || "",
        step1SentMessageId: row[CONTACT_COLS.STEP1_SENT_MESSAGE_ID] || "",
        threadId: row[CONTACT_COLS.THREAD_ID] || "",
        labeled: row[CONTACT_COLS.LABELED] || "No",
        replyReceived: row[CONTACT_COLS.REPLY_RECEIVED] || "No",
        replyDate: row[CONTACT_COLS.REPLY_DATE] || "",
        rowIndex: startRow + i,
        isReady: isReady
      });
    }
    
    // Store in cache (most recent first)
    const cache = CacheService.getUserCache();
    cache.put(CONTACT_CACHE_KEY, JSON.stringify(contacts.reverse()), CONTACT_CACHE_EXPIRY);
    
    console.log("Contact cache POPULATED with " + contacts.length + " contacts");
    return contacts.length;
  } catch (error) {
    console.error("Error populating cache: " + error);
    return 0;
  }
}

/* ======================== CONTACT DATA RETRIEVAL ======================== */

/**
 * Gets ALL contacts data from the sheet.
 * CORRECTED to read all columns including Thread ID and Labeled.
 */
function getAllContactsData() {
  const contacts = [];
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");

  if (!spreadsheetId) {
    logAction("Error", "getAllContactsData: No database connected.");
    return contacts;
  }

  // --- Priority Filter Logic (remains the same) ---
  const userProps = PropertiesService.getUserProperties();
  const priorityFiltersString = userProps.getProperty("PRIORITY_FILTERS");
  let enabledPriorities = [];
  let priorityFilterActive = false;

  if (priorityFiltersString === null || priorityFiltersString === undefined || priorityFiltersString === "") {
    enabledPriorities = ["High", "Medium", "Low"];
    priorityFilterActive = false;
  } else {
    enabledPriorities = priorityFiltersString.split(',').map(p => p.trim());
    priorityFilterActive = true;
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);

    if (!contactsSheet) {
      logAction("Error", "getAllContactsData: Contacts sheet not found.");
      return contacts;
    }

    const lastRow = contactsSheet.getLastRow();
    if (lastRow <= 1) return contacts;

    const numColsToRead = Math.max(contactsSheet.getLastColumn(), Object.keys(CONTACT_COLS).length);
    const dataRange = contactsSheet.getRange(2, 1, lastRow - 1, numColsToRead);
    const dataValues = dataRange.getValues();

    for (let i = 0; i < dataValues.length; i++) {
      const row = dataValues[i];
      if (!row[CONTACT_COLS.EMAIL] && !row[CONTACT_COLS.FIRST_NAME]) {
          continue;
      }

      const contactPriority = row[CONTACT_COLS.PRIORITY] || "Medium";
      if (priorityFilterActive && !enabledPriorities.includes(contactPriority)) {
        continue;
      }

      const isReady = isContactReadyForEmail(row[CONTACT_COLS.NEXT_STEP_DATE]);

      contacts.push({
        firstName: row[CONTACT_COLS.FIRST_NAME] || "",
        lastName: row[CONTACT_COLS.LAST_NAME] || "",
        email: row[CONTACT_COLS.EMAIL] || "",
        company: row[CONTACT_COLS.COMPANY] || "",
        title: row[CONTACT_COLS.TITLE] || "",
        currentStep: parseInt(row[CONTACT_COLS.CURRENT_STEP]) || 1,
        lastEmailDate: row[CONTACT_COLS.LAST_EMAIL_DATE], 
        nextStepDate: row[CONTACT_COLS.NEXT_STEP_DATE],   
        status: row[CONTACT_COLS.STATUS] || "Active",
        notes: row[CONTACT_COLS.NOTES] || "",
        personalPhone: row[CONTACT_COLS.PERSONAL_PHONE] || "",
        workPhone: row[CONTACT_COLS.WORK_PHONE] || "",
        personalCalled: row[CONTACT_COLS.PERSONAL_CALLED] || "No",
        workCalled: row[CONTACT_COLS.WORK_CALLED] || "No",
        priority: contactPriority,
        tags: row[CONTACT_COLS.TAGS] || "",
        sequence: row[CONTACT_COLS.SEQUENCE] || "",
        industry: row[CONTACT_COLS.INDUSTRY] || "",
        step1Subject: row[CONTACT_COLS.STEP1_SUBJECT] || "",
        step1SentMessageId: row[CONTACT_COLS.STEP1_SENT_MESSAGE_ID] || "",
        threadId: row[CONTACT_COLS.THREAD_ID] || "",
        labeled: row[CONTACT_COLS.LABELED] || "No",
        replyReceived: row[CONTACT_COLS.REPLY_RECEIVED] || "No",
        replyDate: row[CONTACT_COLS.REPLY_DATE] || "",
        rowIndex: i + 2, // Row number in sheet (1-indexed, +1 for header)
        isReady: isReady 
      });
    }
    return contacts;
  } catch (error) {
    console.error("Error getting all contacts data: " + error + "\n" + error.stack);
    logAction("Error", "Error in getAllContactsData: " + error.toString());
    return []; 
  }
}

/**
 * Gets a contact by email
 * Now uses getAllContactsData and finds the specific contact. More efficient for single lookups than reading sheet again.
 */
function getContactByEmail(email) {
    if (!email) return null;
    const normalizedEmail = email.toLowerCase().trim();
    const allContacts = getAllContactsData(); // Consider caching this if called very frequently in one execution
    return allContacts.find(contact => contact.email && contact.email.toLowerCase().trim() === normalizedEmail) || null;
}

/**
 * FAST single-contact lookup by email using TextFinder.
 * Optimized for contextual triggers - avoids reading entire sheet.
 * Returns contact object or null if not found.
 * 
 * @param {string} email - The email to search for
 * @param {string} [spreadsheetId] - Optional spreadsheet ID (avoids extra property read if already known)
 * @returns {Object|null} Contact object or null
 */
function getContactByEmailFast(email, spreadsheetId) {
  if (!email) return null;
  
  const normalizedEmail = email.toLowerCase().trim();
  
  // CHECK CACHE FIRST (fast path)
  const cachedContact = getContactFromCache(normalizedEmail);
  if (cachedContact) {
    return cachedContact;
  }
  
  // CACHE MISS - Fall back to TextFinder lookup
  console.log("Cache MISS for: " + email + " - using TextFinder");
  
  // Use provided spreadsheetId or fetch it
  const ssId = spreadsheetId || PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!ssId) return null;
  
  try {
    const spreadsheet = SpreadsheetApp.openById(ssId);
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);
    
    if (!contactsSheet) return null;
    
    // Use TextFinder for O(1) lookup instead of reading all rows
    const emailColumn = contactsSheet.getRange(2, CONTACT_COLS.EMAIL + 1, contactsSheet.getLastRow() - 1, 1);
    const textFinder = emailColumn.createTextFinder(normalizedEmail).matchCase(false).matchEntireCell(true);
    const foundRange = textFinder.findNext();
    
    if (!foundRange) return null;
    
    // Found the email - get the full row
    const rowIndex = foundRange.getRow();
    const numCols = Math.max(contactsSheet.getLastColumn(), Object.keys(CONTACT_COLS).length);
    const rowData = contactsSheet.getRange(rowIndex, 1, 1, numCols).getValues()[0];
    
    // Build contact object
    const isReady = isContactReadyForEmail(rowData[CONTACT_COLS.NEXT_STEP_DATE]);
    
    const contact = {
      firstName: rowData[CONTACT_COLS.FIRST_NAME] || "",
      lastName: rowData[CONTACT_COLS.LAST_NAME] || "",
      email: rowData[CONTACT_COLS.EMAIL] || "",
      company: rowData[CONTACT_COLS.COMPANY] || "",
      title: rowData[CONTACT_COLS.TITLE] || "",
      currentStep: parseInt(rowData[CONTACT_COLS.CURRENT_STEP]) || 1,
      lastEmailDate: rowData[CONTACT_COLS.LAST_EMAIL_DATE],
      nextStepDate: rowData[CONTACT_COLS.NEXT_STEP_DATE],
      status: rowData[CONTACT_COLS.STATUS] || "Active",
      notes: rowData[CONTACT_COLS.NOTES] || "",
      personalPhone: rowData[CONTACT_COLS.PERSONAL_PHONE] || "",
      workPhone: rowData[CONTACT_COLS.WORK_PHONE] || "",
      personalCalled: rowData[CONTACT_COLS.PERSONAL_CALLED] || "No",
      workCalled: rowData[CONTACT_COLS.WORK_CALLED] || "No",
      priority: rowData[CONTACT_COLS.PRIORITY] || "Medium",
      tags: rowData[CONTACT_COLS.TAGS] || "",
      sequence: rowData[CONTACT_COLS.SEQUENCE] || "",
      industry: rowData[CONTACT_COLS.INDUSTRY] || "",
      step1Subject: rowData[CONTACT_COLS.STEP1_SUBJECT] || "",
      step1SentMessageId: rowData[CONTACT_COLS.STEP1_SENT_MESSAGE_ID] || "",
      threadId: rowData[CONTACT_COLS.THREAD_ID] || "",
      labeled: rowData[CONTACT_COLS.LABELED] || "No",
      replyReceived: rowData[CONTACT_COLS.REPLY_RECEIVED] || "No",
      replyDate: rowData[CONTACT_COLS.REPLY_DATE] || "",
      rowIndex: rowIndex,
      isReady: isReady
    };
    
    // Add to cache for future lookups
    addContactToCache(contact);
    
    return contact;
  } catch (error) {
    console.error("Error in getContactByEmailFast: " + error);
    return null;
  }
}

/* ======================== CONTACT FILTERING ======================== */

/**
 * Gets contacts that are ready for the next email
 * Now uses getAllContactsData and filters.
 */
function getContactsReadyForEmail() {
  const allContacts = getAllContactsData();
  const readyContacts = allContacts.filter(contact =>
       contact.status === "Active" && contact.isReady // Use pre-calculated readiness
  );

   // Sort by step, then maybe priority or name?
   readyContacts.sort((a, b) => {
     if (a.currentStep !== b.currentStep) {
         return a.currentStep - b.currentStep;
     }
     // Optional: Secondary sort by priority (High > Medium > Low)
     const priorityOrder = { "High": 1, "Medium": 2, "Low": 3 };
     const priorityA = priorityOrder[a.priority] || 2;
     const priorityB = priorityOrder[b.priority] || 2;
      if (priorityA !== priorityB) {
          return priorityA - priorityB;
      }
     // Optional: Tertiary sort by name
     return (a.lastName + a.firstName).localeCompare(b.lastName + b.firstName);
   });


  return readyContacts;
}

/**
 * Gets contacts in a specific sequence step (filtered from all contacts).
 */
function getContactsInStep(step) {
  const allContacts = getAllContactsData(); // Get all contacts once

  const contactsInStep = allContacts.filter(contact =>
     contact.currentStep === step &&
     (contact.status === "Active" || contact.status === "Paused")
  );

  return contactsInStep;
}

/**
 * Gets contacts that need to be called for a specific step.
 * Filters from all contacts data.
 */
function getContactsForCallStep(step) {
  const allContacts = getAllContactsData(); // Get all contacts once

  const contactsForStep = allContacts.filter(contact => {
     const hasValidPersonal = formatPhoneNumberForDisplay(contact.personalPhone) !== "N/A";
     const hasValidWork = formatPhoneNumberForDisplay(contact.workPhone) !== "N/A";
     const personalNeedsCall = hasValidPersonal && contact.personalCalled !== "Yes";
     const workNeedsCall = hasValidWork && contact.workCalled !== "Yes";

     // Only include contacts who:
     // 1. Are in the correct step
     // 2. Have active or paused status
     // 3. Have received at least the previous step's email (implied by being in this step > 1, or step 1 itself)
     //    - A check for lastEmailDate might be too strict if step 1 allows calling before first email.
     //    - Let's assume if they are in step X > 1, the email for X-1 was sent.
     // 4. Have at least one valid phone number that hasn't been called yet for this step
     return contact.currentStep === step &&
            (contact.status === "Active" || contact.status === "Paused") &&
            (personalNeedsCall || workNeedsCall);
  });

  return contactsForStep;
}

/* ======================== CONTACT STATISTICS ======================== */

/**
 * Gets contact statistics (uses efficient getAllContactsData).
 */
function getContactStats() {
  const stats = {
    total: 0,
    activePaused: 0, // Combined count for step totals
    step1: 0, step2: 0, step3: 0, step4: 0, step5: 0,
    completed: 0, unsubscribed: 0,
    readyForStep1: 0, readyForStep2: 0, readyForStep3: 0, readyForStep4: 0, readyForStep5: 0
  };

  const allContacts = getAllContactsData();
  stats.total = allContacts.length;

  // Count contacts by step and status
  for (const contact of allContacts) {
    const currentStep = contact.currentStep;
    const status = contact.status;

    if (status === "Completed") {
      stats.completed++;
    } else if (status === "Unsubscribed") {
      stats.unsubscribed++;
    } else if (status === "Active" || status === "Paused") {
      stats.activePaused++;
      // Count contacts by step
      if (currentStep >= 1 && currentStep <= CONFIG.SEQUENCE_STEPS) {
        stats["step" + currentStep]++;
      }

      // Check if ready for next email (only active contacts are ready)
      if (contact.isReady && status === "Active") { // Use pre-calculated readiness
         if (currentStep >= 1 && currentStep <= CONFIG.SEQUENCE_STEPS) {
           stats["readyForStep" + currentStep]++;
         }
      }
    }
  }

  return stats;
}

/* ======================== CONTACT READINESS CHECK ======================== */

/**
 * Checks if a contact is ready for the next email based on the nextStepDate.
 * Handles different date representations from Sheets.
 */
function isContactReadyForEmail(nextStepDateVal) {
  const today = new Date();
  today.setHours(0, 0, 0, 0); // Set to beginning of today for accurate comparison

  let nextStepDate = null;

  // Handle Date objects
  if (nextStepDateVal instanceof Date && !isNaN(nextStepDateVal)) {
    nextStepDate = new Date(nextStepDateVal); // Clone to avoid modifying original object
  }
  // Handle valid date strings
  else if (typeof nextStepDateVal === "string" && nextStepDateVal) {
    try {
      nextStepDate = new Date(nextStepDateVal);
      if (isNaN(nextStepDate)) nextStepDate = null; // Invalid date string parsed
    } catch (e) { /* Ignore parsing errors */ }
  }
  // Handle potential numeric timestamps (less common from sheets unless manually entered)
  else if (typeof nextStepDateVal === 'number' && nextStepDateVal > 0) {
      try {
           nextStepDate = new Date(nextStepDateVal);
           if (isNaN(nextStepDate)) nextStepDate = null;
      } catch(e) { /* Ignore errors */ }
  }


  // If no valid next step date is set, they are considered ready (for the first step usually)
  if (!nextStepDate) {
    return true;
  }

  // Normalize the next step date to the beginning of its day
  nextStepDate.setHours(0, 0, 0, 0);

  // Contact is ready if the next step date is today or in the past
  return nextStepDate <= today;
}

/* ======================== DATA MIGRATION & VALIDATION ======================== */

/**
 * Ensures the Contacts sheet has all required columns including new ones.
 * Adds missing columns like Reply Received and Reply Date for existing databases.
 */
function ensureContactsSheetColumns() {
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) return;

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);
    if (!contactsSheet) return;

    const headerRange = contactsSheet.getRange(1, 1, 1, contactsSheet.getLastColumn());
    const headers = headerRange.getValues()[0];

    // Expected headers in order
    const expectedHeaders = [
      "First Name", "Last Name", "Email", "Company", "Title",
      "Current Step", "Last Email Date", "Next Step Date", "Status", "Notes",
      "Personal Phone", "Work Phone", "Personal Called", "Work Called",
      "Priority", "Tags", "Sequence", "Industry", "Step 1 Subject", 
      "Step 1 Sent Message ID", "Connect Sales Link", "Thread ID", "Labeled",
      "Reply Received", "Reply Date"
    ];

    // Check for missing columns at the end
    const currentColumnCount = headers.length;
    const expectedColumnCount = expectedHeaders.length;

    if (currentColumnCount < expectedColumnCount) {
      // Add missing headers
      const missingHeaders = expectedHeaders.slice(currentColumnCount);
      const startCol = currentColumnCount + 1;
      
      contactsSheet.getRange(1, startCol, 1, missingHeaders.length)
        .setValues([missingHeaders])
        .setFontWeight("bold");

      // Add validation for Reply Received column if it was just added
      if (!headers.includes("Reply Received")) {
        const replyReceivedCol = CONTACT_COLS.REPLY_RECEIVED + 1;
        const replyReceivedRange = contactsSheet.getRange(2, replyReceivedCol, contactsSheet.getMaxRows() - 1, 1);
        const replyReceivedRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(["Yes", "No"], true)
          .setAllowInvalid(false)
          .build();
        replyReceivedRange.setDataValidation(replyReceivedRule);
        replyReceivedRange.setNumberFormat("@");
      }

      // Set date format for Reply Date column if it was just added
      if (!headers.includes("Reply Date")) {
        const replyDateCol = CONTACT_COLS.REPLY_DATE + 1;
        contactsSheet.getRange(2, replyDateCol, contactsSheet.getMaxRows() - 1, 1)
          .setNumberFormat("yyyy-mm-dd h:mm:ss");
      }

      SpreadsheetApp.flush();
      logAction("Column Migration", `Added ${missingHeaders.length} missing columns: ${missingHeaders.join(", ")}`);
    }
  } catch (error) {
    console.error("Error ensuring contacts sheet columns: " + error);
  }
}

/**
 * Migration function - Adds default sequences to existing contacts
 */
function migrateExistingContactsToSequences() {
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    logAction("Error", "Migration: No database connected.");
    return { success: false, message: "No database connected." };
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);

    if (!contactsSheet) {
      logAction("Error", "Migration: Contacts sheet not found.");
      return { success: false, message: "Contacts sheet not found." };
    }

    const dataRange = contactsSheet.getDataRange();
    const data = dataRange.getValues();
    
    if (data.length <= 1) {
      return { success: true, message: "No contacts found to migrate." };
    }

    let migratedCount = 0;
    const defaultSequence = "SaaS / B2B Tech";

    // Ensure default sequence sheet exists
    createSequenceSheet(defaultSequence);

    // Process each contact row (skip header)
    for (let i = 1; i < data.length; i++) {
      const currentSequence = data[i][CONTACT_COLS.SEQUENCE] || "";
      
      // Only migrate if sequence is empty
      if (!currentSequence || currentSequence.trim() === "") {
        contactsSheet.getRange(i + 1, CONTACT_COLS.SEQUENCE + 1).setValue(defaultSequence);
        migratedCount++;
      }
    }

    if (migratedCount > 0) {
      SpreadsheetApp.flush();
      logAction("Migration", `Migrated ${migratedCount} contacts to default sequence: ${defaultSequence}`);
    }

    // Clean up old industry filter properties
    const userProps = PropertiesService.getUserProperties();
    userProps.deleteProperty("INDUSTRY_FILTERS");

    return { 
      success: true, 
      message: `Migration completed. ${migratedCount} contacts assigned to default sequence "${defaultSequence}".` 
    };

  } catch (error) {
    console.error("Migration error: " + error + "\n" + error.stack);
    logAction("Error", "Migration error: " + error.toString());
    return { success: false, message: "Migration failed: " + error.message };
  }
}

/**
 * Validates sequence consistency - NEW FUNCTION
 */
function validateSequenceConsistency() {
  const allContacts = getAllContactsData();
  const issues = [];
  
  for (const contact of allContacts) {
    // Check if contact has a sequence assigned
    if (!contact.sequence || contact.sequence.trim() === "") {
      issues.push(`Contact ${contact.email}: No sequence assigned`);
    } else {
      const availableSequences = getAvailableSequences();
      if (!availableSequences.includes(contact.sequence)) {
        issues.push(`Contact ${contact.email}: Invalid sequence "${contact.sequence}"`);
      }
    }
  }
  
  return {
    isValid: issues.length === 0,
    issues: issues,
    totalContacts: allContacts.length,
    contactsWithSequences: allContacts.filter(c => c.sequence && c.sequence.trim() !== "").length
  };
}

/* ======================== CONTACT GROUPING HELPERS ======================== */

/**
 * Checks if a title contains marketing-related keywords (case-insensitive).
 */
function isMarketingTitle(title) {
  if (!title) return false;
  const lowerTitle = title.toLowerCase();
  const keywords = ["market", "marketing", "cmo", "chief marketing", "demand gen", "growth"]; // Add more as needed
  return keywords.some(keyword => lowerTitle.includes(keyword));
}

/**
 * Groups contacts by whether they have a marketing title.
 */
function groupContactsByMarketingTitle(contacts) {
  const marketingContacts = [];
  const otherContacts = [];

  for (const contact of contacts) {
    if (isMarketingTitle(contact.title)) {
      marketingContacts.push(contact);
    } else {
      otherContacts.push(contact);
    }
  }

  // Optional: Sort within groups, e.g., by name
  const sortByName = (a, b) => (a.lastName + a.firstName).localeCompare(b.lastName + b.firstName);
  marketingContacts.sort(sortByName);
  otherContacts.sort(sortByName);


  return {
    marketingContacts: marketingContacts,
    otherContacts: otherContacts
  };
}

/**
 * Groups contacts by company - NEW FUNCTION
 */
function groupContactsByCompany(contacts) {
  const grouped = {};
  
  for (const contact of contacts) {
    const company = contact.company || "No Company";
    if (!grouped[company]) {
      grouped[company] = [];
    }
    grouped[company].push(contact);
  }
  
  return grouped;
}

