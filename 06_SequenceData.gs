/**
 * FILE: Sequence and Template Data Operations
 * 
 * PURPOSE:
 * Handles sequence sheet discovery, creation, and template data retrieval.
 * Manages row-based template format where each sequence has its own sheet.
 * 
 * KEY FUNCTIONS:
 * - getAvailableSequences() - Discover sequences by sheet names
 * - getSequenceSheetName() - Get full sheet name from sequence name
 * - createSequenceSheet() - Create new sequence sheet with default templates
 * - getSequenceTemplateForStep() - Get template for specific sequence and step
 * - getEmailTemplates() - Get all templates (legacy support)
 * - getTemplateForStep() - Get template for step (legacy support)
 * - updateContactSequence() - Change contact's sequence assignment
 * 
 * DEPENDENCIES:
 * - 01_Config.gs: CONFIG, CONTACT_COLS
 * - 03_Database.gs: logAction
 * - 04_ContactData.gs: getContactByEmail
 * 
 * @version 2.3
 */

/* ======================== SEQUENCE DISCOVERY ======================== */

/**
 * Discovers available sequences by looking for sheets named "Sequence-[Name]"
 * @returns {string[]} Array of sequence names (without "Sequence-" prefix)
 */
function getAvailableSequences() {
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    return [];
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();
    const sequences = [];

    for (const sheet of sheets) {
      const sheetName = sheet.getName();
      if (sheetName.startsWith("Sequence-")) {
        const sequenceName = sheetName.substring(9); // Remove "Sequence-" prefix
        sequences.push(sequenceName);
      }
    }

    return sequences.sort(); // Return alphabetically sorted
  } catch (error) {
    console.error("Error discovering sequences: " + error);
    return [];
  }
}

/**
 * Gets the full sheet name for a sequence
 * @param {string} sequenceName The sequence name without prefix
 * @returns {string} Full sheet name with "Sequence-" prefix
 */
function getSequenceSheetName(sequenceName) {
  return "Sequence-" + sequenceName;
}

/* ======================== SEQUENCE CREATION ======================== */

/**
 * Creates a new sequence sheet with ROW-BASED structure
 * @param {string} sequenceName The name of the sequence (without "Sequence-" prefix)
 * @returns {boolean} Success status
 */
function createSequenceSheet(sequenceName) {
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    return false;
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheetName = getSequenceSheetName(sequenceName);
    
    // Check if sheet already exists
    if (spreadsheet.getSheetByName(sheetName)) {
      console.log("Sequence sheet already exists: " + sheetName);
      return true;
    }

    // Create new sheet
    const sheet = spreadsheet.insertSheet(sheetName);
    
    // Set up headers: Step, Name, Subject, Body
    const headers = ["Step", "Name", "Subject", "Body"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
    
    // Add default template data for this sequence
    const defaultTemplates = [
      [1, "Introduction Email", `{{company}} and ${sequenceName} Services?`, `Hi {{firstName}},\n\nI'm reaching out about ${sequenceName} opportunities for {{company}}.\n\nWould you be open to a quick chat to explore how we might help?\n\nBest regards,\n{{senderName}}`],
      [2, "Quick Follow-up", `Re: {{company}} and ${sequenceName} Services?`, `Hi {{firstName}},\n\nJust a quick follow-up on my previous email about ${sequenceName} for {{company}}.\n\nAny thoughts on this?\n\nBest,\n{{senderName}}`],
      [3, "Second Follow-up", `Following up: ${sequenceName} for {{company}}`, `Hi {{firstName}},\n\nJust following up on my earlier note in case this slipped through.\n\nWould a 15-minute chat make sense to discuss ${sequenceName} opportunities?\n\nBest,\n{{senderName}}`],
      [4, "Value Proposition", `What we've done for similar companies`, `Hi {{firstName}},\n\nI wanted to share what we've accomplished for other companies in ${sequenceName}:\n\n1. [Specific benefit 1]\n2. [Specific benefit 2]\n3. [Specific benefit 3]\n\nWould it be helpful to hear how this might apply to {{company}} specifically?\n\nBest,\n{{senderName}}`],
      [5, "Final Outreach", `Final check-in: ${sequenceName} opportunity for {{company}}`, `Hi {{firstName}},\n\nThis will be my final note unless I hear back.\n\nIf now isn't the right time for ${sequenceName} discussions, I completely understand. Feel free to reach out when it makes more sense.\n\nThanks for considering!\n\n{{senderName}}`]
    ];
    
    // Add the template rows
    sheet.getRange(2, 1, defaultTemplates.length, headers.length).setValues(defaultTemplates);
    
    // Format columns
    sheet.autoResizeColumn(1); // Step column
    sheet.autoResizeColumn(2); // Name column
    sheet.setColumnWidth(3, 300); // Subject column
    sheet.setColumnWidth(4, 500); // Body column (wider for content)
    
    logAction("Sequence Creation", "Created new sequence sheet: " + sequenceName);
    return true;
    
  } catch (error) {
    console.error("Error creating sequence sheet: " + error);
    logAction("Error", "Failed to create sequence sheet: " + sequenceName + " - " + error.toString());
    return false;
  }
}

/* ======================== TEMPLATE RETRIEVAL ======================== */

/**
 * Gets the template for a specific step from sequence-specific sheets - ROW-BASED FORMAT
 * Reads from sequence sheets where each row represents a step
 */
function getSequenceTemplateForStep(sequenceName, stepNumber) {
  if (!sequenceName || !stepNumber) {
    return null;
  }

  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    return null;
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheetName = getSequenceSheetName(sequenceName);
    const sequenceSheet = spreadsheet.getSheetByName(sheetName);

    if (!sequenceSheet) {
      console.log(`Sequence sheet not found: ${sheetName}`);
      return null;
    }

    // Get all data from the sheet
    const dataRange = sequenceSheet.getDataRange();
    const data = dataRange.getValues();
    
    if (data.length <= 1) {
      console.log(`No template data found in sequence sheet: ${sheetName}`);
      return null;
    }

    // Look for the row with the matching step number
    // Format: [Step, Name, Subject, Body]
    for (let i = 1; i < data.length; i++) { // Skip header row
      const row = data[i];
      const rowStep = parseInt(row[0]);
      
      if (rowStep === stepNumber) {
        return {
          sequence: sequenceName,
          name: String(row[1] || `Step ${stepNumber} - ${sequenceName}`),
          subject: String(row[2] || `Step ${stepNumber} Follow-up`),
          body: String(row[3] || `Template content for step ${stepNumber}`),
          stepNumber: stepNumber
        };
      }
    }

    console.log(`No template found for Step ${stepNumber} in sequence ${sequenceName}`);
    return null;

  } catch (error) {
    console.error(`Error getting sequence template: ${error}`);
    return null;
  }
}

/**
 * Gets all email templates for a specific sequence from its sheet.
 * @param {string} sequenceName The sequence name (without "Sequence-" prefix)
 * @returns {Array} Array of template objects with step, name, subject, body, rowIndex
 */
function getEmailTemplates(sequenceName) {
  const templates = [];
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");

  if (!spreadsheetId || !sequenceName) {
    return templates;
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheetName = getSequenceSheetName(sequenceName);
    const sequenceSheet = spreadsheet.getSheetByName(sheetName);

    if (!sequenceSheet) {
      console.log("getEmailTemplates: Sequence sheet not found: " + sheetName);
      return templates;
    }

    const lastRow = sequenceSheet.getLastRow();
    if (lastRow <= 1) {
      return templates; // Only header row exists
    }

    const data = sequenceSheet.getRange(2, 1, lastRow - 1, 4).getValues();

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      // Columns: Step, Name, Subject, Body
      const step = parseInt(row[0]) || 0;
      if (step >= 1 && step <= 5) {
        templates.push({
          step: step,
          name: String(row[1] || "").trim(),
          subject: String(row[2] || ""),
          body: String(row[3] || ""),
          rowIndex: i + 2 // 1-based, +1 for header
        });
      }
    }

    // Sort by step number
    templates.sort((a, b) => a.step - b.step);
    return templates;
  } catch (error) {
    console.error("Error getting email templates: " + error);
    return templates;
  }
}

/**
 * Gets the template for a specific step - UPDATED FOR ROW-BASED FORMAT
 */
function getTemplateForStep(step) {
  const templates = getEmailTemplates(); // Get all templates

  // Find the template for the specific step number
  for (const template of templates) {
    if (template.step === step) {
      return template; // Return the matching step template
    }
  }

  return null; // No matching template found
}

/* ======================== SEQUENCE STEP COUNT ======================== */

/**
 * Gets the number of active steps for a sequence (1-5, default 5)
 * @param {string} sequenceName The sequence name
 * @returns {number} Number of steps (1-5)
 */
function getSequenceStepCount(sequenceName) {
  if (!sequenceName) return 5;
  
  const key = "SEQUENCE_STEPS_" + sequenceName.replace(/\s+/g, "_");
  const stored = PropertiesService.getUserProperties().getProperty(key);
  
  if (stored) {
    const count = parseInt(stored);
    if (count >= 1 && count <= 5) {
      return count;
    }
  }
  
  return 5; // Default to 5 steps
}

/**
 * Sets the number of active steps for a sequence
 * @param {string} sequenceName The sequence name
 * @param {number} stepCount Number of steps (1-5)
 * @returns {boolean} Success status
 */
function setSequenceStepCount(sequenceName, stepCount) {
  if (!sequenceName) return false;
  
  const count = parseInt(stepCount);
  if (count < 1 || count > 5) return false;
  
  const key = "SEQUENCE_STEPS_" + sequenceName.replace(/\s+/g, "_");
  PropertiesService.getUserProperties().setProperty(key, count.toString());
  
  logAction("Sequence Config", `Set ${sequenceName} to ${count} steps`);
  return true;
}

/* ======================== SEQUENCE UPDATES ======================== */

/**
 * Updates contact sequence - NEW FUNCTION
 */
function updateContactSequence(email, newSequence) {
  const contact = getContactByEmail(email);
  if (!contact) {
    return { success: false, message: "Contact not found" };
  }

  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    return { success: false, message: "No database connected" };
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);
    
    if (!contactsSheet) {
      return { success: false, message: "Contacts sheet not found" };
    }

    contactsSheet.getRange(contact.rowIndex, CONTACT_COLS.SEQUENCE + 1).setValue(newSequence);
    SpreadsheetApp.flush();
    
    logAction("Update Sequence", `Updated sequence for ${email} to ${newSequence}`);
    return { success: true };
  } catch (error) {
    logAction("Error", `Error updating sequence for ${email}: ${error.toString()}`);
    return { success: false, message: error.message };
  }
}

