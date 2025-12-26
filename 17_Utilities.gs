/**
 * FILE: Utility Functions
 * 
 * PURPOSE:
 * Provides reusable helper functions for formatting, validation, UI elements,
 * and Gmail message processing used throughout the application.
 * 
 * KEY FUNCTIONS:
 * - createNotification() - Create notification responses
 * - formatDate() - Format dates for display
 * - formatPhoneNumberForDisplay() - Clean phone number formatting
 * - truncateText() - Text truncation helper
 * - extractIdFromUrl() - Extract Google Drive file IDs
 * - createIndustryDropdown() - Generate industry dropdown widget
 * - addPaginationButtons() - Add pagination controls to sections
 * - buildErrorCard() - Error message card
 * - buildInfoCard() - Info message card
 * - getMessageDetails_() - Get Gmail message details
 * - getHeader_() - Extract header from message
 * - extractEmailFromHeader_() - Parse email from header string
 * 
 * DEPENDENCIES:
 * - 01_Config.gs: INDUSTRY_OPTIONS
 * - 03_Database.gs: logAction
 * 
 * @version 2.3
 */

/* ======================== NOTIFICATION HELPERS ======================== */

/**
 * Helper function to create a notification
 */
function createNotification(message) {
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText(message))
    .build();
}

/**
 * Builds a simple card to display an error message.
 */
function buildErrorCard(errorMessage) {
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle("Error"));
  var section = CardService.newCardSection();
  section.addWidget(CardService.newTextParagraph().setText(errorMessage));
  section.addWidget(CardService.newTextButton()
       .setText("Open Full Add-on")
       .setOnClickAction(CardService.newAction().setFunctionName("buildAddOn")));
  card.addSection(section);
  return card.build();
}

/**
 * Builds a simple card to display an informational message.
 */
function buildInfoCard(title, message) {
  var card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader().setTitle(title));
  var section = CardService.newCardSection();
  section.addWidget(CardService.newTextParagraph().setText(message));
   section.addWidget(CardService.newTextButton()
       .setText("Open Full Add-on")
       .setOnClickAction(CardService.newAction().setFunctionName("buildAddOn")));
  card.addSection(section);
  return card.build();
}

/* ======================== FORMATTING HELPERS ======================== */

/**
 * Helper function to format a date string or object
 */
function formatDate(date) {
  if (!date) {
    return "";
  }

  let dateObj;
  if (date instanceof Date && !isNaN(date)) {
      dateObj = date;
  } else if (typeof date === "string") {
      try {
          dateObj = new Date(date);
          if (isNaN(dateObj)) {
              return ""; // Invalid date string
          }
      } catch(e) {
          return ""; // Error parsing date string
      }
  } else if (typeof date === 'number') { // Handle potential timestamp numbers
       try {
          dateObj = new Date(date);
          if (isNaN(dateObj)) {
              return ""; // Invalid number
          }
      } catch(e) {
          return ""; // Error creating date from number
      }
  }
   else {
    return ""; // Not a valid date type
  }

   return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MMM dd, yyyy");
}

/**
 * Helper to format phone numbers stored as text, removing potential leading apostrophe.
 */
function formatPhoneNumberForDisplay(phone) {
    if (!phone) {
        return "N/A";
    }
    // Remove leading apostrophe if present (from Sheets text format)
    let formattedPhone = String(phone).trim();
    if (formattedPhone.startsWith("'")) {
        formattedPhone = formattedPhone.substring(1);
    }
    return formattedPhone || "N/A"; // Return N/A if it becomes empty
}

/**
 * Helper function to truncate text for previews.
 */
function truncateText(text, maxLength) {
    if (!text) return "";
    if (text.length <= maxLength) return text;
    return text.substring(0, maxLength) + "...";
}

/**
 * Extracts a Google Drive file ID from a URL or an ID string.
 * @param {string} urlOrId The Google Drive URL or file ID.
 * @returns {string|null} The extracted file ID or null if not found.
 */
function extractIdFromUrl(urlOrId) {
  if (!urlOrId) return null;
  // Check if it's already just the ID (alphanumeric, -, _)
  if (/^[a-zA-Z0-9_-]+$/.test(urlOrId) && urlOrId.length > 20) {
    return urlOrId;
  }
  // Regex to find the ID in a typical Drive URL
  const match = urlOrId.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

/* ======================== UI COMPONENT HELPERS ======================== */

/**
 * Helper function to create an industry dropdown widget
 * @param {string} fieldName The field name for form submission
 * @param {string} title The title/label for the dropdown
 * @param {string} currentValue The currently selected value (empty string for no selection)
 * @param {boolean} includeNoIndustry Whether to include a "No Industry" option (default: true)
 * @return {SelectionInput} The configured dropdown widget
 */
function createIndustryDropdown(fieldName, title, currentValue = "", includeNoIndustry = true) {
  const dropdown = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle(title)
    .setFieldName(fieldName);
  
  if (includeNoIndustry) {
    dropdown.addItem("(No Industry)", "", currentValue === "" || !currentValue);
  }
  
  for (const industry of INDUSTRY_OPTIONS) {
    dropdown.addItem(industry, industry, currentValue === industry);
  }
  
  return dropdown;
}

/**
 * Helper function to add pagination buttons to a section
 * @param {CardSection} section The section to add buttons to.
 * @param {number} currentPage The current page number (1-based).
 * @param {number} totalPages The total number of pages.
 * @param {string} functionName The function to call when buttons are clicked.
 * @param {Object} parameters Additional parameters to pass to the function.
 */
function addPaginationButtons(section, currentPage, totalPages, functionName, parameters) {
  if (totalPages <= 1) {
    return; // No pagination needed
  }

  const buttonSet = CardService.newButtonSet();
  let addedButton = false;

  // Previous Button
  if (currentPage > 1) {
    const prevParams = { ...parameters, page: (currentPage - 1).toString() };
    buttonSet.addButton(CardService.newTextButton()
      .setText("⬅️ Previous")
      .setOnClickAction(CardService.newAction()
        .setFunctionName(functionName)
        .setParameters(prevParams)));
    addedButton = true;
  }

  // Page Indicator
  section.addWidget(CardService.newTextParagraph().setText(`Page ${currentPage} of ${totalPages}`));

  // Next Button
  if (currentPage < totalPages) {
    const nextParams = { ...parameters, page: (currentPage + 1).toString() };
    buttonSet.addButton(CardService.newTextButton()
      .setText("Next ➡️")
      .setOnClickAction(CardService.newAction()
        .setFunctionName(functionName)
        .setParameters(nextParams)));
    addedButton = true;
  }

  if (addedButton) {
      section.addWidget(buttonSet);
  }
}

/* ======================== GMAIL MESSAGE HELPERS ======================== */

/**
 * Helper function to get message details using Gmail API.
 */
function getMessageDetails_(messageId) {
  try {
    const message = GmailApp.getMessageById(messageId);

    if (!message) {
      console.error("Message not found with ID: " + messageId);
      logAction("Error", "Message not found with ID: " + messageId);
      return null;
    }

    // Create a simple object with the needed structure
    const fromHeader = message.getFrom();
    const toHeader = message.getTo();

    // Return a simplified message object with just what we need
    return {
      id: messageId,
      payload: {
        headers: [
          {
            name: "From",
            value: fromHeader
          },
          {
            name: "To",
            value: toHeader
          }
        ]
      }
    };
  } catch (e) {
    console.error("Error getting message details: " + e + "\n" + e.stack);
    logAction("Error", "Error getting message details: " + e);
    return null;
  }
}

/**
 * Helper function to extract a specific header value.
 */
function getHeader_(headers, name) {
  if (!headers) return "";
  name = name.toLowerCase(); // Case-insensitive matching
  for (var i = 0; i < headers.length; i++) {
    if (headers[i].name.toLowerCase() === name) {
      return headers[i].value;
    }
  }
  return "";
}

/**
 * Extracts the email address from a 'From' header string
 */
function extractEmailFromHeader_(fromHeader) {
    if (!fromHeader) return null;
    const match = fromHeader.match(/<([^>]+)>/); // Look for content within <>
    if (match && match[1]) {
        return match[1]; // Return the email address inside <>
    }
    // If no <>, assume the entire string might be the email (less reliable)
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (emailRegex.test(fromHeader)) {
        return fromHeader;
    }
    return null; // Could not extract a valid email
}

