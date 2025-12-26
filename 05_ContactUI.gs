/**
 * FILE: Contact Management UI
 * 
 * PURPOSE:
 * Handles all contact management user interface cards and forms.
 * Includes contact listing, viewing, editing, searching, and adding contacts.
 * 
 * KEY FUNCTIONS:
 * - buildContactManagementCard() - Main contact management UI with pagination
 * - viewContactCard() - Detailed contact view and edit form
 * - searchContacts() - Contact search with pagination
 * - addContact() - Add new contact action handler
 * - saveAllContactChanges() - Save edits from contact view
 * - displayContactSummaryWidget() - Reusable contact summary widget
 * - createLinkedInSearchUrl() - Generate LinkedIn search URL
 * 
 * DEPENDENCIES:
 * - 01_Config.gs: CONFIG, CONTACT_COLS
 * - 04_ContactData.gs: getAllContactsData, getContactByEmail, groupContactsByMarketingTitle
 * - 06_SequenceData.gs: getAvailableSequences, getTemplateForStep
 * - 17_Utilities.gs: formatDate, formatPhoneNumberForDisplay, addPaginationButtons
 * 
 * @version 2.3
 */

/* ======================== CONTACT MANAGEMENT UI ======================== */

/**
 * Builds the Contact Management card with pagination for recent contacts.
 */
function buildContactManagementCard(e) {
  const page = parseInt(e && e.parameters && e.parameters.page || '1');
  const pageSize = CONFIG.PAGE_SIZE;

  const card = CardService.newCardBuilder();

  // Add header with friendly language
  card.setHeader(CardService.newCardHeader()
      .setTitle("👥 Contacts")
      .setSubtitle("Add people to your email campaigns")
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/people_black_48dp.png"));

  // ============================================================
  // ADD CONTACT FORM - Single unified section
  // Required fields first, then optional, button at the bottom
  // ============================================================
  
  const formSection = CardService.newCardSection()
      .setHeader("➕ Add New Contact");

  // REQUIRED FIELDS
  formSection.addWidget(CardService.newTextInput()
      .setFieldName("firstName")
      .setTitle("First Name *")
      .setHint("For personalized emails"));
  
  formSection.addWidget(CardService.newTextInput()
      .setFieldName("lastName")
      .setTitle("Last Name *")
      .setHint("Their last name"));
  
  formSection.addWidget(CardService.newTextInput()
      .setFieldName("email")
      .setTitle("Email *")
      .setHint("e.g., john@company.com"));

  formSection.addWidget(CardService.newTextInput()
      .setFieldName("company")
      .setTitle("Company *")
      .setHint("For {{company}} in emails"));

  // Email Campaign dropdown (required)
  const sequenceDropdown = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setTitle("Email Campaign *")
      .setFieldName("sequence");
  
  const availableSequences = getAvailableSequences();
  if (availableSequences.length === 0) {
    sequenceDropdown.addItem("No campaigns available - Create one first", "", true);
  } else {
    let defaultSet = false;
    for (const sequence of availableSequences) {
        const isDefault = !defaultSet && (sequence === "SaaS / B2B Tech" || sequence === availableSequences[0]);
        sequenceDropdown.addItem(sequence, sequence, isDefault);
        if (isDefault) defaultSet = true;
    }
  }
  formSection.addWidget(sequenceDropdown);

  // OPTIONAL FIELDS (inline, no separate section)
  formSection.addWidget(CardService.newDivider());
  formSection.addWidget(CardService.newTextParagraph()
      .setText("<i>Optional details:</i>"));

  formSection.addWidget(CardService.newTextInput()
      .setFieldName("title")
      .setTitle("Job Title")
      .setHint("For {{title}} in emails"));

  const priorityDropdown = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setTitle("Priority")
      .setFieldName("priority");
  priorityDropdown.addItem("🔥 High", "High", false);
  priorityDropdown.addItem("🟠 Medium", "Medium", true);
  priorityDropdown.addItem("⚪ Low", "Low", false);
  formSection.addWidget(priorityDropdown);

  formSection.addWidget(CardService.newTextInput()
      .setFieldName("tags")
      .setTitle("Tags")
      .setHint("e.g., conference, hot-lead"));

  formSection.addWidget(CardService.newTextInput()
      .setFieldName("personalPhone")
      .setTitle("Personal Phone")
      .setHint("e.g., 555-123-4567"));

  formSection.addWidget(CardService.newTextInput()
      .setFieldName("workPhone")
      .setTitle("Work Phone")
      .setHint("e.g., 555-987-6543"));

  // ADD CONTACT BUTTON - at the very bottom
  formSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("✅ Add Contact")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction()
              .setFunctionName("addContact"))));

  card.addSection(formSection);

  // Search section (collapsible to reduce clutter)
  const searchSection = CardService.newCardSection()
      .setHeader("🔍 Find a Contact")
      .setCollapsible(true)
      .setNumUncollapsibleWidgets(0);
  
  searchSection.addWidget(CardService.newTextInput()
      .setFieldName("searchTerm")
      .setTitle("Search")
      .setHint("Name, email, or company"));
  
  searchSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("Search")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("searchContacts")
               .setParameters({page: '1'}))));

  card.addSection(searchSection);

  // --- Recent Contacts Section with Pagination ---
  const recentSection = CardService.newCardSection()
      .setHeader("📋 Your Contacts");

  const allContacts = getAllContactsData();

  // Sort by last email date (most recent first), handling potential null/invalid dates
  allContacts.sort((a, b) => {
      const dateA = (a.lastEmailDate instanceof Date && !isNaN(a.lastEmailDate)) ? a.lastEmailDate.getTime() : 0;
      const dateB = (b.lastEmailDate instanceof Date && !isNaN(b.lastEmailDate)) ? b.lastEmailDate.getTime() : 0;
      // If dates are equal or both invalid (0), sort by row index descending (newest added)
      if (dateB === dateA) {
          return b.rowIndex - a.rowIndex;
      }
      return dateB - dateA; // Sort descending by time
  });


  const totalContacts = allContacts.length;
  const totalPages = Math.ceil(totalContacts / pageSize);
  const startIndex = (page - 1) * pageSize;
  const endIndex = Math.min(startIndex + pageSize, totalContacts);
  const contactsToShow = allContacts.slice(startIndex, endIndex);

  if (totalContacts === 0) {
    recentSection.addWidget(CardService.newTextParagraph()
        .setText("No contacts found. Add a contact to get started."));
  } else {
    // Group contacts for display (optional, based on original code)
    const groupedContacts = groupContactsByMarketingTitle(contactsToShow); // Group only the current page

    // Display marketing contacts
    if (groupedContacts.marketingContacts.length > 0) {
      recentSection.addWidget(CardService.newTextParagraph()
          .setText("📊 Marketing Contacts:"));
      for (const contact of groupedContacts.marketingContacts) {
        displayContactSummaryWidget(recentSection, contact); // Use helper
         recentSection.addWidget(CardService.newDivider());
      }
       // Add a small spacer if both groups exist
       if (groupedContacts.otherContacts.length > 0) {
           recentSection.addWidget(CardService.newTextParagraph().setText(" ")); // Spacer
       }
    }

     // Display other contacts
     if (groupedContacts.otherContacts.length > 0) {
       if (groupedContacts.marketingContacts.length > 0) { // Add header only if needed
            recentSection.addWidget(CardService.newTextParagraph().setText("📋 Other Contacts:"));
       }
        for (const contact of groupedContacts.otherContacts) {
            displayContactSummaryWidget(recentSection, contact); // Use helper
            recentSection.addWidget(CardService.newDivider());
        }
     }

     // Add Pagination Controls
     addPaginationButtons(recentSection, page, totalPages, "buildContactManagementCard", {});
  }

  card.addSection(recentSection);

  // Add navigation section (Back to Main Menu)
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
      .setText("Back to Main Menu")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("buildAddOn")));
  card.addSection(navSection);

  return card.build();
}

/**
 * Helper function to display a contact summary widget (used in recent/search)
 */
function displayContactSummaryWidget(section, contact) {
   let title = contact.firstName + " " + contact.lastName;
   let priorityIndicator = "";
   if (contact.priority === "High") priorityIndicator = "🔥 ";
   else if (contact.priority === "Medium") priorityIndicator = "🟠 ";
   else if (contact.priority === "Low") priorityIndicator = "⚪ ";
   title = priorityIndicator + title;

   const bottomLabel = `${contact.company || 'N/A'} | ${contact.title || 'N/A'} | ${contact.sequence || 'N/A'}`;

   section.addWidget(CardService.newKeyValue()
       .setTopLabel(contact.email)
       .setContent(title)
       .setBottomLabel(bottomLabel)
       .setOnClickAction(CardService.newAction()
           .setFunctionName("viewContactCard")
           .setParameters({ email: contact.email }))); // viewContactCard shows details, no page needed here
}

/* ======================== ADD CONTACT ======================== */

/**
 * Adds a new contact
 */
function addContact(e) {
  const firstName = e.formInput.firstName;
  const lastName = e.formInput.lastName;
  const email = e.formInput.email;
  const company = e.formInput.company || "";
  const title = e.formInput.title || "";
  const priority = e.formInput.priority || "Medium";
  const personalPhone = e.formInput.personalPhone || "";
  const workPhone = e.formInput.workPhone || "";
  const tags = e.formInput.tags || "";
  const sequence = e.formInput.sequence || ""; // NEW: Get sequence from form
  const industry = e.formInput.industry || "";

  // Validate required fields with specific error messages
  const missingFields = [];
  if (!firstName || firstName.trim() === "") missingFields.push("First Name");
  if (!lastName || lastName.trim() === "") missingFields.push("Last Name");
  if (!email || email.trim() === "") missingFields.push("Email");
  if (!company || company.trim() === "") missingFields.push("Company");
  
  if (missingFields.length > 0) {
    return createNotification("⚠️ Missing required field(s): " + missingFields.join(", ") + ". Please fill in all fields marked with *.");
  }
  
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(email.trim())) {
    return createNotification("⚠️ Invalid email format: \"" + email + "\". Please use format: name@company.com");
  }

  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    return createNotification("No database connected. Please connect to a database first.");
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);

    if (!contactsSheet) {
      logAction("Error", "Add Contact Error: Contacts sheet not found.");
      return createNotification("Contacts sheet not found. Please refresh the add-on.");
    }

    const normalizedEmail = email.toLowerCase().trim();
    const data = contactsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const existingEmail = data[i][CONTACT_COLS.EMAIL];
      if (existingEmail && existingEmail.toString().toLowerCase().trim() === normalizedEmail) {
        return createNotification("Contact with this email already exists.");
      }
    }

    const newRowData = Array(Object.keys(CONTACT_COLS).length).fill("");
    newRowData[CONTACT_COLS.FIRST_NAME] = firstName;
    newRowData[CONTACT_COLS.LAST_NAME] = lastName;
    newRowData[CONTACT_COLS.EMAIL] = email;
    newRowData[CONTACT_COLS.COMPANY] = company;
    newRowData[CONTACT_COLS.TITLE] = title;
    newRowData[CONTACT_COLS.CURRENT_STEP] = 1;
    newRowData[CONTACT_COLS.LAST_EMAIL_DATE] = "";
    newRowData[CONTACT_COLS.NEXT_STEP_DATE] = "";
    newRowData[CONTACT_COLS.STATUS] = "Active";
    newRowData[CONTACT_COLS.NOTES] = "";
    newRowData[CONTACT_COLS.PERSONAL_PHONE] = personalPhone;
    newRowData[CONTACT_COLS.WORK_PHONE] = workPhone;
    newRowData[CONTACT_COLS.PERSONAL_CALLED] = "No";
    newRowData[CONTACT_COLS.WORK_CALLED] = "No";
    newRowData[CONTACT_COLS.PRIORITY] = priority;
    newRowData[CONTACT_COLS.TAGS] = tags;
    newRowData[CONTACT_COLS.SEQUENCE] = sequence; // NEW: Set sequence field
    newRowData[CONTACT_COLS.INDUSTRY] = industry;
    newRowData[CONTACT_COLS.STEP1_SUBJECT] = "";
    newRowData[CONTACT_COLS.STEP1_SENT_MESSAGE_ID] = "";

    // Generate Connect Sales Link
    let connectSalesLinkFormula = "";
    if (email) {
      const emailParts = email.split('@');
      if (emailParts.length === 2) {
        const domain = emailParts[1];
        if (domain) {
          const connectSalesUrl = `https://sales.connect.corp.google.com/search/LEAD/TEXT/type%3ALEAD+${domain}`;
          connectSalesLinkFormula = `=HYPERLINK("${connectSalesUrl}", "${domain}")`;
        }
      }
    }
    newRowData[CONTACT_COLS.CONNECT_SALES_LINK] = connectSalesLinkFormula;

    contactsSheet.appendRow(newRowData);
    SpreadsheetApp.flush();

    logAction("Add Contact", "Added contact: " + firstName + " " + lastName + " (" + email + ") with sequence: [" + sequence + "], tags: [" + tags + "], industry: [" + industry + "]");

    // Check if this was the user's first real contact (for wizard flow)
    const allContacts = getAllContactsData();
    const realContacts = allContacts.filter(c => !c.email.includes("example.com"));
    const isFirstRealContact = realContacts.length === 1; // Just added the first one
    
    if (isFirstRealContact) {
      // Return to wizard to show next steps
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText("🎉 First contact added! Now let's send your first email."))
        .setNavigation(CardService.newNavigation().updateCard(buildFirstContactAddedCard(firstName, email)))
        .build();
    }

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("Contact added successfully!"))
      .setNavigation(CardService.newNavigation().updateCard(buildContactManagementCard({ parameters: { page: '1' } })))
      .build();
  } catch (error) {
    console.error("Error adding contact: " + error + "\n" + error.stack);
    logAction("Error", "Error adding contact: " + error.toString());
    return createNotification("Error adding contact: " + error.message);
  }
}

/* ======================== SEARCH CONTACTS ======================== */

/**
 * Searches for contacts and displays results with pagination.
 */
function searchContacts(e) {
  const searchTerm = e.formInput && e.formInput.searchTerm ? e.formInput.searchTerm.toLowerCase() : (e.parameters.searchTerm ? e.parameters.searchTerm.toLowerCase() : "");
  const page = parseInt(e && e.parameters && e.parameters.page || '1');
  const pageSize = CONFIG.PAGE_SIZE;


  // If search term is empty, just rebuild the main contact management card
  if (!searchTerm) {
    return buildContactManagementCard({ parameters: { page: '1' } });
  }

  const card = CardService.newCardBuilder();

  // Add header
  card.setHeader(CardService.newCardHeader()
      .setTitle("Search Results")
       .setSubtitle(`For: "${e.formInput ? e.formInput.searchTerm : e.parameters.searchTerm}"`) // Show original term
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/search_black_48dp.png"));

  // Get all contacts and filter
  const allContacts = getAllContactsData();
  const filteredContacts = allContacts.filter(contact => {
        const firstName = String(contact.firstName || "").toLowerCase();
        const lastName = String(contact.lastName || "").toLowerCase();
        const email = String(contact.email || "").toLowerCase();
        const company = String(contact.company || "").toLowerCase();
        const title = String(contact.title || "").toLowerCase();
        const notes = String(contact.notes || "").toLowerCase();

        return firstName.includes(searchTerm) ||
               lastName.includes(searchTerm) ||
               email.includes(searchTerm) ||
               company.includes(searchTerm) ||
               title.includes(searchTerm) ||
               notes.includes(searchTerm);
     });

  // --- Pagination Logic ---
  const totalContacts = filteredContacts.length;
  const totalPages = Math.ceil(totalContacts / pageSize);
  const startIndex = (page - 1) * pageSize;
  const endIndex = Math.min(startIndex + pageSize, totalContacts);
  const contactsToShow = filteredContacts.slice(startIndex, endIndex);

  // Add results section
  const resultsSection = CardService.newCardSection()
       .setHeader(`Found ${totalContacts} results`); // Show total count


  if (totalContacts === 0) {
    resultsSection.addWidget(CardService.newTextParagraph()
        .setText("No contacts found matching your search term."));
  } else {
    // Group contacts (optional)
    const groupedContacts = groupContactsByMarketingTitle(contactsToShow);

    // Display marketing contacts
    if (groupedContacts.marketingContacts.length > 0) {
      resultsSection.addWidget(CardService.newTextParagraph().setText("📊 Marketing Contacts:"));
      for (const contact of groupedContacts.marketingContacts) {
          displayContactSummaryWidget(resultsSection, contact);
           resultsSection.addWidget(CardService.newDivider());
      }
       if (groupedContacts.otherContacts.length > 0) {
           resultsSection.addWidget(CardService.newTextParagraph().setText(" ")); // Spacer
       }
    }

    // Display other contacts
    if (groupedContacts.otherContacts.length > 0) {
       if (groupedContacts.marketingContacts.length > 0) {
           resultsSection.addWidget(CardService.newTextParagraph().setText("📋 Other Contacts:"));
       }
        for (const contact of groupedContacts.otherContacts) {
            displayContactSummaryWidget(resultsSection, contact);
            resultsSection.addWidget(CardService.newDivider());
        }
    }

    // Add Pagination Controls - Pass the searchTerm!
     addPaginationButtons(resultsSection, page, totalPages, "searchContacts", { searchTerm: e.formInput ? e.formInput.searchTerm : e.parameters.searchTerm }); // Pass original term back
  }

  card.addSection(resultsSection);

  // Add navigation section
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
      .setText("Back to Contact Management")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("buildContactManagementCard")
           .setParameters({page: '1'}))); // Go back to page 1 of contact mgmt

  card.addSection(navSection);

  return card.build();
}

/* ======================== VIEW & EDIT CONTACT ======================== */

/**
 * Views a single contact's details, allows editing Status, Priority, Notes, and Tags.
 * (No pagination needed here as it's one contact)
 */
function viewContactCard(e) {
    const email = e.parameters.email;
    const editMode = e.parameters.editMode === 'true';
    const contact = getContactByEmail(email);

    if (!contact) {
        logAction("Error", `View Contact Error: Contact not found ${email}`);
        return createNotification("Contact not found: " + email);
    }

    // NEW: Get the auto-send setting
    const userProps = PropertiesService.getUserProperties();
    const autoSendStep1Enabled = userProps.getProperty("AUTO_SEND_STEP_1_ENABLED") === 'true';

    const card = CardService.newCardBuilder();

    // --- Header with Priority ---
    let title = contact.firstName + " " + contact.lastName;
    let priorityIndicator = "";
    if (contact.priority === "High") priorityIndicator = "🔥 ";
    else if (contact.priority === "Medium") priorityIndicator = "🟠 ";
    else if (contact.priority === "Low") priorityIndicator = "⚪ ";
    title = priorityIndicator + title;

    card.setHeader(CardService.newCardHeader()
        .setTitle(title)
        .setSubtitle(contact.email) // Add email to subtitle for clarity
        .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/person_black_48dp.png"));

    // --- Edit Button Section (next to name area) ---
    const editButtonSection = CardService.newCardSection();
    if (editMode) {
        editButtonSection.addWidget(CardService.newButtonSet()
            .addButton(CardService.newTextButton()
                .setText("❌ Cancel Edit")
                .setOnClickAction(CardService.newAction()
                    .setFunctionName("viewContactCard")
                    .setParameters({ email: contact.email, editMode: 'false' }))));
    } else {
        editButtonSection.addWidget(CardService.newButtonSet()
            .addButton(CardService.newTextButton()
                .setText("✏️ Edit")
                .setOnClickAction(CardService.newAction()
                    .setFunctionName("viewContactCard")
                    .setParameters({ email: contact.email, editMode: 'true' }))));
    }
    
    // --- Quick Link Buttons ---
    const linkedInSearchUrl = createLinkedInSearchUrl(contact);
    if (linkedInSearchUrl) {
        editButtonSection.addWidget(CardService.newButtonSet()
            .addButton(CardService.newTextButton()
                .setText("🔗 LinkedIn")
                .setOpenLink(CardService.newOpenLink().setUrl(linkedInSearchUrl))));
    }
    card.addSection(editButtonSection);
    // --- End Edit Button Section ---

    // --- Contact Details Section ---
    const detailsSection = CardService.newCardSection()
        .setHeader("Contact Details");

    if (editMode) {
        // EDITABLE MODE - Show text inputs for all editable fields
        detailsSection.addWidget(CardService.newTextInput()
            .setFieldName("editCompany")
            .setTitle("Company")
            .setValue(contact.company || ""));
        
        detailsSection.addWidget(CardService.newTextInput()
            .setFieldName("editTitle")
            .setTitle("Title")
            .setValue(contact.title || ""));
        
        // Sequence dropdown
        const editSequenceDropdown = CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.DROPDOWN)
            .setTitle("Sequence")
            .setFieldName("sequence");
        const availableSequences = getAvailableSequences();
        for (const sequence of availableSequences) {
            editSequenceDropdown.addItem(sequence, sequence, contact.sequence === sequence);
        }
        detailsSection.addWidget(editSequenceDropdown);
        
        detailsSection.addWidget(CardService.newTextInput()
            .setFieldName("editPersonalPhone")
            .setTitle("Personal Phone")
            .setValue(contact.personalPhone || ""));
        
        detailsSection.addWidget(CardService.newTextInput()
            .setFieldName("editWorkPhone")
            .setTitle("Work Phone")
            .setValue(contact.workPhone || ""));
        
        // Personal Called dropdown
        const personalCalledDropdown = CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.DROPDOWN)
            .setTitle("Personal Called")
            .setFieldName("editPersonalCalled");
        personalCalledDropdown.addItem("No", "No", contact.personalCalled !== "Yes");
        personalCalledDropdown.addItem("Yes", "Yes", contact.personalCalled === "Yes");
        detailsSection.addWidget(personalCalledDropdown);
        
        // Work Called dropdown
        const workCalledDropdown = CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.DROPDOWN)
            .setTitle("Work Called")
            .setFieldName("editWorkCalled");
        workCalledDropdown.addItem("No", "No", contact.workCalled !== "Yes");
        workCalledDropdown.addItem("Yes", "Yes", contact.workCalled === "Yes");
        detailsSection.addWidget(workCalledDropdown);
        
        // Status dropdown
        const statusDropdown = CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.DROPDOWN)
            .setTitle("Status")
            .setFieldName("newStatus");
        const statuses = ["Active", "Paused", "Completed", "Unsubscribed"];
        for (const status of statuses) {
            statusDropdown.addItem(status, status, status === contact.status);
        }
        detailsSection.addWidget(statusDropdown);
        
        // Current Step dropdown
        const stepDropdown = CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.DROPDOWN)
            .setTitle("Current Step")
            .setFieldName("editCurrentStep");
        for (let i = 1; i <= CONFIG.SEQUENCE_STEPS; i++) {
            stepDropdown.addItem("Step " + i, String(i), contact.currentStep === i);
        }
        detailsSection.addWidget(stepDropdown);
        
        // Priority dropdown
        const priorityDropdown = CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.DROPDOWN)
            .setTitle("Priority")
            .setFieldName("newPriority");
        const priorities = ["High", "Medium", "Low"];
        for (const priority of priorities) {
            let displayPriority = priority;
            if (priority === "High") displayPriority = "High 🔥";
            else if (priority === "Medium") displayPriority = "Medium 🟠";
            else if (priority === "Low") displayPriority = "Low ⚪";
            priorityDropdown.addItem(displayPriority, priority, priority === contact.priority);
        }
        detailsSection.addWidget(priorityDropdown);
        
        // Tags input
        detailsSection.addWidget(CardService.newTextInput()
            .setFieldName("tags")
            .setTitle("Tags")
            .setHint("Comma-separated (e.g., vip, follow-up)")
            .setValue(contact.tags || ""));
        
        // Industry dropdown
        const editIndustryDropdown = CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.DROPDOWN)
            .setTitle("Industry")
            .setFieldName("industry");
        editIndustryDropdown.addItem("(No Industry)", "", contact.industry === "" || !contact.industry);
        editIndustryDropdown.addItem("Ecommerce / Retail", "Ecommerce / Retail", contact.industry === "Ecommerce / Retail");
        editIndustryDropdown.addItem("SaaS / B2B Tech", "SaaS / B2B Tech", contact.industry === "SaaS / B2B Tech");
        editIndustryDropdown.addItem("Local Services", "Local Services", contact.industry === "Local Services");
        editIndustryDropdown.addItem("Healthcare & Wellness", "Healthcare & Wellness", contact.industry === "Healthcare & Wellness");
        editIndustryDropdown.addItem("Finance & FinTech", "Finance & FinTech", contact.industry === "Finance & FinTech");
        editIndustryDropdown.addItem("Travel & Tourism", "Travel & Tourism", contact.industry === "Travel & Tourism");
        editIndustryDropdown.addItem("Real Estate & Housing", "Real Estate & Housing", contact.industry === "Real Estate & Housing");
        editIndustryDropdown.addItem("Professional Services", "Professional Services", contact.industry === "Professional Services");
        editIndustryDropdown.addItem("Education & Online Learning", "Education & Online Learning", contact.industry === "Education & Online Learning");
        editIndustryDropdown.addItem("Food & Beverage CPG", "Food & Beverage CPG", contact.industry === "Food & Beverage CPG");
        editIndustryDropdown.addItem("Pets & Animal Services", "Pets & Animal Services", contact.industry === "Pets & Animal Services");
        editIndustryDropdown.addItem("Trades & Industrial / Manufacturing", "Trades & Industrial / Manufacturing", contact.industry === "Trades & Industrial / Manufacturing");
        editIndustryDropdown.addItem("Creative & Media / Agencies", "Creative & Media / Agencies", contact.industry === "Creative & Media / Agencies");
        editIndustryDropdown.addItem("Lifestyle & Personal Services", "Lifestyle & Personal Services", contact.industry === "Lifestyle & Personal Services");
        detailsSection.addWidget(editIndustryDropdown);
        
        // Notes input
        detailsSection.addWidget(CardService.newTextInput()
            .setMultiline(true)
            .setFieldName("notes")
            .setTitle("Notes")
            .setValue(contact.notes || ""));
        
        // Save button
        detailsSection.addWidget(CardService.newButtonSet()
            .addButton(CardService.newTextButton()
                .setText("💾 Save All Changes")
                .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
                .setOnClickAction(CardService.newAction()
                    .setFunctionName("saveAllContactChanges")
                    .setParameters({ email: contact.email }))));
        
    } else {
        // READ-ONLY MODE - Show KeyValue displays
        detailsSection.addWidget(CardService.newKeyValue()
            .setTopLabel("Company").setContent(contact.company || "N/A"));
        detailsSection.addWidget(CardService.newKeyValue()
            .setTopLabel("Title").setContent(contact.title || "N/A"));
        detailsSection.addWidget(CardService.newKeyValue()
            .setTopLabel("Sequence").setContent(contact.sequence || "N/A"));

        const formattedPersonalPhone = formatPhoneNumberForDisplay(contact.personalPhone);
        const formattedWorkPhone = formatPhoneNumberForDisplay(contact.workPhone);

        detailsSection.addWidget(CardService.newKeyValue()
            .setTopLabel("Personal Phone").setContent(formattedPersonalPhone));
        detailsSection.addWidget(CardService.newKeyValue()
            .setTopLabel("Work Phone").setContent(formattedWorkPhone));

        detailsSection.addWidget(CardService.newKeyValue()
            .setTopLabel("Personal Called").setContent(contact.personalCalled === "Yes" ? "Yes ✓" : "No"));
        detailsSection.addWidget(CardService.newKeyValue()
            .setTopLabel("Work Called").setContent(contact.workCalled === "Yes" ? "Yes ✓" : "No"));
        
        detailsSection.addWidget(CardService.newKeyValue()
            .setTopLabel("Sequence Status (Current)")
            .setContent(`Step ${contact.currentStep || 'N/A'} | ${contact.status || 'N/A'}`));
        
        let currentPriorityDisplay = contact.priority || "Medium";
        if (currentPriorityDisplay === "High") currentPriorityDisplay = "High Priority 🔥";
        else if (currentPriorityDisplay === "Medium") currentPriorityDisplay = "Medium Priority 🟠";
        else if (currentPriorityDisplay === "Low") currentPriorityDisplay = "Low Priority ⚪";
        detailsSection.addWidget(CardService.newKeyValue().setTopLabel("Priority (Current)").setContent(currentPriorityDisplay));

        detailsSection.addWidget(CardService.newKeyValue()
            .setTopLabel("Tags")
            .setContent(contact.tags || "N/A")
            .setMultiline(true)); 

        detailsSection.addWidget(CardService.newKeyValue()
            .setTopLabel("Industry")
            .setContent(contact.industry || "N/A")
            .setMultiline(true)); 

        // Display Step 1 Subject and Sent Message ID if available
        if (contact.step1Subject) {
          detailsSection.addWidget(CardService.newKeyValue()
              .setTopLabel("Step 1 Subject")
              .setContent(contact.step1Subject));
        }
        if (contact.step1SentMessageId) {
          detailsSection.addWidget(CardService.newKeyValue()
              .setTopLabel("Step 1 Sent ID")
              .setContent(contact.step1SentMessageId)
              .setOpenLink(CardService.newOpenLink().setUrl(`https://mail.google.com/mail/u/0/#all/${contact.step1SentMessageId}`)));
        }

        if (contact.lastEmailDate) {
            detailsSection.addWidget(CardService.newKeyValue()
                .setTopLabel("Last Email Sent").setContent(formatDate(contact.lastEmailDate)));
        }
        if (contact.nextStepDate) {
            detailsSection.addWidget(CardService.newKeyValue()
                .setTopLabel("Next Step Date").setContent(formatDate(contact.nextStepDate)));
            const readyText = contact.isReady ? "✓ Ready for email" : `⏱ Not ready until ${formatDate(contact.nextStepDate)}`;
            detailsSection.addWidget(CardService.newKeyValue()
                .setTopLabel("Email Readiness").setContent(readyText));
        }
    }

    card.addSection(detailsSection);

    // --- Actions Section (only show when NOT in edit mode) ---
    if (!editMode) {
        const actionsSection = CardService.newCardSection()
            .setHeader("Actions");

        // --- MODIFIED: Dynamic Button Text for Template Email + Preview ---
        if (contact.currentStep !== 2 && getTemplateForStep(contact.currentStep)) {
            let templateButtonText = "Preview & Send Email";
            if (contact.currentStep === 1 && autoSendStep1Enabled) {
                templateButtonText = "Preview & Send Email";
            }

            actionsSection.addWidget(CardService.newTextButton()
                .setText(templateButtonText) // Use dynamic text
                .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
                .setOnClickAction(CardService.newAction()
                    .setFunctionName("previewEmailBeforeSend") // Changed to preview first
                    .setParameters({ email: contact.email })));
        } else if (contact.currentStep !== 2) {
            actionsSection.addWidget(CardService.newTextParagraph().setText("No template found for Step " + contact.currentStep + "."));
        }
        // --- END MODIFICATION ---

        // Compose Email (Custom)
        actionsSection.addWidget(CardService.newTextButton()
            .setText("Compose Custom Email")
            .setOnClickAction(CardService.newAction()
                .setFunctionName("composeEmail")
                .setParameters({ email: contact.email })));

        // Move to Next Step (Manual)
        if (contact.status !== "Completed" && contact.status !== "Unsubscribed" && contact.currentStep < CONFIG.SEQUENCE_STEPS) {
            actionsSection.addWidget(CardService.newTextButton()
                .setText("Move to Next Step Manually")
                .setOnClickAction(CardService.newAction()
                    .setFunctionName("moveContactToNextStep")
                    .setParameters({ email: contact.email })));
        }

        // End Sequence For Company button
        if (contact.company && contact.company.trim() !== "") {
            actionsSection.addWidget(CardService.newDivider());
            actionsSection.addWidget(CardService.newTextButton()
                .setText("End Sequence For Company (" + contact.company + ")")
                .setOnClickAction(CardService.newAction()
                    .setFunctionName("endSequenceForCompany")
                    .setParameters({ email: contact.email })));
        }

        card.addSection(actionsSection);
    }

    // --- Navigation Section ---
    const navSection = CardService.newCardSection();
    navSection.addWidget(CardService.newTextButton()
        .setText("Back to Contact Management")
        .setOnClickAction(CardService.newAction()
            .setFunctionName("buildContactManagementCard")
            .setParameters({ page: '1' }))); // Assuming page '1' is a safe default
    navSection.addWidget(CardService.newTextButton()
        .setText("Back to Main Menu")
        .setOnClickAction(CardService.newAction()
            .setFunctionName("buildAddOn")));
    card.addSection(navSection);

    return card.build();
}

/**
 * Saves all editable changes from the viewContactCard to the database.
 * Handles: Company, Title, Sequence, Personal Phone, Work Phone, Personal Called, Work Called,
 * Status, Current Step, Priority, Tags, Industry, Notes
 */
function saveAllContactChanges(e) {
    const email = e.parameters.email;
    
    // Get all form values - handle both edit mode field names and legacy field names
    const newCompany = e.formInput.editCompany !== undefined ? e.formInput.editCompany : null;
    const newTitle = e.formInput.editTitle !== undefined ? e.formInput.editTitle : null;
    const newPersonalPhone = e.formInput.editPersonalPhone !== undefined ? e.formInput.editPersonalPhone : null;
    const newWorkPhone = e.formInput.editWorkPhone !== undefined ? e.formInput.editWorkPhone : null;
    const newPersonalCalled = e.formInput.editPersonalCalled || null;
    const newWorkCalled = e.formInput.editWorkCalled || null;
    const newCurrentStep = e.formInput.editCurrentStep ? parseInt(e.formInput.editCurrentStep) : null;
    const newStatus = e.formInput.newStatus;
    const newPriority = e.formInput.newPriority;
    const notes = e.formInput.notes !== undefined ? e.formInput.notes : null;
    const newTags = e.formInput.tags !== undefined ? e.formInput.tags : null;
    const newSequence = e.formInput.sequence || null;
    const newIndustry = e.formInput.industry !== undefined ? e.formInput.industry : null;

    const contact = getContactByEmail(email); 

    if (!contact) {
        logAction("Error", `Save All Changes Error: Contact not found ${email}`);
        return createNotification("Contact not found: " + email);
    }

    // Validate required fields if they were submitted
    if (newStatus && !["Active", "Paused", "Completed", "Unsubscribed"].includes(newStatus)) {
        logAction("Error", `Save All Changes Error: Invalid status value '${newStatus}' for ${email}`);
        return createNotification("Invalid status value provided. Please select a valid status.");
    }
    if (newPriority && !["High", "Medium", "Low"].includes(newPriority)) {
        logAction("Error", `Save All Changes Error: Invalid priority value '${newPriority}' for ${email}`);
        return createNotification("Invalid priority value provided. Please select a valid priority.");
    }

    const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
    if (!spreadsheetId) {
        logAction("Error", "Save All Changes Error: No database connected.");
        return createNotification("No database connected. Please connect to a database first.");
    }

    try {
        const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);

        if (!contactsSheet) {
            logAction("Error", `Save All Changes Error: Contacts sheet not found for ${email}`);
            return createNotification("Contacts sheet not found. Please refresh the add-on.");
        }

        const lastCol = Math.max(contactsSheet.getLastColumn(), Object.keys(CONTACT_COLS).length);
        const rowRange = contactsSheet.getRange(contact.rowIndex, 1, 1, lastCol);
        const rowData = rowRange.getValues()[0];

        let changesMade = false;
        let logDetails = [];

        // Company
        if (newCompany !== null) {
            const currentCompany = String(rowData[CONTACT_COLS.COMPANY] || "");
            if (currentCompany !== newCompany) {
                rowData[CONTACT_COLS.COMPANY] = newCompany;
                changesMade = true;
                logDetails.push(`Company to "${newCompany}"`);
            }
        }

        // Title
        if (newTitle !== null) {
            const currentTitle = String(rowData[CONTACT_COLS.TITLE] || "");
            if (currentTitle !== newTitle) {
                rowData[CONTACT_COLS.TITLE] = newTitle;
                changesMade = true;
                logDetails.push(`Title to "${newTitle}"`);
            }
        }

        // Personal Phone
        if (newPersonalPhone !== null) {
            const currentPersonalPhone = String(rowData[CONTACT_COLS.PERSONAL_PHONE] || "");
            if (currentPersonalPhone !== newPersonalPhone) {
                rowData[CONTACT_COLS.PERSONAL_PHONE] = newPersonalPhone;
                changesMade = true;
                logDetails.push(`Personal Phone updated`);
            }
        }

        // Work Phone
        if (newWorkPhone !== null) {
            const currentWorkPhone = String(rowData[CONTACT_COLS.WORK_PHONE] || "");
            if (currentWorkPhone !== newWorkPhone) {
                rowData[CONTACT_COLS.WORK_PHONE] = newWorkPhone;
                changesMade = true;
                logDetails.push(`Work Phone updated`);
            }
        }

        // Personal Called
        if (newPersonalCalled !== null) {
            const currentPersonalCalled = String(rowData[CONTACT_COLS.PERSONAL_CALLED] || "No");
            if (currentPersonalCalled !== newPersonalCalled) {
                rowData[CONTACT_COLS.PERSONAL_CALLED] = newPersonalCalled;
                changesMade = true;
                logDetails.push(`Personal Called to ${newPersonalCalled}`);
            }
        }

        // Work Called
        if (newWorkCalled !== null) {
            const currentWorkCalled = String(rowData[CONTACT_COLS.WORK_CALLED] || "No");
            if (currentWorkCalled !== newWorkCalled) {
                rowData[CONTACT_COLS.WORK_CALLED] = newWorkCalled;
                changesMade = true;
                logDetails.push(`Work Called to ${newWorkCalled}`);
            }
        }

        // Current Step
        if (newCurrentStep !== null && !isNaN(newCurrentStep)) {
            const currentStep = parseInt(rowData[CONTACT_COLS.CURRENT_STEP]) || 1;
            if (currentStep !== newCurrentStep) {
                rowData[CONTACT_COLS.CURRENT_STEP] = newCurrentStep;
                changesMade = true;
                logDetails.push(`Current Step to ${newCurrentStep}`);
            }
        }

        // Status
        if (newStatus && rowData[CONTACT_COLS.STATUS] !== newStatus) {
            rowData[CONTACT_COLS.STATUS] = newStatus;
            changesMade = true;
            logDetails.push(`Status to ${newStatus}`);
        }

        // Priority
        if (newPriority && rowData[CONTACT_COLS.PRIORITY] !== newPriority) {
            rowData[CONTACT_COLS.PRIORITY] = newPriority;
            changesMade = true;
            logDetails.push(`Priority to ${newPriority}`);
        }

        // Notes
        if (notes !== null) {
            const currentNotes = String(rowData[CONTACT_COLS.NOTES] || "");
            if (currentNotes !== notes) {
                rowData[CONTACT_COLS.NOTES] = notes;
                changesMade = true;
                logDetails.push("Notes updated");
            }
        }

        // Tags
        if (newTags !== null) {
            const currentTags = String(rowData[CONTACT_COLS.TAGS] || "");
            if (currentTags !== newTags) {
                rowData[CONTACT_COLS.TAGS] = newTags;
                changesMade = true;
                logDetails.push(`Tags to "${newTags}"`);
            }
        }

        // Sequence
        if (newSequence !== null) {
            const currentSequence = String(rowData[CONTACT_COLS.SEQUENCE] || "");
            if (currentSequence !== newSequence) {
                rowData[CONTACT_COLS.SEQUENCE] = newSequence;
                changesMade = true;
                logDetails.push(`Sequence to "${newSequence}"`);
            }
        }

        // Industry
        if (newIndustry !== null) {
            const currentIndustry = String(rowData[CONTACT_COLS.INDUSTRY] || "");
            if (currentIndustry !== newIndustry) {
                rowData[CONTACT_COLS.INDUSTRY] = newIndustry;
                changesMade = true;
                logDetails.push(`Industry to "${newIndustry}"`);
            }
        }

        if (changesMade) {
            rowRange.setValues([rowData]);
            SpreadsheetApp.flush();
            logAction("Update Contact Batch", `Saved changes for ${contact.firstName} ${contact.lastName} (${email}): ${logDetails.join(', ')}.`);
            return CardService.newActionResponseBuilder()
                .setNotification(CardService.newNotification().setText("Contact details updated successfully!"))
                .setNavigation(CardService.newNavigation().updateCard(viewContactCard({ parameters: { email: email, editMode: 'false' } })))
                .build();
        } else {
            return CardService.newActionResponseBuilder()
                .setNotification(CardService.newNotification().setText("No changes were detected to save."))
                .setNavigation(CardService.newNavigation().updateCard(viewContactCard({ parameters: { email: email, editMode: 'false' } })))
                .build();
        }
    } catch (error) {
        console.error("Error saving all contact changes for " + email + ": " + error + "\n" + error.stack);
        logAction("Error", `Error saving all changes for ${email}: ${error.toString()}`);
        const errorCard = viewContactCard({ parameters: { email: email, editMode: 'false' } });
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText("Error saving contact changes: " + error.message))
            .setNavigation(CardService.newNavigation().updateCard(errorCard))
            .build();
    }
}

/* ======================== HELPER FUNCTIONS ======================== */

/**
 * Creates a LinkedIn search URL for a contact
 */
function createLinkedInSearchUrl(contact) {
  if (!contact) {
    return "";
  }

  const firstName = encodeURIComponent(contact.firstName || "");
  const lastName = encodeURIComponent(contact.lastName || "");
  const company = encodeURIComponent(contact.company || "");

  let searchQuery = "";
  if (firstName && lastName) {
    searchQuery = firstName + "%20" + lastName;
    if (company) {
      searchQuery += "%20" + company;
    }
  } else if (contact.email) { // Check contact.email instead of a local 'email' variable
      searchQuery = encodeURIComponent(contact.email);
  }

  if (!searchQuery) return ""; // Cannot search if no info

  return "https://www.linkedin.com/search/results/all/?keywords=" + searchQuery;
}

