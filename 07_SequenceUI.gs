/**
 * FILE: Sequence Management UI
 * 
 * PURPOSE:
 * Handles sequence management user interface including step views, contact selection,
 * filtering, and navigation between sequence steps. Manages bulk and individual
 * contact progression through sequences.
 * 
 * KEY FUNCTIONS:
 * - buildSequenceManagementCard() - Main sequence dashboard
 * - viewContactsInStep() - Entry point for step views
 * - buildSelectContactsCard() - Standard step view (1,3,4,5)
 * - viewStep2Contacts() - Dedicated Step 2 view
 * - viewContactsReadyForEmail() - View all ready contacts
 * - getVisibleContactsForStepView() - Central filtering/sorting logic
 * 
 * DEPENDENCIES:
 * - 01_Config.gs: CONFIG
 * - 04_ContactData.gs: getAllContactsData, getContactsInStep, getContactsReadyForEmail, getContactStats
 * - 06_SequenceData.gs: getAvailableSequences
 * - 10_SelectionHandling.gs: displayContactWithSelection, displayContactWithSelectionGrouped
 * 
 * @version 2.3
 */

/* ======================== SEQUENCE MANAGEMENT DASHBOARD ======================== */

/**
 * Builds the Sequence Management card (Summaries don't need pagination).
 */
function buildSequenceManagementCard() {
  const card = CardService.newCardBuilder();

  // Add header
  card.setHeader(CardService.newCardHeader()
      .setTitle("Sequence Management")
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/schedule_black_48dp.png"));

  // Add sequence steps summary section
  const stepsSection = CardService.newCardSection()
      .setHeader("Sequence Steps Summary");

  const stats = getContactStats();

  // Create sequence step summaries
  for (let i = 1; i <= CONFIG.SEQUENCE_STEPS; i++) {
    const stepContacts = stats["step" + i] || 0;
    const readyContacts = stats["readyForStep" + i] || 0;
    const stepLabel = `Step ${i}` + (i === 2 ? " (Manual Follow-up)" : "");

    stepsSection.addWidget(CardService.newKeyValue()
        .setTopLabel(stepLabel)
        .setContent(`${stepContacts} Contact${stepContacts === 1 ? '' : 's'}`)
        .setBottomLabel(`${readyContacts} Ready ${i !== 2 ? 'to Send' : 'for Action'}` + (readyContacts > 0 ? " ⚠️" : " ✓")));

    // Button to view contacts in that step (paginated)
    const viewFunctionName = (i === 2) ? "viewStep2Contacts" : "viewContactsInStep";
    stepsSection.addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
            .setText(`View Step ${i} Contacts`)
            .setOnClickAction(CardService.newAction()
                .setFunctionName(viewFunctionName)
                .setParameters({ step: i.toString(), page: '1' })))); // Start on page 1

    // Add divider between steps
    if (i < CONFIG.SEQUENCE_STEPS) {
      stepsSection.addWidget(CardService.newDivider());
    }
  }

  card.addSection(stepsSection);

  // Add bulk actions section
  const bulkSection = CardService.newCardSection()
      .setHeader("Bulk Actions");

  bulkSection.addWidget(CardService.newTextButton()
      .setText("View All Contacts Ready for Email")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("viewContactsReadyForEmail")
          .setParameters({page: '1'}))); // Start on page 1

  card.addSection(bulkSection);

  // Add navigation section
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
      .setText("Back to Main Menu")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("buildAddOn")));

  card.addSection(navSection);

  return card.build();
}

/* ======================== STEP VIEW ENTRY POINT ======================== */

/**
 * Entry point for viewing contacts in a specific step (handles step 2 redirect)
 */
function viewContactsInStep(e) {
    const step = parseInt(e.parameters.step);
    if (step === 2) {
        // Step 2 has its own dedicated view function
        return viewStep2Contacts(e);
    }
    // For steps 1, 3, 4, 5 - build the standard selection card
    return buildSelectContactsCard(e); // Passes event object containing step and page
}

/* ======================== STEP CONTACT FILTERING ======================== */

/**
 * Gets contacts for a specific step, applying sequence and status filters, and sorting.
 * This centralizes the data logic for step-based views.
 * @param {number} step The sequence step number.
 * @param {string} sequenceFilter The name of the sequence to filter by (or "" for all).
 * @param {string} statusFilter The status to filter by ('all' or 'ready').
 * @returns {Array<object>} A filtered and sorted array of contact objects.
 */
function getVisibleContactsForStepView(step, sequenceFilter, statusFilter) {
  // 1. Get base data for the step
  let contacts = getContactsInStep(step);

  // 2. Apply sequence filter
  if (sequenceFilter) {
    contacts = contacts.filter(c => c.sequence === sequenceFilter);
  }

  // 3. Apply the new status filter
  if (statusFilter === 'ready') {
    contacts = contacts.filter(contact => contact.isReady && contact.status === 'Active');
  }

  // 4. Apply consistent sorting
  contacts.sort((a, b) => {
    // Primary sort: Company (with "No Company" at the end)
    const companyA = a.company || "zzzz_No Company";
    const companyB = b.company || "zzzz_No Company";
    if (companyA.toLowerCase() !== companyB.toLowerCase()) {
      return companyA.toLowerCase().localeCompare(companyB.toLowerCase());
    }

    // Secondary sort: Ready status (if not already filtered)
    if (statusFilter !== 'ready') {
        const readyA = a.isReady && a.status === 'Active';
        const readyB = b.isReady && b.status === 'Active';
        if (readyA !== readyB) return readyB - readyA; // Ready contacts first
    }

    // Tertiary sort: Priority
    const priorityOrder = { "High": 1, "Medium": 2, "Low": 3 };
    const priorityA = priorityOrder[a.priority] || 2;
    const priorityB = priorityOrder[b.priority] || 2;
    if (priorityA !== priorityB) return priorityA - priorityB;

    // Final sort: Name
    return (a.lastName + a.firstName).localeCompare(b.lastName + b.firstName);
  });

  return contacts;
}

/* ======================== STANDARD STEP VIEW (1, 3, 4, 5) ======================== */

/**
 * Builds the card for selecting contacts within a specific step (1, 3, 4, 5) with pagination.
 */
function buildSelectContactsCard(e) {
    const step = parseInt(e.parameters.step);
    const page = parseInt(e.parameters.page || '1');
    const isSelectAll = (e.parameters.selectAll === 'true'); 
    let pageSize = (step === 1) ? CONFIG.PAGE_SIZE : 10;

    // NEW: Get the auto-send setting
    const userProps = PropertiesService.getUserProperties();
    const autoSendStep1Enabled = userProps.getProperty("AUTO_SEND_STEP_1_ENABLED") === 'true';

    const statusFilter = e.parameters.statusFilter || 'all'; 
    const sequenceFilter = e.parameters.sequenceFilter || ""; 

    const card = CardService.newCardBuilder();

    card.setHeader(CardService.newCardHeader()
        .setTitle(`Step ${step} Contacts`)
        .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/checklist_rtl_black_48dp.png"));

    const allFilteredContacts = getVisibleContactsForStepView(step, sequenceFilter, statusFilter);
    
    const totalContacts = allFilteredContacts.length;
    const totalPages = Math.ceil(totalContacts / pageSize);
    const startIndex = (page - 1) * pageSize;
    const endIndex = Math.min(startIndex + pageSize, totalContacts);
    const contactsToShow = allFilteredContacts.slice(startIndex, endIndex);

    const filterSection = CardService.newCardSection()
        .setHeader("Filters");
    
    // Dropdown for Sequence
    const sequenceFilterDropdown = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setFieldName("sequenceFilter")
        .setOnChangeAction(CardService.newAction()
            .setFunctionName("applySequenceFilter")
            .setParameters({ step: step.toString(), page: '1' })); // Handler will read other filters from formInput
    
    sequenceFilterDropdown.addItem("All Sequences", "", sequenceFilter === "");
    const availableSequences = getAvailableSequences();
    for (const sequence of availableSequences) {
        sequenceFilterDropdown.addItem(sequence, sequence, sequenceFilter === sequence);
    }
    filterSection.addWidget(sequenceFilterDropdown);

    // Switch for Status
    filterSection.addWidget(CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.SWITCH)
        .setTitle("Show Ready Contacts Only")
        .setFieldName("statusFilter") 
        .addItem("", "ready", statusFilter === 'ready')
        .setOnChangeAction(CardService.newAction()
            .setFunctionName("applyStatusFilter") 
            .setParameters({ step: step.toString() })) // Handler will read other filters from formInput
    );
    card.addSection(filterSection);

    // Main contacts section - single unified section with clear company grouping
    let headerText = `Contacts (${startIndex + 1}-${endIndex} of ${totalContacts})`;
    if (statusFilter === 'ready') {
      headerText = `Ready Contacts (${startIndex + 1}-${endIndex} of ${totalContacts})`;
    }
    const contactsSection = CardService.newCardSection().setHeader(headerText);
    
    if (totalContacts === 0) {
        let noContactsMessage = `No contacts found in step ${step}`;
        if (sequenceFilter) noContactsMessage += ` for sequence "${sequenceFilter}"`;
        if (statusFilter === 'ready') noContactsMessage += ` that are ready for email`;
        noContactsMessage += ".";
        contactsSection.addWidget(CardService.newTextParagraph().setText(noContactsMessage));
    } else {
        // Select all checkbox
        contactsSection.addWidget(CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.CHECK_BOX)
            .setTitle("☑️ Select/Deselect All")
            .setFieldName("select_all")
            .addItem("Select All on This Page", "select_all_value", isSelectAll)
            .setOnChangeAction(CardService.newAction()
                .setFunctionName("handleSelectAllGeneric")
                .setParameters({ step: step.toString(), page: page.toString() }))
        );

        // Group contacts by company with clear visual headers
        const groupedContacts = groupContactsByCompany(contactsToShow);
        
        for (const [companyName, companyContacts] of Object.entries(groupedContacts)) {
            const displayCompanyName = companyName === "No Company" ? "No Company" : companyName;
            const firstContact = companyContacts[0];
            const sequenceName = firstContact.sequence || "No Sequence";
            
            // Bold company header with decorative separator
            contactsSection.addWidget(CardService.newDivider());
            contactsSection.addWidget(CardService.newTextParagraph()
                .setText(`<b>▸ ${displayCompanyName.toUpperCase()}</b>  <font color='#5f6368'>${companyContacts.length} contacts · ${sequenceName}</font>`));
            
            // Display each contact in this company
            for (const contact of companyContacts) {
                displayContactWithSelectionGrouped(contactsSection, contact, isSelectAll, 
                    { type: 'stepView', viewParams: { step: step.toString(), page: page.toString(), sequenceFilter: sequenceFilter, statusFilter: statusFilter } }, "");
            }
        }

        // Action buttons at the bottom of contacts section
        contactsSection.addWidget(CardService.newDivider());
        
        let buttonText = "📧 SEND AS DRAFT";
        if (step === 1 && autoSendStep1Enabled) {
            buttonText = "📧 SEND EMAIL";
        }

        const actionButtonSet = CardService.newButtonSet();
        actionButtonSet.addButton(CardService.newTextButton()
            .setText(buttonText)
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(CardService.newAction()
                .setFunctionName("emailSelectedContacts")
                .setParameters({
                    step: step.toString(),
                    page: page.toString(),
                    sequenceFilter: sequenceFilter,
                    statusFilter: statusFilter
                })));
        if (step < CONFIG.SEQUENCE_STEPS) { 
            actionButtonSet.addButton(CardService.newTextButton()
                .setText("→ Step " + (step + 1))
                .setOnClickAction(CardService.newAction()
                    .setFunctionName("moveSelectedToNextStep")
                    .setParameters({ step: step.toString(), page: page.toString() })));
        }
        contactsSection.addWidget(actionButtonSet);

        addPaginationButtons(contactsSection, page, totalPages, "buildSelectContactsCard", { 
            step: step.toString(), 
            sequenceFilter: sequenceFilter,
            statusFilter: statusFilter 
        });
    }
    card.addSection(contactsSection);

    const navSection = CardService.newCardSection();
    navSection.addWidget(CardService.newTextButton()
        .setText("Back to Sequence Management")
        .setOnClickAction(CardService.newAction()
            .setFunctionName("buildSequenceManagementCard")));
    card.addSection(navSection);

    return card.build();
}

/* ======================== STEP 2 VIEW ======================== */

/**
 * Builds the card to view and select Step 2 contacts for emailing.
 */
function viewStep2Contacts(e) {
    const step = 2; 
    const page = parseInt(e?.parameters?.page || '1');
    const isSelectAll = (e?.parameters?.selectAll === 'true');
    const pageSize = 10;

    const statusFilter = e.parameters.statusFilter || 'all'; 
    const sequenceFilter = e.parameters.sequenceFilter || ""; 

    const card = CardService.newCardBuilder();
    card.setHeader(CardService.newCardHeader()
        .setTitle(`Step ${step} Contacts - Follow-up Replies`)
        .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/reply_all_black_48dp.png")); 
    
    const allFilteredContacts = getVisibleContactsForStepView(step, sequenceFilter, statusFilter);
    
    const totalContacts = allFilteredContacts.length;
    const totalPages = Math.ceil(totalContacts / pageSize);
    const startIndex = (page - 1) * pageSize;
    const endIndex = Math.min(startIndex + pageSize, totalContacts);
    const contactsToShow = allFilteredContacts.slice(startIndex, endIndex);

    const filterSection = CardService.newCardSection().setHeader("Filters");

    const sequenceFilterDropdown = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setFieldName("sequenceFilter")
        .setOnChangeAction(CardService.newAction()
            .setFunctionName("applySequenceFilterStep2")
            .setParameters({ step: step.toString(), page: '1' }));

    sequenceFilterDropdown.addItem("All Sequences", "", sequenceFilter === "");
    const availableSequences = getAvailableSequences();
    for (const sequence of availableSequences) {
        sequenceFilterDropdown.addItem(sequence, sequence, sequenceFilter === sequence);
    }
    filterSection.addWidget(sequenceFilterDropdown);

    filterSection.addWidget(CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.SWITCH)
        .setTitle("Show Ready Contacts Only")
        .setFieldName("statusFilter")
        .addItem("", "ready", statusFilter === 'ready')
        .setOnChangeAction(CardService.newAction()
            .setFunctionName("applyStatusFilter") 
            .setParameters({ step: step.toString() }))
    );
    card.addSection(filterSection);

    // Main contacts section - unified design
    let headerText = `Follow-up Replies (${startIndex + 1}-${endIndex} of ${totalContacts})`;
    if (statusFilter === 'ready') {
      headerText = `Ready Contacts (${startIndex + 1}-${endIndex} of ${totalContacts})`;
    }
    const contactsSection = CardService.newCardSection().setHeader(headerText);

    if (totalContacts === 0) {
        let noContactsMessage = `No contacts currently in step ${step}`;
        if (sequenceFilter) noContactsMessage += ` for sequence "${sequenceFilter}"`;
        if (statusFilter === 'ready') noContactsMessage += ` that are ready for email`;
        noContactsMessage += ".";
        contactsSection.addWidget(CardService.newTextParagraph().setText(noContactsMessage));
    } else {
        // Select all checkbox
        contactsSection.addWidget(CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.CHECK_BOX)
            .setTitle("☑️ Select/Deselect All")
            .setFieldName("select_all")
            .addItem("Select All on This Page", "select_all_value", isSelectAll)
            .setOnChangeAction(CardService.newAction()
                .setFunctionName("handleSelectAllGeneric") 
                .setParameters({ step: step.toString(), page: page.toString() }))
        );

        // Group contacts by company with clear visual headers
        const groupedContacts = groupContactsByCompany(contactsToShow);
        
        for (const [companyName, companyContacts] of Object.entries(groupedContacts)) {
            const displayCompanyName = companyName === "No Company" ? "No Company" : companyName;
            const firstContact = companyContacts[0];
            const sequenceName = firstContact.sequence || "No Sequence";
            
            // Bold company header
            contactsSection.addWidget(CardService.newDivider());
            contactsSection.addWidget(CardService.newTextParagraph()
                .setText(`<b>▸ ${displayCompanyName.toUpperCase()}</b>  <font color='#5f6368'>${companyContacts.length} contacts · ${sequenceName}</font>`));
            
            for (const contact of companyContacts) {
                displayContactWithSelectionGrouped(contactsSection, contact, isSelectAll, 
                    { type: 'stepView', viewParams: { step: step.toString(), page: page.toString(), sequenceFilter: sequenceFilter, statusFilter: statusFilter } }, "");
            }
        }

        // Action buttons
        contactsSection.addWidget(CardService.newDivider());
        const actionButtonSet = CardService.newButtonSet();
        actionButtonSet.addButton(CardService.newTextButton()
            .setText("📧 SEND AS DRAFT REPLY")
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(CardService.newAction()
                .setFunctionName("emailSelectedContacts") 
                .setParameters({
                    step: step.toString(),
                    page: page.toString(),
                    sequenceFilter: sequenceFilter,
                    statusFilter: statusFilter
                }))
        );
        contactsSection.addWidget(actionButtonSet);

        addPaginationButtons(contactsSection, page, totalPages, "viewStep2Contacts", { 
            step: step.toString(), 
            sequenceFilter: sequenceFilter,
            statusFilter: statusFilter
        });
    }
    card.addSection(contactsSection);

    const navSection = CardService.newCardSection();
    navSection.addWidget(CardService.newTextButton()
        .setText("Back to Sequence Management")
        .setOnClickAction(CardService.newAction()
            .setFunctionName("buildSequenceManagementCard")));
    card.addSection(navSection);

    return card.build();
}

/* ======================== READY CONTACTS VIEW ======================== */

/**
 * Views contacts ready for email with pagination.
 */
function viewContactsReadyForEmail(e) {
   const page = parseInt(e && e.parameters && e.parameters.page || '1');
   const pageSize = CONFIG.PAGE_SIZE;
   const isSelectAll = (e?.parameters?.selectAll === 'true');

   // NEW: Get the auto-send setting
   const userProps = PropertiesService.getUserProperties();
   const autoSendStep1Enabled = userProps.getProperty("AUTO_SEND_STEP_1_ENABLED") === 'true';

   const card = CardService.newCardBuilder();

   card.setHeader(CardService.newCardHeader()
       .setTitle("Contacts Ready for Email")
       .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/mark_email_read_black_48dp.png")); 

   const readyContacts = getContactsReadyForEmail(); 
   const totalContacts = readyContacts.length;
   const totalPages = Math.ceil(totalContacts / pageSize);
   const startIndex = (page - 1) * pageSize;
   const endIndex = Math.min(startIndex + pageSize, totalContacts);
   const contactsToShow = readyContacts.slice(startIndex, endIndex);

   const contactsSection = CardService.newCardSection()
       .setHeader(`Ready Contacts (${startIndex + 1}-${endIndex} of ${totalContacts})`);

   if (totalContacts === 0) {
     contactsSection.addWidget(CardService.newTextParagraph()
         .setText("✓ No contacts are currently ready for their next email."));
   } else {
     contactsSection.addWidget(CardService.newTextParagraph()
         .setText("Select contacts to process their step-specific templates:"));

      contactsSection.addWidget(CardService.newSelectionInput()
          .setType(CardService.SelectionInputType.CHECK_BOX)
          .setTitle("Select/Deselect All on This Page")
          .setFieldName("select_all") 
          .addItem("Select/Deselect All", "select_all_value", isSelectAll)
          .setOnChangeAction(CardService.newAction()
              .setFunctionName("handleSelectAllReady") 
              .setParameters({ page: page.toString() })));

      contactsSection.addWidget(CardService.newDivider());

      for (const contact of contactsToShow) {
           displayContactWithSelection(contactsSection, contact, isSelectAll, { type: 'readyView', viewParams: { page: page.toString() } });
           contactsSection.addWidget(CardService.newDivider());
      }

     // --- MODIFIED: Dynamic Button Text ---
     const buttonText = autoSendStep1Enabled ? "PROCESS EMAILS (AUTO-SENDS STEP 1)" : "SEND ALL AS DRAFT";
     // --- END MODIFICATION ---

     contactsSection.addWidget(CardService.newButtonSet()
         .addButton(CardService.newTextButton()
             .setText(buttonText) // Use the dynamic text
             .setOnClickAction(CardService.newAction()
                 .setFunctionName("emailSelectedContacts") 
                       .setParameters({ page: page.toString() }))));

     addPaginationButtons(contactsSection, page, totalPages, "viewContactsReadyForEmail", {});
   }

   card.addSection(contactsSection);

   const navSection = CardService.newCardSection();
   navSection.addWidget(CardService.newTextButton()
       .setText("Back to Sequence Management")
       .setOnClickAction(CardService.newAction()
           .setFunctionName("buildSequenceManagementCard")));
   card.addSection(navSection);

   return card.build();
}

