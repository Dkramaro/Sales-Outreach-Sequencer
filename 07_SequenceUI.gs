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
 * @version 2.7 - Clean dashboard, sticky action footer in contact selection views
 */

/* ======================== SEQUENCE MANAGEMENT DASHBOARD ======================== */

/**
 * Builds the Sequence Management card with clean, scannable UX:
 * - Compact stats header
 * - Single unified step list with visual hierarchy
 * - Circled numbers for quick step identification
 * - FILLED buttons only for urgent steps
 * - Fixed footer for navigation
 */
function buildSequenceManagementCard() {
  const card = CardService.newCardBuilder();
  const stats = getContactStats();

  // Calculate totals
  let totalContacts = 0;
  let totalReady = 0;

  for (let i = 1; i <= CONFIG.SEQUENCE_STEPS; i++) {
    totalContacts += stats["step" + i] || 0;
    totalReady += stats["readyForStep" + i] || 0;
  }

  // Header
  card.setHeader(CardService.newCardHeader()
      .setTitle("Sequence Management")
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/schedule_black_48dp.png"));

  // ═══════════════════════════════════════════════════════════════════
  // SECTION 1: COMPACT STATS BAR
  // ═══════════════════════════════════════════════════════════════════
  const statsSection = CardService.newCardSection();

  // Single line summary with action button
  const summaryText = totalReady > 0
      ? `📊 <b>${totalContacts}</b> Contacts · ⚠️ <b>${totalReady}</b> Ready`
      : `📊 <b>${totalContacts}</b> Contacts · ✓ All caught up`;

  statsSection.addWidget(CardService.newDecoratedText()
      .setText(summaryText)
      .setWrapText(true));

  // Primary action button - only show if there are ready contacts
  if (totalReady > 0) {
    statsSection.addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
            .setText(`📧 Process All ${totalReady} Ready`)
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(CardService.newAction()
                .setFunctionName("viewContactsReadyForEmail")
                .setParameters({ page: '1' }))));
  }

  card.addSection(statsSection);

  // ═══════════════════════════════════════════════════════════════════
  // SECTION 2: PIPELINE STEPS (single unified list with clear separation)
  // ═══════════════════════════════════════════════════════════════════
  const stepsSection = CardService.newCardSection()
      .setHeader("Pipeline Steps");

  for (let i = 1; i <= CONFIG.SEQUENCE_STEPS; i++) {
    const stepContacts = stats["step" + i] || 0;
    const readyContacts = stats["readyForStep" + i] || 0;
    const viewFunctionName = (i === 2) ? "viewStep2Contacts" : "viewContactsInStep";
    const hasReady = readyContacts > 0;
    
    // Step label for TopLabel
    const stepLabel = i === 2 ? `STEP ${i} (Manual)` : `STEP ${i}`;

    if (stepContacts === 0) {
      // Empty step - minimal display
      stepsSection.addWidget(CardService.newDecoratedText()
          .setTopLabel(stepLabel)
          .setText("<font color='#9aa0a6'>Empty</font>")
          .setButton(CardService.newTextButton()
              .setText("View")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName(viewFunctionName)
                  .setParameters({ step: i.toString(), page: '1' }))));
    } else if (hasReady) {
      // Has ready contacts - prominent display with FILLED button
      stepsSection.addWidget(CardService.newDecoratedText()
          .setTopLabel(stepLabel)
          .setText(`<b>${stepContacts}</b> Contact${stepContacts === 1 ? '' : 's'}`)
          .setBottomLabel(`⚠️ ${readyContacts} Ready for action`)
          .setButton(CardService.newTextButton()
              .setText("View ▶")
              .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
              .setOnClickAction(CardService.newAction()
                  .setFunctionName(viewFunctionName)
                  .setParameters({ step: i.toString(), page: '1' }))));
    } else {
      // Has contacts but none ready - normal display
      stepsSection.addWidget(CardService.newDecoratedText()
          .setTopLabel(stepLabel)
          .setText(`<b>${stepContacts}</b> Contact${stepContacts === 1 ? '' : 's'}`)
          .setBottomLabel("✓ All sent")
          .setButton(CardService.newTextButton()
              .setText("View")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName(viewFunctionName)
                  .setParameters({ step: i.toString(), page: '1' }))));
    }

    // Add divider between steps for spacing (not after last step)
    if (i < CONFIG.SEQUENCE_STEPS) {
      stepsSection.addWidget(CardService.newDivider());
    }
  }

  card.addSection(stepsSection);

  // ═══════════════════════════════════════════════════════════════════
  // FIXED FOOTER: Navigation
  // ═══════════════════════════════════════════════════════════════════
  const fixedFooter = CardService.newFixedFooter()
      .setPrimaryButton(CardService.newTextButton()
          .setText("← Main Menu")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("buildAddOn")));
  card.setFixedFooter(fixedFooter);

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
    // Pre-select all contacts by default (user deselects unwanted ones)
    const isSelectAll = e.parameters.selectAll !== 'false';
    let pageSize = (step === 1) ? CONFIG.PAGE_SIZE : 10;

    // NEW: Get the auto-send setting
    const userProps = PropertiesService.getUserProperties();
    const autoSendStep1Enabled = userProps.getProperty("AUTO_SEND_STEP_1_ENABLED") === 'true';

    const statusFilter = e.parameters.statusFilter || 'ready'; 
    const sequenceFilter = e.parameters.sequenceFilter || ""; 

    const card = CardService.newCardBuilder();

    card.setHeader(CardService.newCardHeader()
        .setTitle(`Step ${step} Contacts`)
        .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/checklist_rtl_black_48dp.png"));

    // Instructional header section
    const stepTypeText = step === 1 ? "first email" : `follow-up #${step - 1}`;
    const instructionSection = CardService.newCardSection();
    instructionSection.addWidget(CardService.newDecoratedText()
        .setTopLabel(`📋 SELECT CONTACTS FOR STEP ${step}`)
        .setText(`These contacts are ready for their ${stepTypeText}. Uncheck any you want to skip.`)
        .setWrapText(true));
    card.addSection(instructionSection);

    const allFilteredContacts = getVisibleContactsForStepView(step, sequenceFilter, statusFilter);
    
    const totalContacts = allFilteredContacts.length;
    const totalPages = Math.ceil(totalContacts / pageSize);
    const startIndex = (page - 1) * pageSize;
    const endIndex = Math.min(startIndex + pageSize, totalContacts);
    const contactsToShow = allFilteredContacts.slice(startIndex, endIndex);

    // Initialize SELECT_ALL_STATE for pre-selected contacts (fixes the deselection bug)
    if (isSelectAll && contactsToShow.length > 0) {
        const contactsOnPageEmails = contactsToShow.map(c => c.email);
        userProps.setProperty('SELECT_ALL_STATE', JSON.stringify({
            selectAll: true,
            contactsOnPage: contactsOnPageEmails
        }));
    }

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
        // Select all checkbox with count
        const selectAllLabel = isSelectAll 
            ? `✓ All ${contactsToShow.length} contacts selected (uncheck to deselect)`
            : `Select All (${contactsToShow.length} contacts on this page)`;
        contactsSection.addWidget(CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.CHECK_BOX)
            .setFieldName("select_all")
            .addItem(selectAllLabel, "select_all_value", isSelectAll)
            .setOnChangeAction(CardService.newAction()
                .setFunctionName("handleSelectAllGeneric")
                .setParameters({ step: step.toString(), page: page.toString() }))
        );

        // Group contacts by company with clear visual headers
        const groupedContacts = groupContactsByCompany(contactsToShow);
        
        for (const [companyName, companyContacts] of Object.entries(groupedContacts)) {
            const displayCompanyName = companyName === "No Company" ? "No Company" : companyName;
            
            // Prominent company header with DecoratedText for better visual separation
            contactsSection.addWidget(CardService.newDivider());
            contactsSection.addWidget(CardService.newDecoratedText()
                .setTopLabel("────────────────────────────────")
                .setText(`<b>🏢 ${displayCompanyName.toUpperCase()}</b>`)
                .setBottomLabel(`${companyContacts.length} contact${companyContacts.length === 1 ? '' : 's'}`)
                .setWrapText(true));
            
            // Display each contact in this company
            for (let i = 0; i < companyContacts.length; i++) {
                const contact = companyContacts[i];
                displayContactWithSelectionGrouped(contactsSection, contact, isSelectAll, 
                    { type: 'stepView', viewParams: { step: step.toString(), page: page.toString(), sequenceFilter: sequenceFilter, statusFilter: statusFilter } }, "");
                // Add spacer after each contact for visual separation
                contactsSection.addWidget(CardService.newTextParagraph().setText(" "));
            }
        }

        addPaginationButtons(contactsSection, page, totalPages, "buildSelectContactsCard", { 
            step: step.toString(), 
            sequenceFilter: sequenceFilter,
            statusFilter: statusFilter 
        });
    }
    card.addSection(contactsSection);

    // Fixed footer with action button (only when contacts exist)
    if (totalContacts > 0) {
        let buttonText = "📧 Create Drafts";
        if (step === 1 && autoSendStep1Enabled) {
            buttonText = "📧 Send Emails";
        }

        const fixedFooter = CardService.newFixedFooter()
            .setPrimaryButton(CardService.newTextButton()
                .setText(buttonText)
                .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
                .setOnClickAction(CardService.newAction()
                    .setFunctionName("emailSelectedContacts")
                    .setParameters({
                        step: step.toString(),
                        page: page.toString(),
                        sequenceFilter: sequenceFilter,
                        statusFilter: statusFilter
                    })))
            .setSecondaryButton(CardService.newTextButton()
                .setText("← Back")
                .setOnClickAction(CardService.newAction()
                    .setFunctionName("buildSequenceManagementCard")));
        card.setFixedFooter(fixedFooter);
    } else {
        // No contacts - just show back button in footer
        const fixedFooter = CardService.newFixedFooter()
            .setPrimaryButton(CardService.newTextButton()
                .setText("← Back to Sequence Management")
                .setOnClickAction(CardService.newAction()
                    .setFunctionName("buildSequenceManagementCard")));
        card.setFixedFooter(fixedFooter);
    }

    return card.build();
}

/* ======================== STEP 2 VIEW ======================== */

/**
 * Builds the card to view and select Step 2 contacts for emailing.
 */
function viewStep2Contacts(e) {
    const step = 2; 
    const page = parseInt(e?.parameters?.page || '1');
    // Pre-select all contacts by default (user deselects unwanted ones)
    const isSelectAll = e?.parameters?.selectAll !== 'false';
    const pageSize = 10;

    const statusFilter = e.parameters.statusFilter || 'ready'; 
    const sequenceFilter = e.parameters.sequenceFilter || ""; 

    const card = CardService.newCardBuilder();
    card.setHeader(CardService.newCardHeader()
        .setTitle(`Step ${step} Contacts - Follow-up Replies`)
        .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/reply_all_black_48dp.png"));

    // Instructional header section
    const instructionSection = CardService.newCardSection();
    instructionSection.addWidget(CardService.newDecoratedText()
        .setTopLabel("📋 SELECT CONTACTS FOR STEP 2")
        .setText("These contacts are ready for their follow-up reply. Uncheck any you want to skip.")
        .setWrapText(true));
    card.addSection(instructionSection); 
    
    const allFilteredContacts = getVisibleContactsForStepView(step, sequenceFilter, statusFilter);
    
    const totalContacts = allFilteredContacts.length;
    const totalPages = Math.ceil(totalContacts / pageSize);
    const startIndex = (page - 1) * pageSize;
    const endIndex = Math.min(startIndex + pageSize, totalContacts);
    const contactsToShow = allFilteredContacts.slice(startIndex, endIndex);

    // Initialize SELECT_ALL_STATE for pre-selected contacts (fixes the deselection bug)
    if (isSelectAll && contactsToShow.length > 0) {
        const userProps = PropertiesService.getUserProperties();
        const contactsOnPageEmails = contactsToShow.map(c => c.email);
        userProps.setProperty('SELECT_ALL_STATE', JSON.stringify({
            selectAll: true,
            contactsOnPage: contactsOnPageEmails
        }));
    }

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
        // Select all checkbox with count
        const step2SelectAllLabel = isSelectAll 
            ? `✓ All ${contactsToShow.length} contacts selected (uncheck to deselect)`
            : `Select All (${contactsToShow.length} contacts on this page)`;
        contactsSection.addWidget(CardService.newSelectionInput()
            .setType(CardService.SelectionInputType.CHECK_BOX)
            .setFieldName("select_all")
            .addItem(step2SelectAllLabel, "select_all_value", isSelectAll)
            .setOnChangeAction(CardService.newAction()
                .setFunctionName("handleSelectAllGeneric") 
                .setParameters({ step: step.toString(), page: page.toString() }))
        );

        // Group contacts by company with clear visual headers
        const groupedContacts = groupContactsByCompany(contactsToShow);
        
        for (const [companyName, companyContacts] of Object.entries(groupedContacts)) {
            const displayCompanyName = companyName === "No Company" ? "No Company" : companyName;
            
            // Prominent company header with DecoratedText for better visual separation
            contactsSection.addWidget(CardService.newDivider());
            contactsSection.addWidget(CardService.newDecoratedText()
                .setTopLabel("────────────────────────────────")
                .setText(`<b>🏢 ${displayCompanyName.toUpperCase()}</b>`)
                .setBottomLabel(`${companyContacts.length} contact${companyContacts.length === 1 ? '' : 's'}`)
                .setWrapText(true));
            
            for (let i = 0; i < companyContacts.length; i++) {
                const contact = companyContacts[i];
                displayContactWithSelectionGrouped(contactsSection, contact, isSelectAll, 
                    { type: 'stepView', viewParams: { step: step.toString(), page: page.toString(), sequenceFilter: sequenceFilter, statusFilter: statusFilter } }, "");
                // Add spacer after each contact for visual separation
                contactsSection.addWidget(CardService.newTextParagraph().setText(" "));
            }
        }

        addPaginationButtons(contactsSection, page, totalPages, "viewStep2Contacts", { 
            step: step.toString(), 
            sequenceFilter: sequenceFilter,
            statusFilter: statusFilter
        });
    }
    card.addSection(contactsSection);

    // Fixed footer with action button (only when contacts exist)
    if (totalContacts > 0) {
        const fixedFooter = CardService.newFixedFooter()
            .setPrimaryButton(CardService.newTextButton()
                .setText("📧 Create Drafts")
                .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
                .setOnClickAction(CardService.newAction()
                    .setFunctionName("emailSelectedContacts") 
                    .setParameters({
                        step: step.toString(),
                        page: page.toString(),
                        sequenceFilter: sequenceFilter,
                        statusFilter: statusFilter
                    })))
            .setSecondaryButton(CardService.newTextButton()
                .setText("← Back")
                .setOnClickAction(CardService.newAction()
                    .setFunctionName("buildSequenceManagementCard")));
        card.setFixedFooter(fixedFooter);
    } else {
        // No contacts - just show back button in footer
        const fixedFooter = CardService.newFixedFooter()
            .setPrimaryButton(CardService.newTextButton()
                .setText("← Back to Sequence Management")
                .setOnClickAction(CardService.newAction()
                    .setFunctionName("buildSequenceManagementCard")));
        card.setFixedFooter(fixedFooter);
    }

    return card.build();
}

/* ======================== READY CONTACTS VIEW ======================== */

/**
 * Views contacts ready for email with pagination, grouped by step.
 */
function viewContactsReadyForEmail(e) {
   const page = parseInt(e && e.parameters && e.parameters.page || '1');
   const pageSize = CONFIG.PAGE_SIZE;
   // Pre-select all contacts by default (user deselects unwanted ones)
   const isSelectAll = e?.parameters?.selectAll !== 'false';
   const isDemoMode = e?.parameters?.isDemoMode === 'true';

   // Get the auto-send setting
   const userProps = PropertiesService.getUserProperties();
   const autoSendStep1Enabled = userProps.getProperty("AUTO_SEND_STEP_1_ENABLED") === 'true';

   const card = CardService.newCardBuilder();

   const readyContacts = getContactsReadyForEmail(); 
   const totalContacts = readyContacts.length;
   const totalPages = Math.ceil(totalContacts / pageSize);
   const startIndex = (page - 1) * pageSize;
   const endIndex = Math.min(startIndex + pageSize, totalContacts);
   const contactsToShow = readyContacts.slice(startIndex, endIndex);
   
   // Check if all contacts are demo contacts (for onboarding display)
   const demoContactCount = readyContacts.filter(c => c.email && c.email.includes("example.com")).length;
   const isAllDemoContacts = demoContactCount > 0 && demoContactCount === totalContacts;

   // Dynamic header based on demo mode
   if (isAllDemoContacts || isDemoMode) {
     card.setHeader(CardService.newCardHeader()
         .setTitle("🚀 Bulk Demo Ready!")
         .setSubtitle(`${totalContacts} contacts pre-selected`)
         .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/mark_email_read_black_48dp.png"));
   } else {
     card.setHeader(CardService.newCardHeader()
         .setTitle("Contacts Ready for Email")
         .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/mark_email_read_black_48dp.png"));
   }

   // Initialize SELECT_ALL_STATE for pre-selected contacts (fixes the deselection bug)
   if (isSelectAll && contactsToShow.length > 0) {
       const contactsOnPageEmails = contactsToShow.map(c => c.email);
       userProps.setProperty('SELECT_ALL_STATE', JSON.stringify({
           selectAll: true,
           contactsOnPage: contactsOnPageEmails
       }));
   }

   // Instructional header section - special for demo mode
   const instructionSection = CardService.newCardSection();
   
   if (isAllDemoContacts || isDemoMode) {
     // Demo mode instructions - emphasize bulk processing
     instructionSection.addWidget(CardService.newDecoratedText()
         .setTopLabel("⚡ BULK PROCESSING DEMO")
         .setText(`All <b>${totalContacts} demo contacts</b> are pre-selected and ready!\n\nJust press <b>"📧 Create Drafts"</b> in the sticky footer below to watch all emails get created instantly.`)
         .setWrapText(true));
     
    instructionSection.addWidget(CardService.newTextParagraph()
        .setText(`<b>👇 Scroll down then tap "Create Drafts"</b>`));
   } else {
     // Normal instructions
     instructionSection.addWidget(CardService.newDecoratedText()
         .setTopLabel("📋 SELECT CONTACTS TO EMAIL")
         .setText("All contacts below are ready for their next email. Uncheck any you want to skip from this batch.")
         .setWrapText(true));
   }
   card.addSection(instructionSection);

   const contactsSection = CardService.newCardSection()
       .setHeader(`Ready Contacts (${startIndex + 1}-${endIndex} of ${totalContacts})`);

   if (totalContacts === 0) {
     contactsSection.addWidget(CardService.newTextParagraph()
         .setText("✓ No contacts are currently ready for their next email."));
   } else {
      // Select all checkbox with count
      const selectAllLabel = isSelectAll 
          ? `✓ All ${contactsToShow.length} contacts selected (uncheck to deselect)`
          : `Select All (${contactsToShow.length} contacts on this page)`;
      contactsSection.addWidget(CardService.newSelectionInput()
          .setType(CardService.SelectionInputType.CHECK_BOX)
          .setFieldName("select_all")
          .addItem(selectAllLabel, "select_all_value", isSelectAll)
          .setOnChangeAction(CardService.newAction()
              .setFunctionName("handleSelectAllReady") 
              .setParameters({ page: page.toString() })));

      // Group contacts by step
      const contactsByStep = groupContactsByStep(contactsToShow);
      const sortedSteps = Object.keys(contactsByStep).sort((a, b) => parseInt(a) - parseInt(b));
      
      for (const stepNum of sortedSteps) {
        const stepContacts = contactsByStep[stepNum];
        const step = parseInt(stepNum);
        const stepType = step === 1 ? "First Emails" : `Follow-up #${step - 1}`;
        
        // Enhanced step header with visual prominence
        contactsSection.addWidget(CardService.newDivider());
        contactsSection.addWidget(CardService.newDecoratedText()
            .setTopLabel(`━━━━━━━━━━ STEP ${step} ━━━━━━━━━`)
            .setText(`<b>📮 ${stepType}</b>`)
            .setBottomLabel(`${stepContacts.length} contact${stepContacts.length === 1 ? '' : 's'} ready to send`)
            .setWrapText(true));
        
        // Group contacts by company within each step for better organization
        const companiesInStep = {};
        for (const contact of stepContacts) {
          const companyKey = contact.company && contact.company.trim() ? contact.company : "No Company";
          if (!companiesInStep[companyKey]) {
            companiesInStep[companyKey] = [];
          }
          companiesInStep[companyKey].push(contact);
        }
        
        // Display contacts grouped by company
        for (const [companyName, companyContacts] of Object.entries(companiesInStep)) {
          // Prominent company header with DecoratedText for better visual separation
          contactsSection.addWidget(CardService.newDecoratedText()
              .setTopLabel("────────────────────────────────")
              .setText(`<b>🏢 ${companyName.toUpperCase()}</b>`)
              .setBottomLabel(`${companyContacts.length} contact${companyContacts.length === 1 ? '' : 's'}`)
              .setWrapText(true));
          
          // Display each contact in this company
          for (let i = 0; i < companyContacts.length; i++) {
            const contact = companyContacts[i];
            displayContactWithSelectionSimplified(contactsSection, contact, isSelectAll, { type: 'readyView', viewParams: { page: page.toString() } });
            // Add spacer after each contact for visual separation
            contactsSection.addWidget(CardService.newTextParagraph().setText(" "));
          }
        }
      }

     addPaginationButtons(contactsSection, page, totalPages, "viewContactsReadyForEmail", {});
   }

   card.addSection(contactsSection);

   // Fixed footer with action button (only when contacts exist)
   const wizardCompleted = userProps.getProperty("SKIP_WIZARD") === "true";
   
   if (totalContacts > 0) {
       const buttonText = autoSendStep1Enabled ? "📧 Send Emails" : "📧 Create Drafts";
       
       const fixedFooter = CardService.newFixedFooter()
           .setPrimaryButton(CardService.newTextButton()
               .setText(buttonText)
               .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
               .setOnClickAction(CardService.newAction()
                   .setFunctionName("emailSelectedContacts") 
                   .setParameters({ page: page.toString() })));
       
       // Add back button only if wizard is completed
       if (wizardCompleted) {
           fixedFooter.setSecondaryButton(CardService.newTextButton()
               .setText("← Back")
               .setOnClickAction(CardService.newAction()
                   .setFunctionName("buildSequenceManagementCard")));
       }
       card.setFixedFooter(fixedFooter);
   } else if (wizardCompleted) {
       // No contacts but wizard completed - show back button
       const fixedFooter = CardService.newFixedFooter()
           .setPrimaryButton(CardService.newTextButton()
               .setText("← Back to Sequence Management")
               .setOnClickAction(CardService.newAction()
                   .setFunctionName("buildSequenceManagementCard")));
       card.setFixedFooter(fixedFooter);
   }

   return card.build();
}

/**
 * Groups contacts by their current step number
 * @param {Array} contacts Array of contact objects
 * @returns {Object} Object with step numbers as keys and arrays of contacts as values
 */
function groupContactsByStep(contacts) {
  const grouped = {};
  for (const contact of contacts) {
    const step = contact.currentStep || 1;
    if (!grouped[step]) {
      grouped[step] = [];
    }
    grouped[step].push(contact);
  }
  return grouped;
}

/**
 * Simplified contact display for grouped views - shows rich metadata
 * @param {CardSection} section The section to add the widget to
 * @param {object} contact The contact object
 * @param {boolean} isChecked Whether the checkbox should be initially checked
 * @param {object} originDetails Object indicating the origin view
 */
function displayContactWithSelectionSimplified(section, contact, isChecked, originDetails) {
  const lastEmailDate = formatDate(contact.lastEmailDate);
  
  // Status indicator - only show if not ready or paused
  let statusIcon = contact.status === "Active" && !contact.isReady ? "⏱️" : 
                   contact.status === "Paused" ? "⏸️" : "";

  // Priority indicator
  let priorityIcon = contact.priority === "High" ? "🔥" : 
                     contact.priority === "Medium" ? "🟠" : "⚪";

  // Build checkbox display: priority + status + name + job title (this is now the MAIN focal point)
  let nameDisplay = priorityIcon + statusIcon + " " + contact.firstName + " " + contact.lastName;
  if (contact.title && contact.title.trim()) {
    nameDisplay += " · " + contact.title;
  }

  // Checkbox with NAME+TITLE as display text, EMAIL as value (preserves selection functionality)
  // CRITICAL: fieldName and value must stay as email-based for selection logic to work
  section.addWidget(CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setTitle("")
      .setFieldName("contact_" + contact.email.replace(/[@\\.+-]/g, "_"))
      .addItem(nameDisplay, contact.email, isChecked));

  // Determine if we're in a "ready only" view (ready status is implicit, don't show it)
  const isReadyOnlyView = originDetails.type === 'readyView' || 
                          (originDetails.viewParams && originDetails.viewParams.statusFilter === 'ready');

  // Build bottomLabel: sequence + metadata (date, ready status if applicable, industry, tags)
  let bottomParts = [];
  
  // Sequence first: "Name Sequence" format
  if (contact.sequence && contact.sequence.trim()) {
    bottomParts.push(contact.sequence + " Sequence");
  }
  
  // Date
  if (lastEmailDate && lastEmailDate !== "Never" && lastEmailDate !== "N/A") {
    bottomParts.push("📅 " + lastEmailDate);
  } else {
    bottomParts.push("📅 Never emailed");
  }
  
  // Ready status: only show when viewing ALL contacts (not ready-only view)
  // In ready-only views, everyone is ready so it's implicit
  // In all-contacts views, show "Not ready" for non-ready contacts
  if (!isReadyOnlyView && !contact.isReady) {
    bottomParts.push("⏱️ Not ready");
  }
  
  if (contact.industry && contact.industry.trim()) bottomParts.push(contact.industry);
  if (contact.tags && contact.tags.trim()) bottomParts.push("🏷️ " + contact.tags);
  
  const bottomLabelText = bottomParts.join(" · ");
  
  // DecoratedText: email (main text/same weight), sequence+metadata (bottomLabel/smaller)
  const metadataWidget = CardService.newDecoratedText()
      .setText(contact.email)
      .setBottomLabel(bottomLabelText)
      .setWrapText(true);
  section.addWidget(metadataWidget);

  // Action buttons
  const linkedInSearchUrl = createLinkedInSearchUrl(contact);
  const markCompleteParams = {
    email: contact.email,
    origin: originDetails.type,
    originParamsJson: JSON.stringify(originDetails.viewParams || {})
  };

  // ═══════════════════════════════════════════════════════════════════════════
  // END SEQUENCE BUTTON VISIBILITY RULES - DO NOT MODIFY WITHOUT UNDERSTANDING:
  // ═══════════════════════════════════════════════════════════════════════════
  // During onboarding (SKIP_WIZARD !== "true"):
  //   - User sees demo contacts first, then adds their first real contact
  //   - End Sequence button is HIDDEN for ALL contacts (demo AND real)
  //   - This prevents users from breaking the onboarding flow
  //
  // After onboarding (SKIP_WIZARD === "true"):
  //   - Set when user sends their first real email (completes onboarding)
  //   - End Sequence button is SHOWN for all real contacts
  //   - Demo contacts shouldn't exist after onboarding (auto-deleted)
  //
  // The SKIP_WIZARD property is set in:
  //   1. 09_EmailProcessing.gs - after first real email is sent
  //   2. 02_Core.gs - migration fix for existing users with email history
  // ═══════════════════════════════════════════════════════════════════════════
  const wizardCompleted = PropertiesService.getUserProperties().getProperty("SKIP_WIZARD") === "true";
  const isDemoContact = contact.email && contact.email.toLowerCase().includes("example.com");

  const buttonSet = CardService.newButtonSet();
  buttonSet.addButton(CardService.newTextButton()
      .setText("👤 View")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("viewContactCard")
          .setParameters({ email: contact.email })));
  if (linkedInSearchUrl) {
    buttonSet.addButton(CardService.newTextButton()
        .setText("in")
        .setOpenLink(CardService.newOpenLink().setUrl(linkedInSearchUrl)));
  }
  // Show End Sequence button ONLY after onboarding is complete, and not for demo contacts
  if (wizardCompleted && !isDemoContact) {
    buttonSet.addButton(CardService.newTextButton()
        .setText("❌")
        .setOnClickAction(CardService.newAction()
            .setFunctionName("markContactCompleted")
            .setParameters(markCompleteParams)));
  }
  section.addWidget(buttonSet);
}

