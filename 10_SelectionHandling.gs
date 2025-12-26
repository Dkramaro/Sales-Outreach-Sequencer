/**
 * FILE: Contact Selection and Filtering
 * 
 * PURPOSE:
 * Manages contact selection state, checkbox handling, filter application,
 * and contact display widgets with selection controls. Implements workarounds
 * for CardService framework limitations.
 * 
 * KEY FUNCTIONS:
 * - handleSelectAllGeneric() - Handle select all for any step view
 * - handleSelectAllReady() - Handle select all for ready contacts view
 * - getSelectedEmailsFromInput() - Extract selected emails from form input
 * - emailToInputKey() - Generate checkbox field name from email
 * - displayContactWithSelection() - Display contact with checkbox (standard)
 * - displayContactWithSelectionGrouped() - Display contact with checkbox (grouped)
 * - applySequenceFilter() - Apply sequence filter to step view
 * - applySequenceFilterStep2() - Apply sequence filter to step 2 view
 * - applyStatusFilter() - Apply ready/all status filter
 * 
 * DEPENDENCIES:
 * - 01_Config.gs: CONFIG
 * - 04_ContactData.gs: getContactsReadyForEmail, groupContactsByCompany
 * - 07_SequenceUI.gs: getVisibleContactsForStepView, buildSelectContactsCard, viewStep2Contacts, viewContactsReadyForEmail
 * - 17_Utilities.gs: formatDate, truncateText, createLinkedInSearchUrl
 * 
 * @version 2.3
 */

/* ======================== SELECT ALL HANDLING ======================== */

/**
 * Handles the 'Select All' checkbox change for Steps 1, 2, 3, 4, 5.
 * REVISED to store state in UserProperties to work around the CardService bug.
 * CORRECTED to read all filter states from the live form input.
 */
function handleSelectAllGeneric(e) {
    // Context parameters are still fine to get from e.parameters
    const step = parseInt(e.parameters.step);
    const page = parseInt(e.parameters.page || '1');
    const isSelectAll = !!(e.formInput && e.formInput.select_all);

    // ---  THE CRITICAL FIX ---
    // Read the LIVE state of the other filters directly from e.formInput.
    // This reflects what the user actually sees on the screen.
    const sequenceFilter = e.formInput.sequenceFilter || "";
    // For a switch, the fieldName is only present in formInput if it's ON.
    const statusFilter = e.formInput.statusFilter ? 'ready' : 'all';

    console.log(`handleSelectAllGeneric (CORRECTED) - Step: ${step}, isSelectAll: ${isSelectAll}, sequenceFilter FROM INPUT: '${sequenceFilter}', statusFilter FROM INPUT: '${statusFilter}'`);

    // The rest of this logic was okay, but now it uses the correct filter values.
    const pageSize = (step === 1) ? CONFIG.PAGE_SIZE : 10;
    const visibleContacts = getVisibleContactsForStepView(step, sequenceFilter, statusFilter);
    
    const startIndex = (page - 1) * pageSize;
    const endIndex = Math.min(startIndex + pageSize, visibleContacts.length);
    const contactsOnPage = visibleContacts.slice(startIndex, endIndex);
    const contactsOnPageEmails = contactsOnPage.map(c => c.email);
    
    const userProps = PropertiesService.getUserProperties();
    userProps.setProperty('SELECT_ALL_STATE', JSON.stringify({
        selectAll: isSelectAll,
        contactsOnPage: contactsOnPageEmails
    }));

    const builderFn = (step === 2) ? viewStep2Contacts : buildSelectContactsCard;
    return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation()
            .updateCard(builderFn({
                parameters: {
                    step: step.toString(),
                    page: page.toString(),
                    selectAll: isSelectAll.toString(),
                    // Pass the correctly identified filters FORWARD to the next card.
                    sequenceFilter: sequenceFilter,
                    statusFilter: statusFilter
                }
            })))
        .build();
}

/**
 * Handles the 'Select All' checkbox change for the "Ready Contacts" view.
 * REVISED to store state in UserProperties to work around the CardService bug.
 */
function handleSelectAllReady(e) {
    const page = parseInt(e.parameters.page || '1');
    const isSelectAll = !!(e.formInput && e.formInput.select_all);

    // --- STATE MANAGEMENT FIX ---
    // Get the contacts that are currently visible on the page.
    // This logic must exactly match the data fetching/sorting/pagination in viewContactsReadyForEmail.
    const readyContacts = getContactsReadyForEmail();
    const startIndex = (page - 1) * CONFIG.PAGE_SIZE;
    const endIndex = Math.min(startIndex + CONFIG.PAGE_SIZE, readyContacts.length);
    const contactsOnPage = readyContacts.slice(startIndex, endIndex);
    const contactsOnPageEmails = contactsOnPage.map(c => c.email);

    // Store the state explicitly
    const userProps = PropertiesService.getUserProperties();
    userProps.setProperty('SELECT_ALL_STATE', JSON.stringify({
        selectAll: isSelectAll,
        contactsOnPage: contactsOnPageEmails
    }));
    // --- END STATE MANAGEMENT FIX ---

    // Now rebuild the card with the visual state change
    return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation()
            .updateCard(viewContactsReadyForEmail({
                parameters: {
                    page: page.toString(),
                    selectAll: isSelectAll.toString()
                }
            })))
        .build();
}

/* ======================== SELECTION EXTRACTION ======================== */

/**
 * Generates the expected form input key for a contact's checkbox based on its email.
 * This MUST match the fieldName generation in the displayContactWithSelection... functions.
 * @param {string} email The email address.
 * @returns {string|null} The generated key (e.g., "contact_user_example_com") or null if email is invalid.
 */
function emailToInputKey(email) {
  if (!email || typeof email !== 'string') {
    return null;
  }
  // This regex replaces the common special characters in an email address with an underscore.
  // It is specifically designed to match the regex used in the UI creation functions
  // like displayContactWithSelection and displayContactWithSelectionGrouped.
  const sanitized = email.replace(/[@.+-]/g, '_');
  return `contact_${sanitized}`;
}

/**
 * Determines which emails were selected based on form input.
 * REVISED to use a state stored in UserProperties to robustly handle the "Select All, then deselect"
 * use case, bypassing the CardService framework bug.
 *
 * @param {object} formInput The submitted form data object (e.g., e.formInput).
 * @returns {{selected: string[], cleanup: Function}} An object containing the array of selected emails and a function to clean up temporary properties.
 */
function getSelectedEmailsFromInput(formInput) {
    console.log("Entering getSelectedEmailsFromInput (Stateful Version).");
    console.log("Received formInput:", JSON.stringify(formInput));
    
    const userProps = PropertiesService.getUserProperties();
    const selectAllStateString = userProps.getProperty('SELECT_ALL_STATE');
    
    // Define a cleanup function to be called by the parent function
    const cleanup = () => userProps.deleteProperty('SELECT_ALL_STATE');

    if (selectAllStateString) {
        try {
            const state = JSON.parse(selectAllStateString);
            // This block executes if a "Select All" or "Deselect All" was the immediately preceding action.
            if (state.selectAll === true) {
                console.log("'Select All' state detected. Calculating deselected items.");
                const selectedEmails = new Set(state.contactsOnPage);

                // Now, check for deselected items.
                // A contact is deselected if it was on the page but its key is MISSING from the formInput.
                for (const contactEmail of state.contactsOnPage) {
                    const fieldName = emailToInputKey(contactEmail);
                    if (fieldName && !formInput.hasOwnProperty(fieldName)) {
                        console.log(`Deselecting ${contactEmail} because key ${fieldName} is missing.`);
                        selectedEmails.delete(contactEmail);
                    }
                }
                const finalSelectedEmails = Array.from(selectedEmails);
                console.log("Final selected emails (after deselection):", JSON.stringify(finalSelectedEmails));
                return { selected: finalSelectedEmails, cleanup: cleanup };
            }
        } catch (e) {
            console.error("Error parsing SELECT_ALL_STATE, falling back to default behavior.", e);
        }
    }
    
    // --- Fallback / Normal Individual Selection ---
    // This block runs if 'Select All' was not the last action.
    console.log("No 'Select All' state detected. Processing as individual selections.");
    const selectedEmails = new Set();
    for (const key in formInput) {
        // Find keys that match our contact checkbox pattern.
        if (key.startsWith("contact_")) {
            // The value of the checkbox input is the email itself.
            selectedEmails.add(formInput[key]);
        }
    }
    
    const finalSelectedEmails = Array.from(selectedEmails);
    console.log("Final selected emails (individual mode):", JSON.stringify(finalSelectedEmails));
    return { selected: finalSelectedEmails, cleanup: cleanup };
}

/* ======================== FILTER HANDLERS ======================== */

/**
 * Handles sequence filter change for step contacts view
 */
function applySequenceFilter(e) {
    const sequenceFilter = e.formInput.sequenceFilter || "";
    const statusFilter = e.formInput.statusFilter ? 'ready' : 'all'; // Preserve the status filter!
    const step = e.parameters.step || "1";
    
    console.log(`applySequenceFilter (CORRECTED) - Step: ${step}, New Sequence: ${sequenceFilter}, Preserving Status: '${statusFilter}'`);
    
    return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation()
            .updateCard(buildSelectContactsCard({ 
                parameters: { 
                    step: step, 
                    page: '1', 
                    sequenceFilter: sequenceFilter,
                    statusFilter: statusFilter // Pass the preserved status filter
                } 
            })))
        .build();
}

/**
 * Handles sequence filter change for step 2 contacts view
 */
function applySequenceFilterStep2(e) {
    const sequenceFilter = e.formInput.sequenceFilter || "";
    const statusFilter = e.formInput.statusFilter ? 'ready' : 'all'; // Preserve the status filter!
    const step = e.parameters.step || "2";
    
    console.log(`applySequenceFilterStep2 (CORRECTED) - Step: ${step}, New Sequence: ${sequenceFilter}, Preserving Status: '${statusFilter}'`);
    
    return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation()
            .updateCard(viewStep2Contacts({ 
                parameters: { 
                    step: step, 
                    page: '1', 
                    sequenceFilter: sequenceFilter,
                    statusFilter: statusFilter // Pass the preserved status filter
                } 
            })))
        .build();
}

/**
 * Handles the status filter (Ready/All) switch change for step views.
 */
function applyStatusFilter(e) {
    const step = e.parameters.step || "1";
    const builderFn = (step === '2') ? viewStep2Contacts : buildSelectContactsCard;

    // Read the state of BOTH filters from the live form input.
    const newStatusFilter = e.formInput.statusFilter ? 'ready' : 'all';
    const sequenceFilter = e.formInput.sequenceFilter || ""; // Preserve the sequence filter!

    console.log(`applyStatusFilter (CORRECTED) - Step: ${step}, New Status: ${newStatusFilter}, Preserving Sequence: '${sequenceFilter}'`);

    return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation()
            .updateCard(builderFn({
                parameters: {
                    step: step,
                    page: '1', 
                    sequenceFilter: sequenceFilter, // Pass the preserved sequence filter
                    statusFilter: newStatusFilter
                }
            })))
        .build();
}

/* ======================== CONTACT DISPLAY WIDGETS ======================== */

/**
 * Helper function to display a contact with a checkbox for selection in grouped format.
 * Ultra-compact design: all info visible at a glance, minimal vertical space.
 * @param {CardSection} section The section to add the widget to.
 * @param {object} contact The contact object.
 * @param {boolean} isChecked Whether the checkbox should be initially checked.
 * @param {object} originDetails Object indicating the origin view
 * @param {string} prefix Tree prefix (unused but kept for compatibility)
 */
function displayContactWithSelectionGrouped(section, contact, isChecked, originDetails, prefix = "") {
  const lastEmailDate = formatDate(contact.lastEmailDate);
  
  // Status indicator
  let statusIcon = contact.status === "Active" && contact.isReady ? "✅" : 
                   contact.status === "Active" ? "⏱️" : 
                   contact.status === "Paused" ? "⏸️" : "";

  // Priority indicator
  let priorityIcon = contact.priority === "High" ? "🔥" : 
                     contact.priority === "Medium" ? "🟠" : "⚪";

  // Build rich checkbox title: icons + name
  let checkboxTitle = priorityIcon + statusIcon + " " + contact.firstName + " " + contact.lastName;

  // Checkbox with email as the item
  section.addWidget(CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setTitle(checkboxTitle)
      .setFieldName("contact_" + contact.email.replace(/[@\\.+-]/g, "_"))
      .addItem(contact.email, contact.email, isChecked));

  // Build metadata string - only include what exists
  let infoParts = [];
  if (contact.title && contact.title.trim()) infoParts.push(contact.title);
  if (lastEmailDate && lastEmailDate !== "Never" && lastEmailDate !== "N/A") infoParts.push("📅 " + lastEmailDate);
  if (contact.tags && contact.tags.trim()) infoParts.push("🏷️ " + contact.tags);
  if (contact.industry && contact.industry.trim()) infoParts.push(contact.industry);
  
  // Only show info line if there's metadata
  if (infoParts.length > 0) {
    section.addWidget(CardService.newTextParagraph()
        .setText("<font color='#5f6368'>" + infoParts.join(" · ") + "</font>"));
  }

  // Compact action buttons with view contact link
  const linkedInSearchUrl = createLinkedInSearchUrl(contact);
  const markCompleteParams = {
    email: contact.email,
    origin: originDetails.type,
    originParamsJson: JSON.stringify(originDetails.viewParams || {})
  };

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
  buttonSet.addButton(CardService.newTextButton()
      .setText("❌")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("markContactCompleted")
          .setParameters(markCompleteParams)));
  section.addWidget(buttonSet);
}

/**
 * Helper function to display a contact with a checkbox for selection.
 * Used in "Ready for Email" view - shows full contact info since not grouped by company.
 * @param {CardSection} section The section to add the widget to.
 * @param {object} contact The contact object.
 * @param {boolean} isChecked Whether the checkbox should be initially checked.
 * @param {object} originDetails Object indicating the origin view
 */
function displayContactWithSelection(section, contact, isChecked, originDetails) {
  const lastEmailDate = formatDate(contact.lastEmailDate);
  
  // Status & Priority indicators
  let statusIcon = contact.status === "Active" && contact.isReady ? "✅" : 
                   contact.status === "Active" ? "⏱️" : 
                   contact.status === "Paused" ? "⏸️" : "";
  let priorityIcon = contact.priority === "High" ? "🔥" : 
                     contact.priority === "Medium" ? "🟠" : "⚪";

  // Checkbox title: icons + name
  let checkboxTitle = priorityIcon + statusIcon + " " + contact.firstName + " " + contact.lastName;

  section.addWidget(CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.CHECK_BOX)
      .setTitle(checkboxTitle)
      .setFieldName("contact_" + contact.email.replace(/[@\\.+-]/g, "_"))
      .addItem(contact.email, contact.email, isChecked));

  // Build metadata - company, step, sequence on first line
  let companyInfo = [];
  if (contact.company && contact.company.trim()) companyInfo.push(contact.company);
  companyInfo.push("Step " + contact.currentStep);
  if (contact.sequence && contact.sequence.trim()) companyInfo.push(contact.sequence);
  
  // Additional metadata on second line
  let metaInfo = [];
  if (contact.title && contact.title.trim()) metaInfo.push(contact.title);
  if (lastEmailDate && lastEmailDate !== "Never" && lastEmailDate !== "N/A") metaInfo.push("📅 " + lastEmailDate);
  if (contact.tags && contact.tags.trim()) metaInfo.push("🏷️ " + contact.tags);
  if (contact.industry && contact.industry.trim()) metaInfo.push(contact.industry);

  // Compact info display
  let infoText = "<b>" + companyInfo.join(" · ") + "</b>";
  if (metaInfo.length > 0) {
    infoText += "<br><font color='#5f6368'>" + metaInfo.join(" · ") + "</font>";
  }
  section.addWidget(CardService.newTextParagraph().setText(infoText));

  // Action buttons
  const linkedInSearchUrl = createLinkedInSearchUrl(contact);
  const markCompleteParams = {
    email: contact.email,
    origin: originDetails.type,
    originParamsJson: JSON.stringify(originDetails.viewParams || {})
  };

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
  buttonSet.addButton(CardService.newTextButton()
      .setText("❌")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("markContactCompleted")
          .setParameters(markCompleteParams)));
  section.addWidget(buttonSet);
  
  // Separator
  section.addWidget(CardService.newDivider());
}

/* ======================== DEPRECATED FUNCTIONS ======================== */

/**
 * Helper function to display a step 2 contact with a checkbox.
 * THIS FUNCTION IS NO LONGER USED. displayContactWithSelection is used instead.
 */
function displayStep2Contact(section, contact, isChecked) {
    console.warn("displayStep2Contact is deprecated and should not be called. Use displayContactWithSelection.");
    // displayContactWithSelection(section, contact, isChecked); // displayContactWithSelection is now used by viewStep2Contacts
}

/**
 * Marks selected Step 2 contacts as having their follow-up email drafted and moves them to Step 3.
 * This function is now OBSOLETE as Step 2 emailing is handled by `emailSelectedContacts` via `viewStep2Contacts`.
 */
function markSelectedManualFollowUps(e) {
    console.warn("markSelectedManualFollowUps is OBSOLETE. Emailing for Step 2 is handled by 'emailSelectedContacts' called from 'viewStep2Contacts'.");
    logAction("Obsolete Action", "markSelectedManualFollowUps called. This function is deprecated.");
    return createNotification("This action is obsolete. Step 2 follow-ups are now drafted like other email steps from the 'View Step 2 Contacts' screen.");
}

