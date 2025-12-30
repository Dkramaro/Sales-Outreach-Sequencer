/**
 * FILE: Call Management
 * 
 * PURPOSE:
 * Handles call tracking functionality including phone number display,
 * marking calls as completed, call note management, and call list views
 * grouped by company. Features both Multi-Contact Card View and Focus Mode
 * (single contact power dialer).
 * 
 * KEY FUNCTIONS:
 * - buildCallManagementCard() - Main call management UI router
 * - buildMultiContactView() - Card-based multi-contact display
 * - buildFocusModeView() - Single contact power dialer
 * - markPhoneNumberCalled() - Mark specific phone as called
 * - recordCallOutcome() - Save timestamped call outcome
 * - saveCallNote() - Save call notes
 * 
 * DEPENDENCIES:
 * - 01_Config.gs: CONFIG, CONTACT_COLS
 * - 04_ContactData.gs: getAllContactsData, getContactByEmail
 * - 05_ContactUI.gs: createLinkedInSearchUrl
 * - 16_SharedActions.gs: markContactCompleted
 * - 17_Utilities.gs: formatPhoneNumberForDisplay, truncateText, addPaginationButtons
 * 
 * @version 3.0 - Complete UI Redesign
 */

/* ======================== CALL MANAGEMENT ROUTER ======================== */

/**
 * Main router for Call Management - switches between Focus Mode and Multi-Contact View
 */
function buildCallManagementCard(e) {
  const focusMode = e?.parameters?.focusMode === 'true';
  
  if (focusMode) {
    return buildFocusModeView(e);
  } else {
    return buildMultiContactView(e);
  }
}

/**
 * Gets callable contacts sorted by company, priority, then name
 */
function getCallableContacts_() {
  const allContacts = getAllContactsData();
  const callableContacts = allContacts.filter(contact => {
    const hasValidPersonal = formatPhoneNumberForDisplay(contact.personalPhone) !== "N/A";
    const hasValidWork = formatPhoneNumberForDisplay(contact.workPhone) !== "N/A";
    const personalNeedsCall = hasValidPersonal && contact.personalCalled !== "Yes";
    const workNeedsCall = hasValidWork && contact.workCalled !== "Yes";
    return (contact.status === "Active" || contact.status === "Paused") && (personalNeedsCall || workNeedsCall);
  });

  // Sort: Company -> Priority -> Name
  callableContacts.sort((a, b) => {
    const companyA = a.company || "zzzz_No Company";
    const companyB = b.company || "zzzz_No Company";
    if (companyA.toLowerCase() !== companyB.toLowerCase()) {
      return companyA.toLowerCase().localeCompare(companyB.toLowerCase());
    }
    const priorityOrder = { "High": 1, "Medium": 2, "Low": 3 };
    const priorityA = priorityOrder[a.priority] || 2;
    const priorityB = priorityOrder[b.priority] || 2;
    if (priorityA !== priorityB) return priorityA - priorityB;
    return (a.lastName + a.firstName).localeCompare(b.lastName + b.firstName);
  });

  return callableContacts;
}

/* ======================== MULTI-CONTACT CARD VIEW ======================== */

/**
 * Builds the Multi-Contact Card View - each contact as a distinct visual card
 * Uses native collapsible sections for notes (no refresh on expand/collapse)
 */
function buildMultiContactView(e) {
  const page = parseInt(e?.parameters?.page || '1');
  const localPageSize = 8;

  const card = CardService.newCardBuilder();
  
  // Header with call icon
  card.setHeader(CardService.newCardHeader()
    .setTitle("📞 Call Management")
    .setSubtitle("Power through your call list")
    .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/call_black_48dp.png"));

  // Get data
  const callableContacts = getCallableContacts_();
  const totalContacts = callableContacts.length;
  const totalPages = Math.ceil(totalContacts / localPageSize);
  const startIndex = (page - 1) * localPageSize;
  const endIndex = Math.min(startIndex + localPageSize, totalContacts);
  const contactsToShowOnPage = callableContacts.slice(startIndex, endIndex);

  // ─────────────────────────────────────────────────────────────────
  // CONTROL BAR - Mode Toggle + Progress
  // ─────────────────────────────────────────────────────────────────
  const controlSection = CardService.newCardSection();
  
  if (totalContacts === 0) {
    controlSection.addWidget(CardService.newDecoratedText()
      .setText("🎉 All caught up!")
      .setBottomLabel("No contacts need calls right now")
      .setWrapText(true));
  } else {
    // Progress indicator
    controlSection.addWidget(CardService.newDecoratedText()
      .setTopLabel("📊 CALLS REMAINING")
      .setText(`${totalContacts} contacts to call`)
      .setBottomLabel(`Showing ${startIndex + 1}–${endIndex}`)
      .setWrapText(true));
    
    // Focus Mode toggle button
    controlSection.addWidget(CardService.newTextButton()
      .setText("🎯 Enter Focus Mode")
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor("#1a73e8")
      .setOnClickAction(CardService.newAction()
        .setFunctionName("buildCallManagementCard")
        .setParameters({ focusMode: "true", focusIndex: "0" })));
  }
  
  card.addSection(controlSection);

  // ─────────────────────────────────────────────────────────────────
  // CONTACT CARDS - Each contact is its own collapsible section
  // ─────────────────────────────────────────────────────────────────
  if (totalContacts > 0) {
    for (const contact of contactsToShowOnPage) {
      // Each contact gets its own section with collapsible notes
      buildContactCard_(card, contact, page.toString());
    }

    // ─────────────────────────────────────────────────────────────────
    // PAGINATION
    // ─────────────────────────────────────────────────────────────────
    if (totalPages > 1) {
      const paginationSection = CardService.newCardSection();
      addPaginationButtons(paginationSection, page, totalPages, "buildCallManagementCard", { 
        focusMode: "false"
      });
      card.addSection(paginationSection);
    }
  }

  // ─────────────────────────────────────────────────────────────────
  // NAVIGATION
  // ─────────────────────────────────────────────────────────────────
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
    .setText("← Back to Main Menu")
    .setOnClickAction(CardService.newAction().setFunctionName("buildAddOn")));
  card.addSection(navSection);

  return card.build();
}

/**
 * Builds a single contact card as its own collapsible section
 * Notes are collapsed by default - expand/collapse is client-side (no refresh)
 */
function buildContactCard_(card, contact, page) {
  const email = contact.email;
  const company = contact.company || "No Company";
  
  // Create section for this contact
  const section = CardService.newCardSection()
    .setHeader(`🏢 ${company}`);
  
  let uncollapsibleWidgetCount = 0;
  
  // Priority visual treatment
  let priorityBadge = "";
  if (contact.priority === "High") {
    priorityBadge = "🔥 HIGH";
  } else if (contact.priority === "Medium") {
    priorityBadge = "🟠 MEDIUM";
  } else if (contact.priority === "Low") {
    priorityBadge = "⚪ LOW";
  }

  // ── Contact Header with LinkedIn Button ──
  const displayName = `${contact.firstName} ${contact.lastName}`;
  const infoLine = [
    contact.title,
    `Step ${contact.currentStep || '?'}`,
    contact.tags ? `🏷️${truncateText(contact.tags, 12)}` : null
  ].filter(Boolean).join(" · ");

  const linkedInUrl = createLinkedInSearchUrl(contact);
  
  const nameWidget = CardService.newDecoratedText()
    .setTopLabel(priorityBadge)
    .setText(`<b>${displayName}</b>`)
    .setBottomLabel(infoLine)
    .setWrapText(true)
    .setOnClickAction(CardService.newAction()
      .setFunctionName("viewContactCard")
      .setParameters({ email: email }));
  
  if (linkedInUrl) {
    nameWidget.setButton(CardService.newTextButton()
      .setText("LinkedIn")
      .setOpenLink(CardService.newOpenLink().setUrl(linkedInUrl)));
  }
  
  section.addWidget(nameWidget);
  uncollapsibleWidgetCount++;

  // ── Phone Numbers with Status ──
  const formattedPersonal = formatPhoneNumberForDisplay(contact.personalPhone);
  const formattedWork = formatPhoneNumberForDisplay(contact.workPhone);

  // Personal Phone Row
  if (formattedPersonal !== "N/A") {
    const personalStatus = contact.personalCalled === "Yes";
    const personalIcon = personalStatus ? "✅" : "📱";
    
    if (personalStatus) {
      section.addWidget(CardService.newTextParagraph()
        .setText(`<font color='#34a853'>${personalIcon} ${formattedPersonal} ✓ Done</font>`));
    } else {
      section.addWidget(CardService.newDecoratedText()
        .setText(`📱 ${formattedPersonal}`)
        .setButton(CardService.newTextButton()
          .setText("✓ Called")
          .setOnClickAction(CardService.newAction()
            .setFunctionName("markPhoneNumberCalled")
            .setParameters({ 
              email: email, 
              phoneType: "personal", 
              originPage: page 
            }))));
    }
    uncollapsibleWidgetCount++;
  }

  // Work Phone Row
  if (formattedWork !== "N/A") {
    const workStatus = contact.workCalled === "Yes";
    const workIcon = workStatus ? "✅" : "☎️";
    
    if (workStatus) {
      section.addWidget(CardService.newTextParagraph()
        .setText(`<font color='#34a853'>${workIcon} ${formattedWork} ✓ Done</font>`));
    } else {
      section.addWidget(CardService.newDecoratedText()
        .setText(`☎️ ${formattedWork}`)
        .setButton(CardService.newTextButton()
          .setText("✓ Called")
          .setOnClickAction(CardService.newAction()
            .setFunctionName("markPhoneNumberCalled")
            .setParameters({ 
              email: email, 
              phoneType: "work", 
              originPage: page 
            }))));
    }
    uncollapsibleWidgetCount++;
  }

  // ── +Add Note row with END button on the right ──
  const existingNotes = contact.notes ? truncateText(contact.notes, 25) : "";
  const noteLabel = existingNotes ? `+ Add Note · "${existingNotes}"` : "+ Add Note";
  
  section.addWidget(CardService.newDecoratedText()
    .setText(`<font color='#1a73e8'>${noteLabel}</font>`)
    .setWrapText(true)
    .setButton(CardService.newTextButton()
      .setText("🛑 END")
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor("#d93025")
      .setOnClickAction(CardService.newAction()
        .setFunctionName("quickCompleteFromCallView")
        .setParameters({ email: email, page: page }))));
  uncollapsibleWidgetCount++;

  // ═══════════════════════════════════════════════════════════════
  // COLLAPSIBLE PART - Notes input and save buttons (hidden by default)
  // ═══════════════════════════════════════════════════════════════
  
  // Notes text input
  const noteFieldName = "callNotes_" + email.replace(/[@\.+-]/g, "_");
  section.addWidget(CardService.newTextInput()
    .setFieldName(noteFieldName)
    .setTitle("📝 Note")
    .setValue(contact.notes || "")
    .setMultiline(true)
    .setHint("Add your call notes here..."));

  // Save buttons - only "Save + Keep" and "Save + END"
  const noteButtonRow = CardService.newButtonSet();
  noteButtonRow.addButton(CardService.newTextButton()
    .setText("Save + Keep")
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor("#1a73e8")
    .setOnClickAction(CardService.newAction()
      .setFunctionName("saveNoteKeepCadence")
      .setParameters({ email: email, page: page })));
  noteButtonRow.addButton(CardService.newTextButton()
    .setText("Save + END")
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor("#d93025")
    .setOnClickAction(CardService.newAction()
      .setFunctionName("saveNoteAndEndSequence")
      .setParameters({ email: email, page: page })));
  section.addWidget(noteButtonRow);

  // Make section collapsible - notes hidden by default
  section.setCollapsible(true);
  section.setNumUncollapsibleWidgets(uncollapsibleWidgetCount);

  card.addSection(section);
}

/* ======================== FOCUS MODE - POWER DIALER ======================== */

/**
 * Builds Focus Mode view - single contact power dialer
 */
function buildFocusModeView(e) {
  const focusIndex = parseInt(e?.parameters?.focusIndex || '0');
  const stayOnContact = e?.parameters?.stayOnContact || ''; // Email to stay on even if calls done
  
  const card = CardService.newCardBuilder();
  
  // Get all callable contacts
  const callableContacts = getCallableContacts_();
  let totalContacts = callableContacts.length;
  
  // If stayOnContact is set, we need to show that contact even if all calls are done
  let contact = null;
  let safeIndex = 0;
  let isStayingOnContact = false;
  
  if (stayOnContact) {
    // User just marked last call - stay on this contact for notes/outcome logging
    contact = getContactByEmail(stayOnContact);
    if (contact) {
      isStayingOnContact = true;
      safeIndex = focusIndex;
      // Don't count this in totalContacts since they've completed calls
    }
  }
  
  // Header
  card.setHeader(CardService.newCardHeader()
    .setTitle("🎯 Focus Mode")
    .setSubtitle("One contact at a time")
    .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/call_black_48dp.png"));

  // ─────────────────────────────────────────────────────────────────
  // CONTROL BAR
  // ─────────────────────────────────────────────────────────────────
  const controlSection = CardService.newCardSection();
  
  if (totalContacts === 0 && !isStayingOnContact) {
    controlSection.addWidget(CardService.newDecoratedText()
      .setText("🎉 All caught up!")
      .setBottomLabel("No contacts need calls right now")
      .setWrapText(true));
    
    controlSection.addWidget(CardService.newTextButton()
      .setText("← Exit Focus Mode")
      .setOnClickAction(CardService.newAction()
        .setFunctionName("buildCallManagementCard")
        .setParameters({ focusMode: "false" })));
    
    card.addSection(controlSection);
    return card.build();
  }

  // Get contact from callable list if not staying on a specific one
  if (!isStayingOnContact) {
    safeIndex = Math.min(Math.max(0, focusIndex), totalContacts - 1);
    contact = callableContacts[safeIndex];
  }
  
  // Simple text counter
  if (isStayingOnContact) {
    controlSection.addWidget(CardService.newDecoratedText()
      .setTopLabel("✅ ALL CALLS MADE")
      .setText("Log outcome before moving on")
      .setWrapText(true));
  } else {
    controlSection.addWidget(CardService.newDecoratedText()
      .setTopLabel(`CONTACT ${safeIndex + 1} OF ${totalContacts}`)
      .setText(`${totalContacts - safeIndex} remaining`)
      .setWrapText(true));
  }

  // Exit button only - End Sequence moved to contact section with LinkedIn
  controlSection.addWidget(CardService.newTextButton()
    .setText("✕ Exit")
    .setOnClickAction(CardService.newAction()
      .setFunctionName("buildCallManagementCard")
      .setParameters({ focusMode: "false" })));
  
  card.addSection(controlSection);

  // ─────────────────────────────────────────────────────────────────
  // CONTACT CARD - Full Details
  // ─────────────────────────────────────────────────────────────────
  const contactSection = CardService.newCardSection();
  
  // Priority badge
  let priorityBadge = "🟠 MEDIUM PRIORITY";
  let priorityColor = "#ea8600";
  if (contact.priority === "High") {
    priorityBadge = "🔥 HIGH PRIORITY";
    priorityColor = "#d93025";
  } else if (contact.priority === "Low") {
    priorityBadge = "⚪ LOW PRIORITY";
    priorityColor = "#80868b";
  }

  contactSection.addWidget(CardService.newTextParagraph()
    .setText(`<font color='${priorityColor}'><b>${priorityBadge}</b></font>`));

  // Contact name - large and prominent
  const displayName = `${contact.firstName} ${contact.lastName}`;
  contactSection.addWidget(CardService.newDecoratedText()
    .setText(`<b style="font-size:18px">${displayName}</b>`)
    .setBottomLabel(`${contact.title || ''} · ${contact.company || 'No Company'}`)
    .setWrapText(true)
    .setOnClickAction(CardService.newAction()
      .setFunctionName("viewContactCard")
      .setParameters({ email: contact.email })));

  // Step and sequence info
  contactSection.addWidget(CardService.newTextParagraph()
    .setText(`<font color='#5f6368'>📧 ${contact.sequence || 'No Sequence'} · Step ${contact.currentStep || '?'}</font>`));

  // LinkedIn + End Cadence row
  const linkedInUrl = createLinkedInSearchUrl(contact);
  const actionRow = CardService.newButtonSet();
  
  if (linkedInUrl) {
    actionRow.addButton(CardService.newTextButton()
      .setText("🔗 LinkedIn")
      .setOpenLink(CardService.newOpenLink().setUrl(linkedInUrl)));
  }
  
  actionRow.addButton(CardService.newTextButton()
    .setText("❌ Cadence")
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor("#d93025")
    .setOnClickAction(CardService.newAction()
      .setFunctionName("endSequenceFocus")
      .setParameters({ 
        email: contact.email, 
        focusIndex: safeIndex.toString() 
      })));
  
  contactSection.addWidget(actionRow);

  card.addSection(contactSection);

  // ─────────────────────────────────────────────────────────────────
  // PHONE NUMBERS - Large & Tappable
  // ─────────────────────────────────────────────────────────────────
  const phoneSection = CardService.newCardSection()
    .setHeader("📞 Phone Numbers");

  const formattedPersonal = formatPhoneNumberForDisplay(contact.personalPhone);
  const formattedWork = formatPhoneNumberForDisplay(contact.workPhone);
  const personalNeedsCall = formattedPersonal !== "N/A" && contact.personalCalled !== "Yes";
  const workNeedsCall = formattedWork !== "N/A" && contact.workCalled !== "Yes";

  // Personal Phone
  if (formattedPersonal !== "N/A") {
    if (contact.personalCalled === "Yes") {
      phoneSection.addWidget(CardService.newDecoratedText()
        .setTopLabel("📱 PERSONAL")
        .setText(`<font color='#34a853'>✅ ${formattedPersonal}</font>`)
        .setBottomLabel("Already called")
        .setWrapText(true));
    } else {
      phoneSection.addWidget(CardService.newDecoratedText()
        .setTopLabel("📱 PERSONAL")
        .setText(`<b>${formattedPersonal}</b>`)
        .setWrapText(true)
        .setButton(CardService.newTextButton()
          .setText("✓ Mark Called")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setBackgroundColor("#1a73e8")
          .setOnClickAction(CardService.newAction()
            .setFunctionName("markPhoneNumberCalledFocus")
            .setParameters({ 
              email: contact.email, 
              phoneType: "personal", 
              focusIndex: safeIndex.toString() 
            }))));
    }
  }

  // Work Phone
  if (formattedWork !== "N/A") {
    if (contact.workCalled === "Yes") {
      phoneSection.addWidget(CardService.newDecoratedText()
        .setTopLabel("☎️ WORK")
        .setText(`<font color='#34a853'>✅ ${formattedWork}</font>`)
        .setBottomLabel("Already called")
        .setWrapText(true));
    } else {
      phoneSection.addWidget(CardService.newDecoratedText()
        .setTopLabel("☎️ WORK")
        .setText(`<b>${formattedWork}</b>`)
        .setWrapText(true)
        .setButton(CardService.newTextButton()
          .setText("✓ Mark Called")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setBackgroundColor("#1a73e8")
          .setOnClickAction(CardService.newAction()
            .setFunctionName("markPhoneNumberCalledFocus")
            .setParameters({ 
              email: contact.email, 
              phoneType: "work", 
              focusIndex: safeIndex.toString() 
            }))));
    }
  }

  if (formattedPersonal === "N/A" && formattedWork === "N/A") {
    phoneSection.addWidget(CardService.newTextParagraph()
      .setText("<font color='#80868b'><i>No phone numbers on file</i></font>"));
  }

  card.addSection(phoneSection);

  // ─────────────────────────────────────────────────────────────────
  // NOTES - ABOVE Call Outcomes (write notes first, then log outcome)
  // ─────────────────────────────────────────────────────────────────
  const notesSection = CardService.newCardSection()
    .setHeader("📝 Notes");

  // Show existing notes if any
  if (contact.notes) {
    notesSection.addWidget(CardService.newTextParagraph()
      .setText(`<font color='#5f6368'><i>Existing:</i></font><br>${contact.notes.replace(/\n/g, '<br>')}`));
  }

  const noteFieldName = "focusCallNotes_" + contact.email.replace(/[@\.+-]/g, "_");
  notesSection.addWidget(CardService.newTextInput()
    .setFieldName(noteFieldName)
    .setTitle("Add notes")
    .setMultiline(true)
    .setHint("Type notes here, they auto-save with outcomes below"));

  card.addSection(notesSection);

  // ─────────────────────────────────────────────────────────────────
  // CALL OUTCOME - Quick Buttons (auto-saves any typed notes)
  // ─────────────────────────────────────────────────────────────────
  const outcomeSection = CardService.newCardSection()
    .setHeader("📋 Log Outcome");

  outcomeSection.addWidget(CardService.newTextParagraph()
    .setText("<font color='#5f6368'>Auto-saves notes above:</font>"));

  // Row 1: Spoke (completes), No Answer, VM, Callback
  const outcomeRow1 = CardService.newButtonSet();
  outcomeRow1.addButton(CardService.newTextButton()
    .setText("✅ Spoke")
    .setOnClickAction(CardService.newAction()
      .setFunctionName("recordCallOutcomeAndComplete")
      .setParameters({ 
        email: contact.email, 
        outcome: "Spoke",
        focusIndex: safeIndex.toString()
      })));
  outcomeRow1.addButton(CardService.newTextButton()
    .setText("📵 No Answer")
    .setOnClickAction(CardService.newAction()
      .setFunctionName("recordCallOutcome")
      .setParameters({ 
        email: contact.email, 
        outcome: "No Answer",
        focusIndex: safeIndex.toString()
      })));
  outcomeRow1.addButton(CardService.newTextButton()
    .setText("📝 VM")
    .setOnClickAction(CardService.newAction()
      .setFunctionName("recordCallOutcome")
      .setParameters({ 
        email: contact.email, 
        outcome: "Left VM",
        focusIndex: safeIndex.toString()
      })));
  outcomeRow1.addButton(CardService.newTextButton()
    .setText("🗓 Callback")
    .setOnClickAction(CardService.newAction()
      .setFunctionName("recordCallOutcome")
      .setParameters({ 
        email: contact.email, 
        outcome: "Callback",
        focusIndex: safeIndex.toString()
      })));
  outcomeSection.addWidget(outcomeRow1);

  // Row 2: No Interest (unsubscribes), Bad #
  const outcomeRow2 = CardService.newButtonSet();
  outcomeRow2.addButton(CardService.newTextButton()
    .setText("❌ No Interest")
    .setOnClickAction(CardService.newAction()
      .setFunctionName("recordCallOutcomeAndUnsubscribe")
      .setParameters({ 
        email: contact.email, 
        outcome: "No Interest",
        focusIndex: safeIndex.toString()
      })));
  outcomeRow2.addButton(CardService.newTextButton()
    .setText("🚫 Bad #")
    .setOnClickAction(CardService.newAction()
      .setFunctionName("recordCallOutcome")
      .setParameters({ 
        email: contact.email, 
        outcome: "Bad Number",
        focusIndex: safeIndex.toString()
      })));
  outcomeSection.addWidget(outcomeRow2);

  card.addSection(outcomeSection);

  // ─────────────────────────────────────────────────────────────────
  // NAVIGATION - Next Call
  // ─────────────────────────────────────────────────────────────────
  const navSection = CardService.newCardSection();

  // Next Call button (always show, wraps to beginning if at end)
  navSection.addWidget(CardService.newTextButton()
    .setText("Next Call →")
    .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
    .setBackgroundColor("#1a73e8")
    .setOnClickAction(CardService.newAction()
      .setFunctionName("nextCallFocus")
      .setParameters({ 
        email: contact.email,
        focusIndex: safeIndex.toString(),
        totalContacts: totalContacts.toString()
      })));

  card.addSection(navSection);

  return card.build();
}

/* ======================== FOCUS MODE ACTIONS ======================== */

/**
 * Records a call outcome with timestamp, auto-saves any typed notes, returns to focus mode
 */
function recordCallOutcome(e) {
  const email = e.parameters.email;
  const outcome = e.parameters.outcome;
  const focusIndex = e.parameters.focusIndex || '0';

  const contact = getContactByEmail(email);
  if (!contact) {
    return createNotification("Contact not found: " + email);
  }

  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    return createNotification("No database connected.");
  }

  // Get any typed notes from the form
  const noteFieldName = "focusCallNotes_" + email.replace(/[@\.+-]/g, "_");
  const typedNotes = e.formInput && e.formInput[noteFieldName] ? e.formInput[noteFieldName].trim() : "";

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);

    if (!contactsSheet) {
      return createNotification("Contacts sheet not found.");
    }

    // Format timestamp as dd/mm/yy
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yy");
    
    // Build note: [timestamp] Outcome + any typed notes
    let outcomeNote = `[${timestamp}] ${outcome}`;
    if (typedNotes) {
      outcomeNote += ` - ${typedNotes}`;
    }
    
    // Prepend to existing notes
    const existingNotes = contact.notes || "";
    const newNotes = existingNotes ? `${outcomeNote}\n${existingNotes}` : outcomeNote;

    contactsSheet.getRange(contact.rowIndex, CONTACT_COLS.NOTES + 1).setValue(newNotes);
    SpreadsheetApp.flush();
    removeContactFromCache(contact.email);

    logAction("Call Outcome", `Recorded "${outcome}" for ${contact.firstName} ${contact.lastName}`);

    // Refresh Focus Mode view
    const refreshedCard = buildFocusModeView({ 
      parameters: { 
        focusMode: "true", 
        focusIndex: focusIndex 
      } 
    });

  return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(`✓ Logged: ${outcome}`))
      .setNavigation(CardService.newNavigation().updateCard(refreshedCard))
    .build();

  } catch (error) {
    console.error("Error recording call outcome: " + error);
    return createNotification("Error: " + error.message);
  }
}

/**
 * Records "No Interest" outcome, marks contact complete, and moves to next
 */
function recordCallOutcomeAndComplete(e) {
  const email = e.parameters.email;
  const outcome = e.parameters.outcome;
  const focusIndex = parseInt(e.parameters.focusIndex || '0');

  const contact = getContactByEmail(email);
  if (!contact) {
    return createNotification("Contact not found: " + email);
  }

  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    return createNotification("No database connected.");
  }

  // Get any typed notes from the form
  const noteFieldName = "focusCallNotes_" + email.replace(/[@\.+-]/g, "_");
  const typedNotes = e.formInput && e.formInput[noteFieldName] ? e.formInput[noteFieldName].trim() : "";

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);

    if (!contactsSheet) {
      return createNotification("Contacts sheet not found.");
    }

    // Format timestamp as dd/mm/yy
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yy");
    
    // Build note with outcome + typed notes
    let outcomeNote = `[${timestamp}] ${outcome}`;
    if (typedNotes) {
      outcomeNote += ` - ${typedNotes}`;
    }
    
    const existingNotes = contact.notes || "";
    const newNotes = existingNotes ? `${outcomeNote}\n${existingNotes}` : outcomeNote;

    // Save notes
    contactsSheet.getRange(contact.rowIndex, CONTACT_COLS.NOTES + 1).setValue(newNotes);
    
    // Mark as completed
    contactsSheet.getRange(contact.rowIndex, CONTACT_COLS.STATUS + 1).setValue("Completed");
    
    // Set completion date
    const completionDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    contactsSheet.getRange(contact.rowIndex, CONTACT_COLS.COMPLETION_DATE + 1).setValue(completionDate);
    
    SpreadsheetApp.flush();
    removeContactFromCache(contact.email);

    logAction("Spoke/Complete", `${contact.firstName} ${contact.lastName} - marked complete`);

    // Stay at same index (list will refresh, showing next contact)
    const refreshedCard = buildFocusModeView({ 
      parameters: { 
        focusMode: "true", 
        focusIndex: focusIndex.toString() 
      } 
    });

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(`✓ ${contact.firstName} completed, moving to next...`))
      .setNavigation(CardService.newNavigation().updateCard(refreshedCard))
      .build();

  } catch (error) {
    console.error("Error completing contact: " + error);
    return createNotification("Error: " + error.message);
  }
}

/**
 * Records "No Interest" outcome and marks contact as Unsubscribed
 */
function recordCallOutcomeAndUnsubscribe(e) {
  const email = e.parameters.email;
  const outcome = e.parameters.outcome;
  const focusIndex = parseInt(e.parameters.focusIndex || '0');

  const contact = getContactByEmail(email);
  if (!contact) {
    return createNotification("Contact not found: " + email);
  }

  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    return createNotification("No database connected.");
  }

  // Get any typed notes from the form
  const noteFieldName = "focusCallNotes_" + email.replace(/[@\.+-]/g, "_");
  const typedNotes = e.formInput && e.formInput[noteFieldName] ? e.formInput[noteFieldName].trim() : "";

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);

    if (!contactsSheet) {
      return createNotification("Contacts sheet not found.");
    }

    // Format timestamp as dd/mm/yy
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yy");
    
    // Build note with outcome + typed notes
    let outcomeNote = `[${timestamp}] ${outcome} - UNSUBSCRIBED`;
    if (typedNotes) {
      outcomeNote += ` - ${typedNotes}`;
    }
    
    const existingNotes = contact.notes || "";
    const newNotes = existingNotes ? `${outcomeNote}\n${existingNotes}` : outcomeNote;

    // Save notes
    contactsSheet.getRange(contact.rowIndex, CONTACT_COLS.NOTES + 1).setValue(newNotes);
    
    // Mark as Unsubscribed
    contactsSheet.getRange(contact.rowIndex, CONTACT_COLS.STATUS + 1).setValue("Unsubscribed");
    
    // Set completion date
    const completionDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    contactsSheet.getRange(contact.rowIndex, CONTACT_COLS.COMPLETION_DATE + 1).setValue(completionDate);
    
    SpreadsheetApp.flush();
    removeContactFromCache(contact.email);

    logAction("No Interest", `${contact.firstName} ${contact.lastName} - marked Unsubscribed`);

    // Stay at same index (list will refresh, showing next contact)
    const refreshedCard = buildFocusModeView({ 
      parameters: { 
        focusMode: "true", 
        focusIndex: focusIndex.toString() 
      } 
    });

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(`✓ ${contact.firstName} unsubscribed, moving to next...`))
      .setNavigation(CardService.newNavigation().updateCard(refreshedCard))
      .build();

  } catch (error) {
    console.error("Error unsubscribing contact: " + error);
    return createNotification("Error: " + error.message);
  }
}

/**
 * Move to next contact without completing current one
 */
function nextCallFocus(e) {
  const focusIndex = parseInt(e.parameters.focusIndex || '0');
  
  // Get fresh count of callable contacts (in case some were completed)
  const callableContacts = getCallableContacts_();
  const totalContacts = callableContacts.length;
  
  // If no contacts left, just refresh (will show "all caught up")
  if (totalContacts === 0) {
    const refreshedCard = buildFocusModeView({ 
      parameters: { focusMode: "true", focusIndex: "0" } 
    });
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation().updateCard(refreshedCard))
      .build();
  }
  
  // Move to next, or wrap to 0 if at end
  // Keep same index if contact was removed (list shifted), otherwise increment
  let nextIndex = focusIndex;
  if (focusIndex < totalContacts - 1) {
    nextIndex = focusIndex + 1;
  } else if (focusIndex >= totalContacts) {
    nextIndex = 0; // Wrap if we're past the end
  }
  // If focusIndex == totalContacts - 1, stay at same index (last one) - list may have shifted
  
  const refreshedCard = buildFocusModeView({ 
    parameters: { 
      focusMode: "true", 
      focusIndex: nextIndex.toString() 
    } 
  });

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(refreshedCard))
    .build();
}

/**
 * End sequence for contact (mark complete) and move to next
 */
function endSequenceFocus(e) {
  const email = e.parameters.email;
  const focusIndex = parseInt(e.parameters.focusIndex || '0');

  const contact = getContactByEmail(email);
  if (!contact) {
    return createNotification("Contact not found: " + email);
  }

  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    return createNotification("No database connected.");
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);

    if (!contactsSheet) {
      return createNotification("Contacts sheet not found.");
    }

    // Mark as completed
    contactsSheet.getRange(contact.rowIndex, CONTACT_COLS.STATUS + 1).setValue("Completed");
    
    // Set completion date
    const completionDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    contactsSheet.getRange(contact.rowIndex, CONTACT_COLS.COMPLETION_DATE + 1).setValue(completionDate);
    
    SpreadsheetApp.flush();
    removeContactFromCache(contact.email);

    logAction("Sequence Ended", `${contact.firstName} ${contact.lastName} - sequence ended from Focus Mode`);

    // Stay at same index (list will refresh, showing next contact)
    const refreshedCard = buildFocusModeView({ 
      parameters: { 
        focusMode: "true", 
        focusIndex: focusIndex.toString() 
      } 
    });

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(`✓ ${contact.firstName}'s sequence ended`))
      .setNavigation(CardService.newNavigation().updateCard(refreshedCard))
      .build();

  } catch (error) {
    console.error("Error ending sequence: " + error);
    return createNotification("Error: " + error.message);
  }
}

/**
 * Marks phone called and stays in Focus Mode
 */
function markPhoneNumberCalledFocus(e) {
  const email = e.parameters.email;
  const phoneType = e.parameters.phoneType;
  const focusIndex = e.parameters.focusIndex || '0';

  const contact = getContactByEmail(email);
  if (!contact) {
    return createNotification("Contact not found: " + email);
  }

  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    return createNotification("No database connected.");
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);

    if (!contactsSheet) {
      return createNotification("Contacts sheet not found.");
    }

    const columnToUpdate = (phoneType === "personal") ? CONTACT_COLS.PERSONAL_CALLED + 1 : CONTACT_COLS.WORK_CALLED + 1;
    contactsSheet.getRange(contact.rowIndex, columnToUpdate).setValue("Yes");
    SpreadsheetApp.flush();
    removeContactFromCache(contact.email);

    logAction("Phone Call", `Marked ${phoneType} phone called for ${contact.firstName} ${contact.lastName}`);

    // Check if all calls are now done for this contact
    const updatedContact = getContactByEmail(email);
    const formattedPersonal = formatPhoneNumberForDisplay(updatedContact.personalPhone);
    const formattedWork = formatPhoneNumberForDisplay(updatedContact.workPhone);
    const personalDone = formattedPersonal === "N/A" || updatedContact.personalCalled === "Yes";
    const workDone = formattedWork === "N/A" || updatedContact.workCalled === "Yes";
    const allCallsDone = personalDone && workDone;

    // Build refresh parameters - stay on contact if all calls done
    const refreshParams = { 
      focusMode: "true", 
      focusIndex: focusIndex 
    };
    
    if (allCallsDone) {
      // Stay on this contact so user can log outcome/notes
      refreshParams.stayOnContact = email;
    }

    const refreshedCard = buildFocusModeView({ parameters: refreshParams });

    let notificationText = `✓ ${phoneType.charAt(0).toUpperCase() + phoneType.slice(1)} marked as called`;
    if (allCallsDone) {
      notificationText += " · All calls complete!";
    }

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(notificationText))
      .setNavigation(CardService.newNavigation().updateCard(refreshedCard))
      .build();

  } catch (error) {
    console.error("Error marking phone called: " + error);
    return createNotification("Error: " + error.message);
  }
}

/**
 * Save notes from Focus Mode
 */
function saveCallNoteFocus(e) {
  const email = e.parameters.email;
  const focusIndex = e.parameters.focusIndex || '0';

  const noteFieldName = "focusCallNotes_" + email.replace(/[@\.+-]/g, "_");
  const additionalNotes = e.formInput && e.formInput[noteFieldName] ? e.formInput[noteFieldName] : "";

  if (!additionalNotes.trim()) {
    return createNotification("No notes to save.");
  }

  const contact = getContactByEmail(email);
  if (!contact) {
    return createNotification("Contact not found: " + email);
  }

  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    return createNotification("No database connected.");
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);

    if (!contactsSheet) {
      return createNotification("Contacts sheet not found.");
    }

    // Append to existing notes
    const existingNotes = contact.notes || "";
    const newNotes = existingNotes ? `${existingNotes}\n${additionalNotes}` : additionalNotes;

    contactsSheet.getRange(contact.rowIndex, CONTACT_COLS.NOTES + 1).setValue(newNotes);
    SpreadsheetApp.flush();
    removeContactFromCache(contact.email);

    logAction("Call Note", `Saved notes for ${contact.firstName} ${contact.lastName}`);

    // Refresh Focus Mode view
    const refreshedCard = buildFocusModeView({ 
      parameters: { 
        focusMode: "true", 
        focusIndex: focusIndex 
      } 
    });

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("✓ Notes saved"))
      .setNavigation(CardService.newNavigation().updateCard(refreshedCard))
      .build();

  } catch (error) {
    console.error("Error saving notes: " + error);
    return createNotification("Error: " + error.message);
  }
}

/* ======================== MULTI-VIEW ACTIONS ======================== */

/**
 * Quick complete a contact from call view without notes
 */
function quickCompleteFromCallView(e) {
  const email = e.parameters.email;
  const page = e.parameters.page || '1';
  
  const completeEvent = {
    parameters: {
      email: email,
      origin: "callView",
      originParamsJson: JSON.stringify({ page: page })
    }
  };
  
  return markContactCompleted(completeEvent);
}

/**
 * Saves the call note for a contact from the Call Management card.
 */
function saveCallNote(e) {
  const email = e.parameters.email;
  const page = e.parameters.page || '1';

  const noteFieldName = "callNotes_" + (email ? email.replace(/[@\.+-]/g, "_") : "");
  const newNotes = e.formInput && e.formInput[noteFieldName] ? e.formInput[noteFieldName] : null;

  if (!email) {
    logAction("Error", "saveCallNote: Email parameter missing.");
    return createNotification("Cannot save note: Contact email not provided.");
  }
  if (newNotes === null) {
    const refreshCard = buildMultiContactView({ parameters: { page: page, focusMode: "false" } });
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("No note content to save."))
      .setNavigation(CardService.newNavigation().updateCard(refreshCard))
      .build();
  }

  const contact = getContactByEmail(email);
  if (!contact) {
    logAction("Error", `saveCallNote: Contact not found ${email}`);
    return createNotification("Contact not found: " + email);
  }

  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    logAction("Error", "saveCallNote Error: No database connected.");
    return createNotification("No database connected. Please connect to a database first.");
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);

    if (!contactsSheet) {
      logAction("Error", `saveCallNote Error: Contacts sheet not found for ${email}`);
      return createNotification("Contacts sheet not found. Please refresh the add-on.");
    }

    contactsSheet.getRange(contact.rowIndex, CONTACT_COLS.NOTES + 1).setValue(newNotes);
    SpreadsheetApp.flush();
    removeContactFromCache(contact.email);

    logAction("Call Note Saved", `Saved call note for ${contact.firstName} ${contact.lastName} (${email}).`);

    const updatedCard = buildMultiContactView({ parameters: { page: page, focusMode: "false" } });
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("✓ Note saved for " + contact.firstName))
      .setNavigation(CardService.newNavigation().updateCard(updatedCard))
      .build();

  } catch (error) {
    console.error("Error saving call note: " + error);
    logAction("Error", `Error saving call note for ${email}: ${error.toString()}`);
    return createNotification("Error saving note: " + error.message);
  }
}

/**
 * Saves the call note and keeps the contact in active cadence (explicit keep action).
 */
function saveNoteKeepCadence(e) {
  const email = e.parameters.email;
  const page = e.parameters.page || '1';

  const noteFieldName = "callNotes_" + (email ? email.replace(/[@\.+-]/g, "_") : "");
  const newNotes = e.formInput && e.formInput[noteFieldName] ? e.formInput[noteFieldName] : null;

  if (!email) {
    logAction("Error", "saveNoteKeepCadence: Email parameter missing.");
    return createNotification("Cannot save note: Contact email not provided.");
  }

  const contact = getContactByEmail(email);
  if (!contact) {
    logAction("Error", `saveNoteKeepCadence: Contact not found ${email}`);
    return createNotification("Contact not found: " + email);
  }

  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    logAction("Error", "saveNoteKeepCadence Error: No database connected.");
    return createNotification("No database connected. Please connect to a database first.");
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);

    if (!contactsSheet) {
      logAction("Error", `saveNoteKeepCadence Error: Contacts sheet not found for ${email}`);
      return createNotification("Contacts sheet not found. Please refresh the add-on.");
    }

    // Save notes if provided
    if (newNotes !== null && newNotes.trim() !== "") {
      contactsSheet.getRange(contact.rowIndex, CONTACT_COLS.NOTES + 1).setValue(newNotes);
      logAction("Call Note Saved", `Saved call note for ${contact.firstName} ${contact.lastName} (keeping in cadence)`);
    }
    
    // Ensure contact stays Active (in case it was paused)
    if (contact.status === "Paused") {
      contactsSheet.getRange(contact.rowIndex, CONTACT_COLS.STATUS + 1).setValue("Active");
      logAction("Status Update", `${contact.firstName} ${contact.lastName} set to Active (Keep Cadence)`);
    }
    
    SpreadsheetApp.flush();
    removeContactFromCache(contact.email);

    const updatedCard = buildMultiContactView({ parameters: { page: page, focusMode: "false" } });
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(`✓ Note saved · ${contact.firstName} stays in cadence`))
      .setNavigation(CardService.newNavigation().updateCard(updatedCard))
      .build();

  } catch (error) {
    console.error("Error saving call note (keep cadence): " + error);
    logAction("Error", `Error saving call note for ${email}: ${error.toString()}`);
    return createNotification("Error saving note: " + error.message);
  }
}

/**
 * Saves the call note and marks the contact as completed.
 */
function saveNoteAndEndSequence(e) {
  const email = e.parameters.email;
  const page = e.parameters.page || '1';

  const noteFieldName = "callNotes_" + (email ? email.replace(/[@\.+-]/g, "_") : "");
  const newNotes = e.formInput && e.formInput[noteFieldName] ? e.formInput[noteFieldName] : null;

  if (!email) {
    logAction("Error", "saveNoteAndEndSequence: Email parameter missing.");
    return createNotification("Cannot process: Contact email not provided.");
  }
  
  // Save the Note if content is present
  if (newNotes !== null && newNotes.trim() !== "") {
    const contact = getContactByEmail(email);
    if (contact) {
      const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
      if (spreadsheetId) {
        try {
          const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
          const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);
          if (contactsSheet) {
            contactsSheet.getRange(contact.rowIndex, CONTACT_COLS.NOTES + 1).setValue(newNotes);
            logAction("Call Note Saved", `Saved call note for ${contact.firstName} ${contact.lastName}`);
          }
        } catch (noteError) {
          console.error("Error saving note: " + noteError);
        }
      }
    }
  }

  // Mark the Contact as Completed
  const completeEvent = {
    parameters: {
      email: email,
      origin: "callView",
      originParamsJson: JSON.stringify({ page: page })
    }
  };
  
  return markContactCompleted(completeEvent);
}

/**
 * Marks a phone number as called. Refreshes Call Management card.
 */
function markPhoneNumberCalled(e) {
  const email = e.parameters.email;
  const phoneType = e.parameters.phoneType;
  const originPage = e.parameters.originPage || '1';

  const contact = getContactByEmail(email);
  if (!contact) {
    return createNotification("Contact not found: " + email);
  }
  if (phoneType !== "personal" && phoneType !== "work") {
     return createNotification("Invalid phone type specified.");
  }

  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    return createNotification("No database connected.");
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);

    if (!contactsSheet) {
      logAction("Error", `Mark Phone Called Error: Contacts sheet not found`);
      return createNotification("Contacts sheet not found.");
    }

    const columnToUpdate = (phoneType === "personal") ? CONTACT_COLS.PERSONAL_CALLED + 1 : CONTACT_COLS.WORK_CALLED + 1;
    contactsSheet.getRange(contact.rowIndex, columnToUpdate).setValue("Yes");
    SpreadsheetApp.flush();
    removeContactFromCache(contact.email);

    const updatedContactData = getContactByEmail(email);
    const personalDone = formatPhoneNumberForDisplay(updatedContactData.personalPhone) === "N/A" || updatedContactData.personalCalled === "Yes";
    const workDone = formatPhoneNumberForDisplay(updatedContactData.workPhone) === "N/A" || updatedContactData.workCalled === "Yes";
    const callTaskComplete = personalDone && workDone;

    logAction("Phone Call", `Marked ${phoneType} phone called for ${contact.firstName} ${contact.lastName}`);

    let notificationText = `✓ ${phoneType.charAt(0).toUpperCase() + phoneType.slice(1)} marked as called`;
    if (callTaskComplete) {
      notificationText += ` · All calls complete for ${contact.firstName}!`;
    }

    const refreshedCard = buildMultiContactView({ parameters: { page: originPage, focusMode: "false" } });
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(notificationText))
      .setNavigation(CardService.newNavigation().updateCard(refreshedCard))
      .build();

  } catch (error) {
    console.error(`Error marking ${phoneType} phone called: ` + error);
    logAction("Error", `Error marking ${phoneType} phone called: ${error.toString()}`);
    return createNotification("Error: " + error.message);
  }
}

/* ======================== LEGACY COMPATIBILITY ======================== */

/**
 * Legacy function - redirects to multi-contact view
 */
function displayCallContact(section, contact, callViewPage) {
  // No longer used - kept for compatibility
}

/**
 * Legacy function - redirects back to call management
 */
function showCallNoteCard(e) {
  const page = e.parameters.page || '1';
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .updateCard(buildMultiContactView({ parameters: { page: page, focusMode: "false" } })))
    .build();
}
