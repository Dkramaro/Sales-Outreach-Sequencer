/**
 * FILE: Core Entry Points and Main UI
 * 
 * PURPOSE:
 * Handles add-on initialization, authorization, contextual triggers,
 * and main dashboard UI. This is the primary entry point for the Gmail Add-on.
 * 
 * KEY FUNCTIONS:
 * - onInstall() - Add-on installation handler
 * - onOpen() - Add-on opened handler
 * - getContextualAddOn() - Contextual email trigger
 * - buildAddOn() - Main dashboard builder
 * - setupTriggers() - Daily trigger setup
 * - dailyContactCheck() - Automated daily contact status updates
 * 
 * DEPENDENCIES:
 * - 01_Config.gs: CONFIG, CONTACT_COLS
 * - 03_Database.gs: showConnectDatabaseForm, setupDatabase
 * - 04_ContactData.gs: getContactByEmail, getContactStats
 * - 13_ZoomInfoIntegration.gs: isZoomInfoEmail, buildZoomInfoImportCard
 * 
 * @version 2.4 - Redesigned dashboard: merged stats, primary/secondary action hierarchy
 */

/* ======================== ADD-ON LIFECYCLE ======================== */

/**
 * Called when the add-on is installed
 */
function onInstall(e) {
  setupTriggers();
  return onOpen(e);
}

/**
 * This function gets called when add-on is installed or opened from sidebar
 * on Gmail homepage (not inside an email)
 */
function onHomepage(e) {
  console.log("onHomepage trigger fired");
  return buildAddOn(e);
}

/**
 * Called when the document is opened
 */
function onOpen(e) {
  // Make sure to return a Card here
  return buildAddOn(e);
}

/**
 * Verifies Gmail API authorization
 * This function is specified in the manifest's authorizationCheckFunction
 */
function checkGmailAuthorization() {
  try {
    // Just test that we can access Gmail - force an error if not authorized
    const threads = GmailApp.getInboxThreads(0, 1);
    return { status: "OK" };
  } catch (e) {
    console.error("Gmail authorization check failed: " + e);
    return {
      status: "NOT_AUTHORIZED",
      authorizationUrl: "https://www.googleapis.com/auth/gmail.addons.execute https://www.googleapis.com/auth/gmail.readonly"
    };
  }
}

/* ======================== CONTEXTUAL TRIGGER ======================== */

/**
 * Called to build the contextual add-on view when an email is opened.
 * OPTIMIZED: Eliminates duplicate API calls for faster load times.
 * - Caches Gmail message object (Fix 1)
 * - Uses fast single-contact lookup (Fix 2)
 * - Passes contact data directly to viewContactCard (Fix 3)
 * - Batches property reads (Fix 4)
 */
function getContextualAddOn(e) {
  console.log("Contextual trigger fired.");
  const startTime = new Date().getTime();

  try {
    // More detailed logging to help troubleshoot
    if (e && e.gmail) {
      console.log("Gmail context data available: messageId=" +
                  (e.gmail.messageId ? "present" : "missing") +
                  ", accessToken=" + (e.gmail.accessToken ? "present" : "missing"));
    } else {
      console.log("Gmail context data not available in event object");
    }

    // FIX 4: Batch property reads - get all properties at once
    const userProps = PropertiesService.getUserProperties().getProperties();
    const spreadsheetId = userProps["SPREADSHEET_ID"];
    
    if (!spreadsheetId) {
      console.log("No database connected. Showing setup card.");
      // If not connected, show the setup card
      const setupCard = CardService.newCardBuilder();
      setupCard.setHeader(CardService.newCardHeader()
          .setTitle("SOS with draft Emails")
          .setImageUrl("https://www.gstatic.com/images/branding/product/2x/gmail_48dp.png"));
      const setupSection = CardService.newCardSection()
          .setHeader("Connect Database");
      setupSection.addWidget(CardService.newTextParagraph()
          .setText("To get started, connect to your Google Sheets database."));
      setupSection.addWidget(CardService.newButtonSet()
          .addButton(CardService.newTextButton()
              .setText("Connect to Database")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("showConnectDatabaseForm"))));
      setupCard.addSection(setupSection);
      return setupCard.build();
    }

    // Database is connected, proceed with contextual logic
    const messageId = e.gmail?.messageId;

    if (!messageId) {
      console.error("Missing messageId in event object.");
      logAction("Error", "Contextual Trigger Error: Missing messageId.");
      return buildErrorCard("Could not get email information. Message ID is missing.");
    }

    // FIX 1: Get the message ONCE and reuse it (eliminates duplicate Gmail API call)
    const message = GmailApp.getMessageById(messageId);
    
    if (!message) {
      console.error("Failed to get message");
      return buildErrorCard("Could not fetch email details. Try refreshing or reloading the add-on.");
    }

    // Check if this is a ZoomInfo email (uses cached message object)
    if (isZoomInfoEmail(message)) {
      console.log("ZoomInfo email detected. Load time: " + (new Date().getTime() - startTime) + "ms");
      return buildZoomInfoImportCard(message);
    }

    // FIX 1 continued: Extract headers directly from cached message object
    // Instead of calling getMessageDetails_ which would call GmailApp.getMessageById again
    const fromHeaderValue = message.getFrom();
    const toHeaderValue = message.getTo();

    // --- SENDER/RECIPIENT LOGIC (unchanged) ---
    const currentUserEmail = Session.getActiveUser() ? Session.getActiveUser().getEmail().toLowerCase() : null;
    if (!currentUserEmail) {
        console.error("Could not determine current user's email address.");
        logAction("Error", "Contextual Trigger Error: Could not get current user email.");
        return buildErrorCard("Could not identify the current user. Please ensure you are logged in correctly.");
    }

    const senderEmailRaw = extractEmailFromHeader_(fromHeaderValue);
    const recipientEmailRaw = extractEmailFromHeader_(toHeaderValue);

    let targetEmailForLookup;
    let relevantHeaderForNameExtraction;
    let isCurrentUserSender = false;

    if (senderEmailRaw && senderEmailRaw.toLowerCase() === currentUserEmail) {
        console.log("Email is FROM the current user. Switching to check the 'To' field for contact lookup.");
        targetEmailForLookup = recipientEmailRaw;
        relevantHeaderForNameExtraction = toHeaderValue;
        isCurrentUserSender = true;
    } else {
        console.log("Email is TO the current user (or from another external). Checking 'From' field for contact lookup.");
        targetEmailForLookup = senderEmailRaw;
        relevantHeaderForNameExtraction = fromHeaderValue;
    }

    if (!targetEmailForLookup) {
      console.warn("Could not extract a target email from relevant header: From='" + fromHeaderValue + "', To='" + toHeaderValue + "'");
      const infoMessage = isCurrentUserSender ? "Could not identify recipient email from the 'To' header." : "Could not identify sender email from the 'From' header.";
      return buildInfoCard("Email Not Identified", infoMessage + " The email address might be missing or formatted unexpectedly.");
    }

    console.log("Successfully identified target email for lookup: " + targetEmailForLookup);

    // FIX 2: Use fast single-contact lookup (avoids reading entire sheet)
    // Pass spreadsheetId to avoid extra property read
    const contact = getContactByEmailFast(targetEmailForLookup, spreadsheetId);

    if (contact) {
      // --- Target Email FOUND in contacts ---
      const logMessage = isCurrentUserSender ? "Recipient found in contacts database." : "Sender found in contacts database.";
      console.log(logMessage + " Building contact view card for: " + targetEmailForLookup);
      console.log("Contact lookup complete. Load time: " + (new Date().getTime() - startTime) + "ms");
      
      // FIX 3: Pass contact data directly to avoid duplicate lookup in viewContactCard
      return viewContactCardWithData(contact);
    } else {
      // --- Target Email NOT FOUND in contacts ---
      const logMessage = isCurrentUserSender ? "Recipient NOT found in contacts database." : "Sender NOT found in contacts database.";
      console.log(logMessage + " Building 'Add Contact' card for: " + targetEmailForLookup);
      console.log("Contact not found. Load time: " + (new Date().getTime() - startTime) + "ms");

      const addCard = CardService.newCardBuilder();
      const cardTitle = isCurrentUserSender ? "Recipient Not Found" : "Contact Not Found";
      const cardSubtitle = targetEmailForLookup;
      const paragraphText = isCurrentUserSender ? "This recipient is not in your sequence database." : "This sender is not in your sequence database.";
      const buttonText = isCurrentUserSender ? "Add Recipient to Sequence" : "Add Contact to Sequence";

      addCard.setHeader(CardService.newCardHeader()
        .setTitle(cardTitle)
        .setSubtitle(cardSubtitle)
        .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/person_add_black_48dp.png"));

      const addSection = CardService.newCardSection();
      addSection.addWidget(CardService.newTextParagraph().setText(paragraphText));

      // Extract name from the relevant header if possible
      let firstName = "", lastName = "";
      if (relevantHeaderForNameExtraction) {
        const namePart = relevantHeaderForNameExtraction.split('<')[0].trim();
        if (namePart) {
          const nameParts = namePart.split(/\s+/);
          firstName = nameParts[0] || "";
          if (nameParts.length > 1) {
             lastName = nameParts.slice(1).join(' ') || "";
          }
        }
      }

      // Pre-fill data for the add contact action
      const params = {
        email: targetEmailForLookup,
        firstName: firstName,
        lastName: lastName
      };

      addSection.addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
          .setText(buttonText)
          .setOnClickAction(CardService.newAction()
            .setFunctionName("showAddContactFormWithPrefill")
            .setParameters(params))));

      // Option to go to the main homepage
      addSection.addWidget(CardService.newTextButton()
        .setText("Open Full Add-on")
        .setOnClickAction(CardService.newAction().setFunctionName("buildAddOn")));

      addCard.addSection(addSection);
      return addCard.build();
    }
    // --- END MODIFIED LOGIC ---

  } catch (error) {
    console.error("Error in getContextualAddOn: " + error + "\n" + error.stack);
    logAction("Error", "Error in getContextualAddOn: " + error.toString());
    return buildErrorCard("An error occurred while processing the email context. Error details: " + error.message);
  }
}

/* ======================== MAIN DASHBOARD ======================== */

/**
 * Main entry point for the add-on
 */
function buildAddOn(e) {
  const card = CardService.newCardBuilder();

  // Add header
  card.setHeader(CardService.newCardHeader()
      .setTitle("Sales Outreach Sequence")
      .setSubtitle("Your email campaign manager")
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/email_black_48dp.png"));

  // Check if database is connected
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");

  if (!spreadsheetId) {
    // No database connected - show friendly welcome wizard
    return buildWelcomeWizard();
  }

  // Database is connected, verify it
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    if (!spreadsheet) {
      throw new Error("Could not open spreadsheet");
    }

    // Run background maintenance only on first open of the day
    runDailyStartupTasksIfNeeded();

    // Get stats and contacts
    const stats = getContactStats();
    const contacts = getAllContactsData();
    const realContacts = contacts.filter(c => !c.email.includes("example.com"));
    
    // ============================================================
    // FIRST-TIME USER: Show Getting Started Wizard (unless skipped)
    // ============================================================
    const userProps = PropertiesService.getUserProperties();
    let skipWizard = userProps.getProperty("SKIP_WIZARD") === "true";
    
    // MIGRATION FIX: For users who were using the system BEFORE onboarding was added,
    // they never went through the wizard, so SKIP_WIZARD was never set.
    // If user has real contacts AND has sent at least one email, they're clearly past onboarding.
    // This is a one-time migration that sets SKIP_WIZARD for existing users.
    if (!skipWizard && realContacts.length > 0) {
      const hasEmailHistory = realContacts.some(c => c.lastEmailDate && c.lastEmailDate !== "");
      if (hasEmailHistory) {
        userProps.setProperty("SKIP_WIZARD", "true");
        skipWizard = true;
      }
    }
    
    if (realContacts.length === 0 && !skipWizard) {
      return buildGettingStartedWizard(stats, contacts);
    }

    // ============================================================
    // RETURNING USER: Show Regular Dashboard
    // ============================================================
    
    // Calculate stats
    const totalReady = stats.readyForStep1 + stats.readyForStep2 + stats.readyForStep3 + stats.readyForStep4 + stats.readyForStep5;
    const readyContacts = contacts.filter(c => c.status === "Active" && c.isReady);
    const highPriorityPaused = contacts.filter(c => c.priority === "High" && c.status === "Paused");
    const completedThisWeek = contacts.filter(c => {
      if (c.status === "Completed" && c.lastEmailDate) {
        const weekAgo = new Date();
        weekAgo.setDate(weekAgo.getDate() - 7);
        return c.lastEmailDate >= weekAgo;
      }
      return false;
    }).length;

    // ═══════════════════════════════════════════════════════════════════
    // SECTION 1: WHAT TO DO NEXT (with stats context)
    // ═══════════════════════════════════════════════════════════════════
    const actionSection = CardService.newCardSection()
        .setHeader("🎯 What To Do Next");
    
    // Stats context line (merged from Dashboard)
    const statsText = totalReady > 0
        ? `📊 <b>${stats.total}</b> Contacts · <b>${totalReady}</b> Ready ⚠️`
        : `📊 <b>${stats.total}</b> Contacts · ✓ All caught up`;
    actionSection.addWidget(CardService.newDecoratedText()
        .setText(statsText)
        .setWrapText(true));

    // Action prompt and button
    if (readyContacts.length > 0) {
      actionSection.addWidget(CardService.newTextParagraph()
          .setText(`<b>${readyContacts.length} contact${readyContacts.length > 1 ? 's are' : ' is'} ready for emails!</b>`));
      
      actionSection.addWidget(CardService.newButtonSet()
          .addButton(CardService.newTextButton()
              .setText("📧 Send Emails Now")
              .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("viewContactsReadyForEmail")
                  .setParameters({page: '1'}))));
    } else {
      actionSection.addWidget(CardService.newTextParagraph()
          .setText("No emails to send right now. Check back later!"));
    }

    card.addSection(actionSection);

    // ═══════════════════════════════════════════════════════════════════
    // SECTION 2: SUGGESTIONS (conditional)
    // ═══════════════════════════════════════════════════════════════════
    
    // Onboarding suggestions for new users (< 5 contacts)
    if (realContacts.length < 5) {
      const onboardingSection = CardService.newCardSection()
          .setHeader("💡 Getting Started Tips");

      onboardingSection.addWidget(CardService.newTextParagraph()
          .setText("📖 <b>New here?</b> Read the Help guide to learn all features - bulk sending, call tracking, templates & more!"));

      onboardingSection.addWidget(CardService.newButtonSet()
          .addButton(CardService.newTextButton()
              .setText("📖 Help & Instructions")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("openHelpDocumentation"))));

      onboardingSection.addWidget(CardService.newTextParagraph()
          .setText("➕ <b>Add more contacts</b> to build your pipeline. Have ZoomInfo? Importing is super quick - just open a ZoomInfo email!"));

      onboardingSection.addWidget(CardService.newButtonSet()
          .addButton(CardService.newTextButton()
              .setText("➕ Add Contacts")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("buildContactManagementCard")
                  .setParameters({page: '1'}))));

      card.addSection(onboardingSection);
    } 
    // Regular suggestions for established users (5+ contacts)
    else if (highPriorityPaused.length > 0 || completedThisWeek > 0) {
      const suggestionsSection = CardService.newCardSection()
          .setHeader("💡 Suggestions");

      if (highPriorityPaused.length > 0) {
        suggestionsSection.addWidget(CardService.newTextParagraph()
            .setText(`⚠️ ${highPriorityPaused.length} high-priority contact${highPriorityPaused.length > 1 ? 's are' : ' is'} on hold. Consider resuming.`));
      }

      if (completedThisWeek > 0) {
        suggestionsSection.addWidget(CardService.newTextParagraph()
            .setText(`🎉 Great work! You finished ${completedThisWeek} email campaign${completedThisWeek > 1 ? 's' : ''} this week.`));
      }

      card.addSection(suggestionsSection);
    }

    // ═══════════════════════════════════════════════════════════════════
    // SECTION 3: QUICK ACTIONS (Primary navigation - prominent)
    // ═══════════════════════════════════════════════════════════════════
    const quickActionsSection = CardService.newCardSection()
        .setHeader("📌 Quick Actions");

    quickActionsSection.addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
            .setText("📧 Campaigns")
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(CardService.newAction()
                .setFunctionName("buildSequenceManagementCard"))));

    quickActionsSection.addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
            .setText("➕ Add Contacts")
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(CardService.newAction()
                .setFunctionName("buildContactManagementCard")
                .setParameters({page: '1'}))));

    quickActionsSection.addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
            .setText("✏️ Edit Templates")
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(CardService.newAction()
                .setFunctionName("buildTemplateManagementCard"))));

    // Divider before secondary actions
    quickActionsSection.addWidget(CardService.newDivider());

    // Secondary actions (compact DecoratedText with buttons)
    quickActionsSection.addWidget(CardService.newDecoratedText()
        .setText("📞 Calls")
        .setButton(CardService.newTextButton()
            .setText("Open")
            .setOnClickAction(CardService.newAction()
                .setFunctionName("buildCallManagementCard")
                .setParameters({page: '1'}))));

    quickActionsSection.addWidget(CardService.newDecoratedText()
        .setText("📊 Analytics")
        .setButton(CardService.newTextButton()
            .setText("Open")
            .setOnClickAction(CardService.newAction()
                .setFunctionName("buildAnalyticsHubCard"))));

    quickActionsSection.addWidget(CardService.newDecoratedText()
        .setText("⚙️ Settings")
        .setButton(CardService.newTextButton()
            .setText("Open")
            .setOnClickAction(CardService.newAction()
                .setFunctionName("buildSettingsCard"))));

    quickActionsSection.addWidget(CardService.newDivider());

    quickActionsSection.addWidget(CardService.newDecoratedText()
        .setText("❓ Tutorials")
        .setButton(CardService.newTextButton()
            .setText("Open")
            .setOnClickAction(CardService.newAction()
                .setFunctionName("openHelpDocumentation"))));

    card.addSection(quickActionsSection);

    // Display database info (not collapsible, with clear button)
    const infoSection = CardService.newCardSection()
        .setHeader("📊 Your Data");
    
    infoSection.addWidget(CardService.newTextParagraph()
        .setText("All your contacts are stored in: <b>" + spreadsheet.getName() + "</b>"));

    // Get the Contacts sheet gid to open directly to that tab
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);
    const contactsSheetGid = contactsSheet ? contactsSheet.getSheetId() : 0;

    infoSection.addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
            .setText("📊 Open Contacts Sheet")
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOpenLink(CardService.newOpenLink()
                .setUrl("https://docs.google.com/spreadsheets/d/" + spreadsheetId + "/edit#gid=" + contactsSheetGid))));

    card.addSection(infoSection);

  } catch (error) {
    // Reset spreadsheet ID if connection failed
    PropertiesService.getUserProperties().deleteProperty("SPREADSHEET_ID");
    logAction("Error", "Database connection failed in buildAddOn: " + error);
    return buildConnectionErrorCard(error);
  }

  return card.build();
}

/* ======================== GETTING STARTED WIZARD ======================== */

/**
 * Welcome wizard for users who haven't connected a database yet
 */
function buildWelcomeWizard() {
  const card = CardService.newCardBuilder();

  card.setHeader(CardService.newCardHeader()
      .setTitle("Welcome! 👋")
      .setSubtitle("Let's set up your email outreach tool")
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/celebration_black_48dp.png"));

  // Step 1: Create database
  const step1Section = CardService.newCardSection()
      .setHeader("Step 1 of 3: Create Your Database");

  step1Section.addWidget(CardService.newTextParagraph()
      .setText("First, we'll create a Google Sheet to store your contacts and email campaigns.\n\n<b>This takes just 5 seconds!</b>"));

  step1Section.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("🚀 Create My Database")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction()
              .setFunctionName("createNewDatabase"))));

  step1Section.addWidget(CardService.newTextParagraph()
      .setText("<i>Already have a database?</i>"));

  step1Section.addWidget(CardService.newTextButton()
      .setText("Connect Existing Database →")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("showConnectDatabaseForm")));

  card.addSection(step1Section);

  // What to expect
  const expectSection = CardService.newCardSection()
      .setCollapsible(true)
      .setHeader("What happens next?");

  expectSection.addWidget(CardService.newTextParagraph()
      .setText("After setup, you'll:\n\n" +
          "1️⃣ <b>Add contacts</b> - people you want to email\n\n" +
          "2️⃣ <b>Send emails</b> - using pre-written templates\n\n" +
          "3️⃣ <b>Track progress</b> - see who needs follow-up"));

  card.addSection(expectSection);

  return card.build();
}

/**
 * Getting started wizard for first-time users (database connected, no real contacts)
 */
function buildGettingStartedWizard(stats, contacts) {
  const card = CardService.newCardBuilder();

  card.setHeader(CardService.newCardHeader()
      .setTitle("Getting Started")
      .setSubtitle("3 quick steps to your first email")
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/rocket_launch_black_48dp.png"));

  // Progress indicator
  const progressSection = CardService.newCardSection();
  progressSection.addWidget(CardService.newTextParagraph()
      .setText("✅ Step 1: Database created\n" +
               "⏳ Step 2: Add your first contact\n" +
               "⏳ Step 3: Send your first email"));
  card.addSection(progressSection);

  // If demo contacts exist, encourage trying bulk processing
  const demoContacts = contacts.filter(c => c.email.includes("example.com"));
  const hasDemoContacts = demoContacts.length > 0 && contacts.length === demoContacts.length;
  
  if (hasDemoContacts) {
    const tryItSection = CardService.newCardSection()
        .setHeader("🚀 See Bulk Processing in Action!");
    
    tryItSection.addWidget(CardService.newTextParagraph()
        .setText(`We added <b>${demoContacts.length} demo contacts</b> to show you how quickly this system can process emails in bulk.\n\n` +
                 `<b>What happens next:</b>\n` +
                 `• All ${demoContacts.length} contacts are pre-selected\n` +
                 `• Click the button below to see them\n` +
                 `• Press "<b>Create Drafts</b>" in the sticky footer\n` +
                 `• Watch all ${demoContacts.length} emails get created instantly!`));
    
    tryItSection.addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
            .setText(`📧 Process ${demoContacts.length} Demo Emails`)
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(CardService.newAction()
                .setFunctionName("viewContactsReadyForEmail")
                .setParameters({page: '1', isDemoMode: 'true'}))));
    
    tryItSection.addWidget(CardService.newTextParagraph()
        .setText("<i>Creates drafts in Gmail - nothing sends automatically! Demo contacts are auto-deleted after.</i>"));
    
    card.addSection(tryItSection);
  }

  // Optional: Customize email templates (collapsible)
  const templatesSection = CardService.newCardSection()
      .setCollapsible(true)
      .setHeader("✏️ Customize Templates (Optional)");

  const availableSequences = getAvailableSequences();
  
  if (availableSequences.length > 0) {
    templatesSection.addWidget(CardService.newTextParagraph()
        .setText("Want to personalize your email templates first? You can edit them now or skip to adding contacts."));
    
    // Show first sequence for quick edit
    const firstSequence = availableSequences[0];
    templatesSection.addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
            .setText("✏️ Edit \"" + firstSequence + "\" Templates")
            .setOnClickAction(CardService.newAction()
                .setFunctionName("editTemplatesFromWizard")
                .setParameters({ sequence: firstSequence }))));
    
    if (availableSequences.length > 1) {
      templatesSection.addWidget(CardService.newTextParagraph()
          .setText(`<font color='#666666'>You have ${availableSequences.length} sequences. View all in Template Management.</font>`));
    }
  } else {
    templatesSection.addWidget(CardService.newTextParagraph()
        .setText("No sequences created yet. You can create one or skip to adding contacts."));
    
    templatesSection.addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
            .setText("📧 Create Email Sequence")
            .setOnClickAction(CardService.newAction()
                .setFunctionName("createSequenceFromWizard"))));
  }

  card.addSection(templatesSection);

  // Quick tips
  const tipsSection = CardService.newCardSection()
      .setCollapsible(true)
      .setHeader("💡 Quick Tips");

  tipsSection.addWidget(CardService.newTextParagraph()
      .setText("• <b>Email Campaign</b> = A sequence of emails sent over time\n\n" +
               "• <b>Step 1, 2, 3...</b> = Each email in your campaign\n\n" +
               "• <b>Ready to Send</b> = Contacts due for their next email\n\n" +
               "• Emails are created as <b>drafts</b> so you can review before sending"));

  card.addSection(tipsSection);

  // Navigation (hidden but available)
  const navSection = CardService.newCardSection()
      .setCollapsible(true)
      .setHeader("📁 Skip to Full Menu");

  navSection.addWidget(CardService.newTextButton()
      .setText("Show All Features →")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("buildFullDashboard")));

  card.addSection(navSection);

  return card.build();
}

/* ======================== WIZARD TEMPLATE EDITING ======================== */

/**
 * Opens template editor from the onboarding wizard
 * Marks that we came from wizard so save returns there
 */
function editTemplatesFromWizard(e) {
  const sequenceName = e.parameters.sequence;
  
  // Mark that we came from the wizard
  PropertiesService.getUserProperties().setProperty("EDITING_FROM_WIZARD", "true");
  
  // Build the sequence editor card
  return showSequenceTemplatesEditorFromWizard(sequenceName);
}

/**
 * Shows sequence editor with wizard-aware navigation
 */
function showSequenceTemplatesEditorFromWizard(sequenceName) {
  const card = CardService.newCardBuilder();

  card.setHeader(CardService.newCardHeader()
      .setTitle("📧 " + sequenceName)
      .setSubtitle("Customize your email templates"));

  const templates = getEmailTemplates(sequenceName);
  const currentStepCount = getSequenceStepCount(sequenceName);
  
  // Create a map for quick lookup
  const templateMap = {};
  for (const t of templates) {
    templateMap[t.step] = t;
  }

  // Info section
  const infoSection = CardService.newCardSection();
  infoSection.addWidget(CardService.newTextParagraph()
      .setText("<font color='#1a73e8'><b>📝 Editing from Setup Wizard</b></font>\n" +
               "Customize your templates below. When done, click <b>Return to Setup</b> to continue."));
  card.addSection(infoSection);

  // Show ACTIVE steps (1 to currentStepCount)
  for (let step = 1; step <= currentStepCount; step++) {
    const template = templateMap[step];
    const hasTemplate = template && template.body;
    
    const stepIcon = step === 1 ? "📤" : "↩️";
    const stepType = step === 1 ? "First Email" : "Follow-up #" + (step - 1);
    
    const stepSection = CardService.newCardSection()
        .setHeader(`${stepIcon} Step ${step}: ${stepType}`);
    
    if (hasTemplate) {
      // Subject line for Step 1
      if (step === 1 && template.subject) {
        stepSection.addWidget(CardService.newDecoratedText()
            .setTopLabel("SUBJECT")
            .setText(template.subject)
            .setWrapText(true));
      }
      
      // Preview of the email body
      const bodyPreview = template.body.length > 100 
        ? template.body.substring(0, 100).replace(/\n/g, " ") + "..." 
        : template.body.replace(/\n/g, " ");
      
      stepSection.addWidget(CardService.newDecoratedText()
          .setTopLabel("PREVIEW")
          .setText(bodyPreview)
          .setWrapText(true));
      
      stepSection.addWidget(CardService.newTextButton()
          .setText("✏️ Edit Step " + step)
          .setOnClickAction(CardService.newAction()
              .setFunctionName("editStepTemplateFromWizard")
              .setParameters({ 
                sequence: sequenceName, 
                step: step.toString(),
                rowIndex: template.rowIndex.toString()
              })));
    } else {
      stepSection.addWidget(CardService.newDecoratedText()
          .setText("No email configured")
          .setBottomLabel("Click below to write this email")
          .setWrapText(true));
      
      stepSection.addWidget(CardService.newTextButton()
          .setText("✉️ Write Step " + step + " Email")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction()
              .setFunctionName("editStepTemplateFromWizard")
              .setParameters({ 
                sequence: sequenceName, 
                step: step.toString(),
                rowIndex: "new"
              })));
    }
    
    card.addSection(stepSection);
  }

  // Navigation - Return to wizard
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("✓ Return to Setup")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setBackgroundColor("#137333")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("returnToWizardFromTemplatesUpdated")))
      .addButton(CardService.newTextButton()
          .setText("Skip to Full App")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("skipWizardAndGoToDashboard"))));
  card.addSection(navSection);

  return card.build();
}

/**
 * Edit a step template from the wizard context
 */
function editStepTemplateFromWizard(e) {
  const sequenceName = e.parameters.sequence;
  const step = parseInt(e.parameters.step);
  const rowIndex = e.parameters.rowIndex;
  const isNew = (rowIndex === "new");
  
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheetName = getSequenceSheetName(sequenceName);
  const sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    return createNotification("Sequence sheet not found.");
  }
  
  // Read current template data if editing existing
  let currentName = "";
  let currentSubject = "";
  let currentBody = "";
  
  if (!isNew) {
    const rowData = sheet.getRange(parseInt(rowIndex), 1, 1, 4).getValues()[0];
    currentName = rowData[1] || "";
    currentSubject = rowData[2] || "";
    currentBody = rowData[3] || "";
  }
  
  const card = CardService.newCardBuilder();
  
  const stepType = step === 1 ? "First Email" : "Reply #" + (step - 1);
  card.setHeader(CardService.newCardHeader()
      .setTitle(`Step ${step}: ${stepType}`)
      .setSubtitle(sequenceName));

  // Variable Dictionary
  const varsSection = CardService.newCardSection()
      .setHeader("📋 Variables (copy into your email)");
  
  varsSection.addWidget(CardService.newTextParagraph()
      .setText(
        "<b>Contact:</b> " +
        "<font color='#1a73e8'>{{firstName}}</font> · " +
        "<font color='#1a73e8'>{{lastName}}</font> · " +
        "<font color='#1a73e8'>{{company}}</font> · " +
        "<font color='#1a73e8'>{{title}}</font>\n\n" +
        "<b>Your Info:</b> " +
        "<font color='#1a73e8'>{{senderName}}</font> · " +
        "<font color='#1a73e8'>{{senderCompany}}</font>"
      ));
  card.addSection(varsSection);
  
  // Email Edit Section
  const formSection = CardService.newCardSection()
      .setHeader("Write Your Email");
  
  formSection.addWidget(CardService.newTextInput()
      .setFieldName("templateName")
      .setTitle("Name (for your reference)")
      .setValue(currentName || getDefaultStepName(step))
      .setHint("e.g., Introduction, Follow-up"));
  
  // Only show subject for Step 1
  if (step === 1) {
    formSection.addWidget(CardService.newTextInput()
        .setFieldName("subject")
        .setTitle("Subject Line")
        .setValue(currentSubject || "")
        .setHint("Example: {{company}} and our partnership?"));
  } else {
    formSection.addWidget(CardService.newTextParagraph()
        .setText("ℹ️ <i>This is a reply - it will use \"Re: [Step 1 subject]\" automatically.</i>"));
  }
  
  // Pre-populate body with greeting if empty
  const defaultBody = "Hi {{firstName}},\n\n";
  formSection.addWidget(CardService.newTextInput()
      .setFieldName("body")
      .setTitle("Email Body")
      .setValue(currentBody || defaultBody)
      .setMultiline(true)
      .setHint("Write your message after the greeting"));
  
  formSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("💾 Save Email")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction()
              .setFunctionName("saveStepTemplateFromWizard")
              .setParameters({ 
                sequence: sequenceName, 
                step: step.toString(),
                rowIndex: rowIndex,
                isNew: isNew.toString()
              }))));
  
  card.addSection(formSection);
  
  // Navigation
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
      .setText("← Back to Sequence")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("backToSequenceEditorFromWizard")
          .setParameters({ sequence: sequenceName })));
  card.addSection(navSection);
  
  return card.build();
}

/**
 * Saves step template and returns to wizard sequence editor
 */
function saveStepTemplateFromWizard(e) {
  const sequenceName = e.parameters.sequence;
  const step = parseInt(e.parameters.step);
  const rowIndex = e.parameters.rowIndex;
  const isNew = (e.parameters.isNew === "true");
  const templateName = e.formInput.templateName || getDefaultStepName(step);
  const subject = e.formInput.subject || "";
  const body = e.formInput.body || "";
  
  if (!body) {
    return createNotification("Please write your email body.");
  }
  
  // Step 1 requires subject
  if (step === 1 && !subject) {
    return createNotification("Step 1 requires a subject line.");
  }
  
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheetName = getSequenceSheetName(sequenceName);
  const sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    return createNotification("Sequence sheet not found.");
  }
  
  try {
    if (isNew) {
      // Check if a row for this step already exists
      const existingRow = findRowForStep(sheet, step);
      
      if (existingRow > 0) {
        // Update existing row
        sheet.getRange(existingRow, 2).setValue(templateName);
        sheet.getRange(existingRow, 3).setValue(subject);
        sheet.getRange(existingRow, 4).setValue(body);
      } else {
        // Append new row
        sheet.appendRow([step, templateName, subject, body]);
      }
    } else {
      // Update existing row at specified index
      const row = parseInt(rowIndex);
      sheet.getRange(row, 2).setValue(templateName);
      sheet.getRange(row, 3).setValue(subject);
      sheet.getRange(row, 4).setValue(body);
    }
    
    SpreadsheetApp.flush();
    logAction("Template Update", `Saved Step ${step} template in ${sequenceName} (from wizard)`);
    
    // Return to wizard sequence editor
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("✓ Step " + step + " email saved!"))
      .setNavigation(CardService.newNavigation()
        .updateCard(showSequenceTemplatesEditorFromWizard(sequenceName)))
      .build();
      
  } catch (error) {
    console.error("Error saving template: " + error);
    return createNotification("Error saving: " + error.message);
  }
}

/**
 * Returns to wizard sequence editor
 */
function backToSequenceEditorFromWizard(e) {
  const sequenceName = e.parameters.sequence;
  return showSequenceTemplatesEditorFromWizard(sequenceName);
}

/**
 * Returns to the main onboarding wizard from templates
 */
function returnToWizardFromTemplates() {
  // Clear the wizard editing flag
  PropertiesService.getUserProperties().deleteProperty("EDITING_FROM_WIZARD");
  
  // Return to the getting started wizard
  const stats = getContactStats();
  const contacts = getAllContactsData();
  
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .popToRoot()
      .updateCard(buildGettingStartedWizard(stats, contacts)))
    .build();
}

/**
 * Opens sequence creation from the wizard
 */
function createSequenceFromWizard() {
  // Mark that we came from the wizard
  PropertiesService.getUserProperties().setProperty("EDITING_FROM_WIZARD", "true");
  
  // Go to the template management card to create a sequence
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .pushCard(buildTemplateManagementCard()))
    .build();
}

/**
 * Skips the wizard and goes to full dashboard
 */
function skipWizardAndGoToDashboard() {
  // Clear wizard flags
  PropertiesService.getUserProperties().deleteProperty("EDITING_FROM_WIZARD");
  PropertiesService.getUserProperties().setProperty("SKIP_WIZARD", "true");
  
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .popToRoot()
      .updateCard(buildFullDashboard()))
    .build();
}

/**
 * Opens template editor from the Contact Added card
 */
function editTemplatesFromContactAddedCard(e) {
  const sequenceName = e.parameters.sequence;
  const contactName = e.parameters.contactName;
  
  // Store context so we can return to Contact Added card
  PropertiesService.getUserProperties().setProperty("EDITING_FROM_WIZARD", "true");
  PropertiesService.getUserProperties().setProperty("WIZARD_CONTACT_NAME", contactName);
  
  return showSequenceTemplatesEditorFromWizard(sequenceName);
}

/**
 * Returns to the Contact Added card from templates (if contact name was stored)
 * Otherwise returns to Getting Started wizard
 */
function returnToWizardFromTemplatesUpdated() {
  // Clear the wizard editing flag
  PropertiesService.getUserProperties().deleteProperty("EDITING_FROM_WIZARD");
  
  // Check if we came from Contact Added card
  const contactName = PropertiesService.getUserProperties().getProperty("WIZARD_CONTACT_NAME");
  PropertiesService.getUserProperties().deleteProperty("WIZARD_CONTACT_NAME");
  
  if (contactName) {
    // Return to Contact Added card
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation()
        .popToRoot()
        .updateCard(buildFirstContactAddedCard(contactName, "")))
      .build();
  }
  
  // Return to the getting started wizard
  const stats = getContactStats();
  const contacts = getAllContactsData();
  
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .popToRoot()
      .updateCard(buildGettingStartedWizard(stats, contacts)))
    .build();
}

/**
 * Wizard card shown after user adds their first real contact
 */
function buildFirstContactAddedCard(firstName, email) {
  const card = CardService.newCardBuilder();

  card.setHeader(CardService.newCardHeader()
      .setTitle("🎉 Contact Added!")
      .setSubtitle("One more step to go")
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/check_circle_black_48dp.png"));

  // Progress indicator
  const progressSection = CardService.newCardSection();
  progressSection.addWidget(CardService.newTextParagraph()
      .setText("✅ Step 1: Database created\n" +
               "✅ Step 2: Added " + firstName + "\n" +
               "⏳ Step 3: Send your first email"));
  card.addSection(progressSection);

  // Optional: Customize templates before sending
  const templatesSection = CardService.newCardSection()
      .setHeader("✏️ Customize Templates First? (Optional)");

  const availableSequences = getAvailableSequences();
  
  if (availableSequences.length > 0) {
    templatesSection.addWidget(CardService.newTextParagraph()
        .setText("Before sending, you can customize your email templates:"));
    
    const firstSequence = availableSequences[0];
    templatesSection.addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
            .setText("✏️ Edit \"" + firstSequence + "\" Templates")
            .setOnClickAction(CardService.newAction()
                .setFunctionName("editTemplatesFromContactAddedCard")
                .setParameters({ sequence: firstSequence, contactName: firstName }))));
  }
  
  card.addSection(templatesSection);

  // Main action: Send first email
  const sendSection = CardService.newCardSection()
      .setHeader("📧 Send Your First Email");

  sendSection.addWidget(CardService.newTextParagraph()
      .setText("You're ready to send an email to <b>" + firstName + "</b>!\n\n" +
               "Click below to create your first email draft:"));

  sendSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("📧 Send Email Now")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction()
              .setFunctionName("viewContactsReadyForEmail")
              .setParameters({page: '1'}))));

  sendSection.addWidget(CardService.newTextParagraph()
      .setText("<i>This creates a draft in Gmail for you to review before sending.</i>"));

  card.addSection(sendSection);

  // Add more contacts option
  const moreSection = CardService.newCardSection();
  moreSection.addWidget(CardService.newTextButton()
      .setText("➕ Add More Contacts First")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("buildContactManagementCard")
          .setParameters({page: '1'})));
  
  moreSection.addWidget(CardService.newTextButton()
      .setText("Go to Full Dashboard →")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("buildFullDashboard")));

  card.addSection(moreSection);

  return card.build();
}

/**
 * Success card shown after demo emails are created - guides user to next steps
 * Celebrates bulk processing capability with 10 demo contacts
 */
function buildDemoEmailSuccessCard() {
  const card = CardService.newCardBuilder();

  card.setHeader(CardService.newCardHeader()
      .setTitle("🎉 10 Drafts Created!")
      .setSubtitle("Bulk processing complete!")
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/check_circle_black_48dp.png"));

  // Success message - emphasize bulk processing speed
  const successSection = CardService.newCardSection()
      .setHeader("⚡ That's Bulk Power!");
  
  successSection.addWidget(CardService.newTextParagraph()
      .setText("You just created <b>10 personalized email drafts</b> in seconds!\n\n" +
               "📬 <b>Check your Gmail Drafts folder</b> to see all 10 emails.\n\n" +
               "Each email is personalized with the contact's name, company, and title. Imagine doing this for 100 or 1,000 contacts!"));

  card.addSection(successSection);

  // What happened section
  const detailsSection = CardService.newCardSection()
      .setHeader("What happened");
  
  detailsSection.addWidget(CardService.newTextParagraph()
      .setText("• Created 10 personalized email drafts\n" +
               "• Each uses your email template\n" +
               "• All demo contacts auto-removed\n" +
               "• Ready for real contacts now!"));

  card.addSection(detailsSection);

  // Pro tip about settings
  const tipSection = CardService.newCardSection()
      .setHeader("💡 Pro Tip");
  
  tipSection.addWidget(CardService.newTextParagraph()
      .setText("By default, emails are created as <b>drafts</b> so you can review them before sending.\n\n" +
               "If you prefer, you can enable auto-send for first emails in <b>Settings</b>. (We don't recommend this until you're comfortable with the templates.)"));

  card.addSection(tipSection);

  // Next steps
  const nextSection = CardService.newCardSection()
      .setHeader("🚀 Ready to Start?");

  nextSection.addWidget(CardService.newTextParagraph()
      .setText("🎯 <b>Recommended:</b> Add yourself as your first contact! This way you can see exactly how your emails look before sending to external contacts."));

  nextSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("➕ Add Your First Real Contact")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction()
              .setFunctionName("buildContactManagementCard")
              .setParameters({page: '1'}))));

  nextSection.addWidget(CardService.newTextButton()
      .setText("Go to Dashboard →")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("buildFullDashboard")));

  card.addSection(nextSection);

  return card.build();
}

/**
 * Success card shown after user sends their first REAL email (not demo)
 * This celebrates completion of onboarding and guides to next steps
 */
function buildFirstRealEmailSuccessCard() {
  const card = CardService.newCardBuilder();

  card.setHeader(CardService.newCardHeader()
      .setTitle("🚀 You're Ready!")
      .setSubtitle("Outreach at scale unlocked")
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/celebration_black_48dp.png"));

  // Success message
  const successSection = CardService.newCardSection()
      .setHeader("🎉 First Email Sent!");
  
  successSection.addWidget(CardService.newTextParagraph()
      .setText("Congratulations! You've completed the setup and sent your first email draft.\n\n" +
               "<b>📬 Check your Gmail Drafts folder</b> to review and send it.\n\n" +
               "You're now ready to scale your outreach with automated email sequences!"));

  card.addSection(successSection);

  // Learn more section
  const learnSection = CardService.newCardSection()
      .setHeader("📖 Master All Features");
  
  learnSection.addWidget(CardService.newTextParagraph()
      .setText("Want to get the most out of this tool? Our Help guide covers:\n\n" +
               "• Bulk email sending & follow-ups\n" +
               "• Call management & tracking\n" +
               "• Template customization\n" +
               "• Analytics & reply tracking\n" +
               "• ZoomInfo integration for quick imports"));

  learnSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("📖 Read Help Guide")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction()
              .setFunctionName("openHelpDocumentation"))));

  card.addSection(learnSection);

  // Next steps - go straight to dashboard
  const nextSection = CardService.newCardSection();

  nextSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("🚀 Go to Dashboard")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction()
              .setFunctionName("buildFullDashboard"))));

  card.addSection(nextSection);

  return card.build();
}

/**
 * Full dashboard bypass for users who want to skip the wizard
 */
function buildFullDashboard(e) {
  // Set flag to skip wizard in future
  PropertiesService.getUserProperties().setProperty("SKIP_WIZARD", "true");
  return buildAddOn(e);
}

/**
 * Connection error card
 */
function buildConnectionErrorCard(error) {
  const card = CardService.newCardBuilder();
  
  card.setHeader(CardService.newCardHeader()
      .setTitle("Connection Issue")
      .setSubtitle("Let's fix this")
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/warning_amber_48dp.png"));

  const section = CardService.newCardSection();
  section.addWidget(CardService.newTextParagraph()
      .setText("We couldn't connect to your database. This usually happens when:\n\n" +
               "• The spreadsheet was deleted\n" +
               "• Permissions were changed\n\n" +
               "Let's reconnect:"));

  section.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("🔄 Reconnect Database")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction()
              .setFunctionName("showConnectDatabaseForm"))));

  card.addSection(section);

  return card.build();
}

/* ======================== DAILY STARTUP TASKS ======================== */

/**
 * Runs background maintenance tasks only on the first open of the day.
 * This avoids running expensive operations (migration, column checks, trigger scans)
 * on every single dashboard load.
 */
function runDailyStartupTasksIfNeeded() {
  const userProps = PropertiesService.getUserProperties();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const lastRunDate = userProps.getProperty("LAST_STARTUP_TASKS_DATE");
  
  // Skip if we've already run today
  if (lastRunDate === today) {
    return;
  }
  
  console.log("Running daily startup tasks (first open of " + today + ")");
  
  try {
    // Auto-migrate existing contacts to sequences
    migrateExistingContactsToSequences();
    
    // Ensure all required columns exist (for Reply tracking, etc.)
    ensureContactsSheetColumns();
    
    // Ensure daily triggers exist (for test deployments where onInstall doesn't fire)
    ensureTriggersExist();
    
    // Mark as completed for today
    userProps.setProperty("LAST_STARTUP_TASKS_DATE", today);
    console.log("Daily startup tasks completed successfully");
  } catch (error) {
    console.error("Error in daily startup tasks: " + error);
    // Don't save the date on error - will retry next load
  }
}

/* ======================== AUTOMATED TRIGGERS ======================== */

/**
 * Set up daily triggers for automated operations
 * Called on install and can be manually triggered from Settings
 */
function setupTriggers() {
  // Delete existing triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  const triggersToRemove = ["dailyThreadPreValidation", "dailyMaintenance", "dailyReplyCheck"];
  
  for (const trigger of triggers) {
    if (triggersToRemove.includes(trigger.getHandlerFunction())) {
      try {
         ScriptApp.deleteTrigger(trigger);
         console.log("Deleted existing " + trigger.getHandlerFunction() + " trigger.");
      } catch (e) {
         console.error("Failed to delete trigger: " + e);
      }
    }
  }

  // Daily Maintenance at 4:30 AM
  // Combines: Thread pre-validation + Reply detection
  try {
      ScriptApp.newTrigger("dailyMaintenance")
        .timeBased()
        .everyDays(1)
        .atHour(4)
        .nearMinute(30)
        .create();
      console.log("Created dailyMaintenance trigger (4:30 AM).");
      logAction("Setup Triggers", "Daily maintenance trigger configured successfully");
  } catch(e) {
      console.error("Failed to create dailyMaintenance trigger: " + e);
      logAction("Error", "Failed to create dailyMaintenance trigger: " + e);
  }
}

/**
 * Ensures required triggers exist (for test deployments where onInstall doesn't fire)
 * Only creates missing triggers - doesn't duplicate existing ones
 */
function ensureTriggersExist() {
  const triggers = ScriptApp.getProjectTriggers();
  const existingFunctions = triggers.map(t => t.getHandlerFunction());
  
  // Daily maintenance trigger handles: thread validation + reply detection
  if (existingFunctions.includes("dailyMaintenance")) {
    return; // Essential trigger exists
  }
  
  // Also check for old trigger name and remove it
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "dailyThreadPreValidation") {
      try {
        ScriptApp.deleteTrigger(trigger);
        console.log("Removed old dailyThreadPreValidation trigger.");
      } catch(e) {
        console.error("Failed to remove old trigger: " + e);
      }
    }
  }
  
  console.log("Creating dailyMaintenance trigger...");
  
  try {
    ScriptApp.newTrigger("dailyMaintenance")
      .timeBased()
      .everyDays(1)
      .atHour(4)
      .nearMinute(30)
      .create();
    console.log("Created dailyMaintenance trigger (4:30 AM).");
    logAction("Auto Trigger Setup", "Created dailyMaintenance trigger");
  } catch(e) {
    console.error("Failed to create dailyMaintenance trigger: " + e);
  }
}

/**
 * Daily check of contacts to update statuses and step progression
 */
function dailyContactCheck() {
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    console.log("Daily Check: No database connected.");
    return;
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);

    if (!contactsSheet) {
      logAction("Error", "Daily Check Error: Contacts sheet not found.");
      return;
    }

    const dataRange = contactsSheet.getDataRange();
    const data = dataRange.getValues();
    if (data.length <= 1) {
      console.log("Daily Check: No contacts found.");
      return; // Only header row exists
    }

    const today = new Date();
    today.setHours(0, 0, 0, 0); // Start of day for comparison

    let activatedCount = 0;
    const updates = []; // Store updates to write in batch

    // Process each contact row (skip header)
    for (let i = 1; i < data.length; i++) {
      const status = data[i][CONTACT_COLS.STATUS];
      const nextStepDateVal = data[i][CONTACT_COLS.NEXT_STEP_DATE];
      let nextStepDate = null;
      if (nextStepDateVal instanceof Date && !isNaN(nextStepDateVal)) {
          nextStepDate = new Date(nextStepDateVal);
          nextStepDate.setHours(0,0,0,0); // Normalize date for comparison
      } else if (typeof nextStepDateVal === 'string' && nextStepDateVal) {
          try {
              nextStepDate = new Date(nextStepDateVal);
              if (!isNaN(nextStepDate)) {
                 nextStepDate.setHours(0,0,0,0); // Normalize date for comparison
              } else {
                 nextStepDate = null; // Invalid date string
              }
          } catch(dateErr) {
              nextStepDate = null; // Error parsing date
          }
      }


      // If the contact is paused and next step date is today or earlier, activate them
      if (status === "Paused" && nextStepDate && nextStepDate <= today) {
        // Prepare update: [row index (1-based), column index (1-based), value]
        updates.push([i + 1, CONTACT_COLS.STATUS + 1, "Active"]);
        activatedCount++;

        // Log the action detail immediately or collect details to log later
        const contactName = data[i][CONTACT_COLS.FIRST_NAME] + " " + data[i][CONTACT_COLS.LAST_NAME];
        console.log("Daily Check: Activating contact - " + contactName);
        logAction("Auto Status Update", "Contact activated: " + contactName);
      }
    }

    // Apply batch updates if any changes were made
    if (updates.length > 0) {
       updates.forEach(update => {
           contactsSheet.getRange(update[0], update[1]).setValue(update[2]);
       });
       SpreadsheetApp.flush(); // Ensure changes are saved
       console.log("Daily Check: Activated " + activatedCount + " contacts.");
       logAction("Daily Check", "Activated " + activatedCount + " paused contacts whose next step date was reached.");
    } else {
       console.log("Daily Check: No contacts needed activation.");
    }

  } catch (error) {
    console.error("Error in daily contact check: " + error + "\n" + error.stack);
    logAction("Error", "Error in daily contact check: " + error.toString());
  }
}

/* ======================== ADD CONTACT FROM CONTEXT ======================== */

/**
 * Shows the Add Contact form, pre-filled with data from context.
 * Uses progressive disclosure design for simplicity.
 */
function showAddContactFormWithPrefill(e) {
    const prefillEmail = e.parameters.email || "";
    const prefillFirstName = e.parameters.firstName || "";
    const prefillLastName = e.parameters.lastName || "";

    const card = CardService.newCardBuilder();

    // Add header
    card.setHeader(CardService.newCardHeader()
        .setTitle("➕ Add Contact")
        .setSubtitle("Add to your email campaigns")
        .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/person_add_black_48dp.png"));

    // ============================================================
    // SINGLE UNIFIED FORM - Required fields first, optional after, button at end
    // ============================================================
    const formSection = CardService.newCardSection()
        .setHeader("Contact Details");

    formSection.addWidget(CardService.newTextParagraph()
        .setText("We've filled in what we could from the email:"));

    // REQUIRED FIELDS
    formSection.addWidget(CardService.newTextInput()
        .setFieldName("firstName")
        .setTitle("First Name *")
        .setValue(prefillFirstName)
        .setHint("For personalized emails"));

    formSection.addWidget(CardService.newTextInput()
        .setFieldName("lastName")
        .setTitle("Last Name *")
        .setValue(prefillLastName)
        .setHint("Their last name"));

    formSection.addWidget(CardService.newTextInput()
        .setFieldName("email")
        .setTitle("Email *")
        .setValue(prefillEmail)
        .setHint("Where to send emails"));

    formSection.addWidget(CardService.newTextInput()
        .setFieldName("company")
        .setTitle("Company *")
        .setHint("For {{company}} in emails"));

    // Sequence dropdown (required)
    const sequenceDropdown = CardService.newSelectionInput()
        .setType(CardService.SelectionInputType.DROPDOWN)
        .setTitle("Email Campaign *")
        .setFieldName("sequence");
    
    const availableSequences = getAvailableSequences();
    if (availableSequences.length === 0) {
      sequenceDropdown.addItem("No campaigns available", "", true);
    } else {
      let defaultSet = false;
      for (const sequence of availableSequences) {
          const isDefault = !defaultSet && (sequence === "SaaS / B2B Tech" || sequence === availableSequences[0]);
          sequenceDropdown.addItem(sequence, sequence, isDefault);
          if (isDefault) defaultSet = true;
      }
    }
    formSection.addWidget(sequenceDropdown);

    // OPTIONAL FIELDS (inline, not collapsible)
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

    // ADD CONTACT BUTTON - at the very bottom
    formSection.addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
            .setText("✅ Add Contact")
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
            .setOnClickAction(CardService.newAction()
                .setFunctionName("addContact"))));

    card.addSection(formSection);

    // Navigation section
    const navSection = CardService.newCardSection();
    navSection.addWidget(CardService.newTextButton()
        .setText("View All Contacts")
        .setOnClickAction(CardService.newAction()
            .setFunctionName("buildContactManagementCard")
             .setParameters({page: '1'})));

    navSection.addWidget(CardService.newTextButton()
        .setText("← Back to Dashboard")
        .setOnClickAction(CardService.newAction()
            .setFunctionName("buildAddOn")));

    card.addSection(navSection);

    return card.build();
}

