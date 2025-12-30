/**
 * FILE: Template Management UI
 * 
 * PURPOSE:
 * Handles template and sequence management user interface including
 * template editing, creation, deletion, and sequence creation.
 * 
 * KEY FUNCTIONS:
 * - buildTemplateManagementCard() - Main template management UI
 * - buildTemplateCreationPromptCard() - Missing template prompt
 * - editTemplate() - Template editor card
 * - createNewSequence() - Create new sequence action
 * - createTemplate() - Create new template (legacy)
 * - saveTemplateChanges() - Save template edits
 * - saveTemplateSelection() - Save template selection for step
 * - deleteTemplateConfirmation() - Confirmation before delete
 * - deleteTemplate() - Delete template action
 * 
 * DEPENDENCIES:
 * - 01_Config.gs: INDUSTRY_OPTIONS
 * - 03_Database.gs: logAction
 * - 06_SequenceData.gs: getAvailableSequences, getSequenceSheetName, createSequenceSheet, getEmailTemplates
 * - 08_EmailComposition.gs: replacePlaceholders
 * 
 * @version 2.3
 */

/* ======================== TEMPLATE MANAGEMENT UI ======================== */

/**
 * Builds the Template Management card with paginated sequences.
 * @param {object} e Optional event object containing page parameter
 */
function buildTemplateManagementCard(e) {
  // Pagination setup
  const page = parseInt(e && e.parameters && e.parameters.page || '1');
  const pageSize = 15; // Sequences per page
  
  const card = CardService.newCardBuilder();

  // Add header
  card.setHeader(CardService.newCardHeader()
      .setTitle("SoS with Draft Emails")
      .setSubtitle("Sequence & Template Management")
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/description_black_48dp.png"));

  // ============ CREATE NEW SEQUENCE - TOP & PROMINENT ============
  const createSection = CardService.newCardSection()
      .setHeader("➕ Create New Sequence");

  createSection.addWidget(CardService.newTextParagraph()
      .setText("<b>Start a new email sequence with custom templates</b>"));

  createSection.addWidget(CardService.newTextInput()
      .setFieldName("newSequenceName")
      .setTitle("Sequence Name")
      .setHint('e.g., "Healthcare Outreach" or "Enterprise Sales"'));

  createSection.addWidget(CardService.newTextButton()
      .setText("🚀 Create Sequence")
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor("#1a73e8")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("startSequenceCreationWizard")));

  card.addSection(createSection);

  // ============ EXISTING SEQUENCES (PAGINATED) ============
  const availableSequences = getAvailableSequences();
  const totalSequences = availableSequences.length;
  const totalPages = Math.max(1, Math.ceil(totalSequences / pageSize));
  const validPage = Math.min(Math.max(1, page), totalPages); // Ensure page is within bounds
  const startIndex = (validPage - 1) * pageSize;
  const endIndex = Math.min(startIndex + pageSize, totalSequences);
  const sequencesToShow = availableSequences.slice(startIndex, endIndex);
  
  // Build header with count info
  let headerText = "📋 Your Sequences";
  if (totalSequences > 0) {
    headerText = `📋 Your Sequences (${startIndex + 1}-${endIndex} of ${totalSequences})`;
  }
  
  const sequenceSection = CardService.newCardSection()
      .setHeader(headerText);
  
  if (totalSequences === 0) {
    sequenceSection.addWidget(CardService.newTextParagraph()
        .setText("<font color='#666666'>No sequences yet. Create your first one above!</font>"));
  } else {
    sequenceSection.addWidget(CardService.newTextParagraph()
        .setText(`You have <b>${totalSequences} sequence${totalSequences > 1 ? 's' : ''}</b>. Click Edit to modify templates:`));

    for (const sequenceName of sequencesToShow) {
      // Sequence name as prominent header
      sequenceSection.addWidget(CardService.newTextParagraph()
          .setText(`<b>📋 ${sequenceName}</b>`));
      
      // Edit and Duplicate buttons
      sequenceSection.addWidget(CardService.newButtonSet()
          .addButton(CardService.newTextButton()
              .setText("✏️ Edit Templates")
              .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("showSequenceTemplatesEditor")
                  .setParameters({ sequence: sequenceName })))
          .addButton(CardService.newTextButton()
              .setText("📋 Duplicate")
              .setOnClickAction(CardService.newAction()
                  .setFunctionName("duplicateSequence")
                  .setParameters({ sequence: sequenceName }))));

      sequenceSection.addWidget(CardService.newDivider());
    }
    
    // Add pagination buttons if needed
    addPaginationButtons(sequenceSection, validPage, totalPages, "buildTemplateManagementCard", {});
  }
  card.addSection(sequenceSection);

  // --- Instructions Section ---
  const instructionsSection = CardService.newCardSection()
      .setCollapsible(true)
      .setHeader("How to Use Sequences");
      
  const instructions = `
<b>How Sequences Work:</b>

1. Create a sequence with 1-5 email steps
2. Write a unique email for each step
3. Assign contacts to the sequence
4. Send emails step-by-step with follow-ups

<b>Workflow:</b>
• Step 1: First outreach email (new thread)
• Steps 2-5: Reply emails (same thread)
  `;

  instructionsSection.addWidget(CardService.newTextParagraph()
      .setText(instructions));
  card.addSection(instructionsSection);

  // --- Placeholders Help Section ---
  const helpSection = CardService.newCardSection()
      .setCollapsible(true)
      .setHeader("Available Placeholders");
  const placeholders = [
    "{{firstName}}", "{{lastName}}", "{{email}}",
    "{{company}}", "{{title}}",
    "{{senderName}}", "{{senderCompany}}", "{{senderTitle}}",
    "{{industry}} (from Industry field)",
    "{{department}} (from Notes: 'department: ...')"
  ];
  helpSection.addWidget(CardService.newTextParagraph()
      .setText("Use these placeholders in your templates:\n" + placeholders.join(", ")));
  card.addSection(helpSection);

  // --- Navigation Section ---
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
      .setText("Back to Main Menu")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("buildAddOn")));
  card.addSection(navSection);

  return card.build();
}

/**
 * Builds template creation prompt card - NEW FUNCTION
 */
function buildTemplateCreationPromptCard(sequenceName, stepNumber, suggestedName, spreadsheetUrl) {
  const card = CardService.newCardBuilder();
  
  card.setHeader(CardService.newCardHeader()
    .setTitle("Missing Template")
    .setSubtitle(`Step ${stepNumber} - ${sequenceName}`)
    .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/warning_black_48dp.png"));

  // Problem description
  const problemSection = CardService.newCardSection()
    .setHeader("Template Required");
  problemSection.addWidget(CardService.newTextParagraph()
    .setText(`To send emails for Step ${stepNumber} in the "${sequenceName}" sequence, you need to create a template.`));
  card.addSection(problemSection);

  // Template requirements
  const requirementsSection = CardService.newCardSection()
    .setHeader("Template Requirements");
  requirementsSection.addWidget(CardService.newKeyValue()
    .setTopLabel("Sequence")
    .setContent(sequenceName));
  requirementsSection.addWidget(CardService.newKeyValue()
    .setTopLabel("Template Name")
    .setContent(suggestedName)
    .setBottomLabel("Must start with 'Step " + stepNumber + "'"));
  requirementsSection.addWidget(CardService.newTextParagraph()
    .setText("The template will be automatically selected based on the sequence and step number."));
  card.addSection(requirementsSection);

  // Instructions
  const instructionsSection = CardService.newCardSection()
    .setHeader("How to Create Template");
  instructionsSection.addWidget(CardService.newTextParagraph()
    .setText(`1. Open the "Sequence-${sequenceName}" sheet\n2. Go to column ${stepNumber} (Email ${stepNumber})\n3. In row 3, enter your template in this format:\n\nSubject: Your email subject\n\nBody:\nYour email content here\nMultiple lines supported`));
  card.addSection(instructionsSection);

  // Action buttons
  const actionSection = CardService.newCardSection();
  const buttonSet = CardService.newButtonSet();
  
  if (spreadsheetUrl) {
    buttonSet.addButton(CardService.newTextButton()
      .setText("Open Templates Sheet")
      .setOpenLink(CardService.newOpenLink().setUrl(spreadsheetUrl)));
  }
  
  buttonSet.addButton(CardService.newTextButton()
    .setText("Template Management")
    .setOnClickAction(CardService.newAction()
      .setFunctionName("buildTemplateManagementCard")));
  
  actionSection.addWidget(buttonSet);
  card.addSection(actionSection);

  // Navigation
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
    .setText("Back")
    .setOnClickAction(CardService.newAction()
      .setFunctionName("buildSequenceManagementCard")));
  card.addSection(navSection);

  return card.build();
}

/* ======================== SEQUENCE CREATION WIZARD ======================== */

/**
 * Step 1 of wizard: Validate name and show step count selection
 */
function startSequenceCreationWizard(e) {
  const sequenceName = e.formInput.newSequenceName ? e.formInput.newSequenceName.trim() : "";

  if (!sequenceName) {
    return createNotification("Please enter a sequence name.");
  }

  // Validate sequence name (no special characters that would break sheet names)
  if (!/^[a-zA-Z0-9\s\-&\/]+$/.test(sequenceName)) {
    return createNotification("Sequence name can only contain letters, numbers, spaces, hyphens, ampersands, and forward slashes.");
  }

  // Check if sequence already exists
  const availableSequences = getAvailableSequences();
  const normalizedName = sequenceName.toLowerCase();
  for (const existing of availableSequences) {
    if (existing.toLowerCase() === normalizedName) {
      return createNotification(`A sequence named "${sequenceName}" already exists. Please choose a different name.`);
    }
  }

  // Show step count selection card
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .pushCard(buildStepCountSelectionCard(sequenceName)))
    .build();
}

/**
 * Builds the step count selection card (Step 1 of wizard)
 */
function buildStepCountSelectionCard(sequenceName) {
  const card = CardService.newCardBuilder();

  card.setHeader(CardService.newCardHeader()
      .setTitle("📧 New Sequence Setup")
      .setSubtitle(`Creating: ${sequenceName}`)
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/playlist_add_black_48dp.png"));

  // Step indicator
  const progressSection = CardService.newCardSection();
  progressSection.addWidget(CardService.newTextParagraph()
      .setText("<b>Step 1 of 2:</b> Choose number of emails"));
  card.addSection(progressSection);

  // Step count selection
  const selectionSection = CardService.newCardSection()
      .setHeader("How many emails in this sequence?");

  selectionSection.addWidget(CardService.newTextParagraph()
      .setText("Select how many steps (emails) you want in your outreach sequence:"));

  const stepCountDropdown = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setFieldName("stepCount")
      .setTitle("Number of Emails");

  stepCountDropdown.addItem("1 email (single outreach)", "1", false);
  stepCountDropdown.addItem("2 emails (1 + 1 follow-up)", "2", false);
  stepCountDropdown.addItem("3 emails (1 + 2 follow-ups)", "3", true); // Default
  stepCountDropdown.addItem("4 emails (1 + 3 follow-ups)", "4", false);
  stepCountDropdown.addItem("5 emails (1 + 4 follow-ups)", "5", false);

  selectionSection.addWidget(stepCountDropdown);

  selectionSection.addWidget(CardService.newTextParagraph()
      .setText("<font color='#666666'>💡 <i>Tip: 3-4 emails is typical for most outreach sequences</i></font>"));

  selectionSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("Next: Write Emails →")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setBackgroundColor("#1a73e8")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("proceedToTemplateCreation")
              .setParameters({ sequenceName: sequenceName }))));

  card.addSection(selectionSection);

  // Cancel/Back
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
      .setText("← Cancel")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("buildTemplateManagementCard")));
  card.addSection(navSection);

  return card.build();
}

/**
 * Proceeds from step count selection to template creation
 */
function proceedToTemplateCreation(e) {
  const sequenceName = e.parameters.sequenceName;
  const stepCount = parseInt(e.formInput.stepCount) || 3;

  // Store wizard state in cache
  const wizardState = {
    sequenceName: sequenceName,
    stepCount: stepCount,
    templates: {} // Will hold {1: {name, subject, body}, 2: {...}, etc.}
  };

  // Store in user properties temporarily
  PropertiesService.getUserProperties().setProperty("WIZARD_STATE", JSON.stringify(wizardState));

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .pushCard(buildSequenceWizardTemplatesCard(wizardState)))
    .build();
}

/**
 * Builds the template creation card (Step 2 of wizard)
 * Shows all steps with "Create Email X" buttons - no defaults
 */
function buildSequenceWizardTemplatesCard(wizardState) {
  // If wizardState not passed, load from properties
  if (!wizardState || typeof wizardState === 'string') {
    const stateJson = PropertiesService.getUserProperties().getProperty("WIZARD_STATE");
    if (!stateJson) {
      return buildTemplateManagementCard();
    }
    wizardState = JSON.parse(stateJson);
  }

  const card = CardService.newCardBuilder();
  const sequenceName = wizardState.sequenceName;
  const stepCount = wizardState.stepCount;
  const templates = wizardState.templates || {};

  card.setHeader(CardService.newCardHeader()
      .setTitle("📧 " + sequenceName)
      .setSubtitle(`${stepCount} email${stepCount > 1 ? 's' : ''} in sequence`)
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/drafts_black_48dp.png"));

  // Step indicator
  const progressSection = CardService.newCardSection();
  const completedCount = Object.keys(templates).filter(k => templates[k] && templates[k].body).length;
  progressSection.addWidget(CardService.newTextParagraph()
      .setText(`<b>Step 2 of 2:</b> Write your emails (${completedCount}/${stepCount} complete)`));
  card.addSection(progressSection);

  // Email steps
  for (let step = 1; step <= stepCount; step++) {
    const template = templates[step];
    const hasContent = template && template.body && template.body.trim();
    
    const stepIcon = step === 1 ? "📤" : "↩️";
    const stepType = step === 1 ? "First Email" : "Follow-up #" + (step - 1);
    
    const stepSection = CardService.newCardSection()
        .setHeader(`${stepIcon} Email ${step}: ${stepType}`);

    if (hasContent) {
      // Show preview of created template
      stepSection.addWidget(CardService.newDecoratedText()
          .setTopLabel("✅ CREATED")
          .setText(template.name || `Email ${step}`)
          .setBottomLabel(step === 1 ? `Subject: ${template.subject}` : "(Reply to Email 1)")
          .setWrapText(true));
      
      // Preview snippet
      const bodyPreview = template.body.length > 80 
        ? template.body.substring(0, 80).replace(/\n/g, " ") + "..." 
        : template.body.replace(/\n/g, " ");
      stepSection.addWidget(CardService.newTextParagraph()
          .setText(`<font color='#666666'>${bodyPreview}</font>`));

      stepSection.addWidget(CardService.newTextButton()
          .setText("✏️ Edit Email " + step)
          .setOnClickAction(CardService.newAction()
              .setFunctionName("editWizardStepTemplate")
              .setParameters({ step: step.toString() })));
    } else {
      // Show create button - NOT created yet
      stepSection.addWidget(CardService.newDecoratedText()
          .setTopLabel("⚠️ NOT CREATED")
          .setText(step === 1 ? "First outreach email" : `Follow-up email #${step - 1}`)
          .setBottomLabel("Required - click below to write")
          .setWrapText(true));

      stepSection.addWidget(CardService.newTextButton()
          .setText("✉️ Create Email " + step)
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setBackgroundColor("#ea4335")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("editWizardStepTemplate")
              .setParameters({ step: step.toString() })));
    }

    card.addSection(stepSection);
  }

  // Action buttons
  const actionSection = CardService.newCardSection();
  
  // Check if all templates are complete
  const allComplete = completedCount === stepCount;
  
  if (allComplete) {
    actionSection.addWidget(CardService.newTextParagraph()
        .setText("<font color='#137333'><b>✅ All emails created! Ready to save.</b></font>"));
    
    actionSection.addWidget(CardService.newTextButton()
        .setText("💾 Save Sequence")
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setBackgroundColor("#137333")
        .setOnClickAction(CardService.newAction()
            .setFunctionName("finalizeSequenceCreation")));
  } else {
    actionSection.addWidget(CardService.newTextParagraph()
        .setText(`<font color='#ea4335'><b>⚠️ ${stepCount - completedCount} email${stepCount - completedCount > 1 ? 's' : ''} still needed</b></font>\n` +
                 `<font color='#666666'>Create all emails before saving.</font>`));
    
    // Disabled-looking save button with explanation
    actionSection.addWidget(CardService.newTextButton()
        .setText("💾 Save Sequence (Complete All Emails First)")
        .setOnClickAction(CardService.newAction()
            .setFunctionName("showIncompleteWarning")));
  }
  
  card.addSection(actionSection);

  // Cancel/Back
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("← Back")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("goBackToStepCount")))
      .addButton(CardService.newTextButton()
          .setText("Cancel")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("cancelSequenceWizard"))));
  card.addSection(navSection);

  return card.build();
}

/**
 * Shows warning when trying to save incomplete sequence
 */
function showIncompleteWarning() {
  return createNotification("Please create all email templates before saving the sequence.");
}

/**
 * Goes back to step count selection
 */
function goBackToStepCount(e) {
  const stateJson = PropertiesService.getUserProperties().getProperty("WIZARD_STATE");
  if (!stateJson) {
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation()
        .popCard())
      .build();
  }
  
  const wizardState = JSON.parse(stateJson);
  
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .popCard())
    .build();
}

/**
 * Cancels the sequence creation wizard
 */
function cancelSequenceWizard() {
  // Clear wizard state
  PropertiesService.getUserProperties().deleteProperty("WIZARD_STATE");
  
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .popToRoot()
      .updateCard(buildTemplateManagementCard()))
    .build();
}

/**
 * Edit a specific step template in the wizard
 */
function editWizardStepTemplate(e) {
  const step = parseInt(e.parameters.step);
  
  const stateJson = PropertiesService.getUserProperties().getProperty("WIZARD_STATE");
  if (!stateJson) {
    return createNotification("Session expired. Please start over.");
  }
  
  const wizardState = JSON.parse(stateJson);
  const template = wizardState.templates[step] || {};
  
  const card = CardService.newCardBuilder();
  
  const stepType = step === 1 ? "First Email" : "Follow-up #" + (step - 1);
  card.setHeader(CardService.newCardHeader()
      .setTitle(`✏️ Email ${step}: ${stepType}`)
      .setSubtitle(wizardState.sequenceName)
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/edit_black_48dp.png"));

  // Variable Dictionary
  const varsSection = CardService.newCardSection()
      .setHeader("📋 Available Variables");
  
  varsSection.addWidget(CardService.newTextParagraph()
      .setText(
        "<b>Contact Info:</b> " +
        "<font color='#1a73e8'>{{firstName}}</font> · " +
        "<font color='#1a73e8'>{{lastName}}</font> · " +
        "<font color='#1a73e8'>{{company}}</font> · " +
        "<font color='#1a73e8'>{{title}}</font>\n\n" +
        "<b>Your Info:</b> " +
        "<font color='#1a73e8'>{{senderName}}</font> · " +
        "<font color='#1a73e8'>{{senderCompany}}</font>"
      ));
  card.addSection(varsSection);
  
  // Email form
  const formSection = CardService.newCardSection()
      .setHeader("Write Your Email");
  
  formSection.addWidget(CardService.newTextInput()
      .setFieldName("templateName")
      .setTitle("Email Name (for your reference)")
      .setValue(template.name || getDefaultStepName(step))
      .setHint("e.g., Introduction, First Follow-up"));
  
  // Only show subject for Step 1
  if (step === 1) {
    formSection.addWidget(CardService.newTextInput()
        .setFieldName("subject")
        .setTitle("Subject Line *")
        .setValue(template.subject || "")
        .setHint("e.g., Quick question about {{company}}"));
  } else {
    formSection.addWidget(CardService.newTextParagraph()
        .setText("<font color='#666666'>ℹ️ <i>This email will be sent as a reply to Email 1 (uses \"Re: [subject]\")</i></font>"));
  }
  
  // Pre-populate body with greeting if empty
  const defaultBody = "Hi {{firstName}},\n\n";
  formSection.addWidget(CardService.newTextInput()
      .setFieldName("body")
      .setTitle("Email Body *")
      .setValue(template.body || defaultBody)
      .setMultiline(true)
      .setHint("Write your message after the greeting"));
  
  formSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("💾 Save Email " + step)
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setBackgroundColor("#1a73e8")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("saveWizardStepTemplate")
              .setParameters({ step: step.toString() }))));
  
  card.addSection(formSection);
  
  // Navigation
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
      .setText("← Back to Sequence")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("returnToWizardTemplatesCard")));
  card.addSection(navSection);
  
  return card.build();
}

/**
 * Saves a step template in the wizard
 */
function saveWizardStepTemplate(e) {
  const step = parseInt(e.parameters.step);
  const templateName = e.formInput.templateName || getDefaultStepName(step);
  const subject = e.formInput.subject || "";
  const body = e.formInput.body || "";
  
  // Validation
  if (!body || !body.trim()) {
    return createNotification("Please write your email body.");
  }
  
  if (step === 1 && (!subject || !subject.trim())) {
    return createNotification("Email 1 requires a subject line.");
  }
  
  // Load wizard state
  const stateJson = PropertiesService.getUserProperties().getProperty("WIZARD_STATE");
  if (!stateJson) {
    return createNotification("Session expired. Please start over.");
  }
  
  const wizardState = JSON.parse(stateJson);
  
  // Save template to wizard state
  if (!wizardState.templates) {
    wizardState.templates = {};
  }
  
  wizardState.templates[step] = {
    name: templateName.trim(),
    subject: subject.trim(),
    body: body.trim()
  };
  
  // Save updated state
  PropertiesService.getUserProperties().setProperty("WIZARD_STATE", JSON.stringify(wizardState));
  
  // Return to templates card
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText(`✓ Email ${step} saved!`))
    .setNavigation(CardService.newNavigation()
      .popCard()
      .updateCard(buildSequenceWizardTemplatesCard(wizardState)))
    .build();
}

/**
 * Returns to the wizard templates card
 */
function returnToWizardTemplatesCard() {
  const stateJson = PropertiesService.getUserProperties().getProperty("WIZARD_STATE");
  if (!stateJson) {
    return CardService.newActionResponseBuilder()
      .setNavigation(CardService.newNavigation()
        .popToRoot()
        .updateCard(buildTemplateManagementCard()))
      .build();
  }
  
  const wizardState = JSON.parse(stateJson);
  
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .popCard()
      .updateCard(buildSequenceWizardTemplatesCard(wizardState)))
    .build();
}

/**
 * Finalizes sequence creation - validates all templates and creates the sequence
 */
function finalizeSequenceCreation() {
  const stateJson = PropertiesService.getUserProperties().getProperty("WIZARD_STATE");
  if (!stateJson) {
    return createNotification("Session expired. Please start over.");
  }
  
  const wizardState = JSON.parse(stateJson);
  const sequenceName = wizardState.sequenceName;
  const stepCount = wizardState.stepCount;
  const templates = wizardState.templates || {};
  
  // Validate all templates are complete
  for (let step = 1; step <= stepCount; step++) {
    const template = templates[step];
    if (!template || !template.body || !template.body.trim()) {
      return createNotification(`Email ${step} is missing. Please create all emails before saving.`);
    }
    if (step === 1 && (!template.subject || !template.subject.trim())) {
      return createNotification("Email 1 requires a subject line.");
    }
  }
  
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    return createNotification("No database connected.");
  }
  
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheetName = getSequenceSheetName(sequenceName);
    
    // Check if sheet already exists
    if (spreadsheet.getSheetByName(sheetName)) {
      return createNotification(`A sequence named "${sequenceName}" already exists.`);
    }
    
    // Create new sheet
    const sheet = spreadsheet.insertSheet(sheetName);
    
    // Set up headers: Step, Name, Subject, Body, (empty), Variables, Instructions
    const headers = ["Step", "Name", "Subject", "Body", "", "📋 VARIABLES", "📖 INSTRUCTIONS"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    sheet.setFrozenRows(1);
    
    // All available variables for reference
    const variablesRef = "{{firstName}}\n{{lastName}}\n{{email}}\n{{company}}\n{{title}}\n{{industry}}\n{{senderName}}\n{{senderCompany}}\n{{senderTitle}}";
    const mainInstructions = "⚠️ DO NOT EDIT columns A-D directly unless you know what you're doing!\n\n✅ HOW TO EDIT:\n1. Edit Subject (column C) and Body (column D) only\n2. Use variables from column F by copying them\n3. Keep Step numbers in column A as 1,2,3,4,5\n\n❌ DO NOT:\n• Delete or rename this sheet\n• Change column headers\n• Add extra columns between A-D";
    const replyInstructions = "Step 2+ are REPLY emails.\nThey use 'Re: [Step 1 subject]' automatically.\nLeave the Subject column empty for these steps.";
    
    // Add template rows (only for the configured step count)
    const templateRows = [];
    for (let step = 1; step <= stepCount; step++) {
      const template = templates[step];
      templateRows.push([
        step,
        template.name || getDefaultStepName(step),
        template.subject || "",
        template.body || "",
        "", // Empty separator column
        step === 1 ? variablesRef : "", // Variables only in first row
        step === 1 ? mainInstructions : (step === 2 ? replyInstructions : "") // Instructions
      ]);
    }
    
    sheet.getRange(2, 1, templateRows.length, headers.length).setValues(templateRows);
    
    // Format columns
    sheet.autoResizeColumn(1);
    sheet.autoResizeColumn(2);
    sheet.setColumnWidth(3, 300);
    sheet.setColumnWidth(4, 500);
    sheet.setColumnWidth(5, 20);  // Empty separator column
    sheet.setColumnWidth(6, 150); // Variables column
    sheet.setColumnWidth(7, 300); // Instructions column
    
    // Style the variables and instructions columns
    sheet.getRange(1, 6, 1, 2).setBackground("#e8f0fe").setFontColor("#1a73e8"); // Header style
    sheet.getRange(2, 6, templateRows.length, 1).setBackground("#f8f9fa").setFontFamily("Courier New");
    sheet.getRange(2, 7, templateRows.length, 1).setBackground("#fff8e1").setFontStyle("italic");
    
    // Set word wrap for body and instructions
    sheet.getRange(2, 4, templateRows.length, 1).setWrap(true);
    sheet.getRange(2, 7, templateRows.length, 1).setWrap(true);
    
    SpreadsheetApp.flush();
    
    // Set step count for this sequence
    setSequenceStepCount(sequenceName, stepCount);
    
    // Clear wizard state
    PropertiesService.getUserProperties().deleteProperty("WIZARD_STATE");
    
    logAction("Create Sequence", `Created new sequence "${sequenceName}" with ${stepCount} emails`);
    
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText(`✅ Sequence "${sequenceName}" created with ${stepCount} emails!`))
      .setNavigation(CardService.newNavigation()
        .popToRoot()
        .updateCard(buildTemplateManagementCard()))
      .build();
      
  } catch (error) {
    console.error("Error creating sequence: " + error + "\n" + error.stack);
    logAction("Error", `Error creating sequence '${sequenceName}': ${error.toString()}`);
    return createNotification("Error creating sequence: " + error.message);
  }
}

/**
 * Legacy function - redirects to wizard
 * Kept for backwards compatibility
 */
function createNewSequence(e) {
  return startSequenceCreationWizard(e);
}

/* ======================== EDIT IN SHEETS ======================== */

/**
 * Opens the specific sequence sheet tab in Google Sheets
 * @param {object} e Event object with sequence parameter
 * @returns {ActionResponse} Opens the sheet in a new tab
 */
function openSequenceInSheets(e) {
  const sequenceName = e.parameters.sequence;
  
  if (!sequenceName) {
    return createNotification("Error: No sequence specified.");
  }
  
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    return createNotification("No database connected.");
  }
  
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheetName = getSequenceSheetName(sequenceName);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      return createNotification(`Sequence sheet "${sequenceName}" not found.`);
    }
    
    // Get the sheet's gid (sheet ID) to open the correct tab
    const sheetId = sheet.getSheetId();
    const spreadsheetUrl = spreadsheet.getUrl();
    
    // Construct URL with the specific sheet tab
    const sheetUrl = spreadsheetUrl + "#gid=" + sheetId;
    
    // Return an action response that opens the URL
    return CardService.newActionResponseBuilder()
      .setOpenLink(CardService.newOpenLink()
          .setUrl(sheetUrl)
          .setOpenAs(CardService.OpenAs.FULL_SIZE)
          .setOnClose(CardService.OnClose.NOTHING))
      .setNotification(CardService.newNotification()
          .setText(`Opening "${sequenceName}" in Google Sheets...`))
      .build();
      
  } catch (error) {
    console.error("Error opening sequence in sheets: " + error);
    return createNotification("Error opening sheet: " + error.message);
  }
}

/* ======================== SEQUENCE DUPLICATION ======================== */

/**
 * Duplicates an existing sequence with all its templates
 * Creates a new sequence named "Copy-[original name]"
 */
function duplicateSequence(e) {
  const sourceSequenceName = e.parameters.sequence;
  
  if (!sourceSequenceName) {
    return createNotification("Error: No sequence specified.");
  }
  
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    return createNotification("No database connected.");
  }
  
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sourceSheetName = getSequenceSheetName(sourceSequenceName);
    const sourceSheet = spreadsheet.getSheetByName(sourceSheetName);
    
    if (!sourceSheet) {
      return createNotification(`Source sequence "${sourceSequenceName}" not found.`);
    }
    
    // Generate unique copy name
    let copyNumber = 1;
    let newSequenceName = `Copy-${sourceSequenceName}`;
    let newSheetName = getSequenceSheetName(newSequenceName);
    
    // Keep incrementing if copy already exists
    while (spreadsheet.getSheetByName(newSheetName)) {
      copyNumber++;
      newSequenceName = `Copy${copyNumber}-${sourceSequenceName}`;
      newSheetName = getSequenceSheetName(newSequenceName);
    }
    
    // Duplicate the sheet
    const newSheet = sourceSheet.copyTo(spreadsheet);
    newSheet.setName(newSheetName);
    
    // Copy step count setting
    const sourceStepCount = getSequenceStepCount(sourceSequenceName);
    setSequenceStepCount(newSequenceName, sourceStepCount);
    
    SpreadsheetApp.flush();
    
    logAction("Duplicate Sequence", `Duplicated "${sourceSequenceName}" as "${newSequenceName}"`);
    
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText(`✅ Created "${newSequenceName}" - click Edit to rename and customize`))
      .setNavigation(CardService.newNavigation()
        .updateCard(buildTemplateManagementCard()))
      .build();
      
  } catch (error) {
    console.error("Error duplicating sequence: " + error + "\n" + error.stack);
    logAction("Error", `Error duplicating sequence '${sourceSequenceName}': ${error.toString()}`);
    return createNotification("Error duplicating sequence: " + error.message);
  }
}

/**
 * Shows the rename sequence card
 */
function showRenameSequenceCard(e) {
  const sequenceName = e.parameters.sequence;
  
  const card = CardService.newCardBuilder();
  
  card.setHeader(CardService.newCardHeader()
      .setTitle("✏️ Rename Sequence")
      .setSubtitle(`Current: ${sequenceName}`)
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/edit_black_48dp.png"));
  
  const formSection = CardService.newCardSection();
  
  formSection.addWidget(CardService.newTextInput()
      .setFieldName("newSequenceName")
      .setTitle("New Sequence Name")
      .setValue(sequenceName)
      .setHint("Enter a new name for this sequence"));
  
  formSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("💾 Save New Name")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setBackgroundColor("#1a73e8")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("renameSequence")
              .setParameters({ oldName: sequenceName })))
      .addButton(CardService.newTextButton()
          .setText("Cancel")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("showSequenceTemplatesEditor")
              .setParameters({ sequence: sequenceName }))));
  
  card.addSection(formSection);
  
  return card.build();
}

/**
 * Renames a sequence (sheet name and step count property)
 */
function renameSequence(e) {
  const oldName = e.parameters.oldName;
  const newName = e.formInput.newSequenceName ? e.formInput.newSequenceName.trim() : "";
  
  if (!newName) {
    return createNotification("Please enter a new sequence name.");
  }
  
  if (newName === oldName) {
    return createNotification("The new name is the same as the current name.");
  }
  
  // Validate sequence name
  if (!/^[a-zA-Z0-9\s\-&\/]+$/.test(newName)) {
    return createNotification("Sequence name can only contain letters, numbers, spaces, hyphens, ampersands, and forward slashes.");
  }
  
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    return createNotification("No database connected.");
  }
  
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const oldSheetName = getSequenceSheetName(oldName);
    const newSheetName = getSequenceSheetName(newName);
    
    // Check if target name already exists
    if (spreadsheet.getSheetByName(newSheetName)) {
      return createNotification(`A sequence named "${newName}" already exists.`);
    }
    
    const sheet = spreadsheet.getSheetByName(oldSheetName);
    if (!sheet) {
      return createNotification(`Sequence "${oldName}" not found.`);
    }
    
    // Rename the sheet
    sheet.setName(newSheetName);
    
    // Copy step count to new key and delete old key
    const oldStepCount = getSequenceStepCount(oldName);
    setSequenceStepCount(newName, oldStepCount);
    
    // Delete old step count property
    const oldKey = "SEQUENCE_STEPS_" + oldName.replace(/\s+/g, "_");
    PropertiesService.getUserProperties().deleteProperty(oldKey);
    
    SpreadsheetApp.flush();
    
    logAction("Rename Sequence", `Renamed "${oldName}" to "${newName}"`);
    
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText(`✅ Sequence renamed to "${newName}"`))
      .setNavigation(CardService.newNavigation()
        .popToRoot()
        .updateCard(buildTemplateManagementCard()))
      .build();
      
  } catch (error) {
    console.error("Error renaming sequence: " + error + "\n" + error.stack);
    logAction("Error", `Error renaming sequence: ${error.toString()}`);
    return createNotification("Error renaming sequence: " + error.message);
  }
}

/**
 * Shows confirmation before deleting a sequence
 */
function confirmDeleteSequence(e) {
  const sequenceName = e.parameters.sequence;
  
  const card = CardService.newCardBuilder();
  
  card.setHeader(CardService.newCardHeader()
      .setTitle("⚠️ Delete Sequence?")
      .setSubtitle(sequenceName)
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/warning_amber_48dp.png"));
  
  const section = CardService.newCardSection();
  
  section.addWidget(CardService.newTextParagraph()
      .setText(`<font color='#ea4335'><b>This will permanently delete:</b></font>\n\n` +
               `• The sequence "${sequenceName}"\n` +
               `• All email templates in this sequence\n\n` +
               `<b>This action cannot be undone.</b>`));
  
  section.addWidget(CardService.newTextParagraph()
      .setText(`<font color='#666666'>⚠️ Contacts assigned to this sequence will keep their sequence assignment but won't receive emails.</font>`));
  
  section.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("🗑️ Delete Permanently")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setBackgroundColor("#ea4335")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("deleteSequence")
              .setParameters({ sequence: sequenceName })))
      .addButton(CardService.newTextButton()
          .setText("Cancel")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("showSequenceTemplatesEditor")
              .setParameters({ sequence: sequenceName }))));
  
  card.addSection(section);
  
  return card.build();
}

/**
 * Deletes a sequence (removes sheet and step count property)
 */
function deleteSequence(e) {
  const sequenceName = e.parameters.sequence;
  
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    return createNotification("No database connected.");
  }
  
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheetName = getSequenceSheetName(sequenceName);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      return createNotification(`Sequence "${sequenceName}" not found.`);
    }
    
    // Delete the sheet
    spreadsheet.deleteSheet(sheet);
    
    // Delete step count property
    const key = "SEQUENCE_STEPS_" + sequenceName.replace(/\s+/g, "_");
    PropertiesService.getUserProperties().deleteProperty(key);
    
    SpreadsheetApp.flush();
    
    logAction("Delete Sequence", `Deleted sequence "${sequenceName}"`);
    
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText(`✅ Sequence "${sequenceName}" deleted`))
      .setNavigation(CardService.newNavigation()
        .popToRoot()
        .updateCard(buildTemplateManagementCard()))
      .build();
      
  } catch (error) {
    console.error("Error deleting sequence: " + error + "\n" + error.stack);
    logAction("Error", `Error deleting sequence: ${error.toString()}`);
    return createNotification("Error deleting sequence: " + error.message);
  }
}

/* ======================== TEMPLATE OPERATIONS (LEGACY) ======================== */

/**
 * Saves the template selection for a step in User Properties.
 */
function saveTemplateSelection(e) {
  const step = e.parameters.step;
  const templateName = e.formInput.selected_template; // Get selected value from dropdown

  if (!step || !templateName) {
    console.error("saveTemplateSelection: Missing step or template name parameter/input.");
    return createNotification("Could not save template selection. Missing information.");
  }

  // Save the selection to user properties, keyed by step
  PropertiesService.getUserProperties().setProperty("TEMPLATE_SELECTION_STEP_" + step, templateName);
  logAction("Template Selection", `Saved template '${templateName}' for Step ${step}`);

  // Rebuild the same card to reflect the change in preview and dropdown selection
  // Pass back the current page number if available
  const page = e.parameters.page || '1';

  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification()
      .setText("Template selection saved for Step " + step))
    .setNavigation(CardService.newNavigation()
      // Rebuild the card that shows contacts for this step
      .updateCard(buildSelectContactsCard({ parameters: { step: step, page: page } })))
    .build();
}

/**
 * Shows the template editor card.
 */
function editTemplate(e) {
  const templateName = e.parameters.templateName;

  // Find the template
  const templates = getEmailTemplates();
  const template = templates.find(t => t.name === templateName);

  if (!template) {
    logAction("Error", `Edit Template Error: Template not found '${templateName}'`);
    return createNotification("Template not found: " + templateName);
  }

  const card = CardService.newCardBuilder();

  // Add header
  card.setHeader(CardService.newCardHeader()
      .setTitle("Edit Template")
      .setSubtitle(template.name) // Show name in subtitle
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/edit_black_48dp.png"));

  // --- Editor Section ---
  const editorSection = CardService.newCardSection();
      // Removed header here as it's in the main card header

   // Template Name (Read-only - changing name requires delete/create)
    editorSection.addWidget(CardService.newKeyValue()
         .setTopLabel("Template Name (Read-only)")
         .setContent(template.name));


  editorSection.addWidget(CardService.newTextInput()
      .setFieldName("templateSubject")
      .setTitle("Subject Line")
      .setValue(template.subject));
  editorSection.addWidget(CardService.newTextInput()
      .setMultiline(true)
      .setFieldName("templateBody")
      .setTitle("Email Body")
      .setHint("Use placeholders like {{firstName}}")
      .setValue(template.body));

  editorSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("Save Changes")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("saveTemplateChanges")
              .setParameters({ templateName: template.name }))) // Pass original name for lookup
      .addButton(CardService.newTextButton()
          .setText("Delete Template")
           .setTextButtonStyle(CardService.TextButtonStyle.TEXT) // Less prominent
          .setOnClickAction(CardService.newAction()
              .setFunctionName("deleteTemplateConfirmation") // Go to confirmation step
              .setParameters({ templateName: template.name }))));
  card.addSection(editorSection);


  // --- Live Preview Section --- (Optional but useful)
  const previewSection = CardService.newCardSection()
      .setCollapsible(true)
      .setHeader("Live Preview (using Sample Data)");

  // Create a sample contact for preview
  const sampleContact = {
    firstName: "Alex", lastName: "Sample", email: "alex.sample@example.com",
    company: "Demo Company Inc.", title: "Product Manager",
    industry: "SaaS / B2B Tech", // Use dedicated industry field
    notes: "department: Product" // Example notes format
  };
  // Get sender info for preview
  const userProps = PropertiesService.getUserProperties();
  const senderInfo = {
    name: userProps.getProperty("SENDER_NAME") || "Your Name",
    company: userProps.getProperty("SENDER_COMPANY") || "Your Company",
    title: userProps.getProperty("SENDER_TITLE") || "Your Title"
  };

  // Preview with placeholders replaced using current values (might be slightly off if user hasn't saved yet)
  const previewSubject = replacePlaceholders(template.subject, sampleContact, senderInfo);
  const previewBody = replacePlaceholders(template.body, sampleContact, senderInfo);

  previewSection.addWidget(CardService.newKeyValue()
      .setTopLabel("Preview Subject")
      .setContent(previewSubject)
      .setMultiline(true));
  previewSection.addWidget(CardService.newTextParagraph()
      .setText("<b>Preview Body:</b>\n\n" + previewBody.replace(/\n/g, "<br>"))); // Show with line breaks
  card.addSection(previewSection);

   // --- Placeholders Help Section (Repeat from main view for convenience) ---
    const helpSection = CardService.newCardSection()
        .setCollapsible(true)
        .setHeader("Available Placeholders");
    const placeholders = [
        "{{firstName}}", "{{lastName}}", "{{email}}",
        "{{company}}", "{{title}}",
        "{{senderName}}", "{{senderCompany}}", "{{senderTitle}}",
        "{{industry}} (from Industry field)",
        "{{department}} (from Notes: 'department: ...')"
    ];
    helpSection.addWidget(CardService.newTextParagraph()
        .setText("Use these placeholders in Subject and Body:\n" + placeholders.join("\n")));
    card.addSection(helpSection);


  // --- Navigation Section ---
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
      .setText("Back to Template Management")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("buildTemplateManagementCard")));
  card.addSection(navSection);

  return card.build();
}

/* ======================== TEMPLATE CRUD OPERATIONS ======================== */

/**
 * Creates a new template.
 */
function createTemplate(e) {
  const name = e.formInput.newTemplateName ? e.formInput.newTemplateName.trim() : "";
  const subject = e.formInput.newTemplateSubject || ""; // Allow empty subject/body initially
  const body = e.formInput.newTemplateBody || "";

  if (!name) {
    return createNotification("Template Name cannot be empty.");
  }

  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    return createNotification("No database connected.");
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const templatesSheet = spreadsheet.getSheetByName(CONFIG.TEMPLATES_SHEET_NAME);

    if (!templatesSheet) {
      logAction("Error", "Create Template Error: Templates sheet not found.");
      return createNotification("Templates sheet not found.");
    }

    // Check if template name already exists (case-insensitive check)
    const data = templatesSheet.getDataRange().getValues();
    const normalizedNewName = name.toLowerCase();
    for (let i = 1; i < data.length; i++) {
       const existingName = data[i][0];
      if (existingName && existingName.toLowerCase() === normalizedNewName) {
        return createNotification(`A template named "${name}" already exists. Please choose a unique name.`);
      }
    }

    // Add the new template
    templatesSheet.appendRow([name, subject, body]);
     SpreadsheetApp.flush();

    logAction("Create Template", "Created template: " + name);

    // Refresh the template management card
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText(`Template "${name}" created successfully!`))
      .setNavigation(CardService.newNavigation()
        .updateCard(buildTemplateManagementCard()))
      .build();
  } catch (error) {
    console.error("Error creating template: " + error + "\n" + error.stack);
    logAction("Error", `Error creating template '${name}': ${error.toString()}`);
    return createNotification("Error creating template: " + error.message);
  }
}

/**
 * Saves changes to an existing template's subject and body.
 */
function saveTemplateChanges(e) {
  const templateName = e.parameters.templateName; // Original name passed as parameter
  const newSubject = e.formInput.templateSubject || "";
  const newBody = e.formInput.templateBody || "";

   // No need to validate emptiness here, user might want blank templates

  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    return createNotification("No database connected.");
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const templatesSheet = spreadsheet.getSheetByName(CONFIG.TEMPLATES_SHEET_NAME);

    if (!templatesSheet) {
      logAction("Error", `Save Template Error: Templates sheet not found for '${templateName}'`);
      return createNotification("Templates sheet not found.");
    }

    // Find the template by its original name (case-sensitive find for safety)
    const data = templatesSheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === templateName) {
        rowIndex = i + 1; // 1-based index
        break;
      }
    }

    if (rowIndex === -1) {
       logAction("Error", `Save Template Error: Template '${templateName}' not found for update.`);
      return createNotification("Template not found: " + templateName + ". It might have been deleted.");
    }

    // Update the subject (column 2) and body (column 3)
    templatesSheet.getRange(rowIndex, 2).setValue(newSubject);
    templatesSheet.getRange(rowIndex, 3).setValue(newBody);
     SpreadsheetApp.flush();

    logAction("Update Template", "Updated template: " + templateName);

    // Go back to the main template management view
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText(`Template "${templateName}" updated successfully!`))
      .setNavigation(CardService.newNavigation()
        .updateCard(buildTemplateManagementCard()))
      .build();
  } catch (error) {
    console.error(`Error updating template '${templateName}': ` + error + "\n" + error.stack);
    logAction("Error", `Error updating template '${templateName}': ${error.toString()}`);
    return createNotification("Error updating template: " + error.message);
  }
}

/**
 * Shows a confirmation card before deleting a template.
 */
function deleteTemplateConfirmation(e) {
    const templateName = e.parameters.templateName;

    const card = CardService.newCardBuilder();
    card.setHeader(CardService.newCardHeader()
        .setTitle("Confirm Deletion")
        .setSubtitle(`Template: ${templateName}`)
        .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/warning_amber_48dp.png")); // Warning icon

    const section = CardService.newCardSection();
    section.addWidget(CardService.newTextParagraph()
        .setText(`Are you sure you want to permanently delete the template "${templateName}"? This action cannot be undone.`));

    section.addWidget(CardService.newButtonSet()
        .addButton(CardService.newTextButton()
            .setText("Delete Permanently")
            .setTextButtonStyle(CardService.TextButtonStyle.FILLED) // Make delete button stand out
            .setOnClickAction(CardService.newAction()
                .setFunctionName("deleteTemplate") // Actual delete function
                .setParameters({ templateName: templateName })))
        .addButton(CardService.newTextButton()
            .setText("Cancel")
            .setOnClickAction(CardService.newAction()
                .setFunctionName("editTemplate") // Go back to the editor
                .setParameters({ templateName: templateName }))));

    card.addSection(section);

     // Add navigation section
     const navSection = CardService.newCardSection();
     navSection.addWidget(CardService.newTextButton()
          .setText("Back to Template Management")
          .setOnClickAction(CardService.newAction()
                 .setFunctionName("buildTemplateManagementCard")));
     card.addSection(navSection);


    return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().updateCard(card.build()))
        .build();
}

/**
 * Deletes a template after confirmation.
 */
function deleteTemplate(e) {
    const templateName = e.parameters.templateName;

    const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
    if (!spreadsheetId) {
        return createNotification("No database connected.");
    }

    try {
        const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        const templatesSheet = spreadsheet.getSheetByName(CONFIG.TEMPLATES_SHEET_NAME);

        if (!templatesSheet) {
            logAction("Error", `Delete Template Error: Templates sheet not found for '${templateName}'`);
            return createNotification("Templates sheet not found.");
        }

        // Find the template row
        const data = templatesSheet.getDataRange().getValues();
        let rowIndexToDelete = -1;
        for (let i = 1; i < data.length; i++) { // Start from row 2 (index 1)
            if (data[i][0] === templateName) {
                rowIndexToDelete = i + 1; // 1-based index
                break;
            }
        }

        if (rowIndexToDelete === -1) {
            logAction("Warning", `Delete Template Warning: Template '${templateName}' not found for deletion (already deleted?).`);
            return createNotification(`Template "${templateName}" not found. It might have already been deleted.`);
        }

        // Delete the row
        templatesSheet.deleteRow(rowIndexToDelete);
        SpreadsheetApp.flush();

        logAction("Delete Template", "Deleted template: " + templateName);

        // Go back to the main template management view
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification()
                .setText(`Template "${templateName}" deleted successfully.`))
            .setNavigation(CardService.newNavigation()
                .updateCard(buildTemplateManagementCard()))
            .build();

    } catch (error) {
        console.error(`Error deleting template '${templateName}': ` + error + "\n" + error.stack);
        logAction("Error", `Error deleting template '${templateName}': ${error.toString()}`);
        return createNotification("Error deleting template: " + error.message);
    }
}

/* ======================== IN-SIDEBAR TEMPLATE EDITOR ======================== */

/**
 * Shows all steps for a sequence - edit each step's email
 * Includes step count configuration and delete functionality
 */
function showSequenceTemplatesEditor(e) {
  const sequenceName = e.parameters.sequence;
  const card = CardService.newCardBuilder();

  card.setHeader(CardService.newCardHeader()
      .setTitle("📧 " + sequenceName)
      .setSubtitle("Edit your email sequence"));

  const templates = getEmailTemplates(sequenceName);
  const currentStepCount = getSequenceStepCount(sequenceName);
  
  // Create a map for quick lookup
  const templateMap = {};
  for (const t of templates) {
    templateMap[t.step] = t;
  }

  // Sequence Settings Section
  const configSection = CardService.newCardSection()
      .setHeader("⚙️ Sequence Settings");
  
  // Dropdown for step count selection
  const stepCountDropdown = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setFieldName("stepCount")
      .setTitle("Number of Steps")
      .setOnChangeAction(CardService.newAction()
          .setFunctionName("saveSequenceStepCount")
          .setParameters({ sequence: sequenceName }));
  
  for (let i = 1; i <= 5; i++) {
    const label = i === 1 ? "1 step (single email only)" : `${i} steps (1 email + ${i-1} follow-up${i > 2 ? 's' : ''})`;
    stepCountDropdown.addItem(label, i.toString(), i === currentStepCount);
  }
  
  configSection.addWidget(stepCountDropdown);
  configSection.addWidget(CardService.newTextParagraph()
      .setText(`<font color='#1a73e8'><b>✓ Sequence ends after Step ${currentStepCount}</b></font>`));
  
  // Edit in Sheets, Rename and Delete options
  configSection.addWidget(CardService.newDivider());
  
  // Edit in Sheets button - opens the sequence sheet directly
  configSection.addWidget(CardService.newTextButton()
      .setText("📝 Edit in Google Sheets")
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor("#34a853")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("openSequenceInSheets")
          .setParameters({ sequence: sequenceName })));
  
  configSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("✏️ Rename Sequence")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("showRenameSequenceCard")
              .setParameters({ sequence: sequenceName })))
      .addButton(CardService.newTextButton()
          .setText("🗑️ Delete Sequence")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("confirmDeleteSequence")
              .setParameters({ sequence: sequenceName }))));
  
  card.addSection(configSection);

  // Show ACTIVE steps (1 to currentStepCount)
  for (let step = 1; step <= currentStepCount; step++) {
    const template = templateMap[step];
    const hasTemplate = template && template.body;
    
    const stepIcon = step === 1 ? "📤" : "↩️";
    const stepType = step === 1 ? "First Email" : "Follow-up #" + (step - 1);
    
    if (hasTemplate) {
      // Create collapsible section with cleaner structure
      const stepSection = CardService.newCardSection()
          .setHeader(`${stepIcon} Step ${step}: ${stepType}`)
          .setCollapsible(true)
          .setNumUncollapsibleWidgets(3); // Subject/Preview + Edit button + divider always visible
      
      // Subject line for Step 1
      if (step === 1 && template.subject) {
        stepSection.addWidget(CardService.newDecoratedText()
            .setTopLabel("SUBJECT")
            .setText(template.subject)
            .setWrapText(true));
      }
      
      // Preview of the email body (always visible)
      const bodyPreview = template.body.length > 100 
        ? template.body.substring(0, 100).replace(/\n/g, " ") + "..." 
        : template.body.replace(/\n/g, " ");
      
      stepSection.addWidget(CardService.newDecoratedText()
          .setTopLabel("PREVIEW")
          .setText(bodyPreview)
          .setWrapText(true));
      
      // Edit button (always visible)
      stepSection.addWidget(CardService.newTextButton()
          .setText("✏️ Edit Step " + step)
          .setOnClickAction(CardService.newAction()
              .setFunctionName("editStepTemplate")
              .setParameters({ 
                sequence: sequenceName, 
                step: step.toString(),
                rowIndex: template.rowIndex.toString()
              })));
      
      // Divider before full content
      stepSection.addWidget(CardService.newDivider());
      
      // Full email header
      stepSection.addWidget(CardService.newTextParagraph()
          .setText("<b>📧 FULL EMAIL CONTENT</b>"));
      
      // Full subject for Step 1
      if (step === 1) {
        stepSection.addWidget(CardService.newDecoratedText()
            .setTopLabel("Subject Line")
            .setText(template.subject || "(no subject)")
            .setWrapText(true));
      } else {
        stepSection.addWidget(CardService.newTextParagraph()
            .setText("<font color='#666666'><i>↳ Sends as reply to Step 1</i></font>"));
      }
      
      // Full email body with proper formatting
      const fullBodyFormatted = template.body.replace(/\n/g, "<br>");
      stepSection.addWidget(CardService.newTextParagraph()
          .setText("<font color='#444444'>" + fullBodyFormatted + "</font>"));
      
      card.addSection(stepSection);
    } else {
      // No template - simple section with write prompt
      const stepSection = CardService.newCardSection()
          .setHeader(`${stepIcon} Step ${step}: ${stepType}`);
      
      stepSection.addWidget(CardService.newDecoratedText()
          .setText("No email configured")
          .setBottomLabel("Click below to write this email")
          .setWrapText(true));
      
      stepSection.addWidget(CardService.newTextButton()
          .setText("✉️ Write Step " + step + " Email")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction()
              .setFunctionName("editStepTemplate")
              .setParameters({ 
                sequence: sequenceName, 
                step: step.toString(),
                rowIndex: "new"
              })));
      
      card.addSection(stepSection);
    }
  }

  // Show INACTIVE steps (currentStepCount+1 to 5) - greyed out with delete option
  if (currentStepCount < 5) {
    const inactiveSection = CardService.newCardSection()
        .setHeader("⏸️ Inactive Steps")
        .setCollapsible(true)
        .setNumUncollapsibleWidgets(1);
    
    inactiveSection.addWidget(CardService.newDecoratedText()
        .setText("These steps will NOT be sent")
        .setBottomLabel("Increase step count above to activate")
        .setWrapText(true));
    
    for (let step = currentStepCount + 1; step <= 5; step++) {
      const template = templateMap[step];
      const hasTemplate = template && template.body;
      
      inactiveSection.addWidget(CardService.newDivider());
      
      if (hasTemplate) {
        // Inactive step WITH content - show delete option
        const bodyPreview = template.body.length > 60 
          ? template.body.substring(0, 60).replace(/\n/g, " ") + "..." 
          : template.body.replace(/\n/g, " ");
        
        inactiveSection.addWidget(CardService.newDecoratedText()
            .setTopLabel(`STEP ${step} (inactive)`)
            .setText(bodyPreview)
            .setWrapText(true));
        
        inactiveSection.addWidget(CardService.newTextButton()
            .setText("🗑️ Delete Content")
            .setOnClickAction(CardService.newAction()
                .setFunctionName("deleteStepContentConfirm")
                .setParameters({ 
                  sequence: sequenceName, 
                  step: step.toString(),
                  rowIndex: template.rowIndex.toString()
                })));
      } else {
        // Inactive step without content
        inactiveSection.addWidget(CardService.newDecoratedText()
            .setTopLabel(`STEP ${step}`)
            .setText("Empty")
            .setBottomLabel("No action needed")
            .setWrapText(true));
      }
    }
    
    card.addSection(inactiveSection);
  }

  // Navigation
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
      .setText("← Back to Template Management")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("buildTemplateManagementCard")));
  card.addSection(navSection);

  return card.build();
}

/**
 * Saves the sequence step count when user changes dropdown
 */
function saveSequenceStepCount(e) {
  const sequenceName = e.parameters.sequence;
  const stepCount = parseInt(e.formInput.stepCount);
  
  if (!sequenceName || !stepCount) {
    return createNotification("Error: Missing sequence or step count.");
  }
  
  const success = setSequenceStepCount(sequenceName, stepCount);
  
  if (success) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText(`✓ Sequence set to ${stepCount} step${stepCount > 1 ? 's' : ''}`))
      .setNavigation(CardService.newNavigation()
        .updateCard(showSequenceTemplatesEditor({ parameters: { sequence: sequenceName } })))
      .build();
  } else {
    return createNotification("Error saving step count.");
  }
}

/**
 * Shows confirmation before deleting step content
 */
function deleteStepContentConfirm(e) {
  const sequenceName = e.parameters.sequence;
  const step = e.parameters.step;
  const rowIndex = e.parameters.rowIndex;
  
  const card = CardService.newCardBuilder();
  
  card.setHeader(CardService.newCardHeader()
      .setTitle("⚠️ Confirm Delete")
      .setSubtitle(`Step ${step} in ${sequenceName}`)
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/warning_amber_48dp.png"));
  
  const section = CardService.newCardSection();
  section.addWidget(CardService.newTextParagraph()
      .setText(`Are you sure you want to <b>permanently delete</b> the email content for Step ${step}?\n\n` +
               `This cannot be undone.`));
  
  section.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("🗑️ Delete Permanently")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction()
              .setFunctionName("deleteStepContent")
              .setParameters({ 
                sequence: sequenceName, 
                step: step,
                rowIndex: rowIndex
              })))
      .addButton(CardService.newTextButton()
          .setText("Cancel")
          .setOnClickAction(CardService.newAction()
              .setFunctionName("showSequenceTemplatesEditor")
              .setParameters({ sequence: sequenceName }))));
  
  card.addSection(section);
  
  return card.build();
}

/**
 * Deletes the content of a specific step template
 */
function deleteStepContent(e) {
  const sequenceName = e.parameters.sequence;
  const step = parseInt(e.parameters.step);
  const rowIndex = parseInt(e.parameters.rowIndex);
  
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    return createNotification("No database connected.");
  }
  
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheetName = getSequenceSheetName(sequenceName);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      return createNotification("Sequence sheet not found.");
    }
    
    // Clear the row content (keep step number, clear name, subject, body)
    sheet.getRange(rowIndex, 2, 1, 3).clearContent(); // Clear columns B, C, D (Name, Subject, Body)
    SpreadsheetApp.flush();
    
    logAction("Delete Step", `Deleted Step ${step} content from ${sequenceName}`);
    
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText(`✓ Step ${step} content deleted`))
      .setNavigation(CardService.newNavigation()
        .updateCard(showSequenceTemplatesEditor({ parameters: { sequence: sequenceName } })))
      .build();
      
  } catch (error) {
    console.error("Error deleting step content: " + error);
    return createNotification("Error deleting: " + error.message);
  }
}

/**
 * Shows form to edit a specific step's email
 * Step 1 = New email (needs subject)
 * Steps 2-5 = Replies (no subject, auto-replies to Step 1)
 */
function editStepTemplate(e) {
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

  // Variable Dictionary - ALWAYS VISIBLE
  const varsSection = CardService.newCardSection()
      .setHeader("📋 Variables (copy into your email)");
  
  varsSection.addWidget(CardService.newTextParagraph()
      .setText(
        "<b>Contact:</b>\n" +
        "<font color='#1a73e8'>{{firstName}}</font>  <font color='#1a73e8'>{{lastName}}</font>\n" +
        "<font color='#1a73e8'>{{company}}</font>  <font color='#1a73e8'>{{title}}</font>\n\n" +
        "<b>Your Info:</b>\n" +
        "<font color='#1a73e8'>{{senderName}}</font>  <font color='#1a73e8'>{{senderCompany}}</font>"
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
              .setFunctionName("saveStepTemplate")
              .setParameters({ 
                sequence: sequenceName, 
                step: step.toString(),
                rowIndex: rowIndex,
                isNew: isNew.toString()
              }))));
  
  // Edit in Sheets option
  formSection.addWidget(CardService.newDivider());
  formSection.addWidget(CardService.newTextButton()
      .setText("📝 Edit in Google Sheets Instead")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("openSequenceInSheets")
          .setParameters({ sequence: sequenceName })));
  
  card.addSection(formSection);
  
  // Navigation
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
      .setText("← Back to Sequence")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("showSequenceTemplatesEditor")
          .setParameters({ sequence: sequenceName })));
  card.addSection(navSection);
  
  return card.build();
}

/**
 * Gets default step name
 */
function getDefaultStepName(step) {
  const names = {
    1: "Introduction",
    2: "Quick Follow-up",
    3: "Second Follow-up", 
    4: "Value Proposition",
    5: "Final Outreach"
  };
  return names[step] || "Step " + step;
}

/**
 * Saves step template - updates existing or creates new row
 */
function saveStepTemplate(e) {
  const sequenceName = e.parameters.sequence;
  const step = parseInt(e.parameters.step);
  const rowIndex = e.parameters.rowIndex;
  const isNew = (e.parameters.isNew === "true");
  const templateName = e.formInput.templateName || getDefaultStepName(step);
  const subject = e.formInput.subject || ""; // Empty for replies
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
    logAction("Template Update", `Saved Step ${step} template in ${sequenceName}`);
    
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText("✓ Step " + step + " email saved!"))
      .setNavigation(CardService.newNavigation()
        .updateCard(showSequenceTemplatesEditor({ parameters: { sequence: sequenceName } })))
      .build();
      
  } catch (error) {
    console.error("Error saving template: " + error);
    return createNotification("Error saving: " + error.message);
  }
}

/**
 * Finds the row index for a specific step in the sequence sheet
 */
function findRowForStep(sheet, stepNumber) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return -1;
  
  const steps = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < steps.length; i++) {
    if (parseInt(steps[i][0]) === stepNumber) {
      return i + 2; // +2 for 1-based index and header row
    }
  }
  return -1;
}

/**
 * Redirects to step editor - kept for backwards compatibility
 */
function showAddTemplateForm(e) {
  // Redirect to step 1 editor by default
  return editStepTemplate({
    parameters: {
      sequence: e.parameters.sequence,
      step: "1",
      rowIndex: "new"
    }
  });
}

/**
 * Legacy function - redirects to new save function
 */
function addNewTemplateToSequence(e) {
  return saveStepTemplate({
    parameters: {
      sequence: e.parameters.sequence,
      step: e.formInput.step || "1",
      rowIndex: "new",
      isNew: "true"
    },
    formInput: e.formInput
  });
}

