/**
 * FILE: Analytics and Reporting
 * 
 * PURPOSE:
 * Provides comprehensive analytics for email campaigns including:
 * - Emails sent today view
 * - Reply tracking and detection
 * - Sequence performance metrics
 * - Template effectiveness analytics
 * - Daily maintenance (thread validation + reply detection)
 * 
 * KEY FUNCTIONS:
 * - buildAnalyticsHubCard() - Main analytics navigation
 * - buildTodaysSentEmailsCard() - View emails sent today
 * - buildReplyTrackingCard() - Check and view replies
 * - buildSequenceAnalyticsCard() - Sequence performance metrics
 * - checkForReplies() - Scan for new replies
 * - dailyMaintenance() - Master daily trigger (4:30 AM) combining:
 *     - Thread pre-validation for Step 2+ contacts
 *     - Reply detection for all contacts
 * - runThreadPreValidation() - Finds message IDs, handles changed subjects
 * 
 * DEPENDENCIES:
 * - 01_Config.gs: CONFIG, CONTACT_COLS
 * - 03_Database.gs: logAction
 * - 04_ContactData.gs: getAllContactsData
 * - 06_SequenceData.gs: getAvailableSequences, getSequenceStepCount
 * - 17_Utilities.gs: addPaginationButtons, createNotification
 * 
 * @version 2.5
 */

/* ======================== ANALYTICS HUB ======================== */

/**
 * Main Analytics hub card - entry point for all analytics features
 */
function buildAnalyticsHubCard(e) {
  const card = CardService.newCardBuilder();

  card.setHeader(CardService.newCardHeader()
      .setTitle("📊 Analytics")
      .setSubtitle("Campaign performance & insights")
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/analytics_black_48dp.png"));

  // --- Quick Stats Summary Section ---
  const summarySection = CardService.newCardSection()
      .setHeader("Quick Summary");

  const stats = getAnalyticsSummary();
  
  summarySection.addWidget(CardService.newKeyValue()
      .setTopLabel("Emails Sent Today")
      .setContent(stats.emailsSentToday.toString())
      .setBottomLabel(stats.companiesContactedToday + " companies contacted"));

  summarySection.addWidget(CardService.newKeyValue()
      .setTopLabel("Total Replies Received")
      .setContent(stats.totalReplies.toString())
      .setBottomLabel("Overall reply rate: " + stats.overallReplyRate + "%"));

  summarySection.addWidget(CardService.newKeyValue()
      .setTopLabel("Active Sequences")
      .setContent(stats.activeSequences.toString())
      .setBottomLabel(stats.totalEmailsSent + " total emails sent"));

  card.addSection(summarySection);

  // --- Analytics Navigation Section ---
  const navSection = CardService.newCardSection()
      .setHeader("Analytics Views");

  navSection.addWidget(CardService.newTextButton()
      .setText("📤 Emails Sent Today")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("buildTodaysSentEmailsCard")
          .setParameters({page: '1'})));

  navSection.addWidget(CardService.newTextButton()
      .setText("💬 Reply Tracking")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("buildReplyTrackingCard")
          .setParameters({page: '1'})));

  navSection.addWidget(CardService.newTextButton()
      .setText("📈 Sequence Performance")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("buildSequenceAnalyticsCard")));

  navSection.addWidget(CardService.newTextButton()
      .setText("📋 Template Effectiveness")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("buildTemplateAnalyticsCard")));

  card.addSection(navSection);

  // --- Reply Check Actions Section ---
  const actionsSection = CardService.newCardSection()
      .setHeader("Reply Detection");

  actionsSection.addWidget(CardService.newTextParagraph()
      .setText("Reply detection runs automatically at 4:30 AM daily. Use the button below for an immediate scan."));

  actionsSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("🔍 Check for Replies Now")
          .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
          .setOnClickAction(CardService.newAction()
              .setFunctionName("runManualReplyCheck"))));

  card.addSection(actionsSection);

  // --- Navigation Section ---
  const backSection = CardService.newCardSection();
  backSection.addWidget(CardService.newTextButton()
      .setText("← Back to Main Menu")
      .setOnClickAction(CardService.newAction()
          .setFunctionName("buildAddOn")));

  card.addSection(backSection);

  return card.build();
}

/* ======================== ANALYTICS SUMMARY DATA ======================== */

/**
 * Gets summary statistics for the analytics hub
 */
function getAnalyticsSummary() {
  const allContacts = getAllContactsData();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  let emailsSentToday = 0;
  let companiesContactedToday = new Set();
  let totalReplies = 0;
  let totalEmailsSent = 0;

  for (const contact of allContacts) {
    // Count emails sent today
    if (contact.lastEmailDate) {
      const lastEmailDate = new Date(contact.lastEmailDate);
      if (lastEmailDate >= today && lastEmailDate < tomorrow) {
        emailsSentToday++;
        if (contact.company) {
          companiesContactedToday.add(contact.company);
        }
      }
    }

    // Count total emails sent (based on current step - 1, since step indicates next email to send after initial)
    // If contact is at step 2+, they've received at least 1 email
    // Also count completed contacts
    if (contact.currentStep > 1 || contact.status === "Completed") {
      totalEmailsSent += Math.max(1, contact.currentStep - 1);
      if (contact.status === "Completed") {
        // Completed contacts have received all their sequence emails
        totalEmailsSent++; // Add the final email
      }
    } else if (contact.lastEmailDate) {
      // Step 1 contacts who have a last email date have been emailed once
      totalEmailsSent++;
    }

    // Count replies
    if (contact.replyReceived === "Yes") {
      totalReplies++;
    }
  }

  const availableSequences = getAvailableSequences();
  const overallReplyRate = totalEmailsSent > 0 
    ? Math.round((totalReplies / totalEmailsSent) * 100) 
    : 0;

  return {
    emailsSentToday: emailsSentToday,
    companiesContactedToday: companiesContactedToday.size,
    totalReplies: totalReplies,
    totalEmailsSent: totalEmailsSent,
    activeSequences: availableSequences.length,
    overallReplyRate: overallReplyRate
  };
}

/* ======================== TODAY'S SENT EMAILS ======================== */

/**
 * Gets unique companies to whom emails were sent today, along with contacts emailed.
 * (Moved from 14_Settings.gs for consolidation)
 */
function getTodaysSentCompanies() {
  const allContacts = getAllContactsData();
  const companiesMap = new Map();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  for (const contact of allContacts) {
    if (contact.lastEmailDate) {
      const lastEmailDateObject = new Date(contact.lastEmailDate);

      if (lastEmailDateObject >= today && lastEmailDateObject < tomorrow && contact.company) {
        const companyName = contact.company.trim();
        const contactFullName = `${contact.firstName || ''} ${contact.lastName || ''}`.trim();

        if (!companiesMap.has(companyName)) {
          companiesMap.set(companyName, {
            companyName: companyName,
            mostRecentOverallEmailDate: lastEmailDateObject,
            sampleEmailForDomain: contact.email,
            contactsEmailedToday: []
          });
        }

        const companyEntry = companiesMap.get(companyName);

        if (lastEmailDateObject > companyEntry.mostRecentOverallEmailDate) {
          companyEntry.mostRecentOverallEmailDate = lastEmailDateObject;
        }

        companyEntry.contactsEmailedToday.push({
          fullName: contactFullName || contact.email,
          title: contact.title || "N/A",
          email: contact.email,
          sentDate: lastEmailDateObject,
          step: contact.currentStep,
          sequence: contact.sequence
        });
      }
    }
  }

  const uniqueCompanies = Array.from(companiesMap.values());
  uniqueCompanies.sort((a, b) => b.mostRecentOverallEmailDate - a.mostRecentOverallEmailDate);
  uniqueCompanies.forEach(company => {
    company.contactsEmailedToday.sort((a, b) => b.sentDate - a.sentDate);
  });

  return uniqueCompanies;
}

/**
 * Builds the card to display companies to whom emails were sent today.
 * (Enhanced version of the one from 14_Settings.gs)
 */
function buildTodaysSentEmailsCard(e) {
  const page = parseInt(e && e.parameters && e.parameters.page || '1');
  const pageSize = CONFIG.PAGE_SIZE;

  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader()
    .setTitle("📤 Today's Sent Emails")
    .setSubtitle("Unique companies contacted today")
    .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/send_black_48dp.png"));

  const sentCompaniesData = getTodaysSentCompanies();

  const totalCompanies = sentCompaniesData.length;
  const totalPages = Math.ceil(totalCompanies / pageSize);
  const startIndex = (page - 1) * pageSize;
  const endIndex = Math.min(startIndex + pageSize, totalCompanies);
  const companiesToShow = sentCompaniesData.slice(startIndex, endIndex);

  const companiesSection = CardService.newCardSection();

  if (totalCompanies === 0) {
    companiesSection.addWidget(CardService.newTextParagraph()
      .setText("No emails have been recorded as sent today.\n\nEmails are tracked when you use the bulk email feature to create drafts or send messages."));
  } else {
    companiesSection.setHeader(`Companies Emailed (${startIndex + 1}-${endIndex} of ${totalCompanies})`);
    
    for (const companyData of companiesToShow) {
      companiesSection.addWidget(CardService.newTextParagraph()
        .setText(`<b>${companyData.companyName}</b>`));

      const buttonSet = CardService.newButtonSet();
      let domain = "";
      if (companyData.sampleEmailForDomain) {
        const emailParts = companyData.sampleEmailForDomain.split('@');
        if (emailParts.length === 2) {
          domain = emailParts[1];
        }
      }

      const linkedInUrl = `https://www.linkedin.com/search/results/companies/?keywords=${encodeURIComponent(companyData.companyName)}`;
      buttonSet.addButton(CardService.newTextButton()
        .setText("🔵 LinkedIn")
        .setOpenLink(CardService.newOpenLink().setUrl(linkedInUrl)));

      if (domain) {
        const connectSalesUrl = `https://sales.connect.corp.google.com/search/LEAD/TEXT/type%3ALEAD+${domain}`;
        buttonSet.addButton(CardService.newTextButton()
          .setText("🟢 Connect Sales")
          .setOpenLink(CardService.newOpenLink().setUrl(connectSalesUrl)));

        const goQualifyUrl = `https://sales.connect.corp.google.com/qualify/search?qualify-query=https%3A%2F%2F${encodeURIComponent(domain)}&qualify-country=CA&qualify-source-metadata=Qualify-Extension&selectedTeamId=2148841788`;
        buttonSet.addButton(CardService.newTextButton()
          .setText("🔘 GO/Qualify")
          .setOpenLink(CardService.newOpenLink().setUrl(goQualifyUrl)));
      }
      companiesSection.addWidget(buttonSet);

      if (companyData.contactsEmailedToday && companyData.contactsEmailedToday.length > 0) {
        for (const contactInfo of companyData.contactsEmailedToday) {
          const formattedSentDate = Utilities.formatDate(contactInfo.sentDate, Session.getScriptTimeZone(), "h:mm a");
          const stepInfo = contactInfo.step ? `Step ${contactInfo.step - 1}` : "";
          const contactText = `${contactInfo.fullName} | ${contactInfo.title} | ${stepInfo} | ${formattedSentDate}`;
          
          companiesSection.addWidget(CardService.newDecoratedText()
            .setText(contactInfo.email)
            .setBottomLabel(contactText)
            .setWrapText(true));
        }
      }

      companiesSection.addWidget(CardService.newDivider());
    }

    addPaginationButtons(companiesSection, page, totalPages, "buildTodaysSentEmailsCard", {});
  }

  card.addSection(companiesSection);

  // Navigation
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
    .setText("← Back to Analytics")
    .setOnClickAction(CardService.newAction().setFunctionName("buildAnalyticsHubCard")));
  card.addSection(navSection);

  return card.build();
}

/* ======================== REPLY TRACKING ======================== */

/**
 * Builds the Reply Tracking card showing contacts who have replied
 */
function buildReplyTrackingCard(e) {
  const page = parseInt(e && e.parameters && e.parameters.page || '1');
  const pageSize = CONFIG.PAGE_SIZE;

  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader()
      .setTitle("💬 Reply Tracking")
      .setSubtitle("Contacts who have responded")
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/forum_black_48dp.png"));

  const allContacts = getAllContactsData();
  
  // Filter to contacts with replies
  const contactsWithReplies = allContacts.filter(c => c.replyReceived === "Yes");
  contactsWithReplies.sort((a, b) => {
    // Sort by reply date descending (most recent first)
    const dateA = a.replyDate ? new Date(a.replyDate) : new Date(0);
    const dateB = b.replyDate ? new Date(b.replyDate) : new Date(0);
    return dateB - dateA;
  });

  const totalReplies = contactsWithReplies.length;
  const totalPages = Math.ceil(totalReplies / pageSize);
  const startIndex = (page - 1) * pageSize;
  const endIndex = Math.min(startIndex + pageSize, totalReplies);
  const repliesToShow = contactsWithReplies.slice(startIndex, endIndex);

  // Summary section
  const summarySection = CardService.newCardSection();
  summarySection.addWidget(CardService.newKeyValue()
      .setTopLabel("Total Replies Detected")
      .setContent(totalReplies.toString())
      .setBottomLabel("Click 'Check for Replies' on Analytics hub to scan for new replies"));
  card.addSection(summarySection);

  // Replies list section
  const repliesSection = CardService.newCardSection();

  if (totalReplies === 0) {
    repliesSection.addWidget(CardService.newTextParagraph()
        .setText("No replies detected yet.\n\nUse the 'Check for Replies Now' button on the Analytics hub to scan your inbox for responses from contacts."));
  } else {
    repliesSection.setHeader(`Replies (${startIndex + 1}-${endIndex} of ${totalReplies})`);

    for (const contact of repliesToShow) {
      const fullName = `${contact.firstName} ${contact.lastName}`.trim();
      const replyDateStr = contact.replyDate 
        ? Utilities.formatDate(new Date(contact.replyDate), Session.getScriptTimeZone(), "MMM dd, yyyy h:mm a")
        : "Date unknown";
      
      repliesSection.addWidget(CardService.newDecoratedText()
          .setTopLabel(contact.company || "No Company")
          .setText(`<b>${fullName}</b>`)
          .setBottomLabel(`${contact.sequence || "No Sequence"} • Replied: ${replyDateStr}`)
          .setWrapText(true)
          .setOnClickAction(CardService.newAction()
              .setFunctionName("viewContactCard")
              .setParameters({email: contact.email, page: '1'})));
    }

    addPaginationButtons(repliesSection, page, totalPages, "buildReplyTrackingCard", {});
  }

  card.addSection(repliesSection);

  // Navigation
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
      .setText("← Back to Analytics")
      .setOnClickAction(CardService.newAction().setFunctionName("buildAnalyticsHubCard")));
  card.addSection(navSection);

  return card.build();
}

/* ======================== SEQUENCE ANALYTICS ======================== */

/**
 * Builds the Sequence Analytics card with performance metrics per sequence
 */
function buildSequenceAnalyticsCard(e) {
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader()
      .setTitle("📈 Sequence Performance")
      .setSubtitle("Metrics by email sequence")
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/trending_up_black_48dp.png"));

  const sequenceStats = getSequencePerformanceStats();
  const sequences = Object.keys(sequenceStats).sort();

  if (sequences.length === 0) {
    const emptySection = CardService.newCardSection();
    emptySection.addWidget(CardService.newTextParagraph()
        .setText("No sequence data available yet.\n\nStart sending emails to see performance metrics here."));
    card.addSection(emptySection);
  } else {
    // Create a section for each sequence (with collapsible for step details)
    for (const sequenceName of sequences) {
      const stats = sequenceStats[sequenceName];
      const replyRate = stats.totalSent > 0 
        ? Math.round((stats.totalReplies / stats.totalSent) * 100) 
        : 0;

      const sequenceSection = CardService.newCardSection()
          .setHeader(`📧 ${sequenceName}`)
          .setCollapsible(true)
          .setNumUncollapsibleWidgets(2);

      // Summary metrics
      sequenceSection.addWidget(CardService.newKeyValue()
          .setTopLabel("Total Emails Sent")
          .setContent(stats.totalSent.toString())
          .setBottomLabel(`${stats.activeContacts} active, ${stats.completedContacts} completed`));

      sequenceSection.addWidget(CardService.newKeyValue()
          .setTopLabel("Replies")
          .setContent(`${stats.totalReplies} (${replyRate}% rate)`)
          .setBottomLabel(replyRate >= 10 ? "✅ Good engagement" : replyRate >= 5 ? "🟡 Average" : "🔴 Needs attention"));

      // Per-step breakdown
      sequenceSection.addWidget(CardService.newTextParagraph()
          .setText("<b>Breakdown by Step:</b>"));

      const stepCount = getSequenceStepCount(sequenceName);
      for (let step = 1; step <= stepCount; step++) {
        const stepStats = stats.steps[step] || { sent: 0, replies: 0 };
        const stepReplyRate = stepStats.sent > 0 
          ? Math.round((stepStats.replies / stepStats.sent) * 100) 
          : 0;
        
        const stepLabel = step === 1 ? "Introduction" : `Follow-up ${step - 1}`;
        sequenceSection.addWidget(CardService.newDecoratedText()
            .setText(`Step ${step}: ${stepLabel}`)
            .setBottomLabel(`${stepStats.sent} sent • ${stepStats.replies} replies (${stepReplyRate}%)`));
      }

      card.addSection(sequenceSection);
    }
  }

  // Navigation
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
      .setText("← Back to Analytics")
      .setOnClickAction(CardService.newAction().setFunctionName("buildAnalyticsHubCard")));
  card.addSection(navSection);

  return card.build();
}

/**
 * Gets performance statistics grouped by sequence
 */
function getSequencePerformanceStats() {
  const allContacts = getAllContactsData();
  const stats = {};

  for (const contact of allContacts) {
    const sequenceName = contact.sequence || "Unassigned";
    
    if (!stats[sequenceName]) {
      stats[sequenceName] = {
        totalSent: 0,
        totalReplies: 0,
        activeContacts: 0,
        completedContacts: 0,
        steps: {}
      };
    }

    const seqStats = stats[sequenceName];

    // Count status
    if (contact.status === "Active" || contact.status === "Paused") {
      seqStats.activeContacts++;
    } else if (contact.status === "Completed") {
      seqStats.completedContacts++;
    }

    // Calculate emails sent for this contact
    // If they have a lastEmailDate, at least one email was sent
    if (contact.lastEmailDate) {
      // The step they're on is the NEXT email to send
      // So emails sent = currentStep - 1 (minimum 1 if they have lastEmailDate)
      const emailsSent = Math.max(1, contact.currentStep - 1);
      
      // If completed, they've received all emails up to their sequence's max step
      const finalEmailsSent = contact.status === "Completed" 
        ? getSequenceStepCount(sequenceName)
        : emailsSent;
      
      seqStats.totalSent += finalEmailsSent;

      // Track per-step metrics
      for (let step = 1; step <= finalEmailsSent; step++) {
        if (!seqStats.steps[step]) {
          seqStats.steps[step] = { sent: 0, replies: 0 };
        }
        seqStats.steps[step].sent++;
      }
    }

    // Count replies
    if (contact.replyReceived === "Yes") {
      seqStats.totalReplies++;
      
      // Attribute reply to the step they were on when reply was received
      // This is an approximation - ideally we'd track which step generated the reply
      const replyStep = Math.max(1, contact.currentStep - 1);
      if (!seqStats.steps[replyStep]) {
        seqStats.steps[replyStep] = { sent: 0, replies: 0 };
      }
      seqStats.steps[replyStep].replies++;
    }
  }

  return stats;
}

/* ======================== TEMPLATE ANALYTICS ======================== */

/**
 * Builds the Template Analytics card showing effectiveness by template
 */
function buildTemplateAnalyticsCard(e) {
  const card = CardService.newCardBuilder();
  card.setHeader(CardService.newCardHeader()
      .setTitle("📋 Template Effectiveness")
      .setSubtitle("Performance by email template")
      .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/article_black_48dp.png"));

  const templateStats = getTemplatePerformanceStats();

  if (templateStats.length === 0) {
    const emptySection = CardService.newCardSection();
    emptySection.addWidget(CardService.newTextParagraph()
        .setText("No template data available yet.\n\nStart sending emails to see which templates perform best."));
    card.addSection(emptySection);
  } else {
    // Sort templates by reply rate descending
    templateStats.sort((a, b) => b.replyRate - a.replyRate);

    const templatesSection = CardService.newCardSection()
        .setHeader("Templates Ranked by Reply Rate");

    for (const template of templateStats) {
      const performanceIcon = template.replyRate >= 10 ? "🏆" 
        : template.replyRate >= 5 ? "✅" 
        : template.replyRate >= 1 ? "🟡" 
        : "🔴";

      templatesSection.addWidget(CardService.newDecoratedText()
          .setTopLabel(`${template.sequence} • Step ${template.step}`)
          .setText(`${performanceIcon} ${template.name}`)
          .setBottomLabel(`${template.sent} sent • ${template.replies} replies • ${template.replyRate}% rate`)
          .setWrapText(true));
    }

    card.addSection(templatesSection);

    // Tips section
    const tipsSection = CardService.newCardSection()
        .setHeader("💡 Optimization Tips")
        .setCollapsible(true);

    tipsSection.addWidget(CardService.newTextParagraph()
        .setText("• <b>High performers (🏆 10%+)</b>: Consider using similar approaches in other templates\n\n" +
                 "• <b>Average (🟡 1-5%)</b>: Try A/B testing with different subject lines or opening hooks\n\n" +
                 "• <b>Low performers (🔴 <1%)</b>: Review template content, timing, or target audience"));

    card.addSection(tipsSection);
  }

  // Navigation
  const navSection = CardService.newCardSection();
  navSection.addWidget(CardService.newTextButton()
      .setText("← Back to Analytics")
      .setOnClickAction(CardService.newAction().setFunctionName("buildAnalyticsHubCard")));
  card.addSection(navSection);

  return card.build();
}

/**
 * Gets performance statistics for each template
 */
function getTemplatePerformanceStats() {
  const allContacts = getAllContactsData();
  const availableSequences = getAvailableSequences();
  const templateStats = [];

  // Build a map of template usage
  const templateUsage = new Map(); // Key: "sequence|step", Value: {sent, replies}

  for (const contact of allContacts) {
    if (!contact.sequence || !contact.lastEmailDate) continue;

    // Calculate how many emails were sent
    const emailsSent = contact.status === "Completed"
      ? getSequenceStepCount(contact.sequence)
      : Math.max(0, contact.currentStep - 1);

    for (let step = 1; step <= emailsSent; step++) {
      const key = `${contact.sequence}|${step}`;
      if (!templateUsage.has(key)) {
        templateUsage.set(key, { sent: 0, replies: 0 });
      }
      templateUsage.get(key).sent++;
    }

    // Attribute reply to the step before current
    if (contact.replyReceived === "Yes" && emailsSent > 0) {
      const replyStep = Math.max(1, contact.currentStep - 1);
      const key = `${contact.sequence}|${replyStep}`;
      if (templateUsage.has(key)) {
        templateUsage.get(key).replies++;
      }
    }
  }

  // Convert to array with template names
  for (const sequence of availableSequences) {
    const templates = getEmailTemplates(sequence);
    const stepCount = getSequenceStepCount(sequence);

    for (let step = 1; step <= stepCount; step++) {
      const key = `${sequence}|${step}`;
      const usage = templateUsage.get(key) || { sent: 0, replies: 0 };
      
      const template = templates.find(t => t.step === step);
      const templateName = template ? template.name : `Step ${step}`;

      if (usage.sent > 0) {
        templateStats.push({
          sequence: sequence,
          step: step,
          name: templateName,
          sent: usage.sent,
          replies: usage.replies,
          replyRate: Math.round((usage.replies / usage.sent) * 100)
        });
      }
    }
  }

  return templateStats;
}

/* ======================== REPLY DETECTION ======================== */

/**
 * Manual trigger for reply checking
 */
function runManualReplyCheck() {
  try {
    const result = checkForReplies();
    
    return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
            .setText(`✅ Scan complete! Found ${result.newReplies} new replies.`))
        .setNavigation(CardService.newNavigation()
            .updateCard(buildAnalyticsHubCard()))
        .build();
  } catch (error) {
    console.error("Error in manual reply check: " + error);
    return createNotification("Error checking replies: " + error.message);
  }
}

/**
 * Checks for replies from contacts in the database
 * Scans sent emails for threads with responses from the recipient
 */
function checkForReplies() {
  const allContacts = getAllContactsData();
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  
  if (!spreadsheetId) {
    return { newReplies: 0, error: "No database connected" };
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);
  
  if (!contactsSheet) {
    return { newReplies: 0, error: "Contacts sheet not found" };
  }

  const userEmail = Session.getActiveUser().getEmail().toLowerCase();
  let newRepliesFound = 0;
  const updates = []; // [rowIndex, replyReceived, replyDate]

  // Only check contacts who:
  // 1. Have sent at least one email (have a lastEmailDate)
  // 2. Haven't already been marked as replied
  // 3. Have a step1Subject or threadId to search for
  const contactsToCheck = allContacts.filter(c => 
    c.lastEmailDate && 
    c.replyReceived !== "Yes" && 
    (c.step1Subject || c.threadId)
  );

  console.log(`Checking ${contactsToCheck.length} contacts for replies...`);

  // Process in smaller batches to avoid timeouts
  const BATCH_SIZE = 50;
  const contactsToProcess = contactsToCheck.slice(0, BATCH_SIZE);

  for (const contact of contactsToProcess) {
    try {
      let thread = null;
      
      // Try to get thread by ID first (most reliable)
      if (contact.threadId) {
        try {
          thread = GmailApp.getThreadById(contact.threadId);
        } catch (e) {
          // Thread ID might be invalid, fall through to search
        }
      }

      // If no thread found by ID, search for it
      if (!thread && contact.step1Subject) {
        const searchSubject = contact.step1Subject.replace(/\|/g, ' ').replace(/"/g, '\\"');
        const searchQuery = `to:(${contact.email}) in:sent subject:("${searchSubject}")`;
        
        const threads = GmailApp.search(searchQuery, 0, 3);
        
        for (const t of threads) {
          const messages = t.getMessages();
          if (messages.length > 0) {
            // Verify this is the right thread by checking the first message
            const firstMsg = messages[0];
            if (firstMsg.getSubject() === contact.step1Subject || 
                firstMsg.getTo().toLowerCase().includes(contact.email.toLowerCase())) {
              thread = t;
              break;
            }
          }
        }
      }

      // Check if thread has a reply from the contact
      if (thread) {
        const messages = thread.getMessages();
        
        for (const message of messages) {
          const fromEmail = extractEmailFromHeader_(message.getFrom());
          
          // Check if this message is FROM the contact (not from us)
          if (fromEmail && fromEmail.toLowerCase() === contact.email.toLowerCase()) {
            // This is a reply from the contact!
            updates.push({
              rowIndex: contact.rowIndex,
              replyReceived: "Yes",
              replyDate: message.getDate()
            });
            newRepliesFound++;
            break; // Only need to find one reply per contact
          }
        }
      }
    } catch (error) {
      console.error(`Error checking replies for ${contact.email}: ${error}`);
    }
  }

  // Batch write updates to the sheet
  if (updates.length > 0) {
    for (const update of updates) {
      contactsSheet.getRange(update.rowIndex, CONTACT_COLS.REPLY_RECEIVED + 1).setValue(update.replyReceived);
      contactsSheet.getRange(update.rowIndex, CONTACT_COLS.REPLY_DATE + 1).setValue(update.replyDate);
    }
    SpreadsheetApp.flush();
    logAction("Reply Check", `Found ${newRepliesFound} new replies from ${contactsToProcess.length} contacts checked.`);
  }

  // If there are more contacts to check, schedule a continuation
  if (contactsToCheck.length > BATCH_SIZE) {
    // Store state for continuation if needed
    console.log(`${contactsToCheck.length - BATCH_SIZE} more contacts to check in future runs.`);
  }

  return { newReplies: newRepliesFound, checked: contactsToProcess.length };
}

/* ======================== DAILY MAINTENANCE (MASTER TRIGGER) ======================== */

/**
 * Master daily maintenance function - runs at 4:30 AM
 * Combines all background tasks into one trigger to save quota
 * 
 * Tasks:
 * 1. Thread Pre-Validation - finds message IDs for contacts ready for Step 2+
 * 2. Reply Detection - scans for replies from contacts
 */
function dailyMaintenance() {
  console.log("=== DAILY MAINTENANCE STARTED ===");
  
  let threadValidationResult = { validated: 0, total: 0 };
  let replyCheckResult = { newReplies: 0, checked: 0 };

  // TASK 1: Thread Pre-Validation
  console.log("\n--- TASK 1: Thread Pre-Validation ---");
  try {
    threadValidationResult = runThreadPreValidation();
  } catch (error) {
    console.error("Thread pre-validation failed: " + error);
    logAction("Error", "Thread pre-validation failed: " + error.toString());
  }

  // TASK 2: Reply Detection
  console.log("\n--- TASK 2: Reply Detection ---");
  try {
    replyCheckResult = checkForReplies();
    
    // Send summary email if new replies found
    if (replyCheckResult.newReplies > 0) {
      const userEmail = Session.getActiveUser().getEmail();
      GmailApp.sendEmail(
        userEmail,
        "Sales Outreach: " + replyCheckResult.newReplies + " New Replies Detected",
        `Your daily maintenance found ${replyCheckResult.newReplies} new replies from your contacts.\n\n` +
        `Open the Sales Outreach add-on and go to Analytics > Reply Tracking to view details.`
      );
    }
  } catch (error) {
    console.error("Reply detection failed: " + error);
    logAction("Error", "Reply detection failed: " + error.toString());
  }

  // Summary log
  const summary = `Thread validation: ${threadValidationResult.validated}/${threadValidationResult.total}. ` +
                  `Replies found: ${replyCheckResult.newReplies} (checked ${replyCheckResult.checked}).`;
  console.log("\n=== DAILY MAINTENANCE COMPLETE ===");
  console.log(summary);
  logAction("Daily Maintenance", summary);
}

/**
 * Runs thread pre-validation and returns results
 * Extracted from dailyThreadPreValidation for use in dailyMaintenance
 */
function runThreadPreValidation() {
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    console.log("Thread Pre-Validation: No database connected.");
    return { validated: 0, total: 0 };
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);
  if (!contactsSheet) {
    console.log("Thread Pre-Validation: Contacts sheet not found.");
    return { validated: 0, total: 0 };
  }

  const allContacts = getAllContactsData();
  const userEmail = Session.getActiveUser().getEmail().toLowerCase();
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const contactsNeedingValidation = allContacts.filter(c => {
    if (c.currentStep <= 1 || c.status !== 'Active') return false;
    if (c.step1SentMessageId) return false;
    if (!c.lastEmailDate) return false;
    
    if (!c.nextStepDate) return false;
    const nextDate = new Date(c.nextStepDate);
    nextDate.setHours(0, 0, 0, 0);
    
    return nextDate.getTime() <= today.getTime();
  });

  console.log(`Thread Pre-Validation: ${contactsNeedingValidation.length} contacts need validation.`);

  if (contactsNeedingValidation.length === 0) {
    return { validated: 0, total: 0 };
  }

  let validatedCount = 0;
  let fallbackCount = 0;
  let subjectChangedCount = 0;
  const updates = [];

  for (const contact of contactsNeedingValidation) {
    try {
      console.log(`\n--- Processing: ${contact.email} ---`);
      console.log(`step1Subject: "${contact.step1Subject || '(empty)'}"`);
      console.log(`lastEmailDate: ${contact.lastEmailDate || '(empty)'}`);
      
      let foundMessage = null;
      let usedFallback = false;

      // STRATEGY 1: Try subject line + email search first
      if (contact.step1Subject) {
        const searchSubject = contact.step1Subject.replace(/"/g, '\\"').replace(/\|/g, ' ');
        const subjectQuery = `from:me to:(${contact.email}) is:sent subject:("${searchSubject}")`;
        console.log(`Strategy 1 for ${contact.email} - Searching with subject: "${contact.step1Subject}"`);
        
        const subjectThreads = GmailApp.search(subjectQuery, 0, 5);
        console.log(`Strategy 1 found ${subjectThreads.length} threads`);

        for (const thread of subjectThreads) {
          for (const msg of thread.getMessages()) {
            if (msg.isDraft()) continue;
            
            const toEmail = msg.getTo().toLowerCase();
            const msgSubject = msg.getSubject();
            
            if (toEmail.includes(contact.email.toLowerCase()) && msgSubject === contact.step1Subject) {
              foundMessage = msg;
              console.log(`Strategy 1 SUCCESS for ${contact.email}`);
              break;
            }
          }
          if (foundMessage) break;
        }
        
        if (!foundMessage) {
          console.log(`Strategy 1 FAILED for ${contact.email} - no exact subject match found`);
        }
      } else {
        console.log(`Strategy 1 SKIPPED for ${contact.email} - no step1Subject stored`);
      }

      // STRATEGY 2: Fallback to date + email
      if (!foundMessage && contact.lastEmailDate) {
        usedFallback = true;
        console.log(`Fallback triggered for ${contact.email} - subject search failed`);
        
        const emailDate = new Date(contact.lastEmailDate);
        console.log(`lastEmailDate for ${contact.email}: ${emailDate.toString()}`);
        
        const windowStart = new Date(emailDate);
        windowStart.setDate(windowStart.getDate() - 1);
        const windowEnd = new Date(emailDate);
        windowEnd.setDate(windowEnd.getDate() + 2);

        const startStr = Utilities.formatDate(windowStart, Session.getScriptTimeZone(), "yyyy/MM/dd");
        const endStr = Utilities.formatDate(windowEnd, Session.getScriptTimeZone(), "yyyy/MM/dd");

        const fallbackQuery = `from:me to:(${contact.email}) is:sent after:${startStr} before:${endStr}`;
        console.log(`Fallback search query: ${fallbackQuery}`);
        
        const fallbackThreads = GmailApp.search(fallbackQuery, 0, 10);
        console.log(`Fallback search found ${fallbackThreads.length} threads`);

        let bestMatch = null;
        let closestTimeDiff = Infinity;

        for (const thread of fallbackThreads) {
          for (const msg of thread.getMessages()) {
            if (msg.isDraft()) continue;
            
            const fromEmail = msg.getFrom().toLowerCase();
            const toEmail = msg.getTo().toLowerCase();
            
            if (fromEmail.includes(userEmail) && toEmail.includes(contact.email.toLowerCase())) {
              const msgDate = msg.getDate();
              const timeDiff = Math.abs(msgDate - emailDate);
              
              console.log(`Found candidate message - Subject: "${msg.getSubject()}", Date: ${msgDate.toString()}, TimeDiff: ${Math.round(timeDiff/1000/60)} minutes`);
              
              if (timeDiff < closestTimeDiff) {
                closestTimeDiff = timeDiff;
                bestMatch = msg;
              }
            }
          }
        }

        if (bestMatch) {
          foundMessage = bestMatch;
          fallbackCount++;
          console.log(`Fallback SUCCESS for ${contact.email} - found message with subject: "${bestMatch.getSubject()}"`);
        } else {
          console.log(`Fallback FAILED for ${contact.email} - no matching messages in threads`);
        }
      }

      // Store results
      if (foundMessage) {
        const actualSubject = foundMessage.getSubject();
        const subjectChanged = contact.step1Subject !== actualSubject;
        
        updates.push({
          rowIndex: contact.rowIndex,
          messageId: foundMessage.getId(),
          threadId: foundMessage.getThread().getId(),
          actualSubject: actualSubject,
          subjectChanged: subjectChanged
        });
        
        validatedCount++;
        
        if (subjectChanged) {
          subjectChangedCount++;
          console.log(`Subject changed: "${contact.step1Subject}" → "${actualSubject}"`);
        }
      } else {
        console.warn(`Could not find Step 1 message for ${contact.email}`);
      }

    } catch (e) {
      console.error(`Error processing ${contact.email}: ${e}`);
    }
  }

  // Batch write updates
  if (updates.length > 0) {
    for (const update of updates) {
      contactsSheet.getRange(update.rowIndex, CONTACT_COLS.STEP1_SENT_MESSAGE_ID + 1).setValue(update.messageId);
      contactsSheet.getRange(update.rowIndex, CONTACT_COLS.THREAD_ID + 1).setValue(update.threadId);
      
      if (update.subjectChanged) {
        contactsSheet.getRange(update.rowIndex, CONTACT_COLS.STEP1_SUBJECT + 1).setValue(update.actualSubject);
      }
    }
    SpreadsheetApp.flush();
  }

  const summary = `Validated ${validatedCount}/${contactsNeedingValidation.length}. Fallback: ${fallbackCount}. Subject fixes: ${subjectChangedCount}.`;
  console.log(summary);

  return { validated: validatedCount, total: contactsNeedingValidation.length };
}

