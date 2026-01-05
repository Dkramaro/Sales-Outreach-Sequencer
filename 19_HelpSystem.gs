/**
 * FILE: Help System
 * 
 * PURPOSE:
 * Creates and manages a user-facing help documentation Google Doc.
 * First button press creates the doc; subsequent presses open it.
 * 
 * KEY FUNCTIONS:
 * - openHelpDocumentation() - Main entry point from Help button
 * - createHelpDocument() - Creates comprehensive help Google Doc
 * - getHelpDocId() / setHelpDocId() - Property storage helpers
 * 
 * DEPENDENCIES:
 * - 03_Database.gs: logAction
 * - 17_Utilities.gs: createNotification
 * 
 * @version 1.0
 */

/* ======================== HELP BUTTON HANDLER ======================== */

/**
 * Main entry point - creates doc if needed, then opens it
 */
function openHelpDocumentation() {
  try {
    let docId = getHelpDocId();
    
    if (!docId) {
      // First time - create the help document
      docId = createHelpDocument();
      setHelpDocId(docId);
      logAction("Help", "Created new help documentation: " + docId);
    }
    
    // Verify doc still exists
    try {
      DriveApp.getFileById(docId);
    } catch (e) {
      // Doc was deleted - recreate it
      docId = createHelpDocument();
      setHelpDocId(docId);
      logAction("Help", "Recreated help documentation (previous was deleted): " + docId);
    }
    
    const docUrl = "https://docs.google.com/document/d/" + docId + "/edit";
    
    return CardService.newActionResponseBuilder()
      .setOpenLink(CardService.newOpenLink()
        .setUrl(docUrl)
        .setOpenAs(CardService.OpenAs.FULL_SIZE)
        .setOnClose(CardService.OnClose.NOTHING))
      .setNotification(CardService.newNotification()
        .setText("📖 Opening Help Documentation..."))
      .build();
    
  } catch (error) {
    console.error("Error opening help documentation: " + error);
    logAction("Error", "Failed to open help documentation: " + error.toString());
    return createNotification("Error opening help: " + error.message);
  }
}

/* ======================== PROPERTY HELPERS ======================== */

/**
 * Gets the stored help document ID
 */
function getHelpDocId() {
  return PropertiesService.getUserProperties().getProperty("HELP_DOC_ID");
}

/**
 * Stores the help document ID
 */
function setHelpDocId(docId) {
  PropertiesService.getUserProperties().setProperty("HELP_DOC_ID", docId);
}

/* ======================== DOCUMENT CREATION ======================== */

/**
 * Creates the comprehensive help Google Doc with all sections
 */
function createHelpDocument() {
  const doc = DocumentApp.create("📧 Sales Outreach Sequence - Help Guide");
  const body = doc.getBody();
  
  // Set document styling
  body.setMarginTop(36);
  body.setMarginBottom(36);
  body.setMarginLeft(54);
  body.setMarginRight(54);
  
  // ======================================
  // TITLE & INTRO
  // ======================================
  const title = body.appendParagraph("📧 Sales Outreach Sequence");
  title.setHeading(DocumentApp.ParagraphHeading.TITLE);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  const subtitle = body.appendParagraph("Your Complete User Guide");
  subtitle.setHeading(DocumentApp.ParagraphHeading.SUBTITLE);
  subtitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendParagraph("");
  
  const intro = body.appendParagraph("Welcome! 👋 This guide will help you master your email outreach campaigns. Use the sidebar navigation (View → Show document outline) to jump between sections.");
  intro.setItalic(true);
  
  body.appendHorizontalRule();
  
  // ======================================
  // TABLE OF CONTENTS
  // ======================================
  const tocHeader = body.appendParagraph("📑 Quick Navigation");
  tocHeader.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  body.appendListItem("Bulk Email Sending - The core workflow").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("Adding Contacts - Quick & easy imports").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("Call Management - Phone outreach tracking").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("Create/Edit Templates - Your email content").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("Analytics - Track performance & replies").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("Settings - Customize your experience").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("Google Sheets - What NOT to touch").setGlyphType(DocumentApp.GlyphType.NUMBER);
  
  body.appendHorizontalRule();
  
  // ======================================
  // SECTION 1: BULK EMAIL SENDING
  // ======================================
  const section1 = body.appendParagraph("📤 1. Bulk Email Sending");
  section1.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  body.appendParagraph("This is your main workflow for sending email campaigns. Here's how it works:");
  
  const overviewHeader = body.appendParagraph("The Big Picture 🎯");
  overviewHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  body.appendListItem("Add contacts (manually in the app or via Google Sheets)").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Assign each contact to a sequence (email campaign)").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Dashboard shows \"Ready to Send\" contacts").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Click \"Send Emails Now\" to create drafts in bulk").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Review drafts in Gmail, edit if needed, then send!").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const stepsHeader = body.appendParagraph("Step-by-Step Process 📝");
  stepsHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  const step1H = body.appendParagraph("Step 1: First Emails (Introductions)");
  step1H.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph("• When you add a contact, they start at Step 1\n• Dashboard shows contacts ready for their first email\n• You can draft up to 15 first emails at once\n• All variables ({{firstName}}, {{company}}, etc.) auto-populate\n• Emails go to your Gmail Drafts folder");
  
  const step2H = body.appendParagraph("Step 2-5: Follow-ups");
  step2H.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph("• After Step 1, there's a delay (default: 3 days)\n• When delay passes, contact becomes \"ready\" again\n• Follow-ups reply to the original thread (Re: [subject])\n• You can draft up to 10 follow-ups at once\n• Same process: drafts → review → send");
  
  const delayHeader = body.appendParagraph("⏰ About Time Delays");
  delayHeader.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph("Default delay between emails is 3 days. Change this in Settings → Sequence Settings. The system calculates \"next step date\" automatically when you send an email.");
  
  const tipsHeader = body.appendParagraph("💡 Pro Tips");
  tipsHeader.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendListItem("Always review drafts before sending - personalize where it makes sense").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Step 1 can be set to auto-send in Settings (not recommended for beginners)").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Use priority levels (High/Medium/Low) to focus on important contacts").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  // ======================================
  // SIGNATURE SETUP (SUB-SECTION)
  // ======================================
  const signatureHeader = body.appendParagraph("✍️ Setting Up Your Email Signature");
  signatureHeader.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  body.appendParagraph("Want your signature automatically added to every email? Here's how to set it up correctly:");
  
  const sigSetupH = body.appendParagraph("🚀 Quick Setup");
  sigSetupH.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendListItem("Go to Settings → Email Signature section").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("Click \"📝 Setup Signature\" button").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("A Google Doc will be created automatically with a pre-formatted table").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("Click \"👁️ View Signature\" to customize it").setGlyphType(DocumentApp.GlyphType.NUMBER);
  
  const sigFormatH = body.appendParagraph("⚠️ CRITICAL: Why You MUST Use the Table Format");
  sigFormatH.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  sigFormatH.setForegroundColor("#cc0000");
  
  const warningPara = body.appendParagraph(
    "Pulling content from Google Docs is finicky. If you don't use a table, the formatting WILL BREAK " +
    "when imported into your emails. The pre-populated template uses a 2-column table that preserves formatting perfectly."
  );
  warningPara.setForegroundColor("#cc0000");
  
  const sigCustomizeH = body.appendParagraph("🎨 Customizing Your Signature");
  sigCustomizeH.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  
  body.appendListItem("DELETE all the instruction text in the document (keep only the table)").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Replace placeholder text with your name, title, email, and phone").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("For logo: Select all 4 cells in the first column → Right-click → Merge cells").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Click in merged cell → Insert → Image → Upload your logo").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Keep it SMALL - see the template for recommended size").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Table borders are already set to transparent with 0 width (perfect for emails)").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const sigExampleH = body.appendParagraph("📋 Example Structure");
  sigExampleH.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  
  body.appendParagraph("The table has 2 columns:\n" +
    "• Column 1 (narrow): Your logo (merge all 4 rows)\n" +
    "• Column 2 (wide): Name, Title, Email, Phone (one per row)\n\n" +
    "Example:\n" +
    "[LOGO]  |  John Doe\n" +
    "[LOGO]  |  Account Executive\n" +
    "[LOGO]  |  E: John.Doe@gmail.com\n" +
    "[LOGO]  |  M: 647-123-4567"
  );
  
  const sigToggleH = body.appendParagraph("🔘 Enabling Your Signature");
  sigToggleH.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  
  body.appendParagraph("After creating your signature document, the toggle is OFF by default. This gives you time to customize it first:");
  body.appendListItem("Customize your signature in the Google Doc").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("Return to Settings and flip the \"Append Signature to Emails\" toggle ON").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("The system will process your signature (upload images to Drive for email compatibility)").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("You'll see a confirmation with the signature size in KB").setGlyphType(DocumentApp.GlyphType.NUMBER);
  
  const sigRefreshH = body.appendParagraph("🔄 After Editing Your Signature");
  sigRefreshH.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  
  body.appendParagraph("If you make changes to your signature document later, you MUST click the \"🔄 Refresh\" button in Settings. " +
    "This re-processes your signature and updates the cached version. Without refreshing, your emails will still use the old signature.");
  
  const sigLinksH = body.appendParagraph("🔗 Hyperlinks in Your Signature");
  sigLinksH.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  
  body.appendParagraph("When you add hyperlinks in Google Docs, Google wraps them in redirect URLs (google.com/url?q=...). " +
    "Don't worry - the system automatically strips these and uses your original clean URLs in the final email. " +
    "Your links will appear exactly as you intended (e.g., yourwebsite.com, not google.com/url?q=yourwebsite.com).");
  
  body.appendHorizontalRule();
  
  // ======================================
  // SECTION 2: ADDING CONTACTS
  // ======================================
  const section2 = body.appendParagraph("➕ 2. Adding Contacts");
  section2.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  body.appendParagraph("Three ways to add contacts:");
  
  const method1H = body.appendParagraph("Method 1: In-App (Good for small batches)");
  method1H.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  body.appendListItem("Go to \"Add & Manage Contacts\" from main menu").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("Click \"Add Contact\"").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("Fill in: First Name, Last Name, Email, Company (required)").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("Select a sequence (email campaign) to assign them to").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("Optional: Title, Priority, Phone numbers").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("Click \"Add Contact\" - done!").setGlyphType(DocumentApp.GlyphType.NUMBER);
  
  const method2H = body.appendParagraph("Method 2: Direct in Google Sheets (Bulk imports)");
  method2H.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  body.appendParagraph("Open your connected Google Sheet and add rows directly. See Section 7 for column requirements!");
  
  const mandatoryColumnsH = body.appendParagraph("✅ MANDATORY COLUMNS (Contacts Sheet)");
  mandatoryColumnsH.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  mandatoryColumnsH.setForegroundColor("#137333");
  
  body.appendParagraph("When adding contacts directly in the sheet, these must be filled:");
  body.appendListItem("First Name - Required for emails").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Last Name - Required for emails").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Email - Where to send emails (duh!)").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Company - Required for templates").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Current Step - Set to 1 for new contacts").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Status - Set to \"Active\" for new contacts").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Sequence - Must match an existing sequence name exactly!").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const method3H = body.appendParagraph("Method 3: ZoomInfo Integration (⭐ RECOMMENDED)");
  method3H.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  method3H.setForegroundColor("#1a73e8");
  
  body.appendParagraph("This is my recommended approach for bulk imports! ZoomInfo exports are automatically detected and parsed:");
  
  body.appendListItem("In ZoomInfo, select contacts and export to email").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("Open the ZoomInfo email in Gmail").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("The add-on automatically detects and parses all contact info").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("Review parsed data, select a sequence, and click Import").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("All contacts are added instantly - ready for outreach!").setGlyphType(DocumentApp.GlyphType.NUMBER);
  
  const ziWhyH = body.appendParagraph("Why ZoomInfo is Best:");
  ziWhyH.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  
  body.appendListItem("✅ Fastest way to import bulk contacts").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("✅ Automatic parsing - no manual data entry").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("✅ Includes phone numbers, titles, and all contact details").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("✅ Duplicate detection prevents re-adding existing contacts").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("✅ Contacts immediately show in \"Ready to Send\" queue").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const contextH = body.appendParagraph("🔮 Contextual Adding");
  contextH.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph("When viewing an email in Gmail, if the sender isn't in your database, you'll see an \"Add Contact\" prompt with their info pre-filled!");
  
  body.appendHorizontalRule();
  
  // ======================================
  // SECTION 3: CALL MANAGEMENT
  // ======================================
  const section3 = body.appendParagraph("📞 3. Call Management");
  section3.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  body.appendParagraph("Got phone numbers? This feature helps you track call outreach alongside your emails.");
  
  const howWorksH = body.appendParagraph("How It Works 🔧");
  howWorksH.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  body.appendListItem("Contacts with phone numbers appear in Call Management").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Two view modes: Multi-Contact View or Focus Mode (power dialer)").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Click phone number to call, then mark as \"Called\"").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Once marked called, they leave the call list for this step").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("After the next email is sent, call tracking resets for another round").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const callActionsH = body.appendParagraph("Call Actions 📋");
  callActionsH.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  body.appendListItem("✓ Called - Marks phone as called").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Save + Keep - Save notes, keep in cadence").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Save + END - Save notes & complete their sequence").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Focus Mode outcomes: Spoke, No Answer, VM, Callback, No Interest, Bad #").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const callLimitH = body.appendParagraph("⚠️ Limitations");
  callLimitH.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph("Gmail add-ons have limitations - we can't actually dial for you. This is a tracking system. You make calls externally (phone, VoIP) and log outcomes here. It's simple but effective for call blitzing!");
  
  body.appendHorizontalRule();
  
  // ======================================
  // SECTION 4: TEMPLATES
  // ======================================
  const section4 = body.appendParagraph("✏️ 4. Create/Edit Templates");
  section4.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  body.appendParagraph("Templates are your email content. Each sequence has its own set of templates.");
  
  const createSeqH = body.appendParagraph("Creating a New Sequence");
  createSeqH.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  body.appendListItem("Go to \"Create/Edit Email Templates\"").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("Enter a name (e.g., \"Healthcare Outreach\")").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("Click \"Create Sequence\"").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("Set number of steps (1-5 emails)").setGlyphType(DocumentApp.GlyphType.NUMBER);
  body.appendListItem("Write your email templates for each step").setGlyphType(DocumentApp.GlyphType.NUMBER);
  
  const varsH = body.appendParagraph("📝 Available Variables");
  varsH.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  body.appendParagraph("Use these in your templates - they auto-populate with contact data:");
  body.appendListItem("{{firstName}} - Contact's first name").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("{{lastName}} - Contact's last name").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("{{company}} - Contact's company").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("{{title}} - Contact's job title").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("{{senderName}} - Your name (set in Settings)").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("{{senderCompany}} - Your company (set in Settings)").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("{{senderTitle}} - Your title (set in Settings)").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const stepConfigH = body.appendParagraph("Configuring Steps Per Sequence");
  stepConfigH.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  body.appendParagraph("Each sequence can have 1-5 steps:");
  body.appendListItem("Step 1: First email (requires subject line)").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Steps 2-5: Follow-ups (auto-reply to Step 1's thread)").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("To change step count: Edit sequence → \"Sequence Settings\"").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const templateTipsH = body.appendParagraph("💡 Template Tips");
  templateTipsH.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendListItem("Keep Step 1 short & curiosity-provoking").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Follow-ups should add value, not just \"checking in\"").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Duplicate successful sequences for A/B testing").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  body.appendHorizontalRule();
  
  // ======================================
  // SECTION 5: ANALYTICS
  // ======================================
  const section5 = body.appendParagraph("📊 5. Analytics");
  section5.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  body.appendParagraph("Track your campaign performance and optimize!");
  
  const analyticsViewsH = body.appendParagraph("Analytics Views");
  analyticsViewsH.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  const ev1 = body.appendParagraph("📤 Emails Sent Today");
  ev1.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph("See all companies you contacted today. Quick links to LinkedIn & other tools.");
  
  const ev2 = body.appendParagraph("💬 Reply Tracking");
  ev2.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph("View all contacts who have replied. System scans for replies daily at 4:30 AM, or click \"Check for Replies Now\" for instant scan.");
  
  const ev3 = body.appendParagraph("📈 Sequence Performance");
  ev3.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph("Compare performance across sequences. See reply rates, emails sent, breakdown by step. Great for A/B testing!");
  
  const ev4 = body.appendParagraph("📋 Template Effectiveness");
  ev4.setHeading(DocumentApp.ParagraphHeading.HEADING3);
  body.appendParagraph("Templates ranked by reply rate. Find your winners and losers!");
  
  const metricsH = body.appendParagraph("🎯 Key Metrics");
  metricsH.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendListItem("Reply Rate: % of emails that got responses").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("10%+ = Great 🏆").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("5-10% = Good ✅").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("1-5% = Average 🟡").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("<1% = Needs work 🔴").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  body.appendHorizontalRule();
  
  // ======================================
  // SECTION 6: SETTINGS
  // ======================================
  const section6 = body.appendParagraph("⚙️ 6. Settings");
  section6.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  body.appendParagraph("Customize everything to your workflow:");
  
  const set1H = body.appendParagraph("Email Sending Behavior");
  set1H.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendListItem("Auto-Send Step 1: When ON, first emails send immediately (not drafts). Use with caution!").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const set2H = body.appendParagraph("Sequence Settings");
  set2H.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendListItem("Days Between Emails: Default is 3. How long to wait between steps.").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const set3H = body.appendParagraph("Sender Info (for Template Variables)");
  set3H.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendListItem("Your Name: Populates {{senderName}}").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Your Company: Populates {{senderCompany}}").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Your Title: Populates {{senderTitle}}").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const set4H = body.appendParagraph("Email CC Settings");
  set4H.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendListItem("Add email addresses to automatically CC on outgoing emails").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const set5H = body.appendParagraph("Email Signature from Google Doc");
  set5H.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendListItem("Click \"Setup Signature\" to create a pre-formatted signature document").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Customize the doc, then toggle ON to enable (it's OFF by default)").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("After editing your signature, always click \"Refresh\" to update the cached version").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Images are auto-uploaded to Drive; Google redirect URLs are auto-stripped from links").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const set6H = body.appendParagraph("Step 2 PDF Attachment");
  set6H.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendListItem("Attach a PDF to Step 2 emails (e.g., case study, brochure)").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Select which sequences should get the attachment").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const set7H = body.appendParagraph("Priority Filters");
  set7H.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendListItem("Choose which priority levels to include when drafting emails").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Default: High, Medium, Low (all included)").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const set8H = body.appendParagraph("Manual Actions");
  set8H.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendListItem("Find & Label Sent Emails: Scans your sent folder to find & label email threads").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const set9H = body.appendParagraph("Database Connection");
  set9H.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendListItem("View or disconnect your connected Google Sheet").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Reconnect if you accidentally disconnected").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  body.appendHorizontalRule();
  
  // ======================================
  // SECTION 7: GOOGLE SHEETS
  // ======================================
  const section7 = body.appendParagraph("📊 7. Google Sheets - Important!");
  section7.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  const warningP = body.appendParagraph("⚠️ READ THIS TO AVOID BREAKING THINGS ⚠️");
  warningP.setBold(true);
  warningP.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendParagraph("Your Google Sheet is the database. Mess it up = mess up your campaigns.");
  
  const doNotH = body.appendParagraph("🚫 DO NOT");
  doNotH.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  doNotH.setForegroundColor("#d93025");
  
  body.appendListItem("Delete or rename the header row (Row 1)").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Delete or rename the \"Contacts\" sheet tab").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Rename the sequence sheets (they match sequence names)").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Add columns in the middle (only add at the end if needed)").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Delete columns the system uses (see below)").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const mandatoryH = body.appendParagraph("✅ MANDATORY COLUMNS (Contacts Sheet)");
  mandatoryH.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  mandatoryH.setForegroundColor("#137333");
  
  body.appendParagraph("When adding contacts directly in the sheet, these must be filled:");
  body.appendListItem("First Name - Required for emails").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Last Name - Required for emails").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Email - Where to send emails (duh!)").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Company - Required for templates").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Current Step - Set to 1 for new contacts").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Status - Set to \"Active\" for new contacts").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Sequence - Must match an existing sequence name exactly!").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const canDoH = body.appendParagraph("✅ SAFE TO DO");
  canDoH.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  body.appendListItem("Add new rows (contacts) at the bottom").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Edit existing contact data (names, companies, etc.)").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Delete entire contact rows (not just cell contents)").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Add columns at the very end for your own tracking").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Sort/filter data (just don't delete anything)").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  const statusH = body.appendParagraph("📌 Status Values");
  statusH.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  body.appendListItem("Active - Contact is in active email cadence").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Paused - Temporarily stopped (set manually)").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Completed - Finished all sequence steps").setGlyphType(DocumentApp.GlyphType.BULLET);
  body.appendListItem("Unsubscribed - Opted out / no interest").setGlyphType(DocumentApp.GlyphType.BULLET);
  
  body.appendHorizontalRule();
  
  // ======================================
  // FOOTER
  // ======================================
  body.appendParagraph("");
  const footer = body.appendParagraph("🎉 You're all set! Go crush those outreach campaigns!");
  footer.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendParagraph("");
  const lastUpdated = body.appendParagraph("Last updated: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM dd, yyyy"));
  lastUpdated.setItalic(true);
  lastUpdated.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  lastUpdated.setForegroundColor("#666666");
  
  // Save and return doc ID
  doc.saveAndClose();
  return doc.getId();
}

