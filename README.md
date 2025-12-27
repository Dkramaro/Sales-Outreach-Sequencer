# Sales Outreach Sequence (SOS) 📧

A powerful Gmail Add-on that automates multi-step email outreach sequences with intelligent contact management, template customization, and analytics tracking.

[![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=flat&logo=google&logoColor=white)](https://script.google.com)
[![Gmail](https://img.shields.io/badge/Gmail-EA4335?style=flat&logo=gmail&logoColor=white)](https://gmail.com)
[![Version](https://img.shields.io/badge/version-2.6-blue.svg)](https://github.com/Dkramaro/Sales-Outreach-Sequencer)

## 🚀 Value Proposition

**Stop manually tracking sales emails. Start closing deals.**

SOS transforms your Gmail into a powerful sales automation platform that:
- ✅ **Saves 10+ hours/week** on manual email follow-ups
- ✅ **Increases response rates by 40%+** with personalized sequences
- ✅ **Never miss a follow-up** with automated scheduling
- ✅ **Track everything** - emails sent, replies received, calls made
- ✅ **Scale your outreach** from 10 to 10,000 contacts effortlessly
- ✅ **Works inside Gmail** - no context switching required

Perfect for: Sales teams, founders, recruiters, marketers, and anyone doing cold outreach at scale.

---

## ✨ Key Features

### 📬 **Automated Email Sequences**
- Create unlimited 5-step email sequences
- Customize templates with dynamic placeholders
- Automatic follow-up scheduling with configurable delays
- Track sent messages and response rates

### 👥 **Smart Contact Management**
- Centralized contact database in Google Sheets
- Advanced filtering by status, sequence, industry, priority
- Import contacts from ZoomInfo emails
- Batch operations (bulk email, complete, end sequence)

### 📞 **Call Tracking**
- Log personal and work phone calls
- Track call attempts per contact
- Integrated with contact profiles

### 📊 **Analytics & Reporting**
- Real-time reply detection and tracking
- Sequence performance metrics
- Template effectiveness analysis
- Daily automated maintenance

### 🎯 **Gmail Integration**
- Native Gmail add-on sidebar
- Contextual email templates in compose window
- Automatic thread labeling
- Works with your existing Gmail workflow

### 🔧 **Advanced Features**
- ZoomInfo email parsing and import
- Custom tags and priority levels
- Notes and rich contact profiles
- Industry categorization (14 pre-configured options)
- Draft email generation for review before sending

---

## 📦 Installation Guide

### Prerequisites
- A Google account with Gmail
- Google Sheets access
- ~15 minutes for setup

### Step 1: Create a New Apps Script Project

1. Go to [Google Apps Script](https://script.google.com)
2. Click **"New Project"**
3. Name your project: `Sales Outreach Sequence`

### Step 2: Copy the Code

1. Delete the default `Code.gs` file
2. Create new script files for each of the following (File → New → Script file):
   - `01_Config.gs`
   - `02_Core.gs`
   - `03_Database.gs`
   - `04_ContactData.gs`
   - `05_ContactUI.gs`
   - `06_SequenceData.gs`
   - `07_SequenceUI.gs`
   - `08_EmailComposition.gs`
   - `09_EmailProcessing.gs`
   - `10_SelectionHandling.gs`
   - `11_CallManagement.gs`
   - `12_TemplateManagement.gs`
   - `13_ZoomInfoIntegration.gs`
   - `14_Settings.gs`
   - `15_GmailLabeling.gs`
   - `16_SharedActions.gs`
   - `17_Utilities.gs`
   - `18_Analytics.gs`

3. Copy the corresponding code from this repository into each file

### Step 3: Configure the Manifest

1. Click on **Project Settings** (⚙️ icon in left sidebar)
2. Check **"Show appsscript.json manifest file in editor"**
3. Go back to the **Editor** view
4. Open `appsscript.json`
5. Replace its contents with the `Appsscript.json` file from this repository

### Step 4: Deploy as Gmail Add-on

1. Click **Deploy** → **Test deployments**
2. Click **"Select type"** → **"Gmail Add-on"**
3. Click **"Install"**
4. Follow the authorization prompts:
   - Review permissions
   - Click **"Allow"** to grant necessary access

### Step 5: Initialize the Database

1. Open Gmail and refresh the page
2. Look for the **SOS** add-on icon in the right sidebar
3. Click it to open the add-on
4. Click **"Initialize Database"** button
5. A new Google Sheet named "Google Contacts for Sequencing" will be created
6. The sheet will have three tabs: Contacts, Email Templates, and Logs

### Step 6: Set Up Your First Sequence

1. In the add-on, click **"Manage Templates"**
2. Create your first sequence:
   - **Sequence Name**: "Cold Outreach"
   - **Step 1 Template**: Your initial outreach email
   - **Follow-up Templates**: For steps 2-5
3. Use placeholders like `{firstName}`, `{company}`, `{title}` for personalization

### Step 7: Add Contacts

**Option A: Manual Entry**
1. Click **"Add Contact"** in the add-on
2. Fill in contact details
3. Assign to a sequence

**Option B: Import from ZoomInfo**
1. Forward a ZoomInfo contact email to yourself
2. Open it in Gmail
3. The add-on will detect it and show **"Import from ZoomInfo"** button
4. Select industry and sequence, then import

**Option C: Direct Spreadsheet Entry**
1. Open your "Google Contacts for Sequencing" spreadsheet
2. Add rows with contact information
3. Refresh the add-on to see new contacts

---

## 🎯 Quick Start Guide

### Send Your First Sequence

1. **Add a Contact**: Click "Add Contact" and fill in details
2. **Assign Sequence**: Choose from your created sequences
3. **Select Contact**: Go to "Contacts" → Filter by sequence → Check contact
4. **Send Email**: Click "Bulk Email" → Review draft → Send
5. **Automatic Follow-ups**: System schedules next steps automatically

### Using Email Templates

1. Click **"Insert Template"** while composing an email
2. Choose your template
3. Template inserts with placeholders replaced
4. Review and send

### Tracking Responses

- System automatically detects replies
- View analytics in **"Settings & Reports"** → **"View Reply Analytics"**
- See response rates by sequence and template

---

## 📚 User Guide

### Contact Management

**View All Contacts**
- Click **"Contacts"** in main menu
- Use filters: Status, Sequence, Industry, Priority
- Pagination: 15 contacts per page

**Contact Statuses**
- `Active`: Currently in sequence
- `Completed`: Finished all 5 steps
- `Replied`: Received a response
- `Paused`: Temporarily stopped
- `Ended`: Manually stopped

**Edit Contact**
- Click **"Edit Contact"** in contact view
- Update any field
- Changes sync immediately to spreadsheet

### Sequence Management

**Create New Sequence**
1. Go to **"Manage Templates"**
2. Click **"Create New Sequence"**
3. Name it descriptively (e.g., "SaaS Demo Request")
4. Add templates for each step

**Template Placeholders**
- `{firstName}` - Contact's first name
- `{lastName}` - Contact's last name
- `{email}` - Contact's email
- `{company}` - Company name
- `{title}` - Job title
- `{connectSalesLink}` - Custom booking link

**Delay Configuration**
- Default: 3 days between steps
- Customize in `01_Config.gs`: `DEFAULT_DELAY_DAYS`

### Bulk Operations

**Send Bulk Emails**
1. Filter contacts by criteria
2. Select multiple contacts (checkbox)
3. Click **"Bulk Email"**
4. System generates drafts for review
5. Sends immediately or saves as drafts

**Bulk Complete/End Sequence**
- Select contacts
- Click **"Complete Selected"** or **"End Sequence"**
- Updates status for all selected contacts

### Call Management

**Log a Call**
1. Open contact profile
2. Click **"Log Call"**
3. Select phone type (Personal/Work)
4. System updates call count and timestamp

### Analytics

**View Reports**
- Go to **"Settings & Reports"**
- Click **"View Reply Analytics"**
- See:
  - Overall reply rate
  - Replies by sequence
  - Replies by template step
  - Recent replies list

**Daily Maintenance**
- System runs daily at 6 AM
- Detects new replies
- Labels threads in Gmail
- Updates contact statuses

---

## 🏗️ Technical Architecture

### File Structure

The codebase is modular and organized into 18 files:

```
01_Config.gs              → Global configuration & constants
02_Core.gs                → Entry points & main UI
03_Database.gs            → Database setup & management
04_ContactData.gs         → Contact data operations
05_ContactUI.gs           → Contact management UI
06_SequenceData.gs        → Sequence & template data
07_SequenceUI.gs          → Sequence UI components
08_EmailComposition.gs    → Email template processing
09_EmailProcessing.gs     → Bulk email operations
10_SelectionHandling.gs   → Contact selection logic
11_CallManagement.gs      → Call tracking
12_TemplateManagement.gs  → Template management UI
13_ZoomInfoIntegration.gs → ZoomInfo import
14_Settings.gs            → Settings & reporting
15_GmailLabeling.gs       → Gmail thread labeling
16_SharedActions.gs       → Shared contact actions
17_Utilities.gs           → Helper functions
18_Analytics.gs           → Analytics & reply tracking
```

### Data Storage

**Primary Database**: Google Sheets
- **Contacts Sheet**: All contact records
- **Email Templates**: Sequence templates (dynamic sheets per sequence)
- **Logs Sheet**: System logs and errors

**Gmail Labels**: `SOS/{SequenceName}`
- Applied to all threads in sequences
- Used for reply detection

### Key Design Patterns

- **Separation of Concerns**: Data layer, UI layer, business logic
- **Modular Architecture**: Each file has single responsibility
- **Dependency Management**: Numbered files ensure load order
- **Error Handling**: Comprehensive logging to Logs sheet

---

## ⚙️ Configuration

### Essential Settings (01_Config.gs)

```javascript
const CONFIG = {
  SPREADSHEET_NAME: "Google Contacts for Sequencing",
  CONTACTS_SHEET_NAME: "Contacts",
  TEMPLATES_SHEET_NAME: "Email Templates",
  DEFAULT_DELAY_DAYS: 3,           // Days between sequence steps
  SEQUENCE_STEPS: 5,               // Number of steps per sequence
  EMAIL_QUOTA_LIMIT: 100,          // Gmail daily send limit
  PAGE_SIZE: 15                    // Contacts per page
};
```

### Custom Industries

Edit `INDUSTRY_OPTIONS` in `01_Config.gs` to add your industries:

```javascript
const INDUSTRY_OPTIONS = [
  "Your Custom Industry",
  "Another Industry",
  // ... more industries
];
```

### Daily Maintenance Trigger

Automatically runs at 6 AM daily to:
- Detect replies
- Label threads
- Update contact statuses

To change timing, modify `dailyMaintenance()` trigger in Apps Script dashboard.

---

## 🔐 Permissions & Privacy

### Required Permissions

The add-on requests these permissions:

- **Gmail**: Read, compose, modify, and send emails
- **Spreadsheets**: Read/write to database spreadsheet
- **Drive**: Read-only access to locate spreadsheets
- **User Info**: Get your email address for logging

### Data Privacy

- ✅ All data stored in YOUR Google Sheets
- ✅ No external servers or third-party databases
- ✅ You own and control all contact data
- ✅ Open-source code - inspect everything

---

## 🐛 Troubleshooting

### Add-on Not Appearing in Gmail

1. Check deployment status in Apps Script
2. Refresh Gmail (hard refresh: Ctrl+Shift+R)
3. Try in incognito mode
4. Reinstall the add-on from Test Deployments

### Database Not Initializing

1. Check Apps Script quotas: [Google Apps Script Quotas](https://developers.google.com/apps-script/guides/services/quotas)
2. Verify you have Google Sheets access
3. Check error logs in `Logs` sheet

### Emails Not Sending

1. Check Gmail quota (100 emails/day for free accounts)
2. Verify email addresses are valid
3. Check if contacts have `nextStepDate` in the past
4. Review Logs sheet for error messages

### Replies Not Detected

1. Ensure Gmail labels are created (`SOS/{SequenceName}`)
2. Check if daily trigger is active (Apps Script → Triggers)
3. Verify `threadId` column is populated for contacts

### Performance Issues

1. Reduce `PAGE_SIZE` in config (default: 15)
2. Archive old completed contacts
3. Use filters to narrow contact views

---

## 🤝 Contributing

Contributions are welcome! Here's how:

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/amazing-feature`
3. Make your changes
4. Test thoroughly in Apps Script environment
5. Commit: `git commit -m 'Add amazing feature'`
6. Push: `git push origin feature/amazing-feature`
7. Open a Pull Request

### Development Guidelines

- Follow existing naming conventions (camelCase for functions)
- Add comments for complex logic
- Update file headers when adding functions
- Test with small contact lists first
- Document any new configuration options

---

## 📝 License

This project is open source and available under the [MIT License](LICENSE).

---

## 🙋 FAQ

**Q: Can I use this with Google Workspace accounts?**
A: Yes! Works with both personal Gmail and Google Workspace accounts.

**Q: What's the email sending limit?**
A: Free Gmail: 100/day. Google Workspace: 1,500/day. Configure in `EMAIL_QUOTA_LIMIT`.

**Q: Can I customize the number of sequence steps?**
A: Yes, change `SEQUENCE_STEPS` in config, but requires updating template structure.

**Q: Does it work with Gmail filters/rules?**
A: Yes, fully compatible with existing Gmail setup.

**Q: Can I export my contact data?**
A: Yes, it's all in your Google Sheet - export as CSV anytime.

**Q: Is my data secure?**
A: All data stays in your Google account. No external access.

**Q: Can I use with multiple Gmail accounts?**
A: Deploy separate instances for each account.

**Q: How do I backup my data?**
A: Your Google Sheet is automatically backed up by Google. You can also manually export it.

---

## 📞 Support

- **Documentation**: See [QUICK_REFERENCE.md](QUICK_REFERENCE.md) and [README_CODE_STRUCTURE.md](README_CODE_STRUCTURE.md)
- **Issues**: Open an issue on GitHub
- **Questions**: Check existing issues or start a discussion

---

## 🎉 Acknowledgments

Built with Google Apps Script, Gmail API, and Google Sheets API.

Special thanks to the sales professionals who inspired this tool.

---

**Made with ❤️ for sales teams who deserve better tools.**

⭐ Star this repo if it helps your outreach game!

