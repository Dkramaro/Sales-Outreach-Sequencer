# Quick Reference Guide - Finding Functions Fast

## 🔍 Where to Find What

### Need to work with contacts?
- **Data operations** (get, filter, query) → `04_ContactData.gs`
- **UI & forms** (view, edit, add) → `05_ContactUI.gs`

### Need to work with emails?
- **Individual composition** → `08_EmailComposition.gs`
- **Bulk sending/drafts** → `09_EmailProcessing.gs`
- **Templates & placeholders** → `08_EmailComposition.gs`

### Need to work with sequences?
- **Sequence data** (create, get templates) → `06_SequenceData.gs`
- **Sequence UI** (step views, dashboards) → `07_SequenceUI.gs`
- **Template management UI** → `12_TemplateManagement.gs`

### Need to work with special features?
- **Call tracking** → `11_CallManagement.gs`
- **ZoomInfo import** → `13_ZoomInfoIntegration.gs`
- **Gmail labeling** → `15_GmailLabeling.gs`
- **Settings & config** → `14_Settings.gs`

### Need utilities?
- **Formatting & helpers** → `17_Utilities.gs`
- **Configuration constants** → `01_Config.gs`

### Need entry points?
- **Main dashboard & triggers** → `02_Core.gs`
- **Database setup** → `03_Database.gs`

## 📊 Function Count by File

| File | Functions | Purpose |
|------|-----------|---------|
| 01_Config.gs | 3 constants | Configuration foundation |
| 02_Core.gs | 7 functions | Entry points & main UI |
| 03_Database.gs | 7 functions | Database management |
| 04_ContactData.gs | 10 functions | Contact data operations |
| 05_ContactUI.gs | 7 functions | Contact UI |
| 06_SequenceData.gs | 7 functions | Sequence & template data |
| 07_SequenceUI.gs | 5 functions | Sequence UI |
| 08_EmailComposition.gs | 5 functions | Email composition |
| 09_EmailProcessing.gs | 4 functions | Bulk email processing |
| 10_SelectionHandling.gs | 10 functions | Selection & filtering |
| 11_CallManagement.gs | 5 functions | Call management |
| 12_TemplateManagement.gs | 9 functions | Template management |
| 13_ZoomInfoIntegration.gs | 5 functions | ZoomInfo integration |
| 14_Settings.gs | 4 functions | Settings & reporting |
| 15_GmailLabeling.gs | 3 functions | Gmail labeling |
| 16_SharedActions.gs | 2 functions | Shared actions |
| 17_Utilities.gs | 11 functions | Utilities |

## 🎯 Most Common Tasks

### Adding a new contact field
1. Update `CONTACT_COLS` in `01_Config.gs`
2. Update `setupContactsSheet()` in `03_Database.gs`
3. Update `getAllContactsData()` in `04_ContactData.gs`
4. Update UI forms in `05_ContactUI.gs`

### Adding a new sequence step
1. Update `CONFIG.SEQUENCE_STEPS` in `01_Config.gs`
2. Update template sheets in `06_SequenceData.gs`
3. Update step views in `07_SequenceUI.gs`

### Adding a new placeholder
1. Update `replacePlaceholders()` in `08_EmailComposition.gs`
2. Document in template help sections

### Debugging email issues
1. Check `composeEmailWithTemplate()` in `08_EmailComposition.gs`
2. Check `processBulkEmails()` in `09_EmailProcessing.gs`
3. Check `findSentStep1Messages_Batch()` in `09_EmailProcessing.gs`

## 🚀 For AI Models

When searching for functionality:
1. Check file headers first (each has PURPOSE and KEY FUNCTIONS)
2. Use file numbers to understand load order
3. Follow DEPENDENCIES map in README_CODE_STRUCTURE.md
4. Function names are descriptive and follow camelCase

## 📝 Maintenance Tips

- Keep file headers updated when adding functions
- Maintain dependency documentation
- Group related functions within files using section comments
- Use JSDoc comments for complex functions
- Test modules independently where possible

