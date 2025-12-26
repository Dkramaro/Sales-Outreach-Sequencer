# Sales Outreach Sequence (SOS) - Code Structure

## Overview
This Gmail Add-on automates multi-step email outreach sequences with contact management, template customization, and ZoomInfo integration.

## File Organization

The codebase is organized into 18 modular files, numbered for load order and grouped by functional area:

### Core Foundation (01-03)
- **01_Config.gs** - Global configuration, constants, column mappings
- **02_Core.gs** - Entry points, triggers, main dashboard, contextual email handling
- **03_Database.gs** - Database setup, sheet initialization, connection management, logging

### Contact Management (04-05)
- **04_ContactData.gs** - Contact data operations, queries, filtering, statistics
- **05_ContactUI.gs** - Contact management UI, views, forms, search

### Sequence & Template System (06-07)
- **06_SequenceData.gs** - Sequence discovery, creation, template data retrieval
- **07_SequenceUI.gs** - Sequence management UI, step views, contact selection

### Email Operations (08-09)
- **08_EmailComposition.gs** - Template processing, placeholder replacement, signatures
- **09_EmailProcessing.gs** - Bulk email operations, batch processing, message search

### Specialized Features (10-13)
- **10_SelectionHandling.gs** - Contact selection state, filtering, checkbox handling
- **11_CallManagement.gs** - Call tracking UI and operations
- **12_TemplateManagement.gs** - Template/sequence management UI
- **13_ZoomInfoIntegration.gs** - ZoomInfo email parsing and contact import

### Settings & Integrations (14-15)
- **14_Settings.gs** - User settings, configuration UI
- **15_GmailLabeling.gs** - Background Gmail labeling with batch operations

### Analytics & Reporting (18)
- **18_Analytics.gs** - Analytics hub, reply tracking, sequence/template performance, daily triggers, thread pre-validation

### Shared Components (16-17)
- **16_SharedActions.gs** - Shared contact actions (complete, end sequence)
- **17_Utilities.gs** - Helper functions (formatting, validation, UI components)

## Key Design Patterns

### Separation of Concerns
- **Data Layer**: Files 04, 06 - Pure data operations, no UI
- **UI Layer**: Files 05, 07, 11, 12, 14 - Card building and forms
- **Business Logic**: Files 08, 09, 10, 13, 15, 16 - Processing and workflows

### Naming Conventions
- Functions use camelCase
- Constants use UPPER_SNAKE_CASE
- File headers document PURPOSE, KEY FUNCTIONS, DEPENDENCIES

### Load Order
Files are numbered 01-17 to ensure proper load sequence (Apps Script loads alphabetically).

## Dependencies Map

```
01_Config.gs (no dependencies - foundation)
    ↓
02_Core.gs → 03, 04, 05, 13
03_Database.gs → 01, 06
04_ContactData.gs → 01, 06
05_ContactUI.gs → 01, 04, 06, 17
06_SequenceData.gs → 01, 03, 04
07_SequenceUI.gs → 01, 04, 06, 10
08_EmailComposition.gs → 01, 04, 06
09_EmailProcessing.gs → 01, 04, 06, 08, 10
10_SelectionHandling.gs → 01, 04, 07, 17
11_CallManagement.gs → 01, 04, 05, 16, 17
12_TemplateManagement.gs → 01, 03, 06, 08
13_ZoomInfoIntegration.gs → 01, 03, 06
14_Settings.gs → 01, 03, 04, 06, 17
15_GmailLabeling.gs → 01, 03, 04
16_SharedActions.gs → 01, 03, 04, 05, 07, 11
17_Utilities.gs → 01, 03
18_Analytics.gs → 01, 03, 04, 06, 17
```

## For AI Models

### Semantic Search Optimization
Each file has a clear PURPOSE statement making it easy to find relevant code:
- Looking for contact data? → `04_ContactData.gs`
- Need UI for contacts? → `05_ContactUI.gs`
- Email composition logic? → `08_EmailComposition.gs`
- ZoomInfo parsing? → `13_ZoomInfoIntegration.gs`

### Function Location Guide
- **Add/Edit Contacts**: Files 04, 05
- **Email Sending**: Files 08, 09
- **Sequence Management**: Files 06, 07
- **Call Tracking**: File 11
- **Settings**: File 14
- **ZoomInfo Import**: File 13
- **Utilities**: File 17
- **Analytics & Reporting**: File 18

## Migration Notes

This structure replaces the previous "Block1", "Block2", "Block3", "Block4", and "Zoominfo" files with descriptive, functionally-organized modules.

### What Changed
- ✅ All functions preserved unchanged
- ✅ Logical grouping by feature area
- ✅ File headers with documentation
- ✅ Clear dependency chains
- ✅ Numbered for load order

### What Stayed the Same
- All function signatures
- All business logic
- All data structures
- Configuration constants

## Development Guidelines

1. **Adding New Features**: Place in the most appropriate numbered file or create a new one
2. **Modifying Functions**: Update file headers if dependencies change
3. **Testing**: Test functions in isolation using the modular structure
4. **Documentation**: Keep file headers updated with KEY FUNCTIONS list

## Version
2.6 - Combined thread validation + reply detection into single dailyMaintenance trigger (December 2025)
2.5 - Added thread pre-validation trigger for changed subject lines (December 2025)
2.4 - Added Analytics module with reply tracking (December 2025)

