/**
 * FILE: Configuration and Constants
 * 
 * PURPOSE:
 * Defines all global configuration constants, column mappings, and industry options
 * used throughout the application.
 * 
 * KEY CONTENTS:
 * - CONFIG: Main application configuration object
 * - CONTACT_COLS: Column index mapping for Contacts sheet
 * - INDUSTRY_OPTIONS: Available industry categories
 * - ZOOMINFO_INDUSTRY_OPTIONS: Industry options for ZoomInfo imports
 * 
 * DEPENDENCIES:
 * - None (this is the foundation file)
 * 
 * USED BY:
 * - All other modules depend on these constants
 * 
 * @version 2.3
 */

// ============================================================
// MAIN APPLICATION CONFIGURATION
// ============================================================

const CONFIG = {
  SPREADSHEET_NAME: "Google Contacts for Sequencing",
  CONTACTS_SHEET_NAME: "Contacts",
  TEMPLATES_SHEET_NAME: "Email Templates",
  LOGS_SHEET_NAME: "Logs",
  DEFAULT_DELAY_DAYS: 3,
  SEQUENCE_STEPS: 5,
  EMAIL_QUOTA_LIMIT: 100, // Gmail daily quota limit for programmatic sending
  PAGE_SIZE: 15 // Number of contacts per page
};

// ============================================================
// CONTACTS SHEET COLUMN MAPPING (0-based indices)
// ============================================================

const CONTACT_COLS = {
  FIRST_NAME: 0,
  LAST_NAME: 1,
  EMAIL: 2,
  COMPANY: 3,
  TITLE: 4,
  CURRENT_STEP: 5,
  LAST_EMAIL_DATE: 6,
  NEXT_STEP_DATE: 7,
  STATUS: 8,
  NOTES: 9,
  PERSONAL_PHONE: 10,
  WORK_PHONE: 11,
  PERSONAL_CALLED: 12,
  WORK_CALLED: 13,
  PRIORITY: 14,
  TAGS: 15,
  SEQUENCE: 16,
  INDUSTRY: 17,
  STEP1_SUBJECT: 18,
  STEP1_SENT_MESSAGE_ID: 19,
  CONNECT_SALES_LINK: 20,
  THREAD_ID: 21,
  LABELED: 22,
  REPLY_RECEIVED: 23,
  REPLY_DATE: 24
};

// ============================================================
// INDUSTRY OPTIONS
// ============================================================

const INDUSTRY_OPTIONS = [
  "Ecommerce / Retail",
  "SaaS / B2B Tech", 
  "Local Services",
  "Healthcare & Wellness",
  "Finance & FinTech",
  "Travel & Tourism",
  "Real Estate & Housing",
  "Professional Services",
  "Education & Online Learning",
  "Food & Beverage CPG",
  "Pets & Animal Services",
  "Trades & Industrial / Manufacturing",
  "Creative & Media / Agencies",
  "Lifestyle & Personal Services"
];

const ZOOMINFO_INDUSTRY_OPTIONS = [
  "Ecommerce / Retail",
  "SaaS / B2B Tech",
  "Local Services",
  "Healthcare & Wellness",
  "Finance & FinTech",
  "Travel & Tourism",
  "Real Estate & Housing",
  "Professional Services",
  "Education & Online Learning",
  "Food & Beverage CPG",
  "Pets & Animal Services",
  "Trades & Industrial / Manufacturing",
  "Creative & Media / Agencies",
  "Lifestyle & Personal Services"
];

