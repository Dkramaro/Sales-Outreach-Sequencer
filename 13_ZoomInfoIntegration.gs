/**
 * FILE: ZoomInfo Email Integration
 * 
 * PURPOSE:
 * Handles ZoomInfo email detection, contact extraction, and bulk import.
 * Supports Unicode names, multiple email formats, and editable import fields
 * including tags, industry, and sequence selection.
 * 
 * KEY FUNCTIONS:
 * - isZoomInfoEmail() - Detect ZoomInfo emails
 * - extractZoomInfoContacts() - Parse contacts from email body
 * - buildZoomInfoImportCard() - Build import UI with editable fields
 * - importAllZoomInfoContacts() - Bulk import handler
 * - importSingleZoomInfoContact() - Single import handler
 * - addContactToDatabase() - Add contact to sheet (called from ZoomInfo)
 * 
 * DEPENDENCIES:
 * - 01_Config.gs: CONFIG, CONTACT_COLS, ZOOMINFO_INDUSTRY_OPTIONS
 * - 03_Database.gs: logAction
 * - 06_SequenceData.gs: getAvailableSequences
 * 
 * @version 2.3 - Added Unicode support for names with accents
 */

// ============================================================
// == ZOOMINFO EMAIL DETECTION ==
// ============================================================

/**
* Detects if the current email is from ZoomInfo.
*/
function isZoomInfoEmail(message) {
if (!message) return false;
const from = message.getFrom().toLowerCase();
return from.includes("@zoominfo.com") || from.includes(" zoominfo <");
}

// ============================================================
// == ZOOMINFO CONTACT EXTRACTION ==
// ============================================================

/**
* Enhanced extraction function for ZoomInfo emails
* Handles different email formats without hard-coding any contact information.
* Now supports names with accented characters like "Amanda Dubé".
*/
function extractZoomInfoContacts(message) {
const body = message.getPlainBody();
const contacts = [];

// Clean the body and normalize line breaks
const cleanedBody = body.replace(/\r\n/g, '\n').replace(/ +/g, ' ').trim();

console.log("Processing ZoomInfo email body...");

// Get the expected number of contacts
const recordsMatch = cleanedBody.match(/Number of records exported:?\s*(\d+)/i);
const expectedCount = recordsMatch ? parseInt(recordsMatch[1]) : 0;
console.log(`Expected contacts: ${expectedCount}`);

// --- Unicode-aware Regex Definitions ---
// \p{Lu} - an uppercase letter (Unicode)
// \p{L}  - any kind of letter (Unicode)
// [\p{L}.'-']* - zero or more letters, periods, apostrophes (standard & typographic), or hyphens
const generalNamePartPattern = "\\p{Lu}[\\p{L}.'-']*"; // Allows for names like O'Malley, Dubé, Jean-Luc

// Validates a full name: one or more generalNamePartPattern(s) separated by spaces. 'u' flag for Unicode.
const fullNameValidationPattern = new RegExp(`^${generalNamePartPattern}(?:\\s+${generalNamePartPattern})*$`, 'u');

// Captures a full name that must have at least two parts (e.g., First Last).
// ^(...)$ ensures it matches the whole string/line. 'u' flag for Unicode.
const multiPartNameCapturePatternGlobal = new RegExp(`^(${generalNamePartPattern}(?:\\s+${generalNamePartPattern})+)$`, 'u');

// Specifically for lineScan: expects exactly two name parts. 'u' flag for Unicode.
const lineScanNamePattern = new RegExp(`^${generalNamePartPattern}\\s+${generalNamePartPattern}$`, 'u');

// For Approach 1.5 (modernZoomInfoPattern), name capture group expects at least two parts.
// 'u' flag for Unicode.
const modernZoomInfoNameCaptureGroup = `(${generalNamePartPattern}(?:\\s+${generalNamePartPattern})+)`;
const modernZoomInfoPattern = new RegExp(
    `^${modernZoomInfoNameCaptureGroup}\\s*\\n` + // Name (at least two parts)
    `([^@\\n]+)\\s*\\n` +                         // Title, Company line
    `Add\\s*\\n\\s*` +                            // "Add" literal line
    `([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,})`, // Email
    'gmu' // Global, Multiline, Unicode
);
// --- End Regex Definitions ---

// APPROACH 1: Extract multiple contacts using precise ZoomInfo pattern matching
if (contacts.length < expectedCount || expectedCount === 0) {
  console.log("Trying precise ZoomInfo pattern extraction (Approach 1)...");
   const emailPattern = /([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g;
  const emailMatches = [...cleanedBody.matchAll(emailPattern)];
   console.log(`Found ${emailMatches.length} total email addresses for precise extraction`);
   for (const emailMatch of emailMatches) {
    try {
      const email = emailMatch[1].trim();
      const emailIndex = emailMatch.index;
    
      if (email.toLowerCase().includes('zoominfo.com')) {
        continue;
      }
    
      if (contacts.some(c => c.email.toLowerCase() === email.toLowerCase())) {
        continue;
      }
    
      const textBeforeEmail = cleanedBody.substring(0, emailIndex);
      const lines = textBeforeEmail.split('\n');
    
      let addLineIndex = -1;
      for (let i = lines.length - 1; i >= 0; i--) {
        if (lines[i].trim().toLowerCase() === 'add') {
          addLineIndex = i;
          break;
        }
      }
    
      if (addLineIndex === -1) {
        continue;
      }
    
      let titleCompanyLine = "";
      let fullName = "";
    
      if (addLineIndex > 0) {
        titleCompanyLine = lines[addLineIndex - 1].trim();
      }
    
      if (addLineIndex > 1) {
        fullName = lines[addLineIndex - 2].trim();
      }
    
      // Validate we found a proper name using Unicode-aware pattern
      if (!fullName || !fullNameValidationPattern.test(fullName)) {
        continue;
      }
    
      // Filter out lines that are likely addresses or generic company info
      if (fullName.includes('St,') || fullName.includes('Quebec') || fullName.includes('Montreal') ||
          fullName.includes('Toronto') || fullName.includes('Canada') || fullName.includes('Avenue') ||
          fullName.includes('Road') || fullName.includes('Drive') || fullName.includes('Broadway') ||
          fullName.includes('Vancouver') || fullName.toLowerCase().includes('zoom information') ||
          /\d/.test(fullName) // Skip if name contains numbers (likely an address part)
        ) {
        continue;
      }
    
      const nameParts = fullName.split(/\s+/);
      const firstName = nameParts[0] || "";
      const lastName = nameParts.slice(1).join(" ") || "";
    
      let title = "";
      let company = "";
    
      if (titleCompanyLine) {
        if (titleCompanyLine.includes(', ')) {
          const lastCommaIndex = titleCompanyLine.lastIndexOf(', ');
          if (lastCommaIndex !== -1) {
            title = titleCompanyLine.substring(0, lastCommaIndex).trim();
            company = titleCompanyLine.substring(lastCommaIndex + 2).trim();
          } else {
            title = titleCompanyLine; // Assume it's all title if no comma found as expected
          }
        } else {
          title = titleCompanyLine;
        }
      }
    
      // Look for phone numbers in the context around this contact
      let phoneContextStartIndex = 0;
      if (addLineIndex > 2) { // if fullName line exists
          phoneContextStartIndex = textBeforeEmail.lastIndexOf('\n', lines.slice(0, addLineIndex - 2).join('\n').length) +1;
      } else if (addLineIndex > 1) { // if titleCompanyLine exists (but not fullName)
           phoneContextStartIndex = textBeforeEmail.lastIndexOf('\n', lines.slice(0, addLineIndex - 1).join('\n').length) +1;
      }
      const phoneContextEndIndex = Math.min(cleanedBody.length, emailIndex + email.length + 100); // a bit after the email
      const contactContext = cleanedBody.substring(phoneContextStartIndex, phoneContextEndIndex);
    
      let workPhone = "";
      let personalPhone = "";
    
      const directMatch = contactContext.match(/Direct:?\s*\(?(\d{3})\)?[\s.-]?(\d{3})[\s.-]?(\d{4})/i);
      if (directMatch) {
        workPhone = `${directMatch[1]}-${directMatch[2]}-${directMatch[3]}`;
      }
    
      const personalMatch = contactContext.match(/Personal:?\s*\(?(\d{3})\)?[\s.-]?(\d{3})[\s.-]?(\d{4})/i);
      if (personalMatch) {
        personalPhone = `${personalMatch[1]}-${personalMatch[2]}-${personalMatch[3]}`;
      }
    
      if (firstName && lastName && email) {
        contacts.push({
          firstName: firstName,
          lastName: lastName,
          email: email,
          company: company,
          title: title,
          workPhone: workPhone,
          personalPhone: personalPhone,
          priority: "Medium",
          tags: "",
          industry: ""
        });
        console.log(`Approach 1: Extracted ${firstName} ${lastName}, ${email}, "${title}", "${company}"`);
      }
    
    } catch (error) {
      console.error(`Approach 1: Error processing contact for email: ${error.message}`);
    }
  }
}

// APPROACH 1.5: New specialized approach for modern ZoomInfo format
if (contacts.length < expectedCount || (expectedCount === 0 && contacts.length === 0)) {
  console.log("Trying specialized ZoomInfo format extraction (Approach 1.5)...");
   const modernMatches = [...cleanedBody.matchAll(modernZoomInfoPattern)]; // Uses pre-defined Unicode-aware pattern
   console.log(`Found ${modernMatches.length} contacts using modern ZoomInfo pattern`);
   for (const match of modernMatches) {
    try {
      const fullName = match[1].trim(); // Captured by modernZoomInfoNameCaptureGroup
      const titleCompanyLine = match[2].trim();
      const email = match[3].trim();
    
      if (contacts.some(c => c.email.toLowerCase() === email.toLowerCase()) ||
          email.toLowerCase().includes('zoominfo.com')) {
        continue;
      }
    
      const nameParts = fullName.split(/\s+/);
      const firstName = nameParts[0] || "";
      const lastName = nameParts.slice(1).join(" ") || "";
    
      let title = "";
      let company = "";
      if (titleCompanyLine.includes(', ')) {
        const lastCommaIndex = titleCompanyLine.lastIndexOf(', ');
        if (lastCommaIndex !== -1) {
          title = titleCompanyLine.substring(0, lastCommaIndex).trim();
          company = titleCompanyLine.substring(lastCommaIndex + 2).trim();
        } else {
          title = titleCompanyLine;
        }
      } else {
        title = titleCompanyLine;
      }
    
      // Context for phone numbers is the matched block itself plus some surrounding area
      const contactStartIndex = Math.max(0, match.index - 50); // a bit before the match
      const contactEndIndex = Math.min(cleanedBody.length, match.index + match[0].length + 100); // a bit after the match
      const contactContext = cleanedBody.substring(contactStartIndex, contactEndIndex);
    
      let workPhone = "";
      let personalPhone = "";
    
      const directMatch = contactContext.match(/Direct:?\s*\(?(\d{3})\)?[\s.-]?(\d{3})[\s.-]?(\d{4})/i);
      if (directMatch) {
        workPhone = `${directMatch[1]}-${directMatch[2]}-${directMatch[3]}`;
      }
    
      const personalMatch = contactContext.match(/Personal:?\s*\(?(\d{3})\)?[\s.-]?(\d{3})[\s.-]?(\d{4})/i);
      if (personalMatch) {
        personalPhone = `${personalMatch[1]}-${personalMatch[2]}-${personalMatch[3]}`;
      }
    
      if (firstName && lastName && email) {
        contacts.push({
          firstName: firstName,
          lastName: lastName,
          email: email,
          company: company,
          title: title,
          workPhone: workPhone,
          personalPhone: personalPhone,
          priority: "Medium",
          tags: "",
          industry: ""
        });
        console.log(`Approach 1.5: Extracted ${firstName} ${lastName}, ${email}, "${title}", "${company}"`);
      }
    } catch (error) {
      console.error(`Approach 1.5: Error processing modern ZoomInfo pattern: ${error.message}`);
    }
  }
}

// APPROACH 2: Extract contacts with ** markers (original format)
if (contacts.length < expectedCount || (expectedCount === 0 && contacts.length === 0)) {
  console.log("Trying original format with ** markers (Approach 2)...");
  const contactBlocks = cleanedBody.match(/\*\*([^*]+)\*\*[^@]*?([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g) || [];
  for (const block of contactBlocks) {
    try {
      const nameMatch = block.match(/\*\*([^*]+)\*\*/);
      if (!nameMatch) continue;
      const fullName = nameMatch[1].trim(); // Accents here are fine due to ([^*]+)
    
      // Validate full name with Unicode pattern if needed, though ([^*]+) is broad
      if (!fullNameValidationPattern.test(fullName)) {
          continue;
      }

      const nameParts = fullName.split(/\s+/);
      const firstName = nameParts[0] || "";
      const lastName = nameParts.slice(1).join(" ") || "";
    
      const emailMatch = block.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
      if (!emailMatch) continue;
      const email = emailMatch[1];
      if (contacts.some(c => c.email.toLowerCase() === email.toLowerCase())) { continue; }

      const lines = block.split('\n');
      let title = ""; let company = "";
      const nameLineIndex = lines.findIndex(line => line.includes(`**${fullName}**`));
      if (nameLineIndex >= 0 && nameLineIndex + 1 < lines.length) {
        const titleLine = lines[nameLineIndex + 1].trim();
        if (titleLine && !titleLine.includes('@') && !titleLine.toLowerCase().includes('add')) {
          if (titleLine.includes(', ')) {
            const lastCommaIndex = titleLine.lastIndexOf(", ");
            if (lastCommaIndex !== -1) {
              title = titleLine.substring(0, lastCommaIndex).trim();
              company = titleLine.substring(lastCommaIndex + 2).trim();
            } else {
              title = titleLine;
            }
          }
          else { title = titleLine; }
        }
      }
      let workPhone = ""; let personalPhone = "";
      const directMatch = block.match(/Direct:?\s*\(?(\d{3})\)?[\s.-]?(\d{3})[\s.-]?(\d{4})/i);
      if (directMatch) { workPhone = `${directMatch[1]}-${directMatch[2]}-${directMatch[3]}`; }
      const personalMatch = block.match(/Personal:?\s*\(?(\d{3})\)?[\s.-]?(\d{3})[\s.-]?(\d{4})/i);
      if (personalMatch) { personalPhone = `${personalMatch[1]}-${personalMatch[2]}-${personalMatch[3]}`; }

      if (firstName && lastName && email) {
        contacts.push({
          firstName: firstName, lastName: lastName, email: email, company: company,
          title: title, workPhone: workPhone, personalPhone: personalPhone, priority: "Medium", tags: "", industry: ""
        });
        console.log(`Approach 2: Extracted ${firstName} ${lastName}, ${email}`);
      }
    } catch (error) { console.error(`Approach 2: Error processing ** contact block: ${error.message}`); }
  }
}

// APPROACH 3: Look for plaintext contact blocks without ** markers
if (contacts.length < expectedCount || (expectedCount === 0 && contacts.length === 0)) {
  console.log("Trying improved plaintext contact block approach (Approach 3)...");
   let blocks = cleanedBody.split(/\n\s*\n\s*\n/); // Try triple newlines first
  if (blocks.length <= 1) {
    blocks = cleanedBody.split(/\n\s*\n/); // Fall back to double newlines
  }
   console.log(`Found ${blocks.length} potential contact blocks using plaintext approach`);
   for (const block of blocks) {
    try {
      if (!block.includes('@')) continue;
    
      const emailMatch = block.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
      if (!emailMatch) continue;
      const email = emailMatch[1];
    
      if (contacts.some(c => c.email.toLowerCase() === email.toLowerCase()) ||
          email.toLowerCase().includes('zoominfo.com')) {
        continue;
      }

      const lines = block.split('\n').map(line => line.trim()).filter(line => line.length > 0);
      let fullName = "";
    
      for (let i = 0; i < Math.min(3, lines.length); i++) {
        const line = lines[i];
        const nameMatchResult = line.match(multiPartNameCapturePatternGlobal);
        if (nameMatchResult &&
            !line.includes('St,') && !line.includes('Quebec') && !line.includes('Montreal') &&
            !line.includes('Toronto') && !line.includes('Canada') && !line.includes('Avenue') &&
            !line.includes('@') && !line.includes('Direct:') && !line.includes('Number of') &&
            !line.toLowerCase().includes('zoom information')) {
          fullName = nameMatchResult[1].trim();
          break;
        }
      }
    
      if (!fullName) continue;
    
      const nameParts = fullName.split(/\s+/);
      const firstName = nameParts[0] || "";
      const lastName = nameParts.slice(1).join(" ") || "";
    
      let title = "";
      let company = "";
    
      const nameLineIndex = lines.findIndex(line => line === fullName);
      if (nameLineIndex >= 0 && nameLineIndex + 1 < lines.length) {
        const potentialTitleLine = lines[nameLineIndex + 1].trim();
        if (potentialTitleLine &&
            !potentialTitleLine.includes('@') &&
            !potentialTitleLine.toLowerCase().includes('add') &&
            !potentialTitleLine.toLowerCase().includes('direct:') &&
            !potentialTitleLine.includes('Personal:') &&
            !potentialTitleLine.includes('Home:') &&
            !potentialTitleLine.includes('Work:') &&
            !potentialTitleLine.includes('(') &&
            potentialTitleLine.length > 3 &&
            !/\d{4,}/.test(potentialTitleLine) &&
            !fullNameValidationPattern.test(potentialTitleLine)
            ) {
        
          if (potentialTitleLine.includes(', ')) {
            const lastCommaIndex = potentialTitleLine.lastIndexOf(', ');
            if (lastCommaIndex !== -1) {
              title = potentialTitleLine.substring(0, lastCommaIndex).trim();
              company = potentialTitleLine.substring(lastCommaIndex + 2).trim();
            } else {
              title = potentialTitleLine;
            }
          } else {
            title = potentialTitleLine;
          }
        }
      }
    
      let workPhone = "";
      let personalPhone = "";
      const directMatch = block.match(/Direct:?\s*\(?(\d{3})\)?[\s.-]?(\d{3})[\s.-]?(\d{4})/i);
      if (directMatch) {
        workPhone = `${directMatch[1]}-${directMatch[2]}-${directMatch[3]}`;
      }
      const personalMatch = block.match(/Personal:?\s*\(?(\d{3})\)?[\s.-]?(\d{3})[\s.-]?(\d{4})/i);
      if (personalMatch) {
        personalPhone = `${personalMatch[1]}-${personalMatch[2]}-${personalMatch[3]}`;
      }

      if (firstName && lastName && email) {
        contacts.push({
          firstName: firstName,
          lastName: lastName,
          email: email,
          company: company,
          title: title,
          workPhone: workPhone,
          personalPhone: personalPhone,
          priority: "Medium",
          tags: "",
          industry: ""
        });
        console.log(`Approach 3: Extracted ${firstName} ${lastName}, ${email}, "${title}", "${company}"`);
      }
    } catch (error) {
      console.error(`Approach 3: Error processing plaintext block: ${error.message}`);
    }
  }
}

// APPROACH 4: Last resort - scan line by line for patterns
if (contacts.length < expectedCount || (expectedCount === 0 && contacts.length === 0)) {
  console.log("Trying line-by-line scanning approach (Approach 4)...");
  const lines = cleanedBody.split('\n');
  const processedEmails = contacts.map(c => c.email.toLowerCase());
  for (let i = 0; i < lines.length; i++) {
    try {
      const line = lines[i].trim();
      const emailMatch = line.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
      if (!emailMatch) continue;
      const email = emailMatch[1];
      if (processedEmails.includes(email.toLowerCase()) || email.includes('zoominfo.com') || line.includes('customer service')) { continue; }

      let fullName = "";
      for (let j = i - 1; j >= Math.max(0, i - 3); j--) {
        const prevLine = lines[j].trim();
        if ((lineScanNamePattern.test(prevLine) || multiPartNameCapturePatternGlobal.test(prevLine)) &&
            !prevLine.includes('@') && !prevLine.toLowerCase().includes('direct:') &&
            !prevLine.toLowerCase().includes('personal:') && !prevLine.toLowerCase().includes('add') &&
            !prevLine.toLowerCase().includes('zoom information')) {
          fullName = prevLine;
          break;
        }
      }
      if (!fullName) continue;

      const nameParts = fullName.split(/\s+/);
      const firstName = nameParts[0] || ""; const lastName = nameParts.slice(1).join(" ") || "";
    
      let title = ""; let company = "";
      let titleCompanySearchIndex = -1;
      const nameLineIndex = lines.findIndex(l => l.trim() === fullName);

      if (nameLineIndex !== -1 && nameLineIndex < i -1) {
          titleCompanySearchIndex = nameLineIndex + 1;
      } else if (nameLineIndex !== -1 && nameLineIndex + 1 < i) {
      } else if (nameLineIndex !== -1 && nameLineIndex + 1 < lines.length) {
          titleCompanySearchIndex = nameLineIndex + 1;
      }

      if (titleCompanySearchIndex !== -1 && titleCompanySearchIndex < lines.length) {
          const l = lines[titleCompanySearchIndex].trim();
          if (l && !l.includes('@') && !l.toLowerCase().includes('direct:') &&
              !l.toLowerCase().includes('personal:') && !l.toLowerCase().includes('add') &&
              new RegExp("^\\p{Lu}", "u").test(l) &&
              !fullNameValidationPattern.test(l)
          ) {
            if (l.includes(', ')) {
              const lastCommaIndex = l.lastIndexOf(", ");
              if (lastCommaIndex !== -1) {
                title = l.substring(0, lastCommaIndex).trim();
                company = l.substring(lastCommaIndex + 2).trim();
              } else {
                title = l;
              }
            }
            else { title = l; }
          }
      }
    
      let workPhone = ""; let personalPhone = "";
      for (let j = Math.max(0, i - 3); j < Math.min(lines.length, i + 3); j++) {
        const l = lines[j].trim();
        if (l.includes('Direct:')) { const match = l.match(/Direct:?\s*\(?(\d{3})\)?[\s.-]?(\d{3})[\s.-]?(\d{4})/i); if (match) { workPhone = `${match[1]}-${match[2]}-${match[3]}`; } }
        if (l.includes('Personal:')) { const match = l.match(/Personal:?\s*\(?(\d{3})\)?[\s.-]?(\d{3})[\s.-]?(\d{4})/i); if (match) { personalPhone = `${match[1]}-${match[2]}-${match[3]}`; } }
      }

      if (firstName && lastName && email) {
        contacts.push({
          firstName: firstName, lastName: lastName, email: email, company: company,
          title: title, workPhone: workPhone, personalPhone: personalPhone, priority: "Medium", tags: "", industry: ""
        });
        processedEmails.push(email.toLowerCase());
        console.log(`Approach 4: Extracted ${firstName} ${lastName}, ${email}`);
      }
    } catch (error) { console.error(`Approach 4: Error processing line-by-line scan: ${error.message}`); }
  }
}

// Final validation and company fallback
console.log(`ZoomInfo Extraction Finished: Expected ${expectedCount} contacts, Found ${contacts.length} unique contacts.`);

// If we have contacts but some are missing company info, try to infer from other contacts
if (contacts.length > 0) {
  const companyCounts = {};
  contacts.forEach(c => {
    if (c.company && c.company.length > 0) {
      companyCounts[c.company] = (companyCounts[c.company] || 0) + 1;
    }
  });

  let mostCommonCompany = "";
  let maxCount = 0;
  for (const comp in companyCounts) {
    if (companyCounts[comp] > maxCount) {
      mostCommonCompany = comp;
      maxCount = companyCounts[comp];
    }
  }
   if (mostCommonCompany) {
    let updatedCount = 0;
    for (const contact of contacts) {
      if (!contact.company || contact.company.length === 0) {
        contact.company = mostCommonCompany;
        updatedCount++;
      }
    }
    if (updatedCount > 0) {
      console.log(`Applied most common company ("${mostCommonCompany}") to ${updatedCount} contacts missing company info.`);
    }
  }
}

if (contacts.length === 0 && expectedCount > 0) {
  console.error("Failed to extract any contacts from this email format.");
  console.error("Email format may have changed. Here's a snippet of the email:");
  console.error(cleanedBody.substring(0, 800) + "...");
} else if (contacts.length < expectedCount && expectedCount > 0) {
  console.warn(`Warning: Extracted ${contacts.length} contacts, but email stated ${expectedCount}.`);
  console.warn("Consider reviewing the email body snippet for unparsed contacts:");
  console.warn(cleanedBody.substring(0, 800) + "...");
}

return contacts;
}

// ============================================================
// == ZOOMINFO IMPORT UI ==
// ============================================================

/**
* Builds the card UI with EDITABLE fields for ZoomInfo contacts.
*/
function buildZoomInfoImportCard(message) {
const card = CardService.newCardBuilder();
card.setHeader(CardService.newCardHeader()
  .setTitle("ZoomInfo Contacts Found")
  .setImageUrl("https://d271tldp4c6tdf.cloudfront.net/assets/logos/zoominfo-logo-color-email.png"));

const contacts = extractZoomInfoContacts(message);

if (!contacts || contacts.length === 0) {
const noContactsSection = CardService.newCardSection();
noContactsSection.addWidget(CardService.newTextParagraph()
    .setText("Could not automatically extract contacts from this ZoomInfo email. Please check the email format or add contacts manually."));
card.addSection(noContactsSection);
} else {
// NEW: Sequence selection at TOP (required)
const sequenceSection = CardService.newCardSection()
    .setHeader("Select Sequence for ALL Contacts (Required)");
const sequenceDropdown = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName("selectedSequence");
const availableSequences = getAvailableSequences();
if (availableSequences.length === 0) {
  sequenceDropdown.addItem("No sequences available - Create one first", "", true);
} else {
  let defaultSet = false;
  for (const sequence of availableSequences) {
      const isDefault = !defaultSet && (sequence === "SaaS / B2B Tech" || sequence === availableSequences[0]);
      sequenceDropdown.addItem(sequence, sequence, isDefault);
      if (isDefault) defaultSet = true;
  }
}
sequenceSection.addWidget(sequenceDropdown);
card.addSection(sequenceSection);

// Section for setting default priority
const prioritySection = CardService.newCardSection()
    .setHeader("Set Default Priority for All Contacts");
const priorityDropdown = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setTitle("Priority")
    .setFieldName("defaultPriority");
priorityDropdown.addItem("High 🔥", "High", false);
priorityDropdown.addItem("Medium 🟠", "Medium", true);
priorityDropdown.addItem("Low ⚪", "Low", false);
prioritySection.addWidget(priorityDropdown);
card.addSection(prioritySection);

// Section for bulk import
const bulkSection = CardService.newCardSection().setHeader("Bulk Import");
const recordsMatch = message.getPlainBody().match(/Number of records exported:?\s*(\d+)/i);
const expectedCountInfo = recordsMatch ? ` (Email stated: ${recordsMatch[1]})` : "";
bulkSection.addWidget(CardService.newTextParagraph()
    .setText(`Found ${contacts.length} unique contacts.${expectedCountInfo}`));
const bulkImportAction = CardService.newAction()
    .setFunctionName("importAllZoomInfoContacts")
    .setParameters({
         messageId: message.getId(),
         contactCount: contacts.length.toString()
        });
bulkSection.addWidget(CardService.newButtonSet()
    .addButton(CardService.newTextButton()
        .setText(`Import ${contacts.length} Contacts to [Selected Sequence]`)
        .setOnClickAction(bulkImportAction)
        .setDisabled(contacts.length === 0)));
card.addSection(bulkSection);

// Individual contact sections with EDITABLE fields
for (let i = 0; i < contacts.length; i++) {
  const contact = contacts[i];
  const sectionId = `contact_${i}_`;

  const contactSection = CardService.newCardSection()
      .setHeader(`Contact ${i + 1}: Edit Details`);

  // Helper to add TextInput widgets
  const addTextInput = (label, fieldSuffix, initialValue, hint) => {
    contactSection.addWidget(CardService.newTextInput()
        .setTitle(label)
        .setFieldName(sectionId + fieldSuffix)
        .setValue(initialValue || "")
        .setHint(hint || ""));
  };

  addTextInput("First Name", "firstName", contact.firstName);
  addTextInput("Last Name", "lastName", contact.lastName);
  addTextInput("Email", "email", contact.email);
  addTextInput("Title", "title", contact.title);
  addTextInput("Company", "company", contact.company);
  addTextInput("Work Phone", "workPhone", contact.workPhone);
  addTextInput("Personal Phone", "personalPhone", contact.personalPhone);
  addTextInput("Tags", "tags", contact.tags, "Comma-separated (e.g., lead, hot)");

  // Industry Dropdown
  const industryDropdown = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.DROPDOWN)
      .setFieldName(sectionId + "industry");

   industryDropdown.addItem("(No Industry)", "", contact.industry === "" || !contact.industry);
   for (const industry of ZOOMINFO_INDUSTRY_OPTIONS) {
    industryDropdown.addItem(industry, industry, contact.industry === industry);
  }
  contactSection.addWidget(industryDropdown);

  // Individual import button
  const singleImportAction = CardService.newAction()
      .setFunctionName("importSingleZoomInfoContact")
      .setParameters({
          contactIndex: i.toString(),
          messageId: message.getId()
      });

  contactSection.addWidget(CardService.newButtonSet()
      .addButton(CardService.newTextButton()
          .setText("Import & Add to Step 1")
          .setOnClickAction(singleImportAction)));

  card.addSection(contactSection);
}
}

return card.build();
}

// ============================================================
// == ZOOMINFO IMPORT HANDLERS ==
// ============================================================

/**
* Action handler for bulk import button.
*/
function importAllZoomInfoContacts(e) {
try {
const messageId = e.parameters.messageId;
const contactCount = parseInt(e.parameters.contactCount || "0");

if (!messageId) {
  console.error("Import All Error: messageId parameter missing.");
  return createNotification("Error: Could not identify the email.");
}
if (contactCount === 0) {
   return createNotification("No contacts were initially found to import.");
}

// Get the selected sequence (REQUIRED)
const selectedSequence = e.formInputs && e.formInputs.selectedSequence && e.formInputs.selectedSequence.length > 0
                         ? e.formInputs.selectedSequence[0] : "SaaS / B2B Tech";
// Get the selected default priority
const defaultPriority = e.formInputs && e.formInputs.defaultPriority && e.formInputs.defaultPriority.length > 0
                        ? e.formInputs.defaultPriority[0] : "Medium";
console.log(`Bulk importing ${contactCount} contacts to sequence: ${selectedSequence} with default priority: ${defaultPriority}`);

let successCount = 0;
let failureCount = 0;
const errors = [];

// Iterate based on the number of contacts displayed in the card
for (let i = 0; i < contactCount; i++) {
  const sectionId = `contact_${i}_`;

  // Construct the contact object from FORM INPUTS
  const contact = {
    firstName: e.formInputs?.[sectionId + 'firstName']?.[0] || "",
    lastName: e.formInputs?.[sectionId + 'lastName']?.[0] || "",
    email: e.formInputs?.[sectionId + 'email']?.[0] || "",
    company: e.formInputs?.[sectionId + 'company']?.[0] || "",
    title: e.formInputs?.[sectionId + 'title']?.[0] || "",
    workPhone: e.formInputs?.[sectionId + 'workPhone']?.[0] || "",
    personalPhone: e.formInputs?.[sectionId + 'personalPhone']?.[0] || "",
    tags: e.formInputs?.[sectionId + 'tags']?.[0] || "",
    sequence: selectedSequence,
    industry: e.formInputs?.[sectionId + 'industry']?.[0] || "",
    priority: defaultPriority
  };

   if (!contact.firstName || !contact.lastName || !contact.email) {
      console.warn(`Skipping contact index ${i} due to missing essential fields.`);
      failureCount++;
      errors.push(`Contact ${i+1}: Missing essential info after edit.`);
      continue;
   }

  try {
    const result = addContactToDatabase(contact);
    if (result.success) {
      successCount++;
    } else {
      failureCount++;
      errors.push(`Contact ${i+1} (${contact.email || 'N/A'}): ${result.message}`);
      console.warn(`Failed to import contact index ${i}: ${result.message}`);
    }
  } catch (error) {
    console.error(`Critical error importing contact index ${i}: ${error}`);
    failureCount++;
    errors.push(`Contact ${i+1} (${contact.email || 'N/A'}): ${error.message}`);
  }
}

let notificationText = `Import finished. Success: ${successCount}. Failures: ${failureCount}.`;
if (failureCount > 0) {
    notificationText += ` Errors: ${errors.slice(0, 3).join('; ')}` + (errors.length > 3 ? '...' : '');
}

return CardService.newActionResponseBuilder()
   .setNotification(CardService.newNotification().setText(notificationText))
   .build();

} catch (error) {
console.error("Error in importAllZoomInfoContacts: " + error);
 const notification = CardService.newNotification().setText("Critical error during bulk import: " + error.message);
 return CardService.newActionResponseBuilder().setNotification(notification).build();
}
}

/**
* Action handler for single contact import button.
*/
function importSingleZoomInfoContact(e) {
try {
const contactIndex = e.parameters.contactIndex;
if (contactIndex === undefined || contactIndex === null) {
    console.error("Import Single Error: contactIndex parameter missing.");
    return createNotification("Error: Could not identify which contact to import.");
}

const index = parseInt(contactIndex);
const sectionId = `contact_${index}_`;

const defaultPriority = e.formInputs && e.formInputs.defaultPriority && e.formInputs.defaultPriority.length > 0
                        ? e.formInputs.defaultPriority[0] : "Medium";

 // Get the selected sequence (REQUIRED)
 const selectedSequence = e.formInputs && e.formInputs.selectedSequence && e.formInputs.selectedSequence.length > 0
                          ? e.formInputs.selectedSequence[0] : "SaaS / B2B Tech";

 const contact = {
    firstName: e.formInputs?.[sectionId + 'firstName']?.[0] || "",
    lastName: e.formInputs?.[sectionId + 'lastName']?.[0] || "",
    email: e.formInputs?.[sectionId + 'email']?.[0] || "",
    company: e.formInputs?.[sectionId + 'company']?.[0] || "",
    title: e.formInputs?.[sectionId + 'title']?.[0] || "",
    workPhone: e.formInputs?.[sectionId + 'workPhone']?.[0] || "",
    personalPhone: e.formInputs?.[sectionId + 'personalPhone']?.[0] || "",
    tags: e.formInputs?.[sectionId + 'tags']?.[0] || "",
    sequence: selectedSequence,
    industry: e.formInputs?.[sectionId + 'industry']?.[0] || "",
    priority: defaultPriority
 };

console.log(`Importing single contact index ${index} (${contact.email}) with sequence: ${contact.sequence}`);

 if (!contact.firstName || !contact.lastName || !contact.email) {
    return createNotification(`Cannot import contact: Missing First Name, Last Name, or Email after edits.`);
 }

const result = addContactToDatabase(contact);

if (result.success) {
   return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(`Imported ${contact.firstName} ${contact.lastName} successfully!`))
    .build();
} else {
  console.warn(`Failed to import single contact: ${result.message}`);
  return createNotification(result.message || "Error importing contact.");
}

} catch (error) {
console.error("Error in importSingleZoomInfoContact: " + error);
const notification = CardService.newNotification().setText("Error importing contact: " + error.message);
return CardService.newActionResponseBuilder().setNotification(notification).build();
}
}

// ============================================================
// == DATABASE HELPER FOR ZOOMINFO ==
// ============================================================

/**
* Helper function to add a contact to the database (Google Sheet).
* Receives the (potentially edited) contact object from ZoomInfo import.
*/
function addContactToDatabase(contact) {
if (!CONFIG || !CONFIG.CONTACTS_SHEET_NAME) {
  console.error("AddContactToDB Error: CONFIG.CONTACTS_SHEET_NAME is not defined.");
  return { success: false, message: "Internal Config Error: Sheet name missing." };
}
if (!CONTACT_COLS || CONTACT_COLS.EMAIL === undefined || CONTACT_COLS.TAGS === undefined || CONTACT_COLS.CONNECT_SALES_LINK === undefined) {
  console.error("AddContactToDB Error: CONTACT_COLS mapping is incomplete.");
  return { success: false, message: "Internal Config Error: Column mapping missing." };
}

const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
if (!spreadsheetId) {
console.error("AddContactToDB Error: SPREADSHEET_ID not set in UserProperties.");
return { success: false, message: "Database (Spreadsheet ID) not configured in Add-on properties." };
}

try {
const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);

if (!contactsSheet) {
  console.error(`AddContactToDB Error: Sheet named "${CONFIG.CONTACTS_SHEET_NAME}" not found.`);
  return { success: false, message: `Contacts sheet "${CONFIG.CONTACTS_SHEET_NAME}" not found.` };
}

if (!contact.firstName || !contact.lastName || !contact.email) {
  return { success: false, message: "Cannot add contact: Missing First Name, Last Name, or Email." };
}
 if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(contact.email)) {
     return { success: false, message: `Invalid email format: ${contact.email}` };
 }

const dataRange = contactsSheet.getDataRange();
const data = dataRange.getValues();
const emailColIndex = CONTACT_COLS.EMAIL;

 if (emailColIndex === undefined || emailColIndex < 0 || (data.length > 0 && emailColIndex >= data[0].length)) {
    console.error(`AddContactToDB Error: EMAIL column index invalid.`);
    return { success: false, message: "Internal Config Error: Email column index mismatch." };
}

const lowerCaseEmail = contact.email.toLowerCase();
for (let i = 1; i < data.length; i++) {
  if (data[i][emailColIndex] && data[i][emailColIndex].toString().toLowerCase() === lowerCaseEmail) {
    return { success: false, message: `Contact with email ${contact.email} already exists.` };
  }
}

const formattedWorkPhone = contact.workPhone ? "'" + contact.workPhone.toString().trim() : "";
const formattedPersonalPhone = contact.personalPhone ? "'" + contact.personalPhone.toString().trim() : "";

const numberOfColumns = Math.max(...Object.values(CONTACT_COLS)) + 1;
const newRow = Array(numberOfColumns).fill("");

const safeAssign = (colIndex, value) => {
    if (colIndex !== undefined && colIndex >= 0 && colIndex < newRow.length) {
        newRow[colIndex] = value !== undefined && value !== null ? value : "";
    } else {
        console.warn(`Warning: Column index ${colIndex} is invalid.`);
    }
};

safeAssign(CONTACT_COLS.FIRST_NAME, contact.firstName);
safeAssign(CONTACT_COLS.LAST_NAME, contact.lastName);
safeAssign(CONTACT_COLS.EMAIL, contact.email);
safeAssign(CONTACT_COLS.COMPANY, contact.company);
safeAssign(CONTACT_COLS.TITLE, contact.title);
safeAssign(CONTACT_COLS.CURRENT_STEP, 1);
safeAssign(CONTACT_COLS.LAST_EMAIL_DATE, "");
safeAssign(CONTACT_COLS.NEXT_STEP_DATE, "");
safeAssign(CONTACT_COLS.STATUS, "Active");
safeAssign(CONTACT_COLS.NOTES, "");
safeAssign(CONTACT_COLS.PERSONAL_PHONE, formattedPersonalPhone);
safeAssign(CONTACT_COLS.WORK_PHONE, formattedWorkPhone);
safeAssign(CONTACT_COLS.PERSONAL_CALLED, "No");
safeAssign(CONTACT_COLS.WORK_CALLED, "No");
safeAssign(CONTACT_COLS.PRIORITY, contact.priority || "Medium");
safeAssign(CONTACT_COLS.TAGS, contact.tags || "");
safeAssign(CONTACT_COLS.SEQUENCE, contact.sequence || "SaaS / B2B Tech Outreach");
safeAssign(CONTACT_COLS.INDUSTRY, contact.industry || "");

// Generate Connect Sales Link
let connectSalesLinkFormula = "";
if (contact.email) {
  const emailParts = contact.email.split('@');
  if (emailParts.length === 2) {
    const domain = emailParts[1];
    if (domain) {
      const connectSalesUrl = `https://sales.connect.corp.google.com/search/LEAD/TEXT/type%3ALEAD+${domain}`;
      connectSalesLinkFormula = `=HYPERLINK("${connectSalesUrl}", "${domain}")`;
    }
  }
}
safeAssign(CONTACT_COLS.CONNECT_SALES_LINK, connectSalesLinkFormula);

contactsSheet.appendRow(newRow);

// Add to cache for fast future lookups
if (typeof addContactToCache === "function") {
    addContactToCache(contact);
}

if (typeof logAction === "function") {
    logAction("ZoomInfo Import", `Imported contact: ${contact.firstName} ${contact.lastName} (${contact.email})`);
} else {
    console.log(`Log Action (stub): ZoomInfo Import - ${contact.email}`);
}

return { success: true };
} catch (error) {
console.error("Error adding contact to database: " + error.toString());
return { success: false, message: "Error adding contact: " + error.message };
}
}

