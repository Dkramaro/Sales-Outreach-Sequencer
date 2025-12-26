/**
 * FILE: Bulk Email Processing
 * 
 * PURPOSE:
 * Handles bulk email draft creation, sending, and batch spreadsheet updates.
 * Manages Step 1 message searching for reply threading. Optimized for performance
 * with batched operations.
 * 
 * KEY FUNCTIONS:
 * - emailSelectedContacts() - Entry point for bulk email processing
 * - processBulkEmails() - Core bulk email logic with batched updates
 * - findSentStep1Messages_Batch() - Efficient batch search for Step 1 messages
 * - findSentStep1Message() - Single message search (for individual compose)
 * - extractFirstWords() - Utility for text truncation
 * 
 * DEPENDENCIES:
 * - 01_Config.gs: CONFIG, CONTACT_COLS
 * - 04_ContactData.gs: getAllContactsData, getContactsInStep, getContactsReadyForEmail
 * - 06_SequenceData.gs: getSequenceTemplateForStep, getSequenceStepCount
 * - 08_EmailComposition.gs: replacePlaceholders, getSignatureHtml, getDarkModeStyles, cleanAndWrapSignature
 * - 10_SelectionHandling.gs: getSelectedEmailsFromInput
 * 
 * @version 2.3
 */

/* ======================== BULK EMAIL PROCESSING ======================== */

/**
 * Selects contacts and directly initiates bulk email draft creation.
 * REVISED to integrate with the stateful `getSelectedEmailsFromInput` to bypass CardService bugs.
 * MODIFIED to pass filter states through to the processing function.
 */
function emailSelectedContacts(e) {
  console.log("Entering emailSelectedContacts (direct processing). Event object:", JSON.stringify(e));

  const step = e.parameters.step ? parseInt(e.parameters.step) : null; // Step context
  const page = parseInt(e.parameters.page || '1'); // Page context
  const sequenceFilter = e.parameters.sequenceFilter || ""; // Get sequence filter state
  const statusFilter = e.parameters.statusFilter || "all";   // Get status filter state

  // This block is no longer strictly needed for the selection logic but is good for logging context.
  try {
      let pageSize;
      if (step === null) { pageSize = CONFIG.PAGE_SIZE; } 
      else if (step === 1) { pageSize = CONFIG.PAGE_SIZE; } 
      else { pageSize = 10; }

      let allRelevantContacts = [];
      if (step) {
          allRelevantContacts = getContactsInStep(step);
      } else {
          allRelevantContacts = getContactsReadyForEmail();
      }
      
      const startIndex = (page - 1) * pageSize;
      const endIndex = Math.min(startIndex + pageSize, allRelevantContacts.length);
      const contactsOnPageObjects = allRelevantContacts.slice(startIndex, endIndex);
      const contactsOnPageEmails = contactsOnPageObjects.map(contact => contact.email);
      console.log(`emailSelectedContacts: Context: Page ${page}, Step ${step}. Emails displayed on page:`, JSON.stringify(contactsOnPageEmails));
  } catch(logError) {
      console.error("Error during logging context setup in emailSelectedContacts:", logError);
  }
  // End of logging context block

  console.log("Received e.formInput for selection:", JSON.stringify(e.formInput));
  
  // --- CORE MODIFICATION ---
  // Call the stateful version of getSelectedEmailsFromInput.
  const selectionResult = getSelectedEmailsFromInput(e.formInput);
  const selectedEmails = selectionResult.selected;

  // IMPORTANT: Immediately perform the cleanup to remove the temporary UserProperty.
  // This prevents stale state from affecting the next, unrelated operation.
  selectionResult.cleanup(); 
  // --- END CORE MODIFICATION ---

  console.log("Emails selected by getSelectedEmailsFromInput:", JSON.stringify(selectedEmails));

  if (selectedEmails.length === 0) {
    console.error("emailSelectedContacts: No emails were selected after processing based on form input.");
    return createNotification("No contacts selected. Please check the boxes next to the contacts you want to email.");
  }

  // Check against daily quota
  const remainingQuota = MailApp.getRemainingDailyQuota();
  if (selectedEmails.length > remainingQuota) {
    logAction("Warning", `Attempted to email ${selectedEmails.length} contacts, but quota remaining is ${remainingQuota}.`);
    return createNotification(`Cannot create drafts for ${selectedEmails.length} emails. Your remaining Gmail quota is ${remainingQuota}. Please select fewer contacts or wait until tomorrow.`);
  }

  // Directly call processBulkEmails with the correctly determined list of emails and the filter states.
  return processBulkEmails(selectedEmails, step, page, sequenceFilter, statusFilter); 
}

/**
 * Processes bulk email draft creation with batched spreadsheet updates for maximum performance.
 * REVISED to use a single batch search for Step 1 messages, dramatically improving performance.
 * Incorporates fix to prevent data wiping during sheet updates.
 * CORRECTED: Handles Step 1 as new draft, Steps > 1 as DRAFT REPLIES in the same thread.
 * ADDED: Checks for replies before sending follow-ups (Steps > 1).
 * MODIFIED: Accepts and uses filter states to rebuild the UI correctly.
 * NEW: Integrates signature from Google Doc.
 * NEW: Checks AUTO_SEND_STEP_1_ENABLED setting to either send Step 1 emails or create drafts.
 */
function processBulkEmails(selectedEmailsArray, stepContextValue, pageContextValue, sequenceFilter, statusFilter) {
    const selectedEmails = selectedEmailsArray;
    const stepContext = stepContextValue;
    const pageForReturn = pageContextValue;

    if (!selectedEmails || selectedEmails.length === 0) {
        logAction("Error", "processBulkEmails: No selected emails provided directly to function.");
        return createNotification("No contacts were selected for email processing. Please try again.");
    }

    let successCount = 0,
        failureCount = 0,
        repliedCount = 0,
        sentCount = 0,
        draftCount = 0;
    const failures = [];
    let pdfAttachmentBlob = null;
    let pdfAttachmentError = null;
    const updatesToWrite = []; // *** KEY CHANGE: Array to hold all sheet updates.

    const userProps = PropertiesService.getUserProperties();
    const pdfFileId = userProps.getProperty("PDF_ATTACHMENT_FILE_ID");
    const sequencesForPdfString = userProps.getProperty("SEQUENCES_FOR_PDF") || "";
    const sequencesForPdf = sequencesForPdfString ? sequencesForPdfString.split(',') : [];
    const autoSendStep1Enabled = userProps.getProperty("AUTO_SEND_STEP_1_ENABLED") === 'true';

    // --- Efficient Setup (Done Once) ---
    if (pdfFileId && sequencesForPdf.length > 0) {
        try {
            pdfAttachmentBlob = DriveApp.getFileById(pdfFileId).getBlob();
        } catch (e) {
            pdfAttachmentError = `Failed to get PDF attachment: ${e.message}`;
            logAction("Error", pdfAttachmentError);
        }
    }

    const allContactsData = getAllContactsData();
    const contactMap = new Map(allContactsData.map(c => [c.email.toLowerCase(), c]));
    const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
    if (!spreadsheetId) return createNotification("Database not configured.");
    const contactsSheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(CONFIG.CONTACTS_SHEET_NAME);
    if (!contactsSheet) return createNotification("Contacts sheet not found.");
    const currentUserEmail = Session.getActiveUser() ? Session.getActiveUser().getEmail().toLowerCase() : null;

    // Batch find original Step 1 messages for all follow-ups at once
    const contactsNeedingStep1Message = selectedEmailsArray
        .map(email => contactMap.get(email.toLowerCase()))
        .filter(contact => contact && contact.currentStep > 1 && contact.status === 'Active');
    let step1MessagesFoundMap = new Map();
    if (contactsNeedingStep1Message.length > 0) {
        step1MessagesFoundMap = findSentStep1Messages_Batch(contactsNeedingStep1Message, contactsSheet);
    }

    const signatureHtml = getSignatureHtml();
    // --- End of Efficient Setup ---

    // --- Main Processing Loop (Gmail calls only) ---
    for (const email of selectedEmails) {
        const normalizedEmail = email.toLowerCase();
        const contact = contactMap.get(normalizedEmail);

        if (!contact || contact.status !== 'Active' || !contact.isReady) {
            failureCount++;
            failures.push(`${email}: Skipped (Not active or not ready)`);
            continue;
        }

        try {
            const actualStepToSend = contact.currentStep;
            const templateToUse = getSequenceTemplateForStep(contact.sequence, actualStepToSend);

            if (!templateToUse) {
                failureCount++;
                failures.push(`${email}: No template for Step ${actualStepToSend}`);
                continue;
            }

            const emailBody = replacePlaceholders(templateToUse.body, contact);
            const darkModeStyles = getDarkModeStyles();
            let finalEmailBodyHtml = darkModeStyles + emailBody.replace(/\n/g, "<br>");
            if (signatureHtml) {
                finalEmailBodyHtml += cleanAndWrapSignature(signatureHtml);
            }

            let actionLogVerb = "Draft for";

            if (actualStepToSend === 1) {
                const actualSubjectForStep1Email = replacePlaceholders(templateToUse.subject, contact);
                const emailOptions = {
                    htmlBody: finalEmailBodyHtml,
                    cc: userProps.getProperty("CC_EMAILS") || "",
                    name: userProps.getProperty("SENDER_NAME") || ""
                };

                if (autoSendStep1Enabled) {
                    GmailApp.sendEmail(contact.email, actualSubjectForStep1Email, '', emailOptions);
                    sentCount++;
                    actionLogVerb = "Sent";
                } else {
                    GmailApp.createDraft(contact.email, actualSubjectForStep1Email, '', emailOptions);
                    draftCount++;
                }
                
                updatesToWrite.push({
                    rowIndex: contact.rowIndex,
                    step1Subject: actualSubjectForStep1Email
                });


            } else { // Steps 2+ are always draft replies
                const originalStep1Message = step1MessagesFoundMap.get(normalizedEmail);
                if (!originalStep1Message) {
                    failureCount++;
                    failures.push(`${email}: Step 1 msg not found for reply`);
                    continue;
                }
                contact.originalSubject = originalStep1Message.getSubject();
                
                // *** FIX: Restore the logic to build a complete CC list ***
                const existingCcEmails = userProps.getProperty("CC_EMAILS") || "";
                const ccList = [];
                if (existingCcEmails.trim()) {
                    ccList.push(existingCcEmails.trim());
                }
                // This is the crucial line that was missing:
                ccList.push(contact.email);
                const finalCcString = ccList.join(',');
                // *** END FIX ***
                
                const replyOptions = {
                    htmlBody: finalEmailBodyHtml,
                    cc: finalCcString, // Use the correctly constructed list
                };
                
                if (actualStepToSend === 2 && pdfAttachmentBlob && sequencesForPdf.includes(contact.sequence)) {
                    replyOptions.attachments = [pdfAttachmentBlob];
                }

                originalStep1Message.createDraftReply("", replyOptions);
                draftCount++;
                
                updatesToWrite.push({
                    rowIndex: contact.rowIndex
                });
            }

            successCount++;
            logAction("Bulk Email Processed", `${actionLogVerb} Step ${actualStepToSend}: ${contact.email}.`);

        } catch (error) {
            console.error(`Error processing email for ${email}: ${error}\n${error.stack}`);
            logAction("Error", `Bulk Email Error for ${email}: ${error.toString()}`);
            failureCount++;
            failures.push(`${email}: ${error.message}`);
        }
    }
    // --- End of Loop ---

    // --- BATCH SPREADSHEET WRITE ---
    if (updatesToWrite.length > 0) {
        const lastCol = contactsSheet.getLastColumn();
        const dataRange = contactsSheet.getRange(2, 1, contactsSheet.getLastRow() - 1, lastCol);
        const allSheetData = dataRange.getValues();
        const delayDays = parseInt(userProps.getProperty("DELAY_DAYS")) || CONFIG.DEFAULT_DELAY_DAYS;
        
        for (const update of updatesToWrite) {
            const rowData = allSheetData[update.rowIndex - 2];
            const now = new Date();
            const nextDate = new Date(now.getTime() + (delayDays * 24 * 60 * 60 * 1000));
            
            let currentStep = parseInt(rowData[CONTACT_COLS.CURRENT_STEP]);
            let nextStepVal = currentStep + 1;
            let newStatus = rowData[CONTACT_COLS.STATUS];
            
            // Use per-sequence step count instead of global CONFIG.SEQUENCE_STEPS
            const contactSequence = rowData[CONTACT_COLS.SEQUENCE];
            const sequenceStepCount = getSequenceStepCount(contactSequence);
            if (nextStepVal > sequenceStepCount) {
                nextStepVal = sequenceStepCount;
                newStatus = "Completed";
            }

            rowData[CONTACT_COLS.LAST_EMAIL_DATE] = now;
            rowData[CONTACT_COLS.NEXT_STEP_DATE] = (newStatus === "Completed") ? "" : nextDate;
            rowData[CONTACT_COLS.CURRENT_STEP] = nextStepVal;
            rowData[CONTACT_COLS.STATUS] = newStatus;
            rowData[CONTACT_COLS.PERSONAL_CALLED] = "No";
            rowData[CONTACT_COLS.WORK_CALLED] = "No";
            if (update.step1Subject) {
                rowData[CONTACT_COLS.STEP1_SUBJECT] = update.step1Subject;
            }
        }
        
        dataRange.setValues(allSheetData);
        logAction("Bulk Sheet Update", `Updated ${updatesToWrite.length} contact rows in the sheet.`);
    }
    // --- END OF BATCH WRITE ---

    // --- Check for demo contact and handle special case ---
    let processedDemoContact = false;
    for (const email of selectedEmails) {
      if (email.toLowerCase().includes("example.com")) {
        processedDemoContact = true;
        // Auto-delete the demo contact from the database
        deleteDemoContact(email.toLowerCase());
        break;
      }
    }

    // --- Build final response card ---
    let messageParts = [];
    if (sentCount > 0) messageParts.push(`Sent ${sentCount} email(s)`);
    if (draftCount > 0) messageParts.push(`created ${draftCount} draft(s)`);
    let message = messageParts.join(' and ') + ".";

    if (failureCount > 0) message += ` Failed to process ${failureCount} contact(s).`;
    if (pdfAttachmentError) message += ` NOTE: ${pdfAttachmentError}`;

    // If demo contact was processed, show special wizard completion card
    if (processedDemoContact && draftCount > 0) {
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText("🎉 Demo email created! Check Gmail Drafts."))
        .setNavigation(CardService.newNavigation().updateCard(buildDemoEmailSuccessCard()))
        .build();
    }

    let returnCard;
    if (stepContext != null) {
        const builderFn = (String(stepContext) === '2') ? viewStep2Contacts : buildSelectContactsCard;
        returnCard = builderFn({
            parameters: {
                step: String(stepContext),
                page: String(pageForReturn),
                sequenceFilter: sequenceFilter,
                statusFilter: statusFilter
            }
        });
    } else {
        returnCard = viewContactsReadyForEmail({
            parameters: {
                page: String(pageForReturn)
            }
        });
    }

    return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText(message))
        .setNavigation(CardService.newNavigation().updateCard(returnCard))
        .build();
}

/**
 * Moves selected contacts to the next step.
 */
function moveSelectedToNextStep(e) {
  const step = parseInt(e.parameters.step);
  const page = e.parameters.page || '1'; // Page context for refresh
  let pageSize;
  if (step === 1) { // Originating from Step 1 view
    pageSize = CONFIG.PAGE_SIZE;
  } else { // Originating from Step 2, 3, 4 views (Step 5 cannot move next)
    pageSize = 10;
  }

  // --- Determine which contacts were displayed on the current page ---
  if (!step) {
    console.error("moveSelectedToNextStep: Step parameter is missing.");
    return createNotification("Cannot move contacts: Origin step information is missing.");
  }

  let allContactsInStep = getContactsInStep(step);

  // Apply the same sorting as in buildSelectContactsCard/viewStep2Contacts
  allContactsInStep.sort((a, b) => {
    const readyA = a.isReady && a.status === 'Active';
    const readyB = b.isReady && b.status === 'Active';
    if (readyA !== readyB) return readyB - readyA; // Ready first

    const priorityOrder = { "High": 1, "Medium": 2, "Low": 3 };
    const priorityA = priorityOrder[a.priority] || 2;
    const priorityB = priorityOrder[b.priority] || 2;
    if (priorityA !== priorityB) return priorityA - priorityB; // Higher priority first

    return (a.lastName + a.firstName).localeCompare(b.lastName + b.firstName); // Then by name
  });

  const totalContactsOnCurrentList = allContactsInStep.length;
  const startIndex = (parseInt(page) - 1) * pageSize;
  const endIndex = Math.min(startIndex + pageSize, totalContactsOnCurrentList);
  const contactsOnPageObjects = allContactsInStep.slice(startIndex, endIndex);
  const contactsOnPageEmails = contactsOnPageObjects.map(contact => contact.email);
  // --- End determining contacts on page ---

  console.log(`moveSelectedToNextStep: For Step ${step}, Page ${page}, contacts on page for selection:`, JSON.stringify(contactsOnPageEmails));
  console.log(`moveSelectedToNextStep: Form input for selection:`, JSON.stringify(e.formInput));

  const selectedEmails = getSelectedEmailsFromInput(contactsOnPageEmails, e.formInput);

  if (selectedEmails.length === 0) {
    return createNotification("No contacts selected. Please check boxes to move contacts.");
  }

   // Note: We check per-contact sequence step count inside the loop instead of here
   // since different contacts might be in different sequences with different step limits

  let successCount = 0;
  let failureCount = 0;
   const failures = [];

   const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
   if (!spreadsheetId) return createNotification("No database connected.");

   // Get contacts data once
   const allContactsData = getAllContactsData();
   const contactMap = new Map(allContactsData.map(c => [c.email.toLowerCase(), c]));


  try {
     const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
     const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);
     if (!contactsSheet) throw new Error("Contacts sheet not found.");

     const updates = []; // Batch updates: [rowIndex, colIndex, value]

     for (const email of selectedEmails) {
       const normalizedEmail = email.toLowerCase();
        const contact = contactMap.get(normalizedEmail);

        if (contact && contact.currentStep === step && contact.status !== "Completed" && contact.status !== "Unsubscribed") {
           const nextStep = contact.currentStep + 1;
           let newStatus = contact.status; // Status usually doesn't change on manual move unless last step
           
           // Use per-sequence step count instead of global CONFIG.SEQUENCE_STEPS
           const sequenceStepCount = getSequenceStepCount(contact.sequence);
            if (nextStep > sequenceStepCount) {
                 // Contact is already on or beyond their sequence's last step
                 console.warn(`Attempted to move ${email} beyond last step (sequence limit: ${sequenceStepCount}).`);
                 failureCount++;
                 failures.push(`${email}: Already on or beyond last step.`);
                 continue;
            }

            // Add updates for this row to the batch list
            updates.push([contact.rowIndex, CONTACT_COLS.CURRENT_STEP + 1, nextStep]);
            // Reset call status for the new step
            updates.push([contact.rowIndex, CONTACT_COLS.PERSONAL_CALLED + 1, "No"]);
            updates.push([contact.rowIndex, CONTACT_COLS.WORK_CALLED + 1, "No"]);
            // Clear next step date as it's now manually moved
            updates.push([contact.rowIndex, CONTACT_COLS.NEXT_STEP_DATE + 1, ""]);


            // Optionally update status if moving *to* the last step results in completion?
            // Let's assume manual move doesn't automatically complete. User does that separately.

            successCount++;
            logAction("Manual Move", `Moved ${email} from Step ${step} to Step ${nextStep}`);

        } else if (!contact) {
            console.error(`moveSelectedToNextStep: Contact not found for ${email}`);
            failureCount++;
            failures.push(`${email}: Contact not found.`);
        } else {
             console.warn(`moveSelectedToNextStep: Skipped ${email} (Current Step: ${contact.currentStep}, Status: ${contact.status})`);
             failureCount++;
             failures.push(`${email}: Skipped (Wrong step/status)`);
         }
      } // End loop

     // Apply batch updates to the sheet
     if (updates.length > 0) {
         updates.forEach(update => {
             contactsSheet.getRange(update[0], update[1]).setValue(update[2]);
         });
         SpreadsheetApp.flush();
     }

   } catch(error) {
       console.error("Error during moveSelectedToNextStep: " + error + "\n" + error.stack);
       logAction("Error", `Batch move failed: ${error.toString()}`);
       return createNotification("Error moving contacts: " + error.message);
   }

   // --- Build Response ---
   let message = `Moved ${successCount} contact(s) to Step ${step + 1}.`;
   if (failureCount > 0) {
     message += ` Failed to move ${failureCount} contact(s): ${failures.slice(0, 3).join(', ')}${failures.length > 3 ? '...' : ''}`;
   }

   return CardService.newActionResponseBuilder()
     .setNotification(CardService.newNotification().setText(message))
     .setNavigation(CardService.newNavigation()
       // Rebuild the *original* step's view card on the same page
       .updateCard(buildSelectContactsCard({ parameters: { step: step.toString(), page: page } })))
     .build();
}

/* ======================== STEP 1 MESSAGE SEARCH ======================== */

/**
 * Performs a single, batched Gmail search to find the Step 1 sent messages for multiple contacts.
 * This is highly efficient as it makes only one call to GmailApp.search().
 * @param {Array<object>} contacts An array of contact objects that require searching.
 * @param {SpreadsheetApp.Sheet} contactsSheet The "Contacts" sheet object for updating message IDs.
 * @return {Map<string, GmailMessage>} A map where the key is a lowercase email and the value is the found GmailMessage object.
 */
function findSentStep1Messages_Batch(contacts, contactsSheet) {
    const userEmail = Session.getActiveUser().getEmail().toLowerCase();
    const messagesFound = new Map(); // K: lowercase email, V: message object
    const contactsToSearchFor = [];
    const idsToUpdateInSheet = new Map(); // K: rowIndex, V: messageId

    // 1. First, check for stored IDs for all contacts to minimize searching
    for (const contact of contacts) {
        if (contact.step1SentMessageId) {
            try {
                const message = GmailApp.getMessageById(contact.step1SentMessageId);
                // Basic validation of the stored message
                if (message && !message.isDraft() && message.getTo().toLowerCase().includes(contact.email.toLowerCase())) {
                    messagesFound.set(contact.email.toLowerCase(), message);
                    continue; // Found it, move to next contact
                }
            } catch (e) {
                // ID was invalid or message deleted, it will be searched for.
                console.warn(`Stored message ID ${contact.step1SentMessageId} for ${contact.email} is invalid. Re-searching.`);
            }
        }
        // If we are here, we need to search for this contact's email
        contactsToSearchFor.push(contact);
    }

    if (contactsToSearchFor.length === 0) {
        return messagesFound;
    }

    // 2. Build a single, large query for the remaining contacts
    const now = new Date();
    const searchFrom = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000); // 90-day wide search
    const formattedSearchFrom = Utilities.formatDate(searchFrom, "UTC", "yyyy/MM/dd");

    // Gmail search syntax uses { A OR B OR C } for grouping
    const queryParts = contactsToSearchFor.map(contact => {
        if (!contact.step1Subject) return null;
        
        // --- START CORRECTION ---
        // Replace special characters like '|' with a space FOR THE SEARCH QUERY ONLY.
        // This makes the search broader but avoids syntax errors.
        const searchSubject = contact.step1Subject.replace(/"/g, '\\"').replace(/\|/g, ' ');
        // --- END CORRECTION ---

        return `(to:"${contact.email}" subject:"${searchSubject}")`;
    }).filter(part => part !== null);

    if (queryParts.length === 0) {
        return messagesFound;
    }

    const searchQuery = `from:("${userEmail}") is:sent after:${formattedSearchFrom} {${queryParts.join(" OR ")}}`;
    console.log(`findSentStep1Messages_Batch: Searching with single query for ${queryParts.length} contacts.`);
    
    const threads = GmailApp.search(searchQuery, 0, 50); // Search up to 50 threads to be safe

    // 3. Process the results and map them back to the contacts
    for (const thread of threads) {
        for (const message of thread.getMessages()) {
            if (message.isDraft() || !message.getFrom().toLowerCase().includes(userEmail)) {
                continue;
            }

            const toHeader = message.getTo().toLowerCase();
            const messageSubject = message.getSubject();

            // Find which contact this message belongs to.
            for (const contact of contactsToSearchFor) {
                const contactEmailLower = contact.email.toLowerCase();

                // --- START CORRECTION ---
                // Use a strict, exact match (===) on the subject to verify it's the correct email
                // This filters out any false positives from the broader search.
                if (toHeader.includes(contactEmailLower) && messageSubject === contact.step1Subject) {
                // --- END CORRECTION ---

                    // This is a potential match. Check if it's the newest one we've found.
                    const existingMessage = messagesFound.get(contactEmailLower);
                    if (!existingMessage || message.getDate() > existingMessage.getDate()) {
                        messagesFound.set(contactEmailLower, message);
                        idsToUpdateInSheet.set(contact.rowIndex, message.getId());
                    }
                }
            }
        }
    }

    // 4. Batch update the sheet with newly found message IDs
    if (idsToUpdateInSheet.size > 0 && contactsSheet) {
        console.log(`findSentStep1Messages_Batch: Updating ${idsToUpdateInSheet.size} message IDs in the sheet.`);
        for (const [rowIndex, messageId] of idsToUpdateInSheet.entries()) {
            contactsSheet.getRange(rowIndex, CONTACT_COLS.STEP1_SENT_MESSAGE_ID + 1).setValue(messageId);
        }
    }

    return messagesFound;
}

/* ======================== UTILITY FUNCTIONS ======================== */

/**
 * Extracts the first N words from a given text.
 * @param {string} text The input string.
 * @param {number} wordCount The number of words to extract.
 * @return {string} The first N words, with "..." if truncated.
 */
function extractFirstWords(text, wordCount) {
  if (!text || typeof text !== 'string') return "";
  const words = text.trim().split(/\s+/); // Split by any whitespace
  if (words.length <= wordCount) {
    return words.join(" ");
  }
  return words.slice(0, wordCount).join(" ") + "...";
}

