/**
 * FILE: Gmail Labeling Operations
 * 
 * PURPOSE:
 * Handles background batch labeling of Gmail threads with priority and sequence labels.
 * Uses time-based triggers for long-running batch operations with progress tracking.
 * 
 * KEY FUNCTIONS:
 * - startLabelingProcess() - Initiate batch labeling job
 * - continueLabelingProcess() - Process batches with time-based trigger
 * - getOrCreateLabel() - Gmail label creation/retrieval
 * - deleteTrigger() - Cleanup trigger helper
 * 
 * DEPENDENCIES:
 * - 01_Config.gs: CONTACT_COLS
 * - 03_Database.gs: logAction
 * - 04_ContactData.gs: getAllContactsData
 * 
 * @version 2.3
 */

/* ======================== BATCH LABELING PROCESS ======================== */

/**
 * STARTER FUNCTION
 * Called by the button. It queues up only the necessary contacts for processing.
 */
function startLabelingProcess() {
  const allContacts = getAllContactsData();
  
  // Get a list of contact emails that are NOT yet labeled and might need processing.
  const emailsToProcess = allContacts
    .filter(contact => contact.labeled !== 'Yes' && contact.step1Subject)
    .map(contact => contact.email);

  if (emailsToProcess.length === 0) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("All contacts have already been labeled. Nothing to do!"))
      .build();
  }

  // Store the job details
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('labeling_emailsToProcess', JSON.stringify(emailsToProcess));
  userProperties.setProperty('labeling_currentIndex', '0');

  // Ensure no old trigger is running
  deleteTrigger('continueLabelingProcess');

  // Create a new trigger to start the process
  ScriptApp.newTrigger('continueLabelingProcess')
    .timeBased()
    .after(5 * 1000) // 5 seconds
    .create();

  logAction("Batch Labeling Started", `Queued ${emailsToProcess.length} contacts for labeling.`);
  return CardService.newActionResponseBuilder()
    .setNotification(CardService.newNotification().setText(`Started labeling ${emailsToProcess.length} contacts in the background. You will receive an email when complete.`))
    .build();
}

/**
 * CONTINUATION FUNCTION
 * Processes one batch, applying both Priority and Sequence labels, and reschedules itself.
 */
function continueLabelingProcess() {
  const userProperties = PropertiesService.getUserProperties();
  const emailsToProcess = JSON.parse(userProperties.getProperty('labeling_emailsToProcess'));
  let currentIndex = parseInt(userProperties.getProperty('labeling_currentIndex'));

  const BATCH_SIZE = 200;
  const TIME_LIMIT_MS = 5 * 60 * 1000;
  const startTime = new Date();

  const allContacts = getAllContactsData();
  const contactMap = new Map(allContacts.map(c => [c.email, c]));
  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  const contactsSheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(CONFIG.CONTACTS_SHEET_NAME);
  
  // Get priority labels once
  const highPriorityLabel = getOrCreateLabel("Priority: High 🔥");
  const mediumPriorityLabel = getOrCreateLabel("Priority: Medium 🟠");
  const lowPriorityLabel = getOrCreateLabel("Priority: Low ⚪");

  let updatesForSheet = {};

  for (let i = 0; i < BATCH_SIZE && currentIndex < emailsToProcess.length; i++) {
    if (new Date() - startTime > TIME_LIMIT_MS) break;

    const email = emailsToProcess[currentIndex];
    const contact = contactMap.get(email);
    if (!contact || contact.labeled === 'Yes' || !contact.step1Subject) {
      currentIndex++;
      continue;
    }

    try {
      let threadId = contact.threadId;

      // 1. Find missing threadId if necessary
      if (!threadId) {
        const searchSubjectForQuery = contact.step1Subject.replace(/\|/g, ' ');
        const searchQuery = `to:(${contact.email}) in:sent subject:("${searchSubjectForQuery}")`;
        const threads = GmailApp.search(searchQuery, 0, 5);
        
        let targetThread = null;
        if (threads.length > 0) {
            for (const thread of threads) {
                const messages = thread.getMessages();
                if (messages.length > 0 && messages[0].getSubject() === contact.step1Subject) {
                    targetThread = thread;
                    break;
                }
            }
        }

        if (targetThread) {
            threadId = targetThread.getId();
            if (!updatesForSheet[contact.rowIndex]) updatesForSheet[contact.rowIndex] = {};
            updatesForSheet[contact.rowIndex][CONTACT_COLS.THREAD_ID + 1] = threadId;
        }
      }

      // 2. If we have a thread, apply all labels
      if (threadId) {
        const thread = GmailApp.getThreadById(threadId);

        if (thread) {
          // Apply Priority Label
          let priorityLabelToApply = null;
          if (contact.priority === "High" && highPriorityLabel) priorityLabelToApply = highPriorityLabel;
          else if (contact.priority === "Medium" && mediumPriorityLabel) priorityLabelToApply = mediumPriorityLabel;
          else if (contact.priority === "Low" && lowPriorityLabel) priorityLabelToApply = lowPriorityLabel;
          
          if (priorityLabelToApply) {
            thread.addLabel(priorityLabelToApply);
          }

          // Apply Sequence Label
          if (contact.sequence) {
            const sequenceLabelName = "Seq: " + contact.sequence;
            const sequenceLabel = getOrCreateLabel(sequenceLabelName);
            if (sequenceLabel) {
              thread.addLabel(sequenceLabel);
            }
          }
        }
        
        // 3. Mark as labeled
        if (!updatesForSheet[contact.rowIndex]) updatesForSheet[contact.rowIndex] = {};
        updatesForSheet[contact.rowIndex][CONTACT_COLS.LABELED + 1] = 'Yes';
      }
    } catch (e) { console.error(`Error processing ${email}: ${e}`); }
    
    currentIndex++;
  }

  // Perform batched writes to the spreadsheet
  if (Object.keys(updatesForSheet).length > 0) {
    for (const rowIndex in updatesForSheet) {
      for (const colIndex in updatesForSheet[rowIndex]) {
          contactsSheet.getRange(parseInt(rowIndex), parseInt(colIndex)).setValue(updatesForSheet[rowIndex][colIndex]);
      }
    }
    SpreadsheetApp.flush();
  }

  // Reschedule or Finish
  if (currentIndex < emailsToProcess.length) {
    userProperties.setProperty('labeling_currentIndex', currentIndex.toString());
    console.log(`Processed batch. Next index: ${currentIndex} of ${emailsToProcess.length}`);
  } else {
    // All done! Clean up.
    deleteTrigger('continueLabelingProcess');
    userProperties.deleteProperty('labeling_emailsToProcess');
    userProperties.deleteProperty('labeling_currentIndex');
    
    const userEmail = Session.getActiveUser().getEmail();
    GmailApp.sendEmail(userEmail, "Sales Outreach: Labeling Complete", `The background labeling process has finished for all ${emailsToProcess.length} contacts.`);
    logAction("Batch Labeling Finished", `Successfully processed all ${emailsToProcess.length} contacts.`);
  }
}

/* ======================== LABEL HELPERS ======================== */

/**
 * Gets a Gmail label by name, creating it if it doesn't exist.
 */
function getOrCreateLabel(labelName) {
  try {
    let label = GmailApp.getUserLabelByName(labelName);
    if (!label) {
      label = GmailApp.createLabel(labelName);
    }
    return label;
  } catch (error) {
    console.error("Error getting or creating label '" + labelName + "': " + error);
    return null;
  }
}

/**
 * Helper to delete triggers by function name to prevent duplicates.
 */
function deleteTrigger(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

