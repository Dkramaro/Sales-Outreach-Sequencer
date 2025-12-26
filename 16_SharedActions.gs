/**
 * FILE: Shared Contact Actions
 * 
 * PURPOSE:
 * Handles common contact actions that are used across multiple views including
 * marking contacts as completed and ending sequences for entire companies.
 * 
 * KEY FUNCTIONS:
 * - markContactCompleted() - Mark contact as completed from any view
 * - endSequenceForCompany() - End sequence for all contacts in a company
 * 
 * DEPENDENCIES:
 * - 01_Config.gs: CONTACT_COLS
 * - 03_Database.gs: logAction
 * - 04_ContactData.gs: getContactByEmail, getAllContactsData
 * - 05_ContactUI.gs: viewContactCard
 * - 07_SequenceUI.gs: buildSequenceManagementCard, buildSelectContactsCard, viewStep2Contacts, viewContactsReadyForEmail
 * - 11_CallManagement.gs: buildCallManagementCard
 * 
 * @version 2.3
 */

/* ======================== MARK CONTACT COMPLETED ======================== */

/**
 * Marks a contact as completed in the sequence.
 * Updates status to "Completed", clears Next Step Date.
 * Refreshes the originating card view.
 */
function markContactCompleted(e) {
  const email = e.parameters.email;
  const origin = e.parameters.origin; // e.g., "callView", "stepView", "readyView"
  let originParams = {};
  if (e.parameters.originParamsJson) {
    try {
      originParams = JSON.parse(e.parameters.originParamsJson);
    } catch (err) {
      console.error("markContactCompleted: Could not parse originParamsJson: " + err);
      logAction("Error", "markContactCompleted internal error parsing params for " + email);
      return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification().setText("Error processing request. Returning to main menu."))
        .setNavigation(CardService.newNavigation().updateCard(buildAddOn({})))
        .build();
    }
  }

  // Retrieve and clear any pending notification from script properties
  let pendingNotificationText = PropertiesService.getScriptProperties().getProperty('PENDING_NOTIFICATION_FOR_SAVE_AND_END');
  if (pendingNotificationText) {
    PropertiesService.getScriptProperties().deleteProperty('PENDING_NOTIFICATION_FOR_SAVE_AND_END');
  }


  if (!email) {
    logAction("Error", "markContactCompleted: Email parameter missing.");
    return createNotification("Cannot mark complete: Contact email not provided.");
  }

  const contact = getContactByEmail(email);
  if (!contact) {
    logAction("Error", `markContactCompleted: Contact not found ${email}`);
    return createNotification("Contact not found: " + email);
  }

  if (contact.status === "Completed" && !pendingNotificationText) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("Contact " + contact.firstName + " " + contact.lastName + " is already marked as Completed."))
      .setNavigation(CardService.newNavigation().popCard())
      .build();
  }

  const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
  if (!spreadsheetId) {
    logAction("Error", "markContactCompleted Error: No database connected.");
    return createNotification("No database connected. Please connect to a database first.");
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);

    if (!contactsSheet) {
      logAction("Error", `markContactCompleted Error: Contacts sheet not found for ${email}`);
      return createNotification("Contacts sheet not found. Please refresh the add-on.");
    }

    contactsSheet.getRange(contact.rowIndex, CONTACT_COLS.STATUS + 1).setValue("Completed");
    contactsSheet.getRange(contact.rowIndex, CONTACT_COLS.NEXT_STEP_DATE + 1).setValue("");
    contactsSheet.getRange(contact.rowIndex, CONTACT_COLS.PERSONAL_CALLED + 1).setValue("No");
    contactsSheet.getRange(contact.rowIndex, CONTACT_COLS.WORK_CALLED + 1).setValue("No");
    SpreadsheetApp.flush();

    logAction("Contact Completed", `Marked contact ${contact.firstName} ${contact.lastName} (${email}) as Completed.`);

    let returnCard;
    const pageForReturn = originParams.page || '1';
    const stepForReturn = originParams.step;

    // Determine the card to refresh to
    if (origin === 'callView') {
      returnCard = buildCallManagementCard({ parameters: { page: pageForReturn } });
    } else if (origin === 'readyView') {
      returnCard = viewContactsReadyForEmail({ parameters: { page: pageForReturn } });
    } else if (origin === 'stepView' && stepForReturn) {
      if (stepForReturn === '2' || parseInt(stepForReturn) === 2) {
        returnCard = viewStep2Contacts({ parameters: { step: stepForReturn, page: pageForReturn } });
      } else {
        returnCard = buildSelectContactsCard({ parameters: { step: stepForReturn, page: pageForReturn } });
      }
    } else {
      console.warn("markContactCompleted: Unknown origin. Defaulting to Sequence Management.");
      logAction("Warning", "markContactCompleted: Unknown origin '" + origin + "' for " + email);
      returnCard = buildSequenceManagementCard();
    }
    
    let finalNotificationText = `Contact ${contact.firstName} ${contact.lastName} marked as Completed.`;
    if (pendingNotificationText) {
        finalNotificationText = pendingNotificationText + " " + finalNotificationText;
    }


    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(finalNotificationText))
      .setNavigation(CardService.newNavigation().updateCard(returnCard))
      .build();

  } catch (error) {
    console.error("Error marking contact " + email + " as completed: " + error + "\n" + error.stack);
    logAction("Error", `Error marking contact ${email} as completed: ${error.toString()}`);
    const fallbackCard = (origin && originParams) ?
        ( origin === 'callView' ? buildCallManagementCard({ parameters: { page: originParams.page || '1' } }) :
          origin === 'readyView' ? viewContactsReadyForEmail({ parameters: { page: originParams.page || '1' } }) :
          origin === 'stepView' && originParams.step ?
            (originParams.step === '2' ? viewStep2Contacts({ parameters: { step: originParams.step, page: originParams.page || '1' } }) : buildSelectContactsCard({ parameters: { step: originParams.step, page: originParams.page || '1' } }) ) :
            buildSequenceManagementCard()
        ) : buildSequenceManagementCard();

    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText("Error updating contact: " + error.message))
      .setNavigation(CardService.newNavigation().updateCard(fallbackCard))
      .build();
  }
}

/* ======================== END SEQUENCE FOR COMPANY ======================== */

/**
 * Ends the sequence for all contacts within the same company as the specified contact.
 * Also saves any pending edits (priority, notes, tags) for the current contact.
 */
function endSequenceForCompany(e) {
    const currentContactEmail = e.parameters.email;
    const formInputs = e.formInput;

    if (!currentContactEmail) {
        logAction("Error", "endSequenceForCompany: Current contact email parameter missing.");
        return createNotification("Critical error: Current contact email not provided.");
    }

    const spreadsheetId = PropertiesService.getUserProperties().getProperty("SPREADSHEET_ID");
    if (!spreadsheetId) {
        logAction("Error", "endSequenceForCompany: No database connected.");
        return createNotification("No database connected. Please connect to a database first.");
    }

    try {
        const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        const contactsSheet = spreadsheet.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);
        if (!contactsSheet) {
            logAction("Error", "endSequenceForCompany: Contacts sheet not found.");
            return createNotification("Contacts sheet not found. Please refresh.");
        }

        const allContactsDataWithRowIndex = getAllContactsData(); 
        const currentContactObject = allContactsDataWithRowIndex.find(c => c.email.toLowerCase() === currentContactEmail.toLowerCase());

        if (!currentContactObject) {
            logAction("Error", `endSequenceForCompany: Current contact ${currentContactEmail} not found.`);
            return createNotification(`Contact ${currentContactEmail} not found. Cannot proceed.`);
        }

        const companyName = currentContactObject.company;

        // Safeguard for empty company
        if (!companyName || companyName.trim() === "") {
            logAction("Warning", `endSequenceForCompany: No company for ${currentContactEmail}.`);
            const newPriorityOnly = formInputs.newPriority;
            const notesOnly = formInputs.notes || "";
            const newTagsOnly = formInputs.tags || "";

            if (!newPriorityOnly || !["High", "Medium", "Low"].includes(newPriorityOnly)) {
                return createNotification("Invalid priority value provided.");
            }
            
            const singleRowRange = contactsSheet.getRange(currentContactObject.rowIndex, 1, 1, contactsSheet.getLastColumn());
            const singleRowData = singleRowRange.getValues()[0];

            singleRowData[CONTACT_COLS.PRIORITY] = newPriorityOnly;
            singleRowData[CONTACT_COLS.NOTES] = notesOnly;
            singleRowData[CONTACT_COLS.TAGS] = newTagsOnly;
            singleRowData[CONTACT_COLS.STATUS] = "Completed";
            singleRowData[CONTACT_COLS.NEXT_STEP_DATE] = "";
            singleRowData[CONTACT_COLS.PERSONAL_CALLED] = "No";
            singleRowData[CONTACT_COLS.WORK_CALLED] = "No";
            
            singleRowRange.setValues([singleRowData]);
            SpreadsheetApp.flush();
            logAction("End Sequence (Single/NoCompany)", `Sequence ended for ${currentContactEmail}.`);
            
            const updatedCardView = viewContactCard({ parameters: { email: currentContactEmail } });
            return CardService.newActionResponseBuilder()
                .setNotification(CardService.newNotification().setText(`Sequence ended for ${currentContactObject.firstName}. Edits saved.`))
                .setNavigation(CardService.newNavigation().updateCard(updatedCardView))
                .build();
        }

        // Proceed with company-wide update
        const contactsInCompanyToProcess = allContactsDataWithRowIndex.filter(c => c.company === companyName);
        let updatedCount = 0;
        let logSummary = [];

        const currentContactFormPriority = formInputs.newPriority;
        if (!currentContactFormPriority || !["High", "Medium", "Low"].includes(currentContactFormPriority)) {
            logAction("Error", `endSequenceForCompany: Invalid priority for ${currentContactEmail}.`);
            return createNotification(`Invalid priority value. No changes made.`);
        }

        for (const contact of contactsInCompanyToProcess) {
            const rowRange = contactsSheet.getRange(contact.rowIndex, 1, 1, contactsSheet.getLastColumn());
            const rowData = rowRange.getValues()[0];
            let itemLogDetails = [`${contact.email}:`];
            let changed = false;

            if (contact.email.toLowerCase() === currentContactEmail.toLowerCase()) {
                const notesFromForm = formInputs.notes || "";
                const tagsFromForm = formInputs.tags || "";

                if (rowData[CONTACT_COLS.PRIORITY] !== currentContactFormPriority) {
                    rowData[CONTACT_COLS.PRIORITY] = currentContactFormPriority;
                    itemLogDetails.push(`Priority->${currentContactFormPriority}`);
                    changed = true;
                }
                if (String(rowData[CONTACT_COLS.NOTES] || "") !== notesFromForm) {
                    rowData[CONTACT_COLS.NOTES] = notesFromForm;
                    itemLogDetails.push(`Notes updated`);
                    changed = true;
                }
                if (String(rowData[CONTACT_COLS.TAGS] || "") !== tagsFromForm) {
                    rowData[CONTACT_COLS.TAGS] = tagsFromForm;
                    itemLogDetails.push(`Tags->"${tagsFromForm}"`);
                    changed = true;
                }
            }

            // Mark all as Completed
            if (rowData[CONTACT_COLS.STATUS] !== "Completed") {
                rowData[CONTACT_COLS.STATUS] = "Completed";
                itemLogDetails.push(`Status->Completed`);
                changed = true;
            }
            if (rowData[CONTACT_COLS.NEXT_STEP_DATE] !== "") {
                 rowData[CONTACT_COLS.NEXT_STEP_DATE] = "";
                 changed = true;
            }
            if (rowData[CONTACT_COLS.PERSONAL_CALLED] !== "No") {
                 rowData[CONTACT_COLS.PERSONAL_CALLED] = "No";
                 changed = true;
            }
            if (rowData[CONTACT_COLS.WORK_CALLED] !== "No") {
                rowData[CONTACT_COLS.WORK_CALLED] = "No";
                changed = true;
            }
            
            if (changed) {
                 rowRange.setValues([rowData]);
                 updatedCount++;
                 if (itemLogDetails.length > 1) {
                    logSummary.push(itemLogDetails.join(' '));
                 }
            }
        }

        if (updatedCount > 0) {
            SpreadsheetApp.flush();
            logAction("End Sequence (Company)", `Processed ${updatedCount} contacts for company "${companyName}".`);
        }

        const finalNotificationMessage = `Sequence ended for ${updatedCount} contact(s) in company "${companyName}".`;
        const refreshedCard = viewContactCard({ parameters: { email: currentContactEmail } });
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText(finalNotificationMessage))
            .setNavigation(CardService.newNavigation().updateCard(refreshedCard))
            .build();

    } catch (error) {
        console.error(`Error in endSequenceForCompany: ${error}\n${error.stack}`);
        logAction("Error", `endSequenceForCompany (${currentContactEmail}): ${error.toString()}`);
        const errorFallbackCard = viewContactCard({ parameters: { email: currentContactEmail } });
        return CardService.newActionResponseBuilder()
            .setNotification(CardService.newNotification().setText("Error ending sequence: " + error.message))
            .setNavigation(CardService.newNavigation().updateCard(errorFallbackCard))
            .build();
    }
}

