const config = getConfig();

/**
 * Sanitizes and cleans email body text for Google Docs and LLM consumption.
 */
function cleanEmailBody(rawBody) {
  if (!rawBody) return "[No Body Content]";

  // 1. SIGNATURE REMOVAL (Same as before)
  const signatureDelimiters = ["--", "__", "\n--\n", "\n-- \n",
    "\nregards", "\nBest regards", "\nWith best regards", "\nKind regards",
    "\nMit freundlichen Grüßen", "\nViele Grüße", "\nDanke und Gruß", "\nGrüße",
    "\n>", "Join with Google Meet"];
  let cleanedBody = rawBody;
  for (const delimiter of signatureDelimiters) {
    const index = cleanedBody.indexOf(delimiter);
    if (index !== -1) { cleanedBody = cleanedBody.substring(0, index); }
  }

  // 2. UPDATED REGEX FOR GERMAN CHARACTERS
  // \u00C0-\u00FF covers Ä, ö, ü, ß and other accented European characters
  // \x20-\x7E covers standard English letters/numbers
  // \x0A\x0D\x09 covers newlines and tabs
  cleanedBody = cleanedBody.replace(/[^\x20-\x7E\x0A\x0D\x09\u00C0-\u00FF]/g, "");

  // 3. WHITESPACE & TRUNCATION
  cleanedBody = cleanedBody.replace(/\n{3,}/g, "\n\n");
  const MAX_MESSAGE_LENGTH = config.MAX_MESSAGE_LENGTH;
  if (cleanedBody.length > MAX_MESSAGE_LENGTH) {
    cleanedBody = cleanedBody.substring(0, MAX_MESSAGE_LENGTH) + "\n... [Gekürzt]";
  }

  //Logger.log("cleaned Body: ${cleanedBody.trim()}")
  return cleanedBody.trim();
}

function updateEmailsForLabel(labelName, docBaseName) {

  Logger.log(`Start working on ${labelName}`);

  // --- Configuration ---
  const exportFolderName = config.EXPORT_FOLDER_NAME;
  const maxRetries = config.MAX_RETRIES;
  const retryDelay = config.RETRY_DELAY;
  const batchSize = config.INITIAL_BATCH_SIZE;

  // --- Get the label and emails ---
  const label = GmailApp.getUserLabelByName(labelName);
  if (!label) { return; }

  // We use search() instead of label.getThreads() to filter by time
  //const threads = label.getThreads();
  const threads = GmailApp.search(`label:"${labelName}" newer_than:12m`);

  let newEmailsToProcess = [];
  let lastUpdateTimeProp = 'lastUpdateTime_' + labelName.replace(/\//g, '_');
  let lastUpdateTime = PropertiesService.getUserProperties().getProperty(lastUpdateTimeProp);
  let lastUpdateDate = lastUpdateTime ? new Date(lastUpdateTime) : null;

  for (const thread of threads) {
    const messages = thread.getMessages();
    for (const message of messages) {
      if (!lastUpdateDate || message.getDate() > lastUpdateDate) {
        newEmailsToProcess.push(message);
      }
    }
  }

  Logger.log(`  Found ${newEmailsToProcess.length} to process ...`);

  if (newEmailsToProcess.length === 0) return;

  // --- Get or create the export folder ---
  let exportFolder = DriveApp.getFoldersByName(exportFolderName).next();

  // --- Process and append new emails ---
  for (let i = 0; i < newEmailsToProcess.length; i += batchSize) {
    const batch = newEmailsToProcess.slice(i, i + batchSize);

    // 1. Prepare the batch content string first so we know its size
    let batchContent = "";
    for (const message of batch) {
      const subject = message.getSubject();
      const safeBody = cleanEmailBody(message.getPlainBody());
      const date = message.getDate();
      const sender = message.getFrom();
      const toRecipients = message.getTo(); // Get 'To' recipients
      const ccRecipients = message.getCc(); // Get 'Cc' recipients

      batchContent += `## Email\n`;
      batchContent += `### Subject: ${subject}\n`;
      batchContent += `### Date: ${date}\n`;
      batchContent += `### Sender: ${sender}\n`;
      batchContent += `### To: ${toRecipients}\n`; // Add 'To' recipients
      if (ccRecipients) { // Check if there are 'Cc' recipients
        batchContent += `### Cc: ${ccRecipients}\n`; // Add 'Cc' recipients
      }
      batchContent += `### Body:\n${safeBody}\n\n`;
      batchContent += "---\n\n";
    }

    // 2. Find a document that has enough "Character Space"
    let docToUse = null;
    let part = 1;
    const MAX_CHARS_ALLOWED = config.MAX_CHARS_ALLOWED; // Safety limit below 1 million

    while (!docToUse) {
      let currentDocName = `${docBaseName}_Part${part}${config.FILENAME_SUFFIX}`;
      let files = exportFolder.getFilesByName(currentDocName);
      let docFile;

      if (files.hasNext()) {
        docFile = files.next();
        let tempDoc = DocumentApp.openById(docFile.getId());
        let currentLen = tempDoc.getBody().getText().length;

        // If current text + new batch is under the limit, use this doc
        if (currentLen + batchContent.length < MAX_CHARS_ALLOWED) {
          docToUse = tempDoc;
        } else {
          tempDoc.saveAndClose(); // Close and try Part 2, 3...
          part++;
        }
      } else {
        // Create a brand new Part
        let newDoc = DocumentApp.create(currentDocName);
        let newFile = DriveApp.getFileById(newDoc.getId());
        exportFolder.addFile(newFile);
        DriveApp.getRootFolder().removeFile(newFile);
        docToUse = newDoc;
        Logger.log("Created New Part: " + currentDocName);
      }
    }

    // 3. Append and Save
    let success = false;
    let retries = 0;

    while (retries < maxRetries && !success) {
      try {
        docToUse.getBody().appendParagraph(batchContent);
        docToUse.saveAndClose();

        Utilities.sleep(config.SLEEP_BETWEEN_BATCHES);

        success = true;
        Logger.log(`Successfully appended batch starting at index ${i}`);

      } catch (e) {
        retries++;
        Logger.log(`Error on batch ${i}, attempt ${retries}: ${e}`);
        Utilities.sleep(retryDelay * retries);
        docToUse = DocumentApp.openById(docToUse.getId()); // Re-open
      }
    }

    if (!success) {
      Logger.log(`CRITICAL: Batch starting at index ${i} failed after ${maxRetries} retries.`);
      // We don't break; we try the next batch which might trigger a new "Part"
    }
  }


  // --- Update the last update time ---
  PropertiesService.getUserProperties().setProperty(lastUpdateTimeProp, new Date().toISOString());
}

function resetLastUpdateTime(labelName) {
  let lastUpdateTimeProp = 'lastUpdateTime_' + labelName.replace(/\//g, '_');
  PropertiesService.getUserProperties().deleteProperty(lastUpdateTimeProp);
  Logger.log("Reset successful. The next run for " + labelName + " will process all emails.");
}

// --- Helper functions to run for each label ---
function _run_Update() {
  config.LABEL_CONFIG.forEach(item => {
    updateEmailsForLabel(item.label, item.docBaseName);
  });
}

// -- resets timer on a label to nothing
function label_RESET(label) {
  resetLastUpdateTime(label);
}

function run_Reset() {
  config.LABEL_CONFIG.forEach(item => {
    label_RESET(item.label);
  });
}
