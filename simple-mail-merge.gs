/**
 * ============================================================================
 * Gadi's Simple Mail Merge
 * ============================================================================
 *
 * Gmail mail merge for Google Sheets.
 *
 * @author Gadi Evron (with Claude, and some help from ChatGPT)
 * @version 3.3
 * @updated 2025-11-18
 * @license MIT
 * ============================================================================
 */

// ============================================================================
// CONFIGURATION
// ============================================================================
const CONFIG = {
  SHEETS: {
    CONTACTS: "Contacts",
    EMAIL_DRAFT: "Email Draft",
    INSTRUCTIONS: "Instructions"
  },
  REQUIRED_COLUMNS: ["Name", "Last Name", "Email", "Company", "Title", "Custom1", "Custom2", "Successfully Sent"],
  SAMPLE: {
    NAME: "John",
    LAST_NAME: "Doe",
    EMAIL: "john.doe@example.com",
    COMPANY: "Example Corp",
    TITLE: "Manager",
    CUSTOM1: "Engineering",
    CUSTOM2: "San Francisco"
  }
};

// Pacing: aim for ~300ms total per row.
// If preflight search (~120ms) ran, sleep 180ms; otherwise sleep 300ms.
const PAUSE = {
  SENDING_MS: 0,
  BETWEEN_SENDS_DEFAULT_MS: 300,
  BETWEEN_SENDS_AFTER_SEARCH_MS: 180
};

// Gmail search configuration
const SEARCH_CONFIG = {
  REPLY_MODE_WINDOW: "3d"  // Search threads from last 3 days
};

// ============================================================================
// STATE MANAGEMENT
// ============================================================================
const MailMergeState = {
  draftsCache: new Map(),
  cacheTime: null,
  expandedCache: false,

  refreshDrafts() {
    if (this.cacheTime && (Date.now() - this.cacheTime < 30000)) return;

    const drafts = GmailApp.getDrafts();
    this.draftsCache.clear();

    for (let i = 0; i < Math.min(drafts.length, 10); i++) {
      const subject = drafts[i].getMessage().getSubject().trim();
      if (subject) this.draftsCache.set(subject, drafts[i]);
    }

    this.cacheTime = Date.now();
  },

  expandDraftsCache() {
    const drafts = GmailApp.getDrafts();
    this.draftsCache.clear();

    for (let i = 0; i < Math.min(drafts.length, 50); i++) {
      const subject = drafts[i].getMessage().getSubject().trim();
      if (subject) this.draftsCache.set(subject, drafts[i]);
    }

    this.expandedCache = true;
    this.cacheTime = Date.now();
  },

  getDraft(subject) {
    this.refreshDrafts();
    let draft = this.draftsCache.get(subject.trim());

    // If not found in first 10, try expanding to 50
    if (!draft && !this.expandedCache) {
      this.expandDraftsCache();
      draft = this.draftsCache.get(subject.trim());
    }

    return draft || null;
  }
};

// ============================================================================
// UTILITIES
// ============================================================================
function extractAndValidateEmail(emailField) {
  if (!emailField) return null;

  let email = emailField.toString().trim();

  // Extract email from "Name <email@domain.com>" format
  const match = email.match(/<([^>]+)>/);
  if (match) {
    email = match[1].trim();
  } else {
    // Fallback – pick first email-like token if no angle brackets
    const m = email.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
    if (m) email = m[0];
  }

  // Validate the extracted email
  if (email.length < 5 ||
      email.length > 254 ||
      !email.includes('@') ||
      !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return null;
  }

  return email;
}

function personalizeText(text, name, lastName, contact = null) {
  if (!text) return '';

  let result = text.toString()
    .replace(/\{\{name\}\}/gi, name || '')
    .replace(/\{\{last name\}\}/gi, lastName || '');

  // Enhanced personalization if contact object provided
  if (contact) {
    result = result
      .replace(/\{\{email\}\}/gi, contact.email || '')
      .replace(/\{\{company\}\}/gi, contact.company || '')
      .replace(/\{\{title\}\}/gi, contact.title || '')
      .replace(/\{\{custom1\}\}/gi, contact.custom1 || '')
      .replace(/\{\{custom2\}\}/gi, contact.custom2 || '');
  }

  return result;
}

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

// Robust status writer with tiny retry to avoid missed updates
function writeStatus(statusCell, value, bg) {
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      statusCell.setValue(value);
      if (bg) statusCell.setBackground(bg);
      SpreadsheetApp.flush();
      return true;
    } catch (e) {
      if (attempt < 2) Utilities.sleep(100);
    }
  }
  // Last-chance fallback: value only
  try {
    statusCell.setValue(value);
    SpreadsheetApp.flush();
  } catch (e) {}
  return false;
}

function shouldPreflight(status) {
  const s = String(status || "").toUpperCase();
  return s === "" || s.includes("SENDING") || s.includes("FAILED");
}

/**
 * ============================================================================
 * STATUS UPDATE HELPER (Used by both New Email and Reply Mode)
 * ============================================================================
 *
 * Shared helper to avoid duplication in status updates.
 * Both sendNewEmailMode() and sendReplyMode() call this.
 */
function _updateSuccessStatus({statusCell, successCount, totalCount, contact, messageType, bgColor}) {
  if (!statusCell) return {success: false};

  try {
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd hh:mm:ss z");
    const multiTag = contact && contact.multiEmail ? " | MULTI-EMAIL CELL: used first; others skipped" : "";
    const message = `✅ ${messageType} (${successCount}/${totalCount}) on ${dateStr}${multiTag}`;

    writeStatus(statusCell, message, bgColor);
    return {success: true, message: message};
  } catch (e) {
    return {success: false, error: e.message};
  }
}

// ============================================================================
// REPLY MODE HELPER FUNCTIONS
// ============================================================================

/**
 * Sanitize Gmail search input to prevent query injection
 * @private
 * @param {string} input - User input to sanitize
 * @return {string} Sanitized input safe for Gmail search
 */
function sanitizeGmailSearchInput(input) {
  if (!input) return "";
  return String(input)
    .replace(/\\/g, "\\\\")
    .replace(/"/g, '\\"')
    .replace(/older_than:/gi, "")
    .replace(/newer_than:/gi, "")
    .trim();
}

/**
 * Find Gmail threads matching a subject for reply (3-day-only search)
 * @param {string} originalSubject - Subject line to search for
 * @param {string} searchType - Either "TO" or "BCC" mode
 * @param {string} contactEmail - Email address to search for
 * @return {GmailThread[]} Array of matching threads (empty if none found)
 *
 * Searches only within last 3 days (newer_than:3d) with no fallback to older emails.
 * TO mode: searches for emails sent to contactEmail
 * BCC mode: searches for emails with matching subject
 * All inputs sanitized to prevent query injection.
 */
function findThreadForReply(originalSubject, searchType, contactEmail) {
  try {
    if (!originalSubject || !originalSubject.trim()) return [];
    if (!searchType || !contactEmail) return [];
    if (searchType !== "TO" && searchType !== "BCC") return [];

    // Normalize subject (remove Re: prefix)
    const cleanSubject = originalSubject.replace(/^Re:\s*/i, '').trim();
    if (!cleanSubject) return [];

    // Sanitize to prevent query injection
    const email = sanitizeGmailSearchInput(contactEmail);
    const subject = sanitizeGmailSearchInput(cleanSubject);
    if (!email || !subject) return [];

    // Search both Sent (emails I sent) and Inbox (emails from contact) with exact subject match (3-day window)
    const sentQuery = `in:sent from:me (to:"${email}" OR cc:"${email}" OR bcc:"${email}") subject:"${subject}" newer_than:${SEARCH_CONFIG.REPLY_MODE_WINDOW}`;
    const inboxQuery = `(from:"${email}" OR to:"${email}") subject:"${subject}" newer_than:${SEARCH_CONFIG.REPLY_MODE_WINDOW}`;

    try {
      const sentThreads = GmailApp.search(sentQuery);
      const inboxThreads = GmailApp.search(inboxQuery);

      // Combine and deduplicate by thread ID
      const threadMap = new Map();
      for (const thread of sentThreads) {
        if (thread) threadMap.set(thread.getId(), thread);
      }
      for (const thread of inboxThreads) {
        if (thread) threadMap.set(thread.getId(), thread);  // Overwrites if same ID (dedup)
      }

      // Convert back to array and sort by most recent message date (descending)
      let allThreads = Array.from(threadMap.values());
      allThreads.sort((a, b) => {
        const dateA = a.getLastMessageDate().getTime();
        const dateB = b.getLastMessageDate().getTime();
        return dateB - dateA;  // Most recent first
      });

      return allThreads;
    } catch (searchError) {
      // Search failed, log for debugging
      console.error(`[findThreadForReply] Gmail search failed for ${email}: ${searchError.message}`);
      return [];
    }
  } catch (e) {
    console.error(`[findThreadForReply] Unexpected error: ${e.message}`);
    return [];
  }
}


// ============================================================================
// TEMPLATE VALIDATION FUNCTIONS
// ============================================================================

// Handle escaped tags (\\{{Name}} → literal {{Name}})
function preprocessEscapes(text) {
  if (!text) return {text: '', escapes: 0};

  const escapedPattern = /\\(\{\{[^}]*\}\})/g;
  const escapes = (text.match(escapedPattern) || []).length;
  const processed = text.replace(escapedPattern, '$1');

  return {text: processed, escapes: escapes};
}

// Text snippet generation for error context
function getContextSnippet(text, searchTerm) {
  const index = text.toLowerCase().indexOf(searchTerm.toLowerCase());
  if (index === -1) return '';

  const start = Math.max(0, index - 10);
  const end = Math.min(text.length, index + searchTerm.length + 10);
  const snippet = text.substring(start, end);

  return `...${snippet}...`;
}

// Detect malformed tags with location context
function detectMalformedTags(text, location) {
  const issues = [];

  if (text.includes('{{')) {
    const opens = (text.match(/\{\{/g) || []).length;
    const completes = (text.match(/\{\{[^}]*\}\}/g) || []).length;
    if (opens > completes) {
      const snippet = getContextSnippet(text, '{{');
      issues.push(`Malformed tag in ${location}: unmatched opening braces {{ ${snippet}`);
    }
  }

  if (text.includes('}}')) {
    const closes = (text.match(/\}\}/g) || []).length;
    const completes = (text.match(/\{\{[^}]*\}\}/g) || []).length;
    if (closes > completes) {
      const snippet = getContextSnippet(text, '}}');
      issues.push(`Malformed tag in ${location}: unmatched closing braces }} ${snippet}`);
    }
  }

  const singleOpen = text.replace(/\{\{/g, '').includes('{');
  const singleClose = text.replace(/\}\}/g, '').includes('}');

  if (singleOpen) {
    const snippet = getContextSnippet(text, '{');
    issues.push(`Malformed tag in ${location}: found single opening brace { ${snippet}`);
  }
  if (singleClose) {
    const snippet = getContextSnippet(text, '}');
    issues.push(`Malformed tag in ${location}: found single closing brace } ${snippet}`);
  }

  return issues;
}

// Detect other tag systems with location context
function detectOtherSystems(text, location) {
  const found = [];
  const otherSystemPatterns = [
    { name: "Word Mail Merge", pattern: /<<([^>]+)>>/g },
    { name: "Word Mail Merge", pattern: /«([^»]+)»/g },
    { name: "Mailchimp", pattern: /\*\|([^|]+)\|\*/g },
    { name: "Square Brackets", pattern: /\[([^\]]+)\]/g },
    { name: "Percent Tags", pattern: /%([^%]+)%/g },
    { name: "Dollar Tags", pattern: /\$([^$]+)\$/g }
  ];

  otherSystemPatterns.forEach(pattern => {
    const matches = [...text.matchAll(pattern.pattern)];
    matches.forEach(match => {
      found.push({ system: pattern.name, tag: match[0], content: match[1], location });
    });
  });

  return found;
}

// Detect unknown tags with location context and snippets
function detectUnknownTags(text, validTags, location) {
  const issues = [];
  const ourSystemTags = text.match(/\{\{([^}]*)\}\}/g) || [];

  ourSystemTags.forEach(tag => {
    const tagName = tag.slice(2, -2).trim();
    if (tagName && !validTags.some(valid => valid.toLowerCase() === tagName.toLowerCase())) {
      const suggestions = validTags.filter(valid =>
        valid.toLowerCase().includes(tagName.toLowerCase()) ||
        tagName.toLowerCase().includes(valid.toLowerCase())
      );
      const suggestion = suggestions.length > 0 ? ` (did you mean {{${suggestions[0]}}?)` : '';
      const snippet = getContextSnippet(text, tag);
      issues.push(`Unknown tag "${tag}" in ${location}: ${snippet}${suggestion}`);
    } else if (!tagName) {
      const snippet = getContextSnippet(text, tag);
      issues.push(`Empty tag "${tag}" in ${location}: ${snippet}`);
    }
  });

  return issues;
}

// Detect missing data with location context and snippets
function detectMissingData(text, contact, validTags, location) {
  const issues = [];
  const dataMap = {
    'name': contact.name, 'last name': contact.lastName, 'email': contact.email,
    'company': contact.company, 'title': contact.title, 'custom1': contact.custom1, 'custom2': contact.custom2
  };

  validTags.forEach(tagName => {
    const tagPattern = new RegExp(`\\{\\{${tagName.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&')}\\}\\}`, 'gi');
    if (tagPattern.test(text)) {
      const value = dataMap[tagName.toLowerCase()];
      if (!value || String(value).trim() === '') {
        const snippet = getContextSnippet(text, `{{${tagName}}}`);
        issues.push(`Tag {{${tagName}}} in ${location}: ${snippet} - no data provided for contact`);
      }
    }
  });

  return issues;
}

// Detect unbracketed tags with location context and snippets
function detectUnbracketedTags(text, validTags, location) {
  const issues = [];

  validTags.forEach(tagName => {
    const escapedTag = tagName.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&');
    const plainTextRegex = new RegExp(`\\b${escapedTag}\\b(?![^{]*\\}\\})`, 'gi');
    if (plainTextRegex.test(text)) {
      const snippet = getContextSnippet(text, tagName);
      issues.push(`Found "${tagName}" without brackets in ${location}: ${snippet} - did you mean {{${tagName}}}?`);
    }
  });

  return issues;
}

// Enhanced validateEmailTemplate with location context and snippets
function validateEmailTemplate(subject, body, contact = null) {
  const subjectProcessed = preprocessEscapes(subject || '');
  const bodyProcessed = preprocessEscapes(body || '');

  const hardErrors = [];
  const softWarnings = [];
  const validTags = ['Name', 'Last Name', 'Email', 'Company', 'Title', 'Custom1', 'Custom2'];

  // 1. Detect malformed tags with location context
  const subjectMalformed = detectMalformedTags(subjectProcessed.text, 'subject line');
  const bodyMalformed = detectMalformedTags(bodyProcessed.text, 'email body');
  hardErrors.push(...subjectMalformed, ...bodyMalformed);

  // 2. Detect other tag systems with location context
  const subjectOtherTags = detectOtherSystems(subjectProcessed.text, 'subject line');
  const bodyOtherTags = detectOtherSystems(bodyProcessed.text, 'email body');
  [...subjectOtherTags, ...bodyOtherTags].forEach(tag => {
    hardErrors.push(`Found ${tag.system} tag "${tag.tag}" in ${tag.location} - this system uses {{Name}}`);
  });

  // 3. Detect unknown tags with location context and snippets
  const subjectTags = detectUnknownTags(subjectProcessed.text, validTags, 'subject line');
  const bodyTags = detectUnknownTags(bodyProcessed.text, validTags, 'email body');
  hardErrors.push(...subjectTags, ...bodyTags);

  // 4. Detect missing data with location context
  if (contact) {
    const subjectMissing = detectMissingData(subjectProcessed.text, contact, validTags, 'subject line');
    const bodyMissing = detectMissingData(bodyProcessed.text, contact, validTags, 'email body');
    hardErrors.push(...subjectMissing, ...bodyMissing);
  }

  // 5. Detect unbracketed tags with location context
  const subjectUnbracketed = detectUnbracketedTags(subjectProcessed.text, validTags, 'subject line');
  const bodyUnbracketed = detectUnbracketedTags(bodyProcessed.text, validTags, 'email body');
  softWarnings.push(...subjectUnbracketed, ...bodyUnbracketed);

  return {
    hardErrors: [...new Set(hardErrors)],
    softWarnings: [...new Set(softWarnings)],
    isValid: hardErrors.length === 0,
    hasWarnings: softWarnings.length > 0,
    hardErrorMessage: hardErrors.length > 0 ? "Template Errors:\n\n" + hardErrors.join('\n') + "\n\nPlease fix these issues before sending." : "",
    warningMessage: softWarnings.length > 0 ? "Template Warnings:\n\n" + softWarnings.join('\n') + "\n\nContinue anyway?" : ""
  };
}

// ============================================================================
// CORE FUNCTIONS
// ============================================================================
function getContacts(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  // Standard 8-column format - status is always in column 8
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const contacts = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const emailFieldRaw = row[2] ? row[2].toString().trim() : "";

    if (!emailFieldRaw && !row[0] && !row[1]) continue;

    // Detect multiple emails; take first
    const emailsFound = emailFieldRaw.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/ig) || [];
    let validEmail = null;
    let primaryEmailRaw = null;

    if (emailsFound.length > 0) {
      primaryEmailRaw = emailsFound[0];
      validEmail = extractAndValidateEmail(primaryEmailRaw);
    }

    const contact = {
      rowNumber: i + 2,
      name: (row[0] || "").toString().trim(),
      lastName: (row[1] || "").toString().trim(),
      email: validEmail,
      company: (row[3] || "").toString().trim(),
      title: (row[4] || "").toString().trim(),
      custom1: (row[5] || "").toString().trim(),
      custom2: (row[6] || "").toString().trim(),
      status: row[7] || "",
      statusColumn: 8,
      multiEmail: emailsFound.length > 1,
      invalidEmail: emailFieldRaw && !validEmail,
      rawEmailField: emailFieldRaw
    };

    contacts.push(contact);
  }

  return contacts;
}

function getEmailDraft(sheet) {
  const subject = sheet.getRange("B1").getValue();
  if (!subject) return null;

  const subjectStr = subject.toString().trim();
  if (subjectStr === "Enter your Gmail draft subject here") {
    return null;
  }

  const draft = MailMergeState.getDraft(subjectStr);
  if (!draft) return null;

  const result = {
    subject: subjectStr,
    body: draft.getMessage().getBody(),
    attachments: draft.getMessage().getAttachments()
  };

  // Check for additional recipient fields (only if they exist)
  try {
    if (sheet.getLastRow() >= 6) {
      const senderName = sheet.getRange("B3").getValue();
      const additionalTo = sheet.getRange("B4").getValue();
      const cc = sheet.getRange("B5").getValue();
      const bcc = sheet.getRange("B6").getValue();

      result.senderName = senderName ? senderName.toString().trim() : "";
      result.additionalTo = additionalTo ? additionalTo.toString().trim() : "";
      result.cc = cc ? cc.toString().trim() : "";
      result.bcc = bcc ? bcc.toString().trim() : "";
    }
  } catch (e) {
    // Fields don't exist - this is fine
    result.senderName = "";
    result.additionalTo = "";
    result.cc = "";
    result.bcc = "";
  }

  // Check for reply mode fields (B9-B12, new layout)
  try {
    if (sheet.getLastRow() >= 9) {
      const replyModeRaw = sheet.getRange("B9").getValue().toString().trim();
      const validModes = ["New Email", "Reply: TO Mode", "Reply: BCC Mode"];
      result.replyMode = validModes.includes(replyModeRaw) ? replyModeRaw : "New Email";

      const b10Raw = sheet.getRange("B10").getValue().toString().trim();
      result.includeRecipients = ["YES", "Y", "TRUE", "1"].includes(b10Raw.toUpperCase());

      result.originalSubject = sheet.getRange("B11").getValue().toString().trim() || "";
      result.originalTo = sheet.getRange("B12").getValue().toString().trim() || "";
    } else {
      // Old sheet without reply fields - use defaults
      result.replyMode = "New Email";
      result.originalSubject = "";
      result.originalTo = "";
      result.includeRecipients = false;
    }
  } catch (e) {
    // Reply fields don't exist - this is fine (old sheet)
    result.replyMode = "New Email";
    result.originalSubject = "";
    result.originalTo = "";
    result.includeRecipients = false;
  }

  return result;
}

// ============================================================================
// MAIN FUNCTIONS
// ============================================================================
function sendEmails() {
  const ui = SpreadsheetApp.getUi();

  const contactsSheet = getSheet(CONFIG.SHEETS.CONTACTS);
  const draftSheet = getSheet(CONFIG.SHEETS.EMAIL_DRAFT);

  if (!contactsSheet || !draftSheet) {
    ui.alert("Error", "Required sheets not found. Run 'Create Merge Sheets' first.", ui.ButtonSet.OK);
    return;
  }

  const contacts = getContacts(contactsSheet);
  if (contacts.length === 0) {
    ui.alert("Error", "No valid contacts found", ui.ButtonSet.OK);
    return;
  }

  const emailDraft = getEmailDraft(draftSheet);
  if (!emailDraft) {
    ui.alert("Error", "Gmail draft not found. Check subject in B1 and ensure it matches your Gmail draft exactly.", ui.ButtonSet.OK);
    return;
  }

  const validation = validateEmailTemplate(emailDraft.subject, emailDraft.body, contacts[0] || null);
  if (!validation.isValid) {
    ui.alert("Template Errors", validation.hardErrorMessage, ui.ButtonSet.OK);
    return;
  }
  if (validation.hasWarnings && ui.alert("Template Warnings", validation.warningMessage, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
    return;
  }

  // Detect truly fresh run (no statuses present at all) → no preflight on blanks
  const isFreshRun = contacts.every(c => String(c.status || "").trim() === "");

  const toSend = contacts.filter(c => {
    const s = (c.status || "").toString().toUpperCase();
    return !s.includes("SENT SUCCESSFULLY");
  });

  // Build cross-run duplicate set once
  const previouslySentEmails = new Set(
    contacts
      .filter(c => String(c.status || "").toUpperCase().includes("SENT SUCCESSFULLY"))
      .map(c => String(c.email || "").toLowerCase())
  );

  if (toSend.length === 0) {
    ui.alert("Complete", "All emails already sent!", ui.ButtonSet.OK);
    return;
  }

  // Build confirmation message with mode info if in reply mode
  const modeInfo = (emailDraft.replyMode && emailDraft.replyMode !== "New Email")
    ? `\n\nMode: ${emailDraft.replyMode}\nSearching: "${emailDraft.originalSubject.substring(0, 47)}${emailDraft.originalSubject.length > 47 ? "..." : ""}"`
    : "";

  if (ui.alert("Confirm", `Send ${toSend.length} emails?${modeInfo}\n\nWatch the Status column for live progress...`, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
    return;
  }

  // Make sure we're viewing the Contacts sheet so user can see updates
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(contactsSheet);

  let successCount = 0;
  let duplicateCount = 0;
  const seenEmails = new Set();

  for (let i = 0; i < toSend.length; i++) {
    const contact = toSend[i];
    const statusCell = contactsSheet.getRange(contact.rowNumber, contact.statusColumn);

    if (contact.invalidEmail) {
      writeStatus(statusCell, `❌ INVALID EMAIL: ${contact.rawEmailField}`, "#ffcdd2");
      continue;
    }

    // Check duplicates
    const emailLower = contact.email.toLowerCase();

    // Cross-run duplicate: already marked sent somewhere → skip
    if (previouslySentEmails.has(emailLower)) {
      writeStatus(statusCell, `Duplicate: Previously sent`, "#ead1ff");
      continue;
    }

    // In-run duplicate (within this batch)
    if (seenEmails.has(emailLower)) {
      writeStatus(statusCell, `[SKIP] Duplicate within batch`, "#ead1ff");
      duplicateCount++;
      continue;
    }
    seenEmails.add(emailLower);

    // REPLY MODE - USING HELPER
    if (emailDraft.replyMode !== "New Email") {
      const result = _sendReplyMode({
        contact,
        emailDraft,
        contactsSheet,
        statusCell,
        toSendLength: toSend.length,
        ui,
        successCount
      });

      // Write status from result (only after we know the outcome, not before checking for threads)
      if (result.statusUpdate) {
        writeStatus(result.statusUpdate.cell, result.statusUpdate.message, result.statusUpdate.bgColor);
      }

      // Handle abort case
      if (result.abortBatch) {
        ui.alert("Error", result.error, ui.ButtonSet.OK);
        return;  // ABORT batch
      }

      // Handle success case
      if (result.success) {
        successCount += result.successIncrement;
        if (result.emailAdded) seenEmails.add(result.emailAdded);
      }
      // else: Skip case (abortBatch=false, success=false) - status already written above

      // Pacing after reply
      if (i < toSend.length - 1) Utilities.sleep(PAUSE.BETWEEN_SENDS_AFTER_SEARCH_MS);

      continue;  // Move to next contact
    }

    // NEW EMAIL MODE - USING HELPER
    writeStatus(statusCell, `⏳ SENDING ${i + 1} of ${toSend.length}... PLEASE WAIT`, "#ffeb3b");
    Utilities.sleep(PAUSE.SENDING_MS);

    let didPreflightSearch = false;

    const result = _sendNewEmailMode({
      contact,
      emailDraft,
      contactsSheet,
      statusCell,
      toSendLength: toSend.length,
      isFreshRun,
      ui,
      i,
      successCount
    });

    // Handle success case
    if (result.success) {
      successCount += result.successIncrement;
      if (result.emailAdded) seenEmails.add(result.emailAdded);
      writeStatus(result.statusUpdate.cell, result.statusUpdate.message, result.statusUpdate.bgColor);

      // Track if preflight search happened (for pacing)
      if (result.didPreflightSearch) didPreflightSearch = true;
    }

    // Pacing after send
    if (i < toSend.length - 1) {
      Utilities.sleep(didPreflightSearch ? PAUSE.BETWEEN_SENDS_AFTER_SEARCH_MS : PAUSE.BETWEEN_SENDS_DEFAULT_MS);
    }
  }

  const duplicateText = duplicateCount > 0 ? ` | Duplicates found: ${duplicateCount}` : '';
  ui.alert("Complete", `${successCount} emails sent successfully!${duplicateText}`, ui.ButtonSet.OK);
}

function previewEmail() {
  const ui = SpreadsheetApp.getUi();

  const draftSheet = getSheet(CONFIG.SHEETS.EMAIL_DRAFT);
  if (!draftSheet) {
    ui.alert("Error", "Email Draft sheet not found", ui.ButtonSet.OK);
    return;
  }

  // Check Reply Mode (B9)
  const replyMode = draftSheet.getRange("B9").getValue().toString().trim();

  // ========== NEW EMAIL MODE (or default) ==========
  if (replyMode === "New Email" || !replyMode) {
    _previewNewEmailMode({draftSheet, ui, sampleContact: CONFIG.SAMPLE});
    return;
  }

  // ========== REPLY MODE ==========
  _previewReplyMode({draftSheet, ui, replyMode});
}

/**
 * Shows preview of new email with personalization and validation summary
 */
function _previewNewEmailMode({draftSheet, ui, sampleContact}) {
  const emailDraft = getEmailDraft(draftSheet);
  if (!emailDraft) {
    ui.alert("Error", "Gmail draft not found. Check subject in B1 and ensure it matches your Gmail draft exactly.", ui.ButtonSet.OK);
    return;
  }

  const previewSubject = personalizeText(emailDraft.subject, sampleContact.name, sampleContact.lastName, sampleContact);
  const previewBody = personalizeText(emailDraft.body, sampleContact.name, sampleContact.lastName, sampleContact);

  // Fast basic HTML rendering - prioritizes speed over complex formatting
  let cleanBody = previewBody
    // Essential line breaks (most important for readability)
    .replace(/<\/(p|div|br)>/gi, '\n')
    .replace(/<\/(h[1-6]|li)>/gi, '\n')
    // Remove all remaining HTML tags (fastest approach)
    .replace(/<[^>]*>/g, '')
    // Essential entity decoding only
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&#39;/g, "'")
    .replace(/&quot;/g, '"')
    // Quick URL formatting (most visual impact)
    .replace(/(https?:\/\/[^\s]+)/gi, '\n🔗 $1\n')
    // Simple cleanup (single pass)
    .replace(/\n{3,}/g, '\n\n')
    .replace(/^\s+|\s+$/g, '');

  const preview = `📧 To: ${sampleContact.email}\n📋 Subject: ${previewSubject}\n\n${cleanBody}`;

  const validation = validateEmailTemplate(emailDraft.subject, emailDraft.body, sampleContact);
  const validationSummary = validation.isValid ? "\n\n✅ Template validated" :
    `\n\n⚠️ Issues: ${validation.hardErrors.concat(validation.softWarnings).slice(0, 2).join('; ')}${validation.hardErrors.length + validation.softWarnings.length > 2 ? '...' : ''}`;

  ui.alert("Email Preview", preview + validationSummary, ui.ButtonSet.OK);
}

/**
 * Shows preview of thread that will be replied to
 */
function _previewReplyMode({draftSheet, ui, replyMode}) {
  const searchSubject = draftSheet.getRange("B11").getValue().toString().trim();
  if (!searchSubject) {
    ui.alert("Error", "Enter search subject in B11 (the subject to find in email threads)", ui.ButtonSet.OK);
    return;
  }

  const searchEmail = draftSheet.getRange("B12").getValue().toString().trim();

  // If B12 blank, show preview of what will happen at send time (searches each contact individually)
  if (!searchEmail) {
    const emailDraft = getEmailDraft(draftSheet);
    if (!emailDraft) {
      ui.alert("Error", "Gmail draft not found. Check subject in B1.", ui.ButtonSet.OK);
      return;
    }

    const displaySubject = searchSubject.substring(0, 47) + (searchSubject.length > 47 ? "..." : "");
    const preview = `✅ Will reply to matching threads\n\nDraft: ${emailDraft.subject}\n\nSearch: "${displaySubject}"\n\nThis will search each contact's threads individually.`;
    ui.alert("Reply Preview", preview, ui.ButtonSet.OK);
    return;
  }

  const threads = findThreadForReply(searchSubject, replyMode.includes("TO") ? "TO" : "BCC", searchEmail);

  if (threads.length === 0) {
    ui.alert("Preview", "⚠️ No thread found (0 results)\n\nEnsure B11 has the correct subject to search for", ui.ButtonSet.OK);
  } else if (threads.length === 1) {
    const threadSubject = threads[0].getFirstMessageSubject();
    const preview = `✅ Will reply to this thread\n\nSubject: ${threadSubject}\n\nMode: ${replyMode}\nRecipient: ${searchEmail || "(all recipients)"}`;
    ui.alert("Thread Preview", preview, ui.ButtonSet.OK);
  } else {
    ui.alert("Preview", `⚠️ Multiple threads found (${threads.length} results)\n\nBe more specific in B11 to identify the correct thread`, ui.ButtonSet.OK);
  }
}

function sendPreviewTest() {
  const ui = SpreadsheetApp.getUi();

  const draftSheet = getSheet(CONFIG.SHEETS.EMAIL_DRAFT);
  if (!draftSheet) {
    ui.alert("Error", "Email Draft sheet not found", ui.ButtonSet.OK);
    return;
  }

  // Check Reply Mode (B9)
  const replyMode = draftSheet.getRange("B9").getValue().toString().trim();

  // ========== NEW EMAIL MODE (or default) ==========
  if (replyMode === "New Email" || !replyMode) {
    _sendTestNewEmailMode({draftSheet, ui});
    return;
  }

  // ========== REPLY MODE - WITH WARNING ==========
  _sendTestReplyMode({draftSheet, ui, replyMode});
}

/**
 * Sends test email to B2 address with full personalization
 */
function _sendTestNewEmailMode({draftSheet, ui}) {
  const testEmail = draftSheet.getRange("B2").getValue();
  const validTestEmail = extractAndValidateEmail(testEmail);
  if (!validTestEmail || testEmail.toString().trim() === "your-email@example.com") {
    ui.alert("Error", "Enter valid email in B2", ui.ButtonSet.OK);
    return;
  }

  const emailDraft = getEmailDraft(draftSheet);
  if (!emailDraft) {
    ui.alert("Error", "Gmail draft not found. Check subject in B1 and ensure it matches your Gmail draft exactly.", ui.ButtonSet.OK);
    return;
  }

  if (ui.alert("Confirm", `Send test to: ${validTestEmail}?`, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
    return;
  }

  try {
    const sampleContact = {
      name: CONFIG.SAMPLE.NAME,
      lastName: CONFIG.SAMPLE.LAST_NAME,
      email: CONFIG.SAMPLE.EMAIL,
      company: CONFIG.SAMPLE.COMPANY,
      title: CONFIG.SAMPLE.TITLE,
      custom1: CONFIG.SAMPLE.CUSTOM1,
      custom2: CONFIG.SAMPLE.CUSTOM2
    };

    const personalizedSubject = personalizeText(emailDraft.subject, CONFIG.SAMPLE.NAME, CONFIG.SAMPLE.LAST_NAME, sampleContact);
    const personalizedBody = personalizeText(emailDraft.body, CONFIG.SAMPLE.NAME, CONFIG.SAMPLE.LAST_NAME, sampleContact);

    const emailOptions = { htmlBody: personalizedBody };
    if (emailDraft.senderName) {
      emailOptions.name = emailDraft.senderName;
    }
    if (emailDraft.attachments && emailDraft.attachments.length > 0) {
      emailOptions.attachments = emailDraft.attachments;
    }

    GmailApp.sendEmail(validTestEmail, personalizedSubject, "", emailOptions);

    ui.alert("Success", `Test sent to: ${validTestEmail}`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert("Error", `Test failed: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Sends test reply to thread specified in B11 with warning dialog
 */
function _sendTestReplyMode({draftSheet, ui, replyMode}) {
  const warning = "⚠️ WARNING - Reply Mode Test\n\n" +
    "This will send an ACTUAL REPLY to a real Gmail thread!\n\n" +
    "To safely test:\n" +
    "1. Email yourself with the test subject\n" +
    "2. Put that subject in B11\n" +
    "3. Put your test email in B2\n\n" +
    "Continue?";

  if (ui.alert("Warning", warning, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
    return;
  }

  const testEmail = draftSheet.getRange("B2").getValue();
  const validTestEmail = extractAndValidateEmail(testEmail);
  if (!validTestEmail || testEmail.toString().trim() === "your-email@example.com") {
    ui.alert("Error", "Enter valid email in B2", ui.ButtonSet.OK);
    return;
  }

  const searchSubject = draftSheet.getRange("B11").getValue().toString().trim();
  if (!searchSubject) {
    ui.alert("Error", "Enter search subject in B11 (the subject to find in email threads)", ui.ButtonSet.OK);
    return;
  }

  try {
    const threads = findThreadForReply(searchSubject, replyMode.includes("TO") ? "TO" : "BCC", validTestEmail);

    if (threads.length === 0) {
      ui.alert("Error", "❌ No thread found\n\nCheck B11 subject matches your test thread", ui.ButtonSet.OK);
    } else if (threads.length === 1) {
      threads[0].reply("Test reply - checking Reply Mode functionality");
      ui.alert("Success", "✅ Test reply sent to thread", ui.ButtonSet.OK);
    } else {
      ui.alert("Error", `❌ Multiple threads found (${threads.length})\n\nBe more specific in B11 to identify the correct thread`, ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert("Error", `Reply failed: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * ============================================================================
 * EXTRACTED HELPER: _sendNewEmailMode
 * ============================================================================
 * Extracted from sendEmails() lines 712-810 for code deduplication
 *
 * NEW EMAIL MODE: Send new emails to single contact
 * Handles: preflight search, personalization, sending, post-error verification
 *
 * @param {Object} params - Parameters
 * @param {Object} params.contact - Contact object
 * @param {Object} params.emailDraft - Email draft object
 * @param {Sheet} params.contactsSheet - Contacts sheet
 * @param {Range} params.statusCell - Status cell to write to
 * @param {number} params.toSendLength - Total contacts to send (for progress counter)
 * @param {boolean} params.isFreshRun - Is this a fresh run (affects preflight)
 * @param {Object} params.ui - SpreadsheetApp UI object
 *
 * @returns {Object} {
 *   success: boolean,
 *   successIncrement: number (0 or 1),
 *   emailAdded: string or null (email to add to seenEmails),
 *   statusUpdate: {cell, message, bgColor},
 *   error: string or null
 * }
 */
function _sendNewEmailMode({contact, emailDraft, contactsSheet, statusCell, toSendLength, isFreshRun, ui, i, successCount}) {
  let didPreflightSearch = false;
  let lastPersonalizedSubject = null;

  try {
    const personalizedSubject = personalizeText(emailDraft.subject, contact.name, contact.lastName, contact);
    const personalizedBody = personalizeText(emailDraft.body, contact.name, contact.lastName, contact);
    lastPersonalizedSubject = personalizedSubject;

    // Preflight (resume-only) for blank/SENDING/FAILED
    if (!isFreshRun && shouldPreflight(contact.status)) {
      try {
        const safeSubj = String(personalizedSubject).replace(/["""]/g, "");
        const query = `in:sent from:me to:${contact.email} subject:"${safeSubj}" newer_than:3d`;
        didPreflightSearch = true;
        const threads = GmailApp.search(query);
        if (threads && threads.length > 0) {
          const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd hh:mm:ss z");
          const multiTag = contact.multiEmail ? " | MULTI-EMAIL CELL: used first; others skipped" : "";
          const successBgPreflight = contact.multiEmail ? "#ffe0b2" : "#ead1ff";
          return {
            success: true,
            successIncrement: 1,
            emailAdded: contact.email.toLowerCase(),
            statusUpdate: {
              cell: statusCell,
              message: `✅ SENT SUCCESSFULLY (${successCount + 1}/${toSendLength}) on ${dateStr} (verified prior send)${multiTag}`,
              bgColor: successBgPreflight
            },
            didPreflightSearch: true
          };
        }
      } catch (e) {
        // search failed, continue with send
      }
    }

    // Build recipient list
    let toRecipients = contact.email;
    if (emailDraft.additionalTo) {
      toRecipients += "," + emailDraft.additionalTo;
    }

    // Prepare email options
    const emailOptions = { htmlBody: personalizedBody };
    if (emailDraft.senderName) emailOptions.name = emailDraft.senderName;
    if (emailDraft.attachments && emailDraft.attachments.length > 0) emailOptions.attachments = emailDraft.attachments;
    if (emailDraft.cc) emailOptions.cc = emailDraft.cc;
    if (emailDraft.bcc) emailOptions.bcc = emailDraft.bcc;

    // Send the email
    GmailApp.sendEmail(toRecipients, personalizedSubject, "", emailOptions);

    // Success
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd hh:mm:ss z");
    const multiTag = contact.multiEmail ? " | MULTI-EMAIL CELL: used first; others skipped" : "";
    const successBg = contact.multiEmail ? "#ffe0b2" : "#d5f4e6";

    return {
      success: true,
      successIncrement: 1,
      emailAdded: contact.email.toLowerCase(),
      statusUpdate: {
        cell: statusCell,
        message: `✅ SENT SUCCESSFULLY (${successCount + 1}/${toSendLength}) on ${dateStr}${multiTag}`,
        bgColor: successBg
      },
      didPreflightSearch: didPreflightSearch
    };

  } catch (error) {
    // Verify if Gmail actually sent before marking failure
    let verified = false;
    try {
      if (lastPersonalizedSubject) {
        const safeSubj = String(lastPersonalizedSubject).replace(/["""]/g, "");
        const query = `in:sent from:me to:${contact.email} subject:"${safeSubj}" newer_than:3d`;
        const threads = GmailApp.search(query);
        if (threads && threads.length > 0) {
          const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd hh:mm:ss z");
          const multiTag = contact.multiEmail ? " | MULTI-EMAIL CELL: used first; others skipped" : "";
          const successBgPostError = contact.multiEmail ? "#ffe0b2" : "#ead1ff";
          return {
            success: true,
            successIncrement: 1,
            emailAdded: contact.email.toLowerCase(),
            statusUpdate: {
              cell: statusCell,
              message: `✅ SENT SUCCESSFULLY (${successCount + 1}/${toSendLength}) on ${dateStr} (verified after error)${multiTag}`,
              bgColor: successBgPostError
            },
            verified: true
          };
        }
      }
    } catch (e) {
      // ignore search failures
    }

    // Return failure
    return {
      success: false,
      successIncrement: 0,
      emailAdded: null,
      statusUpdate: {
        cell: statusCell,
        message: `❌ FAILED: ${error.message}`,
        bgColor: "#ffcdd2"
      },
      error: error.message
    };
  }
}

/**
 * ============================================================================
 * EXTRACTED HELPER: _sendReplyMode
 * ============================================================================
 * Extracted from sendEmails() lines 662-710 for code deduplication
 *
 * REPLY MODE: Send reply to existing email thread
 * Handles: thread search, validation, personalization, reply, error handling
 *
 * @param {Object} params - Parameters
 * @param {Object} params.contact - Contact object
 * @param {Object} params.emailDraft - Email draft object
 * @param {Sheet} params.contactsSheet - Contacts sheet
 * @param {Range} params.statusCell - Status cell to write to
 * @param {number} params.toSendLength - Total contacts to send (for progress counter)
 * @param {Object} params.ui - SpreadsheetApp UI object
 *
 * @returns {Object} {
 *   success: boolean,
 *   successIncrement: number (0 or 1),
 *   emailAdded: string or null (email to add to seenEmails),
 *   statusUpdate: {cell, message, bgColor},
 *   abortBatch: boolean (abort entire batch if true),
 *   error: string or null
 * }
 */
function _sendReplyMode({contact, emailDraft, contactsSheet, statusCell, toSendLength, ui, successCount}) {
  try {
    // Search for thread to reply to
    // Use contact email if search email (B12) is blank - allows replying to each contact's thread
    const searchEmail = emailDraft.originalTo || contact.email;

    const threads = findThreadForReply(emailDraft.originalSubject,
      emailDraft.replyMode.includes("TO") ? "TO" : "BCC",
      searchEmail);

    // Handle results
    if (threads.length === 0) {
      return {
        success: false,
        successIncrement: 0,
        emailAdded: null,
        statusUpdate: {
          cell: statusCell,
          message: `⚠️ No thread found (0 results)`,
          bgColor: "#ffeb3b"
        },
        abortBatch: true,
        error: `No threads found matching "${emailDraft.originalSubject}".\n\nCheck B11 value or try different subject.`
      };
    }

    if (threads.length > 1) {
      return {
        success: false,
        successIncrement: 0,
        emailAdded: null,
        statusUpdate: {
          cell: statusCell,
          message: `⚠️ Multiple threads (${threads.length}) - SKIPPED`,
          bgColor: "#ffeb3b"
        },
        abortBatch: false  // Skip this contact, continue batch
      };
    }

    // 1 thread found - reply to it
    const personalizedBody = personalizeText(emailDraft.body, contact.name, contact.lastName, contact);
    const replyOptions = { htmlBody: personalizedBody };

    // Add CC/BCC if requested
    if (emailDraft.includeRecipients) {
      const msg = threads[0].getMessages()[0];
      if (msg.getCc()) replyOptions.cc = msg.getCc();
      if (msg.getBcc()) replyOptions.bcc = msg.getBcc();
    }

    threads[0].reply("", replyOptions);

    // Success
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd hh:mm:ss z");
    const multiTag = contact.multiEmail ? " | MULTI-EMAIL CELL: used first; others skipped" : "";

    return {
      success: true,
      successIncrement: 1,
      emailAdded: contact.email.toLowerCase(),
      statusUpdate: {
        cell: statusCell,
        message: `✅ REPLY SENT (${successCount + 1}/${toSendLength}) on ${dateStr}${multiTag}`,
        bgColor: "#c8e6c9"
      },
      abortBatch: false
    };

  } catch (error) {
    return {
      success: false,
      successIncrement: 0,
      emailAdded: null,
      statusUpdate: {
        cell: statusCell,
        message: `❌ REPLY FAILED: ${error.message}`,
        bgColor: "#ffcdd2"
      },
      abortBatch: true,
      error: `Reply failed for ${contact.email}: ${error.message}`
    };
  }
}

function testScriptSimple() {
  const ui = SpreadsheetApp.getUi();

  const emailResponse = ui.prompt("Test", "Enter your email:", ui.ButtonSet.OK_CANCEL);
  if (emailResponse.getSelectedButton() !== ui.Button.OK) return;

  const testEmailInput = emailResponse.getResponseText().trim();
  const validTestEmail = extractAndValidateEmail(testEmailInput);
  if (!validTestEmail) {
    ui.alert("Error", "Invalid email format", ui.ButtonSet.OK);
    return;
  }

  if (ui.alert("Confirm", `Send test to: ${validTestEmail}?`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  try {
    GmailApp.sendEmail(validTestEmail, "Mail Merge Test", "Gmail integration working!");
    ui.alert("Success", `Test sent to: ${validTestEmail}`, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert("Error", `Test failed: ${error.message}`, ui.ButtonSet.OK);
  }
}

function resetContactsSheet() {
  const ui = SpreadsheetApp.getUi();

  if (ui.alert("Confirm", "Clear all contacts?", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const contactsSheet = getSheet(CONFIG.SHEETS.CONTACTS);
  if (!contactsSheet) {
    ui.alert("Error", "Contacts sheet not found", ui.ButtonSet.OK);
    return;
  }

  // Clear and setup enhanced format using standard columns
  contactsSheet.clear();
  const data = [
    CONFIG.REQUIRED_COLUMNS,
    [CONFIG.SAMPLE.NAME, CONFIG.SAMPLE.LAST_NAME, CONFIG.SAMPLE.EMAIL, CONFIG.SAMPLE.COMPANY, CONFIG.SAMPLE.TITLE, CONFIG.SAMPLE.CUSTOM1, CONFIG.SAMPLE.CUSTOM2, ""],
    ["Jane", "Smith", "jane.smith@example.com", "Tech Inc", "Developer", "Product", "New York", ""]
  ];

  contactsSheet.getRange(1, 1, data.length, CONFIG.REQUIRED_COLUMNS.length).setValues(data);
  contactsSheet.getRange(1, 1, 1, CONFIG.REQUIRED_COLUMNS.length).setFontWeight("bold");
  contactsSheet.setFrozenRows(1);
  contactsSheet.autoResizeColumns(1, CONFIG.REQUIRED_COLUMNS.length);

  ui.alert("Success", "Contacts cleared", ui.ButtonSet.OK);
}

function clearSentStatus() {
  const ui = SpreadsheetApp.getUi();

  if (ui.alert("Confirm", "Clear all sent statuses?", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  const contactsSheet = getSheet(CONFIG.SHEETS.CONTACTS);
  if (!contactsSheet) {
    ui.alert("Error", "Contacts sheet not found", ui.ButtonSet.OK);
    return;
  }

  // Standard 8-column format - status is always in column 8
  const lastRow = contactsSheet.getLastRow();
  const clearToRow = Math.max(lastRow, 100); // Minimum 100 rows to catch scattered data

  if (clearToRow > 1) {
    contactsSheet.getRange(2, 8, clearToRow - 1, 1).clearContent();
    ui.alert("Success", "Statuses cleared", ui.ButtonSet.OK);
  }
}

function showHelp() {
  const ui = SpreadsheetApp.getUi();

  const helpText = `MAIL MERGE HELP\n\n1. Run 'Create Merge Sheets'\n2. Create Gmail draft with {{Name}} {{Last Name}}\n3. Enter draft subject in Email Draft sheet B1\n4. Add contacts to Contacts sheet\n5. Use 'Send Emails'\n\nCheck Instructions sheet for more details.`;

  ui.alert("Help", helpText, ui.ButtonSet.OK);
}

// ============================================================================
// SHEET SETUP
// ============================================================================
function createMergeSheets() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheets = [
    { name: CONFIG.SHEETS.CONTACTS, setup: setupContactsSheet },
    { name: CONFIG.SHEETS.EMAIL_DRAFT, setup: setupEmailDraftSheet },
    { name: CONFIG.SHEETS.INSTRUCTIONS, setup: setupInstructionsSheet }
  ];

  sheets.forEach(sheetConfig => {
    let sheet = ss.getSheetByName(sheetConfig.name);

    if (sheet) {
      if (ui.alert("Exists", `Reset ${sheetConfig.name}?`, ui.ButtonSet.YES_NO) === ui.Button.YES) {
        sheet.clear();
        sheetConfig.setup(sheet);
      }
    } else {
      sheet = ss.insertSheet(sheetConfig.name);
      sheetConfig.setup(sheet);
    }
  });

  const contactsSheet = ss.getSheetByName(CONFIG.SHEETS.CONTACTS);
  if (contactsSheet) ss.setActiveSheet(contactsSheet);

  ui.alert("Complete", "Sheets created successfully!", ui.ButtonSet.OK);
}

function setupContactsSheet(sheet) {
  const data = [
    CONFIG.REQUIRED_COLUMNS,
    [CONFIG.SAMPLE.NAME, CONFIG.SAMPLE.LAST_NAME, CONFIG.SAMPLE.EMAIL, CONFIG.SAMPLE.COMPANY, CONFIG.SAMPLE.TITLE, CONFIG.SAMPLE.CUSTOM1, CONFIG.SAMPLE.CUSTOM2, ""],
    ["Jane", "Smith", "jane.smith@example.com", "Tech Inc", "Developer", "Product", "New York", ""]
  ];

  sheet.getRange(1, 1, data.length, CONFIG.REQUIRED_COLUMNS.length).setValues(data);
  sheet.getRange(1, 1, 1, CONFIG.REQUIRED_COLUMNS.length).setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, CONFIG.REQUIRED_COLUMNS.length);
}

function setupEmailDraftSheet(sheet) {
  const data = [
    ["Gmail draft subject: (New Email only)", "Enter your Gmail draft subject here"],
    ["Test email:", "your-email@example.com"],
    ["Sender name:", ""],
    ["Additional TO:", ""],
    ["CC:", ""],
    ["BCC:", ""],
    ["", ""],
    ["--- REPLY MODE (Optional - leave below as default for regular sends) ---", ""],
    ["Reply mode:", "New Email"],
    ["Include recipients: (Reply Mode only)", "YES or NO (default: NO)"],
    ["Search subject: (Reply Mode only)", "Searches sent and received emails (3-day window)"],
    ["Search email: (Reply Mode only)", "Leave as-is to reply by each contact's thread"]
  ];

  sheet.getRange(1, 1, data.length, 2).setValues(data);

  // Bold all labels in column A
  sheet.getRange("A1:A12").setFontWeight("bold");

  // Make section header (A8) italic and slightly larger
  sheet.getRange("A8").setFontStyle("italic").setFontSize(11);

  // Gray out the helper text for mode-specific fields
  sheet.getRange("A1").setFontColor("#999999");     // New Email only
  sheet.getRange("A10").setFontColor("#999999");    // Reply Mode only
  sheet.getRange("A11").setFontColor("#999999");    // Reply Mode only
  sheet.getRange("A12").setFontColor("#999999");    // Reply Mode only

  // Add light background to Reply Mode section to visually separate
  sheet.getRange("A8:B12").setBackground("#f9f9f9");

  // Add data validation for B9 (Reply mode dropdown)
  const replyModeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['New Email', 'Reply: TO Mode', 'Reply: BCC Mode'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange("B9").setDataValidation(replyModeRule);

  // Resize columns A and B with minimums
  sheet.setColumnWidth(1, Math.max(280, sheet.getColumnWidth(1)));
  sheet.setColumnWidth(2, Math.max(400, sheet.getColumnWidth(2)));
}

// Complete instructions and documentation
function setupInstructionsSheet(sheet) {
  const instructions = [
    [" GADI'S MAIL MERGE SYSTEM - COMPLETE GUIDE"],
    [""],
    ["═══════════════════════════════════════════════════════════════════════"],
    [" QUICK START (4 STEPS)"],
    ["═══════════════════════════════════════════════════════════════════════"],
    ["1. Create Gmail draft with personalization tags"],
    ["2. Enter your Gmail draft subject in 'Email Draft' sheet cell B1"],
    ["3. Add your contacts to the 'Contacts' sheet"],
    ["4. Click 'Mail Merge' → 'Send Emails'"],
    [""],
    ["═══════════════════════════════════════════════════════════════════════"],
    [" PERSONALIZATION TAGS (Use in Gmail Draft)"],
    ["═══════════════════════════════════════════════════════════════════════"],
    ["AVAILABLE TAGS:"],
    ["  • {{Name}} - First name"],
    ["  • {{Last Name}} - Last name"],
    ["  • {{Email}} - Recipient's email address"],
    ["  • {{Company}} - Company name"],
    ["  • {{Title}} - Job title"],
    ["  • {{Custom1}} - Flexible field (department, location, etc.)"],
    ["  • {{Custom2}} - Second flexible field"],
    [""],
    ["CASE INSENSITIVE: {{Name}}, {{name}}, {{NAME}} all work!"],
    ["ESCAPING: Use \\{{Name}} to show literal {{Name}} in emails"],
    [""],
    ["EXAMPLE GMAIL DRAFT:"],
    ["Subject: Welcome to {{Company}}, {{Name}}!"],
    [""],
    ["Hi {{Name}},"],
    [""],
    ["Welcome to your new role as {{Title}} at {{Company}}!"],
    ["Your {{Custom1}} team in {{Custom2}} is excited to work with you."],
    [""],
    ["Best regards,"],
    ["The Team"],
    [""],
    ["═══════════════════════════════════════════════════════════════════════"],
    [" CONTACT SHEET FORMAT & EMAIL EXAMPLES"],
    ["═══════════════════════════════════════════════════════════════════════"],
    ["STANDARD FORMAT (8 columns):"],
    ["  1. Name  2. Last Name  3. Email  4. Company"],
    ["  5. Title  6. Custom1  7. Custom2  8. Successfully Sent"],
    [""],
    ["EMAIL FORMAT EXAMPLES:"],
    ["  • Simple: john.doe@company.com"],
    ["  • With name: John Doe <john.doe@company.com>"],
    ["  • Display format: \"John Doe\" <john@company.com>"],
    [""],
    ["MULTI-EMAIL CELL HANDLING:"],
    ["If a cell contains multiple emails separated by commas or spaces:"],
    ["  • System automatically uses the FIRST valid email address"],
    ["  • Other emails in the same cell are skipped"],
    ["  • Status will show: '| MULTI-EMAIL CELL: used first; others skipped'"],
    ["  • Example: 'john@company.com, jane@company.com' → uses john@company.com"],
    [""],
    ["Both 'Create Merge Sheets' and 'Reset Contacts Sheet' use this format."],
    ["All personalization tags are available with this format."],
    [""],
    ["═══════════════════════════════════════════════════════════════════════"],
    [" TEMPLATE VALIDATION"],
    ["═══════════════════════════════════════════════════════════════════════"],
    ["Templates are validated before sending:"],
    [""],
    ["ERRORS (will stop sending):"],
    ["  • Malformed tags: {{Name, Name}}, {Name}"],
    ["  • Unknown tags: {{FirstName}}, {{Compny}}"],
    ["  • Other tag systems: <<Name>>, [Email], %Company%"],
    ["  • Missing data: {{Company}} when company field is empty"],
    [""],
    ["WARNINGS (will ask if you want to continue):"],
    ["  • Unbracketed tags: 'Name' without brackets"],
    [""],
    ["ERROR MESSAGE IMPROVEMENTS:"],
    ["Errors now include location context and text snippets:"],
    ["  • Before: 'Unknown tag {{Compny}} found'"],
    ["  • After: 'Unknown tag {{Compny}} in email body: ...{{Name}} at {{Compny}} for contact... (did you mean {{Company}}?)'"],
    [""],
    ["═══════════════════════════════════════════════════════════════════════"],
    [" CASE INSENSITIVE & ESCAPING EXAMPLES"],
    ["═══════════════════════════════════════════════════════════════════════"],
    ["CASE INSENSITIVE TAG MATCHING:"],
    ["All of these work identically:"],
    ["  • {{Name}} - Standard format"],
    ["  • {{name}} - Lowercase"],
    ["  • {{NAME}} - Uppercase"],
    ["  • {{NaMe}} - Mixed case"],
    ["  • {{last name}} - Works for multi-word tags too"],
    ["  • {{LAST NAME}} - Also works"],
    [""],
    ["ESCAPING EXAMPLES:"],
    ["To show literal brackets in your email:"],
    ["  • \\{{Name}} → displays as: {{Name}}"],
    ["  • \\{{Company}} → displays as: {{Company}}"],
    ["  • Welcome to \\{{YourCompany}} → displays as: Welcome to {{YourCompany}}"],
    [""],
    ["USE CASE: Documentation or examples in emails"],
    ["'To personalize, use \\{{Name}} in your template'"],
    ["→ displays as: 'To personalize, use {{Name}} in your template'"],
    [""],
    ["═══════════════════════════════════════════════════════════════════════"],
    [" ATTACHMENTS & ADVANCED FEATURES"],
    ["═══════════════════════════════════════════════════════════════════════"],
    ["ATTACHMENTS:"],
    ["  • Attach files to your Gmail draft"],
    ["  • All attachments automatically included in every sent email"],
    ["  • No additional setup required"],
    [""],
    ["EMAIL OPTIONS:"],
    ["  • Sender name: Your display name for outgoing emails"],
    ["  • Additional TO: Add extra recipients to every email"],
    ["  • CC: Carbon copy recipients"],
    ["  • BCC: Blind carbon copy recipients"],
    ["  • Fill these in the 'Email Draft' sheet"],
    [""],
    ["═══════════════════════════════════════════════════════════════════════"],
    [" REPLY MODE"],
    ["═══════════════════════════════════════════════════════════════════════"],
    ["Send replies to existing email threads instead of new emails."],
    ["Leave all Reply Mode fields blank/default to send regular new emails."],
    [""],
    ["CONFIGURATION (in Email Draft sheet - rows 8-12):"],
    ["  ROW 8: Section Header (--- REPLY MODE ---) - for clarity only"],
    ["  ROW 9: Reply Mode - Options: 'New Email' (default) | 'Reply: TO Mode' | 'Reply: BCC Mode'"],
    ["  ROW 10: Include Recipients - 'YES' or 'NO' (default: NO)"],
    ["  ROW 11: Search Subject - What subject to search for in Gmail threads (searches last 3 days)"],
    ["  ROW 12: Search Email - Optional: specific email to search for"],
    [""],
    ["HOW IT WORKS:"],
    ["  1. System searches for threads with subject from ROW 11 (last 3 days)"],
    ["  2. If 1 thread found: Replies using mode from ROW 9"],
    ["  3. If 0 threads found: Status shows '⚠️ No thread found' - skips contact"],
    ["  4. If 2+ threads found: Status shows '⚠️ Multiple threads' - skips (too ambiguous)"],
    [""],
    ["QUICK START:"],
    ["  1. Keep ROW 9 as 'New Email' to send regular emails (default)"],
    ["  2. To reply to threads:"],
    ["     a. Change ROW 9 to 'Reply: TO Mode' or 'Reply: BCC Mode'"],
    ["     b. Fill ROW 11 with the thread subject to search for"],
    ["     c. Optional: Fill ROW 12 if searching for specific recipient"],
    ["  3. Run Send Emails normally"],
    [""],
    ["═══════════════════════════════════════════════════════════════════════"],
    [" STATUS MESSAGES & COLORS"],
    ["═══════════════════════════════════════════════════════════════════════"],
    [""],
    ["SUCCESS STATUS MESSAGES:"],
    ["  ✅ SENT SUCCESSFULLY (X/Y) on DATE TIME — email sent successfully (green #d5f4e6)"],
    ["  ✅ SENT SUCCESSFULLY (X/Y) on DATE TIME (verified prior send) — found in Sent Mail; not resent (lavender #ead1ff)"],
    ["  ✅ SENT SUCCESSFULLY (X/Y) on DATE TIME (verified after error) — error occurred but Gmail confirms sent (lavender #ead1ff)"],
    ["  ✅ SENT SUCCESSFULLY (X/Y) on DATE TIME | MULTI-EMAIL CELL: used first; others skipped — cell had multiple emails (peach #ffe0b2)"],
    ["  ✅ REPLY SENT (X/Y) on DATE TIME — reply sent to thread successfully (mint green #c8e6c9)"],
    [""],
    ["IN-PROGRESS STATUS MESSAGES:"],
    ["  ⏳ SENDING X of Y... PLEASE WAIT — new email is being sent (yellow #ffeb3b)"],
    ["  ⏳ REPLYING X of Y... PLEASE WAIT — reply is being sent (yellow #ffeb3b)"],
    [""],
    ["ERROR STATUS MESSAGES:"],
    ["  ❌ INVALID EMAIL: [address] — email format is invalid (red #ffcdd2)"],
    ["  ❌ FAILED: [error details] — send operation failed (red #ffcdd2)"],
    ["  ❌ REPLY FAILED: [error details] — reply operation failed (red #ffcdd2)"],
    [""],
    ["DUPLICATE STATUS MESSAGES:"],
    ["  Duplicate: Previously sent — email already sent in previous run (lavender #ead1ff)"],
    ["  [SKIP] Duplicate within batch — email appears twice in current batch (lavender #ead1ff)"],
    [""],
    ["REPLY MODE STATUS MESSAGES:"],
    ["  ⚠️ No thread found (0 results) — no matching Gmail thread found; contact skipped (yellow #ffeb3b)"],
    ["  ⚠️ Multiple threads (N) - SKIPPED — multiple threads found; ambiguous; contact skipped (yellow #ffeb3b)"],
    [""],
    ["COMPLETE COLOR REFERENCE:"],
    ["  • #d5f4e6 - Green (normal success)"],
    ["  • #ffe0b2 - Peach (multi-email cell success)"],
    ["  • #c8e6c9 - Mint Green (reply success)"],
    ["  • #ead1ff - Lavender (Gmail-verified or duplicates)"],
    ["  • #ffeb3b - Yellow (in-progress or warnings)"],
    ["  • #ffcdd2 - Red (errors/failures)"],
    [""],
    ["═══════════════════════════════════════════════════════════════════════"],
    [" DIALOG BOXES & ERROR MESSAGES"],
    ["═══════════════════════════════════════════════════════════════════════"],
    [""],
    ["CONFIRMATION DIALOGS:"],
    ["  • \"Send X emails?\" - Confirmation before batch send (shows mode info for Reply Mode)"],
    ["  • \"Send test to: [email]?\" - Confirmation for test email send"],
    ["  • \"Clear all contacts?\" - Confirmation to reset Contacts sheet"],
    ["  • \"Clear all sent statuses?\" - Confirmation to clear status column"],
    ["  • \"Reset [Sheet Name]?\" - Confirmation when sheet already exists"],
    [""],
    ["TEMPLATE VALIDATION DIALOGS:"],
    ["  • \"Template Errors\" - Shows malformed tags, unknown tags, missing data, other tag systems"],
    ["    Examples: Unknown tag {{Compny}} in email body (did you mean {{Company}}?)"],
    ["  • \"Template Warnings\" - Shows unbracketed tags; user can continue or abort"],
    [""],
    ["ERROR DIALOGS:"],
    ["  • \"Required sheets not found. Run 'Create Merge Sheets' first.\""],
    ["  • \"No valid contacts found\""],
    ["  • \"Gmail draft not found. Check subject in B1...\""],
    ["  • \"Email Draft sheet not found\""],
    ["  • \"Contacts sheet not found\""],
    ["  • \"Enter valid email in B2\""],
    ["  • \"Enter search subject in B11 (the subject to find in email threads)\""],
    ["  • \"Invalid email format\""],
    ["  • \"Failed at [email]: [error message]\""],
    ["  • \"Reply failed: [error message]\""],
    ["  • \"Test failed: [error message]\""],
    [""],
    ["SUCCESS/COMPLETION DIALOGS:"],
    ["  • \"All emails already sent!\" - All contacts have status"],
    ["  • \"X emails sent successfully! | Duplicates found: X\" - Batch complete"],
    ["  • \"Sheets created successfully!\" - Sheet creation complete"],
    ["  • \"Test sent to: [email]\" - Test email sent"],
    ["  • \"Contacts cleared\" - Contacts reset"],
    ["  • \"Statuses cleared\" - Status column cleared"],
    [""],
    ["PREVIEW DIALOGS:"],
    ["  • \"Email Preview\" - Shows personalized preview with validation summary"],
    ["  • \"Thread Preview\" - Shows thread that will be replied to"],
    ["  • \"No thread found (0 results)...\" - No matching thread for reply preview"],
    ["  • \"Multiple threads found (X results)...\" - Too many threads for reply preview"],
    [""],
    ["SPECIAL DIALOGS:"],
    ["  • \"WARNING - Reply Mode Test\" - Warning that test will send actual reply to Gmail thread"],
    ["  • \"MAIL MERGE HELP\" - Help information (see Help menu)"],
    [""],
    ["═══════════════════════════════════════════════════════════════════════"],
    [""],
    ["════════════════════════════════════════════════════════════════════════"],
    ["Version: 3.3 | Author: Gadi Evron | Updated: 2025-11-18"],
    [""],
  ];

  sheet.getRange(1, 1, instructions.length, 1).setValues(instructions);
  sheet.getRange("A1").setFontWeight("bold").setFontSize(12);
  sheet.autoResizeColumn(1);
  sheet.setColumnWidth(1, Math.max(600, sheet.getColumnWidth(1)));
}

// ============================================================================
// MENU
// ============================================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Mail Merge")
    .addItem("Create Merge Sheets", "createMergeSheets")
    .addItem("Send Emails", "sendEmails")
    .addSeparator()
    .addItem("Popup Email Preview", "previewEmail")
    .addItem("Send Preview Email", "sendPreviewTest")
    .addSeparator()
    .addItem("Test Script", "testScriptSimple")
    .addSeparator()
    .addItem("Reset Contacts Sheet", "resetContactsSheet")
    .addItem("Clear Sent Status", "clearSentStatus")
    .addItem("Help", "showHelp")
    .addToUi();
}
