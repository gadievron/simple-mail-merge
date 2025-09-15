/**
 * ============================================================================
 * Gadi's Simple Mail Merge
 * ============================================================================
 * 
 * Gmail mail merge for Google Sheets.
 * 
 * @author Gadi Evron (with Claude, and some help from ChatGPT)
 * @version 2.7
 * @updated 2025-09-16
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
  
  if (ui.alert("Confirm", `Send ${toSend.length} emails?\n\nWatch the Status column for live progress...`, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
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
    
    // Show progress
    writeStatus(statusCell, `⏳ SENDING ${i + 1} of ${toSend.length}... PLEASE WAIT`, "#ffeb3b");
    Utilities.sleep(PAUSE.SENDING_MS);
    
    let didPreflightSearch = false; // track per-row
    let lastPersonalizedSubject = null; // for post-error verification
    
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
            successCount++;
            const multiTag = contact.multiEmail ? " | MULTI-EMAIL CELL: used first; others skipped" : "";
            // Gmail verified success → peach if multi-email else lavender
            const successBgPreflight = contact.multiEmail ? "#ffe0b2" : "#ead1ff";
            writeStatus(statusCell, `✅ SENT SUCCESSFULLY (${successCount}/${toSend.length}) on ${dateStr} (verified prior send)${multiTag}`, successBgPreflight);
            // pace: search (~120ms) + 180ms sleep ≈ 300ms total
            if (i < toSend.length - 1) Utilities.sleep(PAUSE.BETWEEN_SENDS_AFTER_SEARCH_MS);
            continue;
          }
        } catch (e) {
          // if search fails, proceed with send
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
      
      // Success status with stable date
      successCount++;
      const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd hh:mm:ss z");
      const multiTag = contact.multiEmail ? " | MULTI-EMAIL CELL: used first; others skipped" : "";
      // Normal success: multi-email gets PEACH (#ffe0b2), normal success stays GREEN
      const successBg = contact.multiEmail ? "#ffe0b2" : "#d5f4e6";
      writeStatus(statusCell, `✅ SENT SUCCESSFULLY (${successCount}/${toSend.length}) on ${dateStr}${multiTag}`, successBg);
      
      // pace: if we searched preflight, only sleep 180; else 300
      if (i < toSend.length - 1) {
        Utilities.sleep(didPreflightSearch ? PAUSE.BETWEEN_SENDS_AFTER_SEARCH_MS
                                           : PAUSE.BETWEEN_SENDS_DEFAULT_MS);
      }
      
    } catch (error) {
      // Before marking failure, verify if Gmail actually sent it
      let verified = false;
      try {
        if (lastPersonalizedSubject) {
          const safeSubj = String(lastPersonalizedSubject).replace(/["""]/g, "");
          const query = `in:sent from:me to:${contact.email} subject:"${safeSubj}" newer_than:3d`;
          const threads = GmailApp.search(query);
          if (threads && threads.length > 0) {
            const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd hh:mm:ss z");
            successCount++;
            const multiTag = contact.multiEmail ? " | MULTI-EMAIL CELL: used first; others skipped" : "";
            // Gmail verified success → peach if multi-email else lavender
            const successBgPostError = contact.multiEmail ? "#ffe0b2" : "#ead1ff";
            writeStatus(statusCell, `✅ SENT SUCCESSFULLY (${successCount}/${toSend.length}) on ${dateStr} (verified after error)${multiTag}`, successBgPostError);
            verified = true;
          }
        }
      } catch (e) {
        // ignore search failures; fall back to failure status
      }

      if (!verified) {
        writeStatus(statusCell, `❌ FAILED: ${error.message}`, "#ffcdd2");
      }

      const duplicateText = duplicateCount > 0 ? ` | Duplicates: ${duplicateCount}` : '';
      ui.alert(verified ? "Notice" : "Error",
        verified
          ? `Send already completed earlier for ${contact.email}.\n\nSent: ${successCount}${duplicateText}`
          : `Failed at ${contact.email}: ${error.message}\n\nSent: ${successCount}${duplicateText}`,
        ui.ButtonSet.OK);
      return;
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
  
  const emailDraft = getEmailDraft(draftSheet);
  if (!emailDraft) {
    ui.alert("Error", "Gmail draft not found. Check subject in B1 and ensure it matches your Gmail draft exactly.", ui.ButtonSet.OK);
    return;
  }
  
  const sampleContact = {
    name: CONFIG.SAMPLE.NAME,
    lastName: CONFIG.SAMPLE.LAST_NAME,
    email: CONFIG.SAMPLE.EMAIL,
    company: CONFIG.SAMPLE.COMPANY,
    title: CONFIG.SAMPLE.TITLE,
    custom1: CONFIG.SAMPLE.CUSTOM1,
    custom2: CONFIG.SAMPLE.CUSTOM2
  };
  
  const previewSubject = personalizeText(emailDraft.subject, CONFIG.SAMPLE.NAME, CONFIG.SAMPLE.LAST_NAME, sampleContact);
  const previewBody = personalizeText(emailDraft.body, CONFIG.SAMPLE.NAME, CONFIG.SAMPLE.LAST_NAME, sampleContact);
  
  // Fast basic HTML rendering - prioritizes speed over complex formatting
  let cleanBody = previewBody
    // Essential line breaks (most important for readability)
    .replace(/<\/?(p|div|br)(\s[^>]*)?>/gi, '\n')
    .replace(/<\/?(h[1-6]|li)(\s[^>]*)?>/gi, '\n')
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
  
  const preview = `📧 To: ${CONFIG.SAMPLE.EMAIL}\n📋 Subject: ${previewSubject}\n\n${cleanBody}`;
  
  const validation = validateEmailTemplate(emailDraft.subject, emailDraft.body, sampleContact);
  const validationSummary = validation.isValid ? "\n\n✅ Template validated" : 
    `\n\n⚠️ Issues: ${validation.hardErrors.concat(validation.softWarnings).slice(0, 2).join('; ')}${validation.hardErrors.length + validation.softWarnings.length > 2 ? '...' : ''}`;
  
  ui.alert("Email Preview", preview + validationSummary, ui.ButtonSet.OK);
}

function sendPreviewTest() {
  const ui = SpreadsheetApp.getUi();
  
  const draftSheet = getSheet(CONFIG.SHEETS.EMAIL_DRAFT);
  if (!draftSheet) {
    ui.alert("Error", "Email Draft sheet not found", ui.ButtonSet.OK);
    return;
  }
  
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
    ["Gmail draft subject:", "Enter your Gmail draft subject here"],
    ["Test email:", "your-email@example.com"],
    ["Sender name:", ""],
    ["Additional TO:", ""],
    ["CC:", ""],
    ["BCC:", ""]
  ];
  
  sheet.getRange(1, 1, data.length, 2).setValues(data);
  sheet.getRange("A1:A6").setFontWeight("bold");
  
  // Resize columns A and B with minimums
  sheet.setColumnWidth(1, Math.max(150, sheet.getColumnWidth(1)));
  sheet.setColumnWidth(2, Math.max(300, sheet.getColumnWidth(2)));
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
    [" STATUS LEGEND"],
    ["═══════════════════════════════════════════════════════════════════════"],
    [" SENT SUCCESSFULLY (X/Y) on YYYY-MM-DD HH:MM:SS TZ — normal success (green)"],
    [" … (verified prior send) — found in Sent Mail recently; not resent (lavender; peach if multi-email)"],
    [" … (verified after error) — send errored but Gmail shows it sent (lavender; peach if multi-email)"],
    [" … | MULTI-EMAIL CELL: used first; others skipped — success from a cell with multiple emails (peach)"],
    [" SENDING N of M… PLEASE WAIT — in progress (yellow)"],
    [" FAILED: <e> — send failed (red)"],
    ["Duplicate: Previously sent — duplicate across runs/sheet (lavender)"],
    ["[SKIP] Duplicate within batch — duplicate within this run (lavender)"],
    [""],
    ["COLORS:"],
    ["  • Success (normal): #d5f4e6 (green)"],
    ["  • Success (multi-email): #ffe0b2 (peach)"],
    ["  • Duplicates: #ead1ff (lavender)"],
    ["  • Gmail-verified success: #ead1ff (lavender) or #ffe0b2 (peach if multi-email)"],
    ["  • Sending: #ffeb3b (yellow)"],
    ["  • Failed: #ffcdd2 (red)"],
    [""],
    ["════════════════════════════════════════════════════════════════════════"],
    ["Version: 2.7.0 | Author: Gadi Evron | Updated: 2025-09-14"]
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
