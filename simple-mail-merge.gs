/**
 * ============================================================================
 * Gadi's Simple Mail Merge
 * ============================================================================
 * 
 * Gmail mail merge for Google Sheets.
 * 
 * @author Gadi Evron (with Claude, and some help from ChatGPT)
 * @version 2.6
 * @updated 2025-08-21
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

// ============================================================================
// STATE MANAGEMENT
// ============================================================================
const MailMergeState = {
  draftsCache: new Map(),
  cacheTime: null,
  
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
  
  getDraft(subject) {
    this.refreshDrafts();
    return this.draftsCache.get(subject.trim()) || null;
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
    // NEW: fallback – pick first email-like token if no angle brackets
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
    .split('{{Name}}').join(name || '')
    .split('{{Last Name}}').join(lastName || '');
  
  // Enhanced personalization if contact object provided
  if (contact) {
    result = result
      .split('{{Email}}').join(contact.email || '')
      .split('{{Company}}').join(contact.company || '')
      .split('{{Title}}').join(contact.title || '')
      .split('{{Custom1}}').join(contact.custom1 || '')
      .split('{{Custom2}}').join(contact.custom2 || '');
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

    // NEW: detect multiple emails; take first
    const emailsFound = emailFieldRaw.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/ig) || [];
    if (emailsFound.length === 0) continue;
    const primaryEmailRaw = emailsFound[0];

    const validEmail = extractAndValidateEmail(primaryEmailRaw);
    if (!validEmail) continue;
    
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
      // NEW: flag for multi-email cell
      multiEmail: emailsFound.length > 1
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
    
    // Check duplicates
    const emailLower = contact.email.toLowerCase();

    // Cross-run duplicate: if this email was already sent in any prior row, skip
    if (previouslySentEmails.has(emailLower)) {
      writeStatus(statusCell, `Duplicate: Previously sent`, "#fff3cd");
      continue;
    }

    // In-run duplicate (within this batch)
    if (seenEmails.has(emailLower)) {
      writeStatus(statusCell, `[SKIP] Duplicate within batch`, "#fff3cd"); // Light yellow
      duplicateCount++;
      continue;
    }
    seenEmails.add(emailLower);
    
    // Show progress
    writeStatus(statusCell, `⏳ SENDING ${i + 1} of ${toSend.length}... PLEASE WAIT`, "#ffeb3b");
    Utilities.sleep(300);
    
    try {
      const personalizedSubject = personalizeText(emailDraft.subject, contact.name, contact.lastName, contact);
      const personalizedBody = personalizeText(emailDraft.body, contact.name, contact.lastName, contact);

      // Preflight (only for blank/SENDING/FAILED)
      if (shouldPreflight(contact.status)) {
        try {
          const safeSubj = String(personalizedSubject).replace(/["“”]/g, "");
          const query = `from:me to:${contact.email} subject:"${safeSubj}" newer_than:3d`;
          const threads = GmailApp.search(query);
          if (threads && threads.length > 0) {
            const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
            successCount++;
            const multiTag = contact.multiEmail ? " | MULTI-EMAIL CELL: used first; others skipped" : "";
            const successBg = contact.multiEmail ? "#cfe8ff" : "#d5f4e6";
            writeStatus(statusCell, `✅ SENT SUCCESSFULLY (${successCount}/${toSend.length}) on ${dateStr} (verified prior send)${multiTag}`, successBg);
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
      const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
      const multiTag = contact.multiEmail ? " | MULTI-EMAIL CELL: used first; others skipped" : "";
      const successBg = contact.multiEmail ? "#cfe8ff" : "#d5f4e6";
      writeStatus(statusCell, `✅ SENT SUCCESSFULLY (${successCount}/${toSend.length}) on ${dateStr}${multiTag}`, successBg);
      
      if (i < toSend.length - 1) Utilities.sleep(300);
      
    } catch (error) {
      // Before marking failure, verify if Gmail actually sent it
      let verified = false;
      try {
        const safeSubj = String(personalizedSubject).replace(/["“”]/g, "");
        const query = `from:me to:${contact.email} subject:"${safeSubj}" newer_than:3d`;
        const threads = GmailApp.search(query);
        if (threads && threads.length > 0) {
          const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
          successCount++;
          const multiTag = contact.multiEmail ? " | MULTI-EMAIL CELL: used first; others skipped" : "";
          const successBg = contact.multiEmail ? "#cfe8ff" : "#d5f4e6";
          writeStatus(statusCell, `✅ SENT SUCCESSFULLY (${successCount}/${toSend.length}) on ${dateStr} (verified after error)${multiTag}`, successBg);
          verified = true;
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
    // Quick URL formatting (most visual impact)
    .replace(/(https?:\/\/[^\s]+)/gi, '\n🔗 $1\n')
    // Simple cleanup (single pass)
    .replace(/\n{3,}/g, '\n\n')
    .replace(/^\s+|\s+$/g, '');
  
  const preview = `📧 To: ${CONFIG.SAMPLE.EMAIL}\n📋 Subject: ${previewSubject}\n\n${cleanBody}`;
  
  ui.alert("Email Preview", preview, ui.ButtonSet.OK);
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
  
  const helpText = `MAIL MERGE HELP

1. Run 'Create Merge Sheets'
2. Create Gmail draft with {{Name}} {{Last Name}}
3. Enter draft subject in Email Draft sheet B1
4. Add contacts to Contacts sheet
5. Use 'Send Emails'

Check Instructions sheet for more details.`;
  
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

function setupInstructionsSheet(sheet) {
  const instructions = [
    ["🚀 GADI'S MAIL MERGE SYSTEM - COMPLETE GUIDE"],
    [""],
    ["═══════════════════════════════════════════════════════════════════════"],
    ["📋 QUICK START (4 STEPS)"],
    ["═══════════════════════════════════════════════════════════════════════"],
    ["1. Create Gmail draft with personalization tags"],
    ["2. Enter your Gmail draft subject in 'Email Draft' sheet cell B1"],
    ["3. Add your contacts to the 'Contacts' sheet"],
    ["4. Click 'Mail Merge' → 'Send Emails'"],
    [""],
    ["═══════════════════════════════════════════════════════════════════════"],
    ["🏷️ PERSONALIZATION TAGS (Use in Gmail Draft)"],
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
    ["📊 CONTACT SHEET FORMAT"],
    ["═══════════════════════════════════════════════════════════════════════"],
    ["STANDARD FORMAT (8 columns):"],
    ["  1. Name  2. Last Name  3. Email  4. Company"],
    ["  5. Title  6. Custom1  7. Custom2  8. Successfully Sent"],
    [""],
    ["Both 'Create Merge Sheets' and 'Reset Contacts Sheet' use this format."],
    ["All personalization tags are available with this format."],
    [""],
    ["═══════════════════════════════════════════════════════════════════════"],
    ["📎 ATTACHMENTS & ADVANCED FEATURES"],
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
    ["⚡ MENU FUNCTIONS"],
    ["═══════════════════════════════════════════════════════════════════════"],
    ["CREATE MERGE SHEETS: Set up sheets with standard 8-column format"],
    ["SEND EMAILS: Main function - sends personalized emails"],
    [""],
    ["POPUP EMAIL PREVIEW: See how your email will look"],
    ["SEND PREVIEW EMAIL: Send test email to yourself"],
    [""],
    ["TEST SCRIPT: Quick Gmail integration test"],
    [""],
    ["RESET CONTACTS SHEET: Clear all contacts, reset to standard format"],
    ["CLEAR SENT STATUS: Remove all status markers (allows re-sending)"],
    ["HELP: Show quick help dialog"],
    [""],
    ["════════════════════════════════════════════════════════════════════════"],
    ["Version: 2.6 | Author: Gadi Evron | Updated: 2025-08-21"]
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
