/**
 * ============================================================================
 * Gadi's Simple Mail Merge
 * ============================================================================
 * 
 * Gmail mail merge for Google Sheets.
 * 
 * @author Gadi Evron (with Claude, and some help from ChatGPT)
 * @version 2.4.0
 * @updated 2025-01-30
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
 REQUIRED_COLUMNS: ["Name", "Last Name", "Email", "Successfully Sent"],
 SAMPLE: {
 NAME: "John",
 LAST_NAME: "Doe", 
 EMAIL: "sample@example.com"
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
function validateEmail(email) {
 if (!email) return false;
 const trimmed = email.toString().trim();
 return trimmed.length > 4 && 
 trimmed.length < 255 && 
 trimmed.includes('@') && 
 /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(trimmed);
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

// ============================================================================
// CORE FUNCTIONS
// ============================================================================
function getContacts(sheet) {
 const lastRow = sheet.getLastRow();
 if (lastRow <= 1) return [];
 
 // Simple detection: if column 4 header is "Company", it's enhanced format
 const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
 const isEnhanced = headers.length >= 8 && headers[3] === "Company";
 const statusCol = isEnhanced ? 8 : 4;
 const numCols = isEnhanced ? 8 : 4;
 
 const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
 const contacts = [];
 
 for (let i = 0; i < data.length; i++) {
 const row = data[i];
 const email = row[2] ? row[2].toString().trim() : "";
 
 if (!email || !validateEmail(email)) continue;
 
 const contact = {
 rowNumber: i + 2,
 name: (row[0] || "").toString().trim(),
 lastName: (row[1] || "").toString().trim(),
 email: email,
 status: row[statusCol - 1] || "",
 statusColumn: statusCol
 };
 
 // Add enhanced fields if available
 if (isEnhanced) {
 contact.company = (row[3] || "").toString().trim();
 contact.title = (row[4] || "").toString().trim();
 contact.custom1 = (row[5] || "").toString().trim();
 contact.custom2 = (row[6] || "").toString().trim();
 }
 
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
 if (sheet.getLastRow() >= 5) {
 const additionalTo = sheet.getRange("B3").getValue();
 const cc = sheet.getRange("B4").getValue();
 const bcc = sheet.getRange("B5").getValue();
 
 result.additionalTo = additionalTo ? additionalTo.toString().trim() : "";
 result.cc = cc ? cc.toString().trim() : "";
 result.bcc = bcc ? bcc.toString().trim() : "";
 }
 } catch (e) {
 // Fields don't exist - this is fine
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
 
 const toSend = contacts.filter(c => 
 !c.status || !c.status.toString().includes("Sent successfully")
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
 
 // Check for duplicates at runtime
 const emailLower = contact.email.toLowerCase();
 if (seenEmails.has(emailLower)) {
 const statusCell = contactsSheet.getRange(contact.rowNumber, contact.statusColumn);
 statusCell.setValue(`Duplicate: Email skipped`);
 statusCell.setBackground("#fff3cd"); // Light yellow background for duplicates
 SpreadsheetApp.flush();
 duplicateCount++;
 continue;
 }
 seenEmails.add(emailLower);
 
 // Show progress in the status column immediately with more obvious text
 const statusCell = contactsSheet.getRange(contact.rowNumber, contact.statusColumn);
 statusCell.setValue(`⏳ SENDING ${i + 1} of ${toSend.length}... PLEASE WAIT`);
 statusCell.setBackground("#ffeb3b"); // Yellow background to make it obvious
 SpreadsheetApp.flush();
 
 // Add a small delay so user can see the "sending" status
 Utilities.sleep(500);
 
 try {
 const personalizedSubject = personalizeText(emailDraft.subject, contact.name, contact.lastName, contact);
 const personalizedBody = personalizeText(emailDraft.body, contact.name, contact.lastName, contact);
 
 // Build recipient list
 let toRecipients = contact.email;
 if (emailDraft.additionalTo) {
 toRecipients += "," + emailDraft.additionalTo;
 }
 
 // Prepare email options
 const emailOptions = { htmlBody: personalizedBody };
 if (emailDraft.attachments && emailDraft.attachments.length > 0) {
 emailOptions.attachments = emailDraft.attachments;
 }
 if (emailDraft.cc) {
 emailOptions.cc = emailDraft.cc;
 }
 if (emailDraft.bcc) {
 emailOptions.bcc = emailDraft.bcc;
 }
 
 // Send the email
 GmailApp.sendEmail(toRecipients, personalizedSubject, "", emailOptions);
 
 // Immediately update with success status
 successCount++;
 statusCell.setValue(`✅ SENT SUCCESSFULLY (${successCount}/${toSend.length}) on ${new Date().toLocaleDateString()}`);
 statusCell.setBackground("#d5f4e6"); // Light green background for success
 SpreadsheetApp.flush();
 
 console.log(`✅ Sent to ${contact.email} (${successCount}/${toSend.length})`);
 
 if (i < toSend.length - 1) Utilities.sleep(300);
 
 } catch (error) {
 console.log(`❌ Failed to send to ${contact.email}: ${error.message}`);
 
 // Update failure status
 statusCell.setValue(`❌ FAILED: ${error.message}`);
 statusCell.setBackground("#ffcdd2"); // Light red background for errors
 SpreadsheetApp.flush();
 
 const duplicateText = duplicateCount > 0 ? ` | Duplicates: ${duplicateCount}` : '';
 ui.alert("Error", `Failed at ${contact.email}: ${error.message}\n\nSent: ${successCount}${duplicateText}`, ui.ButtonSet.OK);
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
 company: "Example Corp",
 title: "Manager",
 custom1: "Engineering", 
 custom2: "San Francisco"
 };
 
 const previewSubject = personalizeText(emailDraft.subject, CONFIG.SAMPLE.NAME, CONFIG.SAMPLE.LAST_NAME, sampleContact);
 const previewBody = personalizeText(emailDraft.body, CONFIG.SAMPLE.NAME, CONFIG.SAMPLE.LAST_NAME, sampleContact);
 
 // Strip HTML tags for cleaner preview
 const cleanBody = previewBody.replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');
 
 const preview = `PREVIEW\n\nTo: ${CONFIG.SAMPLE.EMAIL}\nSubject: ${previewSubject}\n\nContent:\n${cleanBody}`;
 
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
 if (!testEmail || !validateEmail(testEmail) || testEmail.toString().trim() === "your-email@example.com") {
 ui.alert("Error", "Enter valid email in B2", ui.ButtonSet.OK);
 return;
 }
 
 const emailDraft = getEmailDraft(draftSheet);
 if (!emailDraft) {
 ui.alert("Error", "Gmail draft not found. Check subject in B1 and ensure it matches your Gmail draft exactly.", ui.ButtonSet.OK);
 return;
 }
 
 if (ui.alert("Confirm", `Send test to: ${testEmail}?`, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
 return;
 }
 
 try {
 const sampleContact = {
 email: CONFIG.SAMPLE.EMAIL,
 company: "Example Corp",
 title: "Manager",
 custom1: "Engineering",
 custom2: "San Francisco"
 };
 
 const personalizedSubject = personalizeText(emailDraft.subject, CONFIG.SAMPLE.NAME, CONFIG.SAMPLE.LAST_NAME, sampleContact);
 const personalizedBody = personalizeText(emailDraft.body, CONFIG.SAMPLE.NAME, CONFIG.SAMPLE.LAST_NAME, sampleContact);
 
 const emailOptions = { htmlBody: personalizedBody };
 if (emailDraft.attachments && emailDraft.attachments.length > 0) {
 emailOptions.attachments = emailDraft.attachments;
 }
 
 GmailApp.sendEmail(testEmail, personalizedSubject, "", emailOptions);
 
 ui.alert("Success", `Test sent to: ${testEmail}`, ui.ButtonSet.OK);
 } catch (error) {
 ui.alert("Error", `Test failed: ${error.message}`, ui.ButtonSet.OK);
 }
}

function testScriptSimple() {
 const ui = SpreadsheetApp.getUi();
 
 const emailResponse = ui.prompt("Test", "Enter your email:", ui.ButtonSet.OK_CANCEL);
 if (emailResponse.getSelectedButton() !== ui.Button.OK) return;
 
 const testEmail = emailResponse.getResponseText().trim();
 if (!validateEmail(testEmail)) {
 ui.alert("Error", "Invalid email format", ui.ButtonSet.OK);
 return;
 }
 
 if (ui.alert("Confirm", `Send test to: ${testEmail}?`, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
 
 try {
 GmailApp.sendEmail(testEmail, "Mail Merge Test", "Gmail integration working!");
 ui.alert("Success", `Test sent to: ${testEmail}`, ui.ButtonSet.OK);
 } catch (error) {
 ui.alert("Error", `Test failed: ${error.message}`, ui.ButtonSet.OK);
 }
}

function resetContactsSheet() {
 const ui = SpreadsheetApp.getUi();
 
 if (ui.alert("Confirm", "Clear all contacts and reset to enhanced format?", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
 
 const contactsSheet = getSheet(CONFIG.SHEETS.CONTACTS);
 if (!contactsSheet) {
 ui.alert("Error", "Contacts sheet not found", ui.ButtonSet.OK);
 return;
 }
 
 // Clear and setup enhanced format
 contactsSheet.clear();
 const data = [
 ["Name", "Last Name", "Email", "Company", "Title", "Custom1", "Custom2", "Successfully Sent"],
 ["John", "Doe", "john.doe@example.com", "Example Corp", "Manager", "Engineering", "San Francisco", ""],
 ["Jane", "Smith", "jane.smith@example.com", "Tech Inc", "Developer", "Product", "New York", ""]
 ];
 
 contactsSheet.getRange(1, 1, data.length, 8).setValues(data);
 contactsSheet.getRange(1, 1, 1, 8).setFontWeight("bold");
 contactsSheet.setFrozenRows(1);
 contactsSheet.autoResizeColumns(1, 8);
 
 ui.alert("Success", "Contacts sheet reset to enhanced format", ui.ButtonSet.OK);
}

function clearSentStatus() {
 const ui = SpreadsheetApp.getUi();
 
 if (ui.alert("Confirm", "Clear all sent statuses?", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
 
 const contactsSheet = getSheet(CONFIG.SHEETS.CONTACTS);
 if (!contactsSheet) {
 ui.alert("Error", "Contacts sheet not found", ui.ButtonSet.OK);
 return;
 }
 
 // Simple detection: if column 4 header is "Company", status is in column 8, else column 4
 const headers = contactsSheet.getRange(1, 1, 1, contactsSheet.getLastColumn()).getValues()[0];
 const statusCol = headers.length >= 8 && headers[3] === "Company" ? 8 : 4;
 const lastRow = contactsSheet.getLastRow();
 
 if (lastRow > 1) {
 contactsSheet.getRange(2, statusCol, lastRow - 1, 1).clearContent();
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
 ["John", "Doe", "john.doe@example.com", ""],
 ["Jane", "Smith", "jane.smith@example.com", ""]
 ];
 
 sheet.getRange(1, 1, data.length, 4).setValues(data);
 sheet.getRange(1, 1, 1, 4).setFontWeight("bold");
 sheet.setFrozenRows(1);
 sheet.autoResizeColumns(1, 4);
}

function setupEmailDraftSheet(sheet) {
 const data = [
 ["Gmail draft subject:", "Enter your Gmail draft subject here"],
 ["Test email:", "your-email@example.com"],
 ["Additional TO:", ""],
 ["CC:", ""],
 ["BCC:", ""]
 ];
 
 sheet.getRange(1, 1, data.length, 2).setValues(data);
 sheet.getRange("A1:A5").setFontWeight("bold");
 
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
 ["BASIC TAGS (work in both 4-column and 8-column formats):"],
 ["  • {{Name}} - First name"],
 ["  • {{Last Name}} - Last name"],
 [""],
 ["ENHANCED TAGS (only work with 8-column contact format):"],
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
 ["📊 CONTACT SHEET FORMATS"],
 ["═══════════════════════════════════════════════════════════════════════"],
 ["BASIC FORMAT (4 columns) - Compatible with old versions:"],
 ["  1. Name  2. Last Name  3. Email  4. Successfully Sent"],
 [""],
 ["ENHANCED FORMAT (8 columns) - New features:"],
 ["  1. Name  2. Last Name  3. Email  4. Company"],
 ["  5. Title  6. Custom1  7. Custom2  8. Successfully Sent"],
 [""],
 ["The system auto-detects your format and works with both!"],
 ["Use 'Reset Contacts Sheet' to switch to enhanced format."],
 [""],
 ["═══════════════════════════════════════════════════════════════════════"],
 ["📎 ATTACHMENTS & ADVANCED FEATURES"],
 ["═══════════════════════════════════════════════════════════════════════"],
 ["ATTACHMENTS:"],
 ["  • Attach files to your Gmail draft"],
 ["  • All attachments automatically included in every sent email"],
 ["  • No additional setup required"],
 [""],
 ["MULTIPLE RECIPIENTS:"],
 ["  • Additional TO: Add extra recipients to every email"],
 ["  • CC: Carbon copy recipients"],
 ["  • BCC: Blind carbon copy recipients"],
 ["  • Fill these in the 'Email Draft' sheet"],
 [""],
 ["═══════════════════════════════════════════════════════════════════════"],
 ["⚡ MENU FUNCTIONS"],
 ["═══════════════════════════════════════════════════════════════════════"],
 ["CREATE MERGE SHEETS: Set up the three required sheets"],
 ["SEND EMAILS: Main function - sends personalized emails"],
 [""],
 ["POPUP EMAIL PREVIEW: See how your email will look"],
 ["SEND PREVIEW TEST: Send test email to yourself"],
 [""],
 ["TEST SCRIPT: Quick Gmail integration test"],
 ["RESET CONTACTS SHEET: Clear all contacts, switch to enhanced format"],
 [""],
 ["CLEAR SENT STATUS: Remove all status markers (allows re-sending)"],
 ["HELP: Show quick help dialog"],
 [""],
 ["════════════════════════════════════════════════════════════════════════"],
 ["Version: 2.4.0 | Author: Gadi Evron | Updated: 2025-01-30"]
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
 .addItem("Send Preview Test", "sendPreviewTest")
 .addSeparator()
 .addItem("Test Script", "testScriptSimple")
 .addItem("Reset Contacts Sheet", "resetContactsSheet")
 .addSeparator()
 .addItem("Clear Sent Status", "clearSentStatus")
 .addItem("Help", "showHelp")
 .addToUi();
}
