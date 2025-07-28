/**
 * ============================================================================
 * simple-mail-merge
 * ============================================================================
 * 
 * A simple Gmail mail merge script for Google Sheets.
 * 
 * @author Gadi Evron (with Claude and some ChatGPT)
 * @version 2.3.0
 * @updated 2025-01-28
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

function personalizeText(text, name, lastName) {
 if (!text) return '';
 return text.toString()
 .split('{{Name}}').join(name || '')
 .split('{{Last Name}}').join(lastName || '');
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
 
 const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
 const contacts = [];
 const seenEmails = new Set();
 
 for (let i = 0; i < data.length; i++) {
 const row = data[i];
 const email = row[2] ? row[2].toString().trim() : "";
 
 if (!email || !validateEmail(email)) continue;
 
 const emailLower = email.toLowerCase();
 if (seenEmails.has(emailLower)) continue;
 seenEmails.add(emailLower);
 
 contacts.push({
 rowNumber: i + 2,
 name: (row[0] || "").toString().trim(),
 lastName: (row[1] || "").toString().trim(),
 email: email,
 status: row[3] || ""
 });
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
 
 return {
 subject: subjectStr,
 body: draft.getMessage().getBody()
 };
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
 
 if (ui.alert("Confirm", `Send ${toSend.length} emails?`, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
 return;
 }
 
 let successCount = 0;
 
 for (let i = 0; i < toSend.length; i++) {
 const contact = toSend[i];
 
 try {
 const personalizedSubject = personalizeText(emailDraft.subject, contact.name, contact.lastName);
 const personalizedBody = personalizeText(emailDraft.body, contact.name, contact.lastName);
 
 GmailApp.sendEmail(contact.email, personalizedSubject, "", { htmlBody: personalizedBody });
 
 contactsSheet.getRange(contact.rowNumber, 4).setValue(`Sent successfully on ${new Date().toLocaleDateString()}`);
 successCount++;
 
 if (i < toSend.length - 1) Utilities.sleep(300);
 
 } catch (error) {
 contactsSheet.getRange(contact.rowNumber, 4).setValue(`FAILED: ${error.message}`);
 ui.alert("Error", `Failed at ${contact.email}: ${error.message}\n\nSent: ${successCount}`, ui.ButtonSet.OK);
 return;
 }
 }
 
 ui.alert("Complete", `${successCount} emails sent successfully!`, ui.ButtonSet.OK);
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
 
 const previewSubject = personalizeText(emailDraft.subject, CONFIG.SAMPLE.NAME, CONFIG.SAMPLE.LAST_NAME);
 const previewBody = personalizeText(emailDraft.body, CONFIG.SAMPLE.NAME, CONFIG.SAMPLE.LAST_NAME);
 
 const preview = PREVIEW\n\nTo: ${CONFIG.SAMPLE.EMAIL}\nSubject: ${previewSubject}\n\nContent:\n${previewBody};
 
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
 const personalizedSubject = personalizeText(emailDraft.subject, CONFIG.SAMPLE.NAME, CONFIG.SAMPLE.LAST_NAME);
 const personalizedBody = personalizeText(emailDraft.body, CONFIG.SAMPLE.NAME, CONFIG.SAMPLE.LAST_NAME);
 
 GmailApp.sendEmail(testEmail, personalizedSubject, "", { htmlBody: personalizedBody });
 
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

function clearSentStatus() {
 const ui = SpreadsheetApp.getUi();
 
 if (ui.alert("Confirm", "Clear all sent statuses?", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
 
 const contactsSheet = getSheet(CONFIG.SHEETS.CONTACTS);
 if (!contactsSheet) {
 ui.alert("Error", "Contacts sheet not found", ui.ButtonSet.OK);
 return;
 }
 
 const lastRow = contactsSheet.getLastRow();
 if (lastRow > 1) {
 contactsSheet.getRange(2, 4, lastRow - 1, 1).clearContent();
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
}

function setupEmailDraftSheet(sheet) {
 const data = [
 ["Gmail draft subject:", "Enter your Gmail draft subject here"],
 ["Test email:", "your-email@example.com"]
 ];
 
 sheet.getRange(1, 1, data.length, 2).setValues(data);
 sheet.getRange("A1:A2").setFontWeight("bold");
}

function setupInstructionsSheet(sheet) {
 const instructions = [
 ["MAIL MERGE INSTRUCTIONS"],
 [""],
 ["SETUP:"],
 ["1. Create Gmail draft with {{Name}} {{Last Name}}"],
 ["2. Enter draft subject in Email Draft sheet B1"],
 ["3. Add contacts to Contacts sheet"],
 ["4. Use Send Emails"],
 [""],
 ["FEATURES:"],
 ["Stops on first failure"],
 ["Skips already sent emails"],
 ["Simple validation"]
 ];
 
 sheet.getRange(1, 1, instructions.length, 1).setValues(instructions);
 sheet.getRange("A1").setFontWeight("bold");
 sheet.getRange("A3").setFontWeight("bold");
 sheet.getRange("A9").setFontWeight("bold");
}

// ============================================================================
// MENU
// ============================================================================
function onOpen() {
 SpreadsheetApp.getUi().createMenu("Mail Merge")
 .addItem("Create Merge Sheets", "createMergeSheets")
 .addItem("Send Emails", "sendEmails")
 .addSeparator()
 .addItem("Preview Email", "previewEmail")
 .addItem("Send Preview Test", "sendPreviewTest")
 .addSeparator()
 .addItem("Test Script", "testScriptSimple")
 .addSeparator()
 .addItem("Clear Sent Status", "clearSentStatus")
 .addItem("Help", "showHelp")
 .addToUi();
}
