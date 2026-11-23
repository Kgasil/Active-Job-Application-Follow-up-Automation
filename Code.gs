/**
 * Active Job Application & Follow-up Automation — Gmail + Google Sheets
 *
 * HOW TO PREPARE:
 * 1) Import your contacts spreadsheet into Google Sheets.
 * 2) Keep the source tab named "Contatos RH" with columns A=Company, B=HR Email.
 * 3) Open Extensions > Apps Script and paste this file.
 * 4) Fill in the Portuguese and English resume file IDs in CONFIG.
 * 5) Run setupOutreach() once; then review the "Controle de envios" tab.
 * 6) Run runDailyOutreach() manually for the first test batch.
 * 7) Only after verifying the test, run installDailyTrigger().
 *
 * The script sends individual messages, avoids BCC, prevents duplicates, and tracks status.
 * Follow-ups are sent in the same thread after N business days only if no reply is detected.
 */

const CONFIG = {
  SOURCE_SHEET: 'Contatos RH',
  CONTROL_SHEET: 'Controle de envios',
  RESUME_PT_FILE_ID: 'YOUR_PT_RESUME_FILE_ID',
  RESUME_EN_FILE_ID: 'YOUR_EN_RESUME_FILE_ID',
  SENDER_EMAIL: 'your-email@gmail.com', // Update if your Gmail account differs.
  SENDER_NAME: 'Your Name',
  SENDER_PHONE: '(00) 00000-0000',
  SENDER_LOCATION: 'Your City - State/Country',
  LINKEDIN_URL: 'https://linkedin.com/in/your-profile/',
  GITHUB_URL: 'https://github.com/your-username',
  PT_FILE_NAME: 'Resume_PT.pdf',
  EN_FILE_NAME: 'Resume_EN.pdf',
  DAILY_LIMIT: 15,
  SEND_ENABLED: false, // Change to true only after testing and reviewing the list.
  SUBJECT: 'Professional Introduction | Data, BI, Processes & Operations',
  FOLLOWUP_BUSINESS_DAYS: 5,
  TIME_TRIGGER_HOUR: 9
};

const HEADERS = [
  'Empresa', 'Email de RH', 'Status', 'Data do primeiro envio',
  'Data do follow-up', 'ID da conversa Gmail', 'Observação', 'Idioma (PT/EN)'
];

function setupOutreach() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  if (!source) {
    throw new Error('Could not find sheet "' + CONFIG.SOURCE_SHEET + '". Check the imported tab name.');
  }

  let control = ss.getSheetByName(CONFIG.CONTROL_SHEET);
  if (control && control.getLastRow() > 1) {
    throw new Error(
      'The control sheet already contains data. Do not run setupOutreach() again to avoid wiping send history.'
    );
  }
  if (!control) control = ss.insertSheet(CONFIG.CONTROL_SHEET);
  control.clearContents();
  control.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  control.setFrozenRows(1);

  const values = source.getDataRange().getValues();
  const seen = new Set();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const company = String(values[i][0] || '').trim();
    const email = String(values[i][1] || '').trim().toLowerCase();
    if (!company || !isValidEmail_(email) || seen.has(email)) continue;
    seen.add(email);
    rows.push([company, email, 'PENDENTE', '', '', '', '', 'PT']);
  }
  if (rows.length) control.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
  control.autoResizeColumns(1, HEADERS.length);
  SpreadsheetApp.getUi().alert(
    'Control sheet created with ' + rows.length + ' unique contacts. Review the list before sending.'
  );
}

function runDailyOutreach() {
  validateConfig_();
  if (!CONFIG.SEND_ENABLED) {
    throw new Error('Batch sending is locked. Test first and set SEND_ENABLED to true.');
  }
  const today = new Date().getDay();
  if (today === 0 || today === 6) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Weekend: no emails sent today.', 'Outreach', 5);
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.CONTROL_SHEET);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('Run setupOutreach() first.');

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();

  processFollowups_(sheet, data);

  let sent = 0;
  for (let i = 0; i < data.length && sent < CONFIG.DAILY_LIMIT; i++) {
    const row = data[i];
    const company = String(row[0] || '').trim();
    const email = String(row[1] || '').trim().toLowerCase();
    const status = String(row[2] || '').trim();
    const language = normalizeLanguage_(row[7]);
    if (status !== 'PENDENTE' || !isValidEmail_(email)) continue;

    const resume = getResumeFile_(language);
    const body = buildFirstEmail_(company, language);
    const subject = language === 'PT'
      ? 'Apresentação profissional | Dados, BI, Processos e Operações'
      : CONFIG.SUBJECT;
    const attachmentName = language === 'EN' ? CONFIG.EN_FILE_NAME : CONFIG.PT_FILE_NAME;
    GmailApp.sendEmail(email, subject, body, {
      attachments: [resume.getBlob().setName(attachmentName)],
      name: CONFIG.SENDER_NAME,
      replyTo: CONFIG.SENDER_EMAIL
    });

    const sentAt = new Date();
    Utilities.sleep(1200);
    const threadId = findSentThreadId_(email, subject);
    const followupAt = addBusinessDays_(sentAt, CONFIG.FOLLOWUP_BUSINESS_DAYS);
    const sheetRow = i + 2;
    const fallbackNote = 'Sent; Gmail thread ID not found. Verify manually before follow-up.';
    sheet.getRange(sheetRow, 3, 1, 5).setValues([[
      'ENVIADO', sentAt, followupAt, threadId || '', threadId ? '' : fallbackNote
    ]]);
    sent++;
    Utilities.sleep(1200);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Batch finished: ' + sent + ' emails sent.', 'Outreach', 8);
}

function sendTestEmail() {
  validateConfig_();
  const resume = getResumeFile_('PT');
  GmailApp.sendEmail(
    CONFIG.SENDER_EMAIL,
    '[TEST] Apresentação profissional | Dados, BI, Processos e Operações',
    buildFirstEmail_('Test Company', 'PT'),
    {
      attachments: [resume.getBlob().setName(CONFIG.PT_FILE_NAME)],
      name: CONFIG.SENDER_NAME,
      replyTo: CONFIG.SENDER_EMAIL
    }
  );
  SpreadsheetApp.getUi().alert(
    'Portuguese test email sent to ' + CONFIG.SENDER_EMAIL + '. Check copy, signature, and PDF before enabling sends.'
  );
}

function sendTestEmailEnglish() {
  validateConfig_();
  const resume = getResumeFile_('EN');
  GmailApp.sendEmail(
    CONFIG.SENDER_EMAIL,
    '[TEST] ' + CONFIG.SUBJECT,
    buildFirstEmail_('Test Company', 'EN'),
    {
      attachments: [resume.getBlob().setName(CONFIG.EN_FILE_NAME)],
      name: CONFIG.SENDER_NAME,
      replyTo: CONFIG.SENDER_EMAIL
    }
  );
  SpreadsheetApp.getUi().alert(
    'English test email sent to ' + CONFIG.SENDER_EMAIL + '. Review the message and attached PDF.'
  );
}

function processFollowups_(sheet, data) {
  const now = new Date();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const status = String(row[2] || '').trim();
    const due = row[4] instanceof Date ? row[4] : null;
    const threadId = String(row[5] || '').trim();
    const email = String(row[1] || '').trim().toLowerCase();
    const company = String(row[0] || '').trim();
    const language = normalizeLanguage_(row[7]);
    if (status !== 'ENVIADO' || !due || due > now) continue;
    const sheetRow = i + 2;
    if (!threadId) {
      sheet.getRange(sheetRow, 7).setValue('Follow-up pending: Gmail thread not found; check manually.');
      continue;
    }

    const thread = GmailApp.getThreadById(threadId);
    if (!thread) {
      sheet.getRange(sheetRow, 7).setValue('Follow-up pending: thread not found; check manually.');
      continue;
    }

    const hasReply = thread.getMessages().some(message => {
      const from = extractEmail_(message.getFrom());
      return from === email && message.getDate() > row[3];
    });
    if (hasReply) {
      sheet.getRange(sheetRow, 3).setValue('RESPONDEU');
      sheet.getRange(sheetRow, 7).setValue('Reply received; automated sequence stopped.');
      continue;
    }

    thread.reply(buildFollowupEmail_(company, language));
    sheet.getRange(sheetRow, 3).setValue('FOLLOW-UP ENVIADO');
    sheet.getRange(sheetRow, 7).setValue('Follow-up sent automatically.');
  }
}

function installDailyTrigger() {
  if (!CONFIG.SEND_ENABLED) {
    throw new Error('Set SEND_ENABLED to true after reviewing the test before installing the daily trigger.');
  }
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runDailyOutreach') ScriptApp.deleteTrigger(trigger);
  });
  ScriptApp.newTrigger('runDailyOutreach')
    .timeBased()
    .everyDays(1)
    .atHour(CONFIG.TIME_TRIGGER_HOUR)
    .create();
  SpreadsheetApp.getUi().alert(
    'Daily trigger installed. The script will send up to ' + CONFIG.DAILY_LIMIT +
    ' new emails per day and process due follow-ups.'
  );
}

function removeDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runDailyOutreach') ScriptApp.deleteTrigger(trigger);
  });
  SpreadsheetApp.getUi().alert('Daily trigger removed.');
}

function buildFirstEmail_(company, language) {
  if (normalizeLanguage_(language) === 'EN') {
    return 'Dear ' + company + ' HR team,\n\n' +
      'I am reaching out to express my interest in currently open positions in Data, Business Intelligence, ' +
      'Process Improvement, Operations, Planning, or Back Office. I would also be glad to be considered for ' +
      'relevant opportunities that may arise in the future.\n\n' +
      'I am currently pursuing a degree in Production Engineering, hold a technical qualification in Industrial ' +
      'Chemistry, and have experience in corporate environments combining business operations, process knowledge, ' +
      'and technical development.\n\n' +
      'My background includes SQL, Python, Excel, and Power BI; data preparation and organization; dashboards and ' +
      'performance indicators; process mapping and improvement; cost analysis; operational support; and workflow ' +
      'automation using Google Apps Script, APIs, webhooks, n8n, and Make.\n\n' +
      'I am looking for an opportunity where I can apply my analytical and technical skills, contribute to ' +
      'process efficiency, and continue developing professionally in remote or hybrid roles.\n\n' +
      'Please find my resume attached. You can also learn more about my background and projects through my ' +
      'LinkedIn and GitHub profiles. Thank you for your time and consideration. I would be glad to discuss ' +
      'opportunities that may match my profile.\n\n' +
      'Best regards,\n' + CONFIG.SENDER_NAME + '\n' +
      CONFIG.SENDER_LOCATION + '\n' +
      CONFIG.SENDER_PHONE + '\n' +
      CONFIG.SENDER_EMAIL + '\n' +
      'LinkedIn: ' + CONFIG.LINKEDIN_URL + '\n' +
      'GitHub: ' + CONFIG.GITHUB_URL + '\n\n' +
      'If you prefer not to receive further messages, please reply to let me know.';
  }
  return 'Hello ' + company + ' HR team,\n\n' +
    'I am sharing my resume for consideration in roles across Data, Business Intelligence, Processes, ' +
    'Operations, Planning, or Back Office.\n\n' +
    'I am a Production Engineering student and Industrial Chemistry technician with corporate experience ' +
    'combining operational execution, process understanding, and technical development.\n\n' +
    'My background and experience cover:\n\n' +
    '• Data & Business Intelligence: SQL, Python, Excel, and Power BI, working on data cleaning, joining, ' +
    'organization, dashboard building, KPI tracking, and routine automation.\n\n' +
    '• Processes & Operations: workflow mapping and improvement, control standardization, information ' +
    'organization, cost analysis, operational support, and process automation.\n\n' +
    '• Automation & Tech: building solutions for task optimization and system integration using Google Apps ' +
    'Script, APIs, webhooks, and tools such as n8n and Make.\n\n' +
    'I am looking for an opportunity to apply my analytical and technical skills, contribute to process ' +
    'efficiency, and keep growing professionally in remote or hybrid roles.\n\n' +
    'I have attached my resume and also share my LinkedIn and GitHub profiles where you can view more about ' +
    'my background and projects.\n\n' +
    'Thank you for your time and I remain available to discuss roles aligned with my profile.\n\n' +
    'Best regards,\n\n' +
    CONFIG.SENDER_NAME + '\n' +
    CONFIG.SENDER_LOCATION + '\n' +
    CONFIG.SENDER_PHONE + '\n' +
    CONFIG.SENDER_EMAIL + '\n\n' +
    'LinkedIn: ' + CONFIG.LINKEDIN_URL + '\n' +
    'GitHub: ' + CONFIG.GITHUB_URL;
}

function buildFollowupEmail_(company, language) {
  if (normalizeLanguage_(language) === 'EN') {
    return 'Hello ' + company + ' HR team,\n\n' +
      'I am following up briefly to check whether my resume, sent previously, reached the appropriate team.\n\n' +
      'I remain available to discuss opportunities in Data, BI, Process Improvement, Operations, or Planning ' +
      'that may match my background.\n\n' +
      'Thank you again for your time.\n\nBest regards,\n' + CONFIG.SENDER_NAME + '\n' + CONFIG.SENDER_EMAIL;
  }
  return 'Hello ' + company + ' HR team,\n\n' +
    'I am reaching back out briefly to confirm whether my resume, sent earlier, reached the responsible team.\n\n' +
    'I remain available to discuss opportunities in Data, BI, Processes, Operations, or Planning aligned with ' +
    'my profile.\n\n' +
    'Thank you again for your attention.\n\nBest regards,\n' + CONFIG.SENDER_NAME + '\n' + CONFIG.SENDER_EMAIL;
}

function normalizeLanguage_(value) {
  return String(value || 'PT').trim().toUpperCase() === 'EN' ? 'EN' : 'PT';
}

function getResumeFile_(language) {
  const fileId = normalizeLanguage_(language) === 'EN' ? CONFIG.RESUME_EN_FILE_ID : CONFIG.RESUME_PT_FILE_ID;
  if (!fileId) throw new Error('Resume file ID is not configured for language ' + language + '.');
  const file = DriveApp.getFileById(fileId);
  if (file.getMimeType() !== MimeType.PDF) throw new Error('The selected resume file is not a PDF.');
  return file;
}

function findSentThreadId_(email, subject) {
  const query = 'in:sent to:' + email + ' subject:"' + subject.replace(/"/g, '') + '" newer_than:3d';
  const threads = GmailApp.search(query, 0, 5);
  return threads.length ? threads[0].getId() : '';
}

function addBusinessDays_(date, days) {
  const result = new Date(date);
  let added = 0;
  while (added < days) {
    result.setDate(result.getDate() + 1);
    const day = result.getDay();
    if (day !== 0 && day !== 6) added++;
  }
  return result;
}

function extractEmail_(from) {
  const match = String(from || '').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return match ? match[0].toLowerCase() : String(from || '').toLowerCase();
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function validateConfig_() {
  if (!CONFIG.RESUME_PT_FILE_ID || !CONFIG.RESUME_EN_FILE_ID) {
    throw new Error('Check the Portuguese and English resume file IDs in CONFIG.');
  }
  if (!isValidEmail_(CONFIG.SENDER_EMAIL)) throw new Error('Check SENDER_EMAIL in CONFIG.');
}
