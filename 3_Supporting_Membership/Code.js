// =========================================================================
// OPEN COMMUNITY RADIO STARTER KIT
// Module 4: Supporting Membership & SEPA Workflow
//
// Paste this file into "Code.gs" of the Apps Script project that belongs to
// your Supporting Members spreadsheet (Extensions -> Apps Script).
// Released under the MIT License.
// =========================================================================


// =========================================================================
// CONFIGURATION – fill in your own values
// (see the manual, chapter "Locating Your Setup IDs")
// =========================================================================
const STATION_NAME = 'Your Station Name';        // used in emails and on the PDF
const LEGAL_NAME = 'Your Station Name e.V.';     // official name of the association (PDF + SEPA mandate)
const SHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';     // the Supporting Members spreadsheet
const SHEET_NAME = 'Members';
const ADMIN_EMAIL = 'admin@yourstation.com';     // the board – receives the veto notifications
const SENDER_ALIAS = 'hello@yourstation.com';    // must be set up under Gmail "Send mail as"

// Shared mailing list (newsletter) – the same ID as in Module 1
const MAILING_LIST_SHEET_ID = 'YOUR_MAILING_LIST_SHEET_ID_HERE';
const MAILING_LIST_TAB_NAME = 'Mailing List';

// PDFs in Google Drive that are attached to the welcome email
const STATUTES_FILE_ID = 'YOUR_STATUTES_PDF_ID';
const PRIVACY_POLICY_FILE_ID = 'YOUR_PRIVACY_POLICY_PDF_ID';

// SEPA direct debit mandate – both values are required details of the mandate and are printed on every PDF.
// Creditor identifier (Gläubiger-Identifikationsnummer) from the Deutsche Bundesbank, e.g. 'DE98ZZZ09999999999'.
const CREDITOR_ID = '';
// Postal address of the association (the creditor), e.g. 'Example Street 1, 12345 City, Germany'.
const CREDITOR_ADDRESS = 'YOUR_ASSOCIATION_ADDRESS_HERE';

// Only needed for resendPendingNotifications(): the live web app URL ending in /exec.
// Leave empty during normal operation.
const WEB_APP_URL = '';

const MIN_AMOUNT = 1;
const MAX_AMOUNT = 100000;

// Column layout of the Members tab (1 = column A).
// Column C (mandate reference) is filled in by the board; column R is written by the script.
const COL = {
  LAST_NAME: 1, FIRST_NAME: 2, ADDRESS: 4, EMAIL: 5, IBAN: 6, BIC: 7, BANK: 8,
  MONTHLY: 9, QUARTERLY: 10, YEARLY: 11, RECEIPT: 12, MEMBER_TYPE: 13, MOTIVATION: 14,
  STATUS: 15, COMPANY: 16, MAILING_LIST: 17, TOKEN: 18
};
const TOKEN_HEADER = 'Accept Token (do not edit)';

const STATUS_PENDING = 'Pending Veto';
const STATUS_PDF_SENT = 'PDF sent (Waiting for signature)';


// =========================================================================
// SECTION 1: WEB APP ROUTING (Loading the Form & Handling Accept Button)
// =========================================================================

/**
 * Triggers when someone visits the Web App URL.
 * Loads the application form, or handles the "Accept" link from the admin email.
 */
function doGet(e) {
  const params = (e && e.parameter) || {};

  if (params.action === 'accept') {
    try {
      return handleAccept(params.token);
    } catch (err) {
      console.error('Accept failed: ' + err);
      return renderMessagePage(
        'Something went wrong',
        'The application could not be accepted (' + err.message + '). The status in the sheet was not changed, so you can try the link again. Details are in the Executions log in Apps Script.',
        true
      );
    }
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(`Supporting Membership - ${LEGAL_NAME}`)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// =========================================================================
// SECTION 2: FORM SUBMISSION (Logging Data & Veto Phase Notification)
// =========================================================================

/**
 * Triggers when someone submits the supporting membership form.
 * Returns { ok: true } or { ok: false, error: '...' }.
 */
function submitMembership(input) {
  try {
    const check = validateMembership(input);
    if (check.error) return { ok: false, error: check.error };
    const m = check.data;

    const token = Utilities.getUuid();
    const lock = LockService.getScriptLock();
    let rowNumber;

    try {
      lock.waitLock(30000); // so the row number always belongs to this application
      const sheet = getMemberSheet();
      ensureTokenColumn(sheet);
      sheet.appendRow([
        m.lastName,
        m.firstName,
        '',
        m.street + ', ' + m.zipCity,
        m.email,
        m.iban,
        m.bic,
        m.bankName,
        m.interval === 'monthly' ? m.amount : '',
        m.interval === 'quarterly' ? m.amount : '',
        m.interval === 'yearly' ? m.amount : '',
        m.receipt ? 'Yes' : 'No',
        memberTypeLabel(m.memberType),
        m.motivation,
        STATUS_PENDING,
        m.companyName,
        m.mailingList ? 'Yes' : 'No',
        token
      ].map(toSafeCell));
      rowNumber = sheet.getLastRow();
      sheet.getRange(rowNumber, COL.STATUS).setBackground('#FFF2CC'); // yellow
    } finally {
      lock.releaseLock();
    }

    if (m.mailingList) {
      const fullName = `${m.firstName} ${m.lastName}`;
      addToMailingList({
        email: m.email,
        name: m.companyName ? `${fullName} (${m.companyName})` : fullName,
        alias: '',
        source: 'Supporting Member'
      });
    }

    // If this mail fails, the application is still saved – resendPendingNotifications() can resend it.
    try {
      sendAdminNotification({
        firstName: m.firstName,
        lastName: m.lastName,
        companyName: m.companyName,
        isCommercial: m.memberType === 'commercial',
        memberTypeLabel: memberTypeLabel(m.memberType),
        email: m.email,
        motivation: m.motivation,
        mailingList: m.mailingList
      }, rowNumber, token);
    } catch (mailErr) {
      console.error('Admin notification failed for row ' + rowNumber + ': ' + mailErr);
    }

    return { ok: true };
  } catch (err) {
    console.error('submitMembership failed: ' + err);
    return { ok: false, error: 'Something went wrong on our side. Please try again later or contact us at ' + ADMIN_EMAIL + '.' };
  }
}

/**
 * Server-side validation – never trust what the browser sends.
 */
function validateMembership(input) {
  const f = input || {};
  const m = {
    memberType: f.memberType === 'commercial' ? 'commercial' : (f.memberType === 'individual' ? 'individual' : ''),
    companyName: cleanLine(f.companyName, 200),
    firstName: cleanLine(f.firstName, 100),
    lastName: cleanLine(f.lastName, 100),
    street: cleanLine(f.address, 200),
    zipCity: cleanLine(f.zipCity, 200).replace(/,/g, ' ').replace(/\s{2,}/g, ' ').trim(),
    email: cleanLine(f.email, 254),
    motivation: cleanText(f.motivation, 2000),
    amount: Number(String(f.amount == null ? '' : f.amount).replace(',', '.')),
    interval: ['monthly', 'quarterly', 'yearly'].indexOf(f.interval) !== -1 ? f.interval : '',
    iban: String(f.iban || '').replace(/\s+/g, '').toUpperCase(),
    bic: String(f.bic || '').replace(/\s+/g, '').toUpperCase(),
    bankName: cleanLine(f.bankName, 200),
    coc: f.coc === 'on',
    receipt: f.receipt === 'on',
    mailingList: f.mailingList === 'on'
  };

  if (!m.memberType) return { error: 'Please choose whether you are joining as an individual or a business.' };
  if (m.memberType === 'commercial' && !m.companyName) return { error: 'Please enter your company name.' };
  if (m.memberType !== 'commercial') m.companyName = '';
  if (!m.firstName || !m.lastName) return { error: 'Please enter your first and last name.' };
  if (!m.street || !m.zipCity) return { error: 'Please enter your full address.' };
  if (!isValidEmail(m.email)) return { error: 'Please enter a valid email address.' };
  if (!m.motivation) return { error: 'Please tell us briefly what motivates you to join.' };
  if (!Number.isFinite(m.amount) || m.amount < MIN_AMOUNT || m.amount > MAX_AMOUNT) {
    return { error: `Please enter an amount of at least ${MIN_AMOUNT} €.` };
  }
  m.amount = Math.round(m.amount * 100) / 100;
  if (!m.interval) return { error: 'Please choose a payment interval.' };
  if (!isValidIban(m.iban)) return { error: 'Please check your IBAN – it does not seem to be valid.' };
  if (m.bic && !/^[A-Z]{6}[A-Z0-9]{2}([A-Z0-9]{3})?$/.test(m.bic)) {
    return { error: 'Please check your BIC (8 or 11 characters).' };
  }
  if (!m.bankName) return { error: 'Please enter the name of your bank.' };
  if (!m.coc) return { error: 'Please accept the Code of Conduct.' };
  return { data: m };
}

/**
 * Adds someone to the shared mailing list.
 * Columns: Email | Name | DJ Alias | Source | Added on
 * If the address is already on the list, nothing is duplicated –
 * only empty Name / DJ Alias cells are filled in.
 */
function addToMailingList(entry) {
  try {
    if (!MAILING_LIST_SHEET_ID || MAILING_LIST_SHEET_ID.indexOf('YOUR_') === 0) {
      console.warn('MAILING_LIST_SHEET_ID is not set – newsletter sign-up not saved for ' + entry.email);
      return;
    }
    const sheet = SpreadsheetApp.openById(MAILING_LIST_SHEET_ID).getSheetByName(MAILING_LIST_TAB_NAME);
    if (!sheet) {
      console.error(`Tab "${MAILING_LIST_TAB_NAME}" not found in the mailing list.`);
      return;
    }
    const target = String(entry.email).toLowerCase().trim();
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
      for (let i = 0; i < rows.length; i++) {
        if (String(rows[i][0]).toLowerCase().trim() !== target) continue;
        if (!rows[i][1] && entry.name) sheet.getRange(i + 2, 2).setValue(toSafeCell(entry.name));
        if (!rows[i][2] && entry.alias) sheet.getRange(i + 2, 3).setValue(toSafeCell(entry.alias));
        return;
      }
    }
    sheet.appendRow([entry.email, entry.name || '', entry.alias || '', entry.source, new Date()].map(toSafeCell));
  } catch (err) {
    console.error('Could not add to mailing list: ' + err);
  }
}

/**
 * Notification email to the board (veto phase). Every applicant value is HTML-escaped.
 */
function sendAdminNotification(n, rowNumber, token) {
  const e = escapeHtml;
  const displayName = n.isCommercial
    ? `${n.companyName} (Contact: ${n.firstName} ${n.lastName})`
    : `${n.firstName} ${n.lastName}`;

  const acceptUrl = `${getWebAppUrl()}?action=accept&token=${encodeURIComponent(token)}`;
  const declineSubject = encodeURIComponent(`Your application to ${STATION_NAME}`);
  const declineBody = encodeURIComponent(`Hi ${n.firstName},\n\nThank you for your interest. Unfortunately, we cannot accept your application at this time...\n\nBest,\n${STATION_NAME} Team`);
  const declineUrl = `mailto:${encodeURIComponent(n.email).replace(/%40/g, '@')}?subject=${declineSubject}&body=${declineBody}`;

  const adminBody = `
    <div style="font-family: Arial, sans-serif; padding: 20px; color: #333;">
      <h2 style="color: #000;">New Supporting Member Application</h2>
      <p><strong>Applicant:</strong> ${e(displayName)}</p>
      <p><strong>Email:</strong> ${e(n.email)}</p>
      <p><strong>Type:</strong> ${e(n.memberTypeLabel)}</p>
      <p><strong>Motivation:</strong> ${e(n.motivation).replace(/\n/g, '<br>')}</p>
      <p><strong>Newsletter:</strong> ${n.mailingList ? 'Yes' : 'No'}</p>
      <br>
      <p>The data has been added to the sheet (currently row ${rowNumber}). If you accept this person/company, click Accept. Only then will the PDF be generated and sent for signature.</p>
      <br>
      <a href="${e(acceptUrl)}" style="background-color: #166534; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: bold; margin-right: 15px; display: inline-block;">Accept & Send PDF</a>
      <a href="${e(declineUrl)}" style="background-color: #991b1b; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: bold; display: inline-block;">Veto (Open Draft)</a>
    </div>
  `;

  const subject = `New Application: ${n.firstName} ${n.lastName}`.replace(/[\r\n]+/g, ' ');
  GmailApp.sendEmail(ADMIN_EMAIL, subject, '', {
    htmlBody: adminBody,
    from: SENDER_ALIAS,
    cc: Session.getEffectiveUser().getEmail() // the script owner – see "Google Groups Email Trick" in the manual
  });
}


// =========================================================================
// SECTION 3: ACCEPT ACTION (PDF Generation & Emailing Applicant)
// =========================================================================

/**
 * Triggers when the board clicks "Accept & Send PDF".
 * The link only carries a random token; all data is read from the sheet.
 * Only applications with status "Pending Veto" are processed, so a second
 * click never sends the PDF twice.
 */
function handleAccept(token) {
  if (!isValidToken(token)) {
    return renderMessagePage(
      'Link not valid',
      'This accept link is invalid or incomplete. If a notification email got lost, run resendPendingNotifications() once in the Apps Script editor to get new links.',
      true
    );
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getMemberSheet();
    const row = findRowByToken(sheet, COL.TOKEN, token);
    if (!row) {
      return renderMessagePage('Application not found', 'No application in the sheet matches this link. The row may have been deleted.', true);
    }

    const values = sheet.getRange(row, 1, 1, COL.TOKEN).getValues()[0];
    const status = String(values[COL.STATUS - 1]).trim();
    if (status !== STATUS_PENDING) {
      return renderMessagePage('Already handled', `This application already has the status "${status}". No email was sent.`, false);
    }

    const data = rowToMember(values);

    // 1. Generate the pre-filled application PDF
    const applicationPdf = buildApplicationPdf(data);

    // 2. Load statutes and privacy policy from Google Drive
    const statutesPdf = DriveApp.getFileById(STATUTES_FILE_ID).getAs('application/pdf');
    const privacyPdf = DriveApp.getFileById(PRIVACY_POLICY_FILE_ID).getAs('application/pdf');

    // 3. Email to the applicant
    const applicantBody = `
      Hi ${escapeHtml(data.firstName)},<br><br>
      Fantastic news! We have reviewed your application and are thrilled to welcome you as a supporting member of ${escapeHtml(LEGAL_NAME)}.<br><br>
      To complete the process and activate your SEPA mandate, we just need your signature.<br><br>
      <strong>Next steps:</strong><br>
      1. Attached to this email, you will find our statutes and privacy policy (DSGVO). Please read through them.<br>
      2. Open the attached, pre-filled membership application.<br>
      3. Manually or digitally check the missing boxes for the statutes, DSGVO and the SEPA direct debit mandate on the application.<br>
      4. Sign the document (signature of the account holder) and simply reply to this email with the signed PDF.<br><br>
      As soon as we receive the PDF, you are officially on board!<br><br>
      Best,<br>
      ${escapeHtml(STATION_NAME)} Team
    `;

    GmailApp.sendEmail(data.email, `Action Required: Sign your ${STATION_NAME} Membership`, '', {
      htmlBody: applicantBody,
      from: SENDER_ALIAS,
      attachments: [applicationPdf, statutesPdf, privacyPdf]
    });

    // 4. Status -> PDF sent (darker yellow)
    const statusCell = sheet.getRange(row, COL.STATUS);
    statusCell.setValue(STATUS_PDF_SENT);
    statusCell.setBackground('#FFE599');

    return renderMessagePage(
      'PDFs sent successfully!',
      'The applicant has just received the application, statutes, and privacy policy. The status in the sheet has been updated. Once the signed document is returned via email, set the status to "Active".',
      false
    );
  } finally {
    lock.releaseLock();
  }
}

function buildApplicationPdf(data) {
  if (!data.creditorId || !data.creditorAddress) {
    console.warn('CREDITOR_ID or CREDITOR_ADDRESS is not set – the SEPA mandate on this PDF is incomplete.');
  }
  const template = HtmlService.createTemplateFromFile('PdfTemplate');
  template.data = data;
  const html = template.evaluate().getContent();
  const pdf = Utilities.newBlob(html, 'text/html', 'Membership_Application.html').getAs('application/pdf');
  const baseName = data.isCommercial && data.companyName ? data.companyName : `${data.firstName}_${data.lastName}`;
  pdf.setName(`${safeFileName(STATION_NAME)}_Membership_${safeFileName(baseName)}.pdf`);
  return pdf;
}

/** Returns '' for empty settings and for placeholders that were not replaced (e.g. 'YOUR_..._HERE'). */
function configValue(value) {
  const s = String(value || '').trim();
  return /^YOUR_/i.test(s) ? '' : s;
}

/** Turns a sheet row into the data object used by the PDF template and the emails. */
function rowToMember(values) {
  const get = col => values[col - 1];
  const address = splitAddress(get(COL.ADDRESS));
  const typeLabel = String(get(COL.MEMBER_TYPE) || '');
  const monthly = get(COL.MONTHLY);
  const quarterly = get(COL.QUARTERLY);
  const yearly = get(COL.YEARLY);

  return {
    lastName: String(get(COL.LAST_NAME) || '').trim(),
    firstName: String(get(COL.FIRST_NAME) || '').trim(),
    street: address.street,
    zipCity: address.zipCity,
    email: String(get(COL.EMAIL) || '').trim(),
    iban: String(get(COL.IBAN) || ''),
    bic: String(get(COL.BIC) || ''),
    bankName: String(get(COL.BANK) || ''),
    monthly: monthly,
    quarterly: quarterly,
    yearly: yearly,
    amountDisplay: formatAmount(monthly || quarterly || yearly),
    receipt: get(COL.RECEIPT) === 'Yes',
    memberTypeLabel: typeLabel,
    isCommercial: typeLabel.indexOf('commercial') !== -1,
    motivation: String(get(COL.MOTIVATION) || ''),
    companyName: String(get(COL.COMPANY) || '').trim(),
    mailingList: get(COL.MAILING_LIST) === 'Yes',
    creditorId: configValue(CREDITOR_ID),
    creditorAddress: configValue(CREDITOR_ADDRESS),
    stationName: STATION_NAME,
    legalName: LEGAL_NAME
  };
}


// =========================================================================
// SECTION 4: ONE-TIME MIGRATION / RESEND
// =========================================================================

/**
 * Run this ONCE manually in the Apps Script editor after deploying the update
 * (select the function at the top, then click "Run").
 *
 * For every application that is still "Pending Veto", the board gets a new
 * notification email with a secure Accept button. Old links stop working.
 * Can also be used later if a notification email ever got lost.
 */
function resendPendingNotifications() {
  const baseUrl = getWebAppUrl();
  if (!baseUrl || /\/dev$/.test(baseUrl)) {
    throw new Error('Please paste the live web app URL (ending in /exec) into WEB_APP_URL and run this again.');
  }

  const sheet = getMemberSheet();
  ensureTokenColumn(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    console.log('No applications found.');
    return;
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, COL.TOKEN).getValues();
  let count = 0;
  rows.forEach((values, i) => {
    if (String(values[COL.STATUS - 1]).trim() !== STATUS_PENDING) return;
    const rowNumber = i + 2;
    let token = String(values[COL.TOKEN - 1] || '').trim();
    if (!isValidToken(token)) {
      token = Utilities.getUuid();
      sheet.getRange(rowNumber, COL.TOKEN).setValue(token);
    }
    sendAdminNotification(rowToMember(values), rowNumber, token);
    count++;
  });
  console.log(`Sent ${count} notification(s) with new accept links. Web app URL used: ${baseUrl}`);
}


// =========================================================================
// SECTION 5: HELPERS
// =========================================================================

function getMemberSheet() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Tab "${SHEET_NAME}" not found.`);
  return sheet;
}

function getWebAppUrl() {
  return WEB_APP_URL || ScriptApp.getService().getUrl();
}

/** Makes sure column 18 exists and has a header (only if the header cell is empty). */
function ensureTokenColumn(sheet) {
  if (sheet.getMaxColumns() < COL.TOKEN) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), COL.TOKEN - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() < 1) return;
  const header = sheet.getRange(1, COL.TOKEN);
  if (header.getValue() === '') header.setValue(TOKEN_HEADER);
}

/** Finds the row (>= 2) whose token column matches exactly. Returns 0 if not found. */
function findRowByToken(sheet, column, token) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const cell = sheet.getRange(2, column, lastRow - 1, 1)
    .createTextFinder(token)
    .matchEntireCell(true)
    .findNext();
  return cell ? cell.getRow() : 0;
}

function memberTypeLabel(type) {
  return 'Supporting Member (' + type + ')';
}

/** "Street 1, Hinterhaus, 52062 Aachen" -> split at the LAST comma. */
function splitAddress(fullAddress) {
  const s = String(fullAddress || '');
  const i = s.lastIndexOf(',');
  if (i === -1) return { street: s.trim(), zipCity: '' };
  return { street: s.slice(0, i).trim(), zipCity: s.slice(i + 1).trim() };
}

function formatAmount(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return String(value || '');
  return Number.isInteger(n) ? String(n) : n.toFixed(2).replace('.', ',');
}

function safeFileName(name) {
  return String(name).replace(/[^\p{L}\p{N}_-]+/gu, '_').replace(/^_+|_+$/g, '') || 'Application';
}

/** IBAN check incl. the mod-97 checksum (catches typos before the board sees the application). */
function isValidIban(iban) {
  if (!/^[A-Z]{2}\d{2}[A-Z0-9]{11,30}$/.test(iban)) return false;
  const rearranged = iban.slice(4) + iban.slice(0, 4);
  let remainder = 0;
  for (const ch of rearranged) {
    const digits = ch >= 'A' && ch <= 'Z' ? String(ch.charCodeAt(0) - 55) : ch;
    for (const d of digits) remainder = (remainder * 10 + Number(d)) % 97;
  }
  return remainder === 1;
}

function renderMessagePage(title, message, isError) {
  const color = isError ? '#991b1b' : '#166534';
  return HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; text-align: center; padding: 50px 20px; max-width: 600px; margin: 0 auto;">
      <h2 style="color: ${color};">${escapeHtml(title)}</h2>
      <p style="color: #444; font-size: 16px; line-height: 1.5;">${escapeHtml(message)}</p>
    </div>
  `).setTitle(title).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Makes user input safe to put into HTML (emails, pages). */
function escapeHtml(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/** Prevents formula injection: text starting with = + - @ is stored as plain text. */
function toSafeCell(value) {
  if (typeof value !== 'string') return value;
  return /^[=+\-@\t\r]/.test(value) ? "'" + value : value;
}

function cleanText(value, maxLength) {
  return String(value == null ? '' : value).trim().slice(0, maxLength);
}

function cleanLine(value, maxLength) {
  return cleanText(value, maxLength).replace(/[\r\n\t]+/g, ' ');
}

function isValidEmail(email) {
  const s = String(email || '');
  return s.length <= 254 && /^[^\s@<>"]+@[^\s@<>"]+\.[^\s@<>".]{2,}$/.test(s);
}

function isValidToken(token) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(String(token || ''));
}
