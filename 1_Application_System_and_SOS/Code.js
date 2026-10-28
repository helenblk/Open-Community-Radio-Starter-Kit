// =========================================================================
// OPEN COMMUNITY RADIO STARTER KIT
// Module 1 + 3: Show Application System & SOS Alerts
//
// Paste this file into "Code.gs" of the Apps Script project that belongs to
// your Applications spreadsheet (Extensions -> Apps Script).
// Released under the MIT License.
// =========================================================================


// =========================================================================
// CONFIGURATION – fill in your own values
// (see the manual, chapter "Locating Your Setup IDs")
// =========================================================================
const STATION_NAME = 'Your Station Name';            // used in all emails
const CALENDAR_ID = 'YOUR_CALENDAR_ID_HERE';          // your broadcast calendar
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';    // the Applications spreadsheet
const APPLICATIONS_TAB_NAME = 'Applications';
const SOS_TAB_NAME = 'SOS List';

// Shared mailing list (newsletter) – the same ID goes into the Supporting Membership module
const MAILING_LIST_SHEET_ID = 'YOUR_MAILING_LIST_SHEET_ID_HERE';
const MAILING_LIST_TAB_NAME = 'Mailing List';

const ADMIN_EMAIL = 'admin@yourstation.com';   // receives the application notifications
const SENDER_ALIAS = 'hello@yourstation.com';  // must be set up under Gmail "Send mail as"
const DAYS_IN_FUTURE = 60;                     // how far ahead DJs can book slots

// Broadcast hours shown under each slot card (English weekday name -> hours).
// Change or remove them to match your own broadcast days.
const BROADCAST_HOURS = {
  Wednesday: '18:00 - 22:00',
  Saturday: '16:00 - 20:00',
  Sunday: '14:00 - 18:00'
};

// Column layout of the applications tab (1 = column A).
// Columns 15-17 (O-Q) are written by the script.
const COL = {
  TIMESTAMP: 1, SLOT: 2, NAME: 3, ALIAS: 4, EMAIL: 5, PHONE: 6, INSTAGRAM: 7,
  CONCEPT: 8, TECH: 9, OTHER: 10, HOW_FOUND: 11, COC: 12, NEWSLETTER: 13, SOS: 14,
  STATUS: 15, EVENT_ID: 16, TOKEN: 17
};
const SCRIPT_HEADERS = ['Status', 'Calendar Event ID', 'Accept Token (do not edit)']; // headers for columns 15-17

const STATUS_PENDING = 'Pending';
const STATUS_CONFIRMED = 'Confirmed';

// Google Calendar color IDs
const COLOR_LAVENDER = '1';
const COLOR_GREEN = '2';
const COLOR_YELLOW = '5';
const COLOR_GREEN_ALT = '10';

const WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const DAY_MS = 24 * 60 * 60 * 1000;


// =========================================================================
// SECTION 1: SOS ALERT SYSTEM (Runs directly inside Google Sheets)
// =========================================================================

/**
 * Builds the custom menu when you open the Google Sheet.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚨 Free Slot')
    .addItem('🚨 Send SOS Alert to DJs', 'showSosDialog')
    .addToUi();
}

/**
 * Opens the input dialog asking the manager which slot became available.
 */
function showSosDialog() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Send SOS Alert to Standby DJs',
    'Which slot just became available? (e.g., "Tomorrow, Friday 16:00 - 18:00"):',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK) {
    const slotDetails = response.getResponseText();
    if (slotDetails.trim() === '') {
      ui.alert('Error', 'Please enter details about the open slot.', ui.ButtonSet.OK);
      return;
    }
    sendSosAlert(slotDetails);
  }
}

/**
 * Fetches emails from the SOS List and sends one BCC email to everyone.
 */
function sendSosAlert(slotDetails) {
  const ui = SpreadsheetApp.getUi();
  try {
    const sosSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SOS_TAB_NAME);
    if (!sosSheet) {
      ui.alert('Error', `The tab "${SOS_TAB_NAME}" does not exist yet or has no sign-ups.`, ui.ButtonSet.OK);
      return;
    }

    const rows = sosSheet.getDataRange().getValues();
    const seen = new Set();
    const emails = [];

    // Emails are in column B (index 1); row 1 is the header.
    // Duplicates are removed case-insensitively.
    for (let i = 1; i < rows.length; i++) {
      const email = String(rows[i][1] || '').trim();
      const key = email.toLowerCase();
      if (!isValidEmail(email) || seen.has(key)) continue;
      seen.add(key);
      emails.push(email);
    }

    if (emails.length === 0) {
      ui.alert('Info', 'There are currently no emails on the SOS List.', ui.ButtonSet.OK);
      return;
    }

    const subject = `🚨 ${STATION_NAME} SOS: Spontaneous Open Slot!`;
    const body = `Hey ${STATION_NAME} DJ,\n\nA slot just became available on short notice and we are looking for someone to jump in!\n\nSlot: ${slotDetails}\n\nIf you have time and want to play, just reply directly to this email. First come, first served!\n\nCheers,\nYour ${STATION_NAME} Team`;

    GmailApp.sendEmail(ADMIN_EMAIL, subject, body, {
      bcc: emails.join(','),
      from: SENDER_ALIAS
    });

    ui.alert('Success!', `The SOS Alert was successfully sent via BCC to ${emails.length} DJs!`, ui.ButtonSet.OK);

  } catch (err) {
    ui.alert('Error', 'Could not send SOS Alert: ' + err.toString(), ui.ButtonSet.OK);
  }
}


// =========================================================================
// SECTION 2: WEB APP ROUTING (Loading the Form & Handling Accept Button)
// =========================================================================

/**
 * Triggers when someone visits the Web App URL.
 * Loads the application form, or handles the "Accept" link from the admin email.
 */
function doGet(e) {
  const params = (e && e.parameter) || {};

  if (params.action === 'accept') {
    try {
      return handleShowAccept(params.token);
    } catch (err) {
      console.error('Accept failed: ' + err);
      return renderMessagePage(
        'Something went wrong',
        'The application could not be accepted completely (' + err.message + '). Please check the calendar and the sheet, or look at the Executions log in Apps Script.',
        true
      );
    }
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(`${STATION_NAME} - Show Application`)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// =========================================================================
// SECTION 3: FETCHING AVAILABLE CALENDAR SLOTS
// =========================================================================

/**
 * Scans the calendar for free "Slot" events and skips blocked days.
 * Used by the frontend to build the slot cards.
 */
function getAvailableSlots() {
  const calendar = getCalendar();
  const now = new Date();
  const endDate = new Date(now.getTime() + DAYS_IN_FUTURE * DAY_MS);
  const events = calendar.getEvents(now, endDate);

  const blockedDays = new Set();
  const bookedMinutes = new Set();
  const freeSlots = [];

  // First pass: blocked days, booked times and free placeholder slots
  events.forEach(event => {
    const title = event.getTitle();
    if (event.isAllDayEvent()) {
      if (!isBroadcastDayTitle(title)) addBlockedDays(event, blockedDays); // e.g. summer break
      return;
    }
    if (isBroadcastDayTitle(title)) return;
    if (isFreeSlotEvent(event)) {
      freeSlots.push(event);
    } else {
      bookedMinutes.add(minuteKey(event.getStartTime())); // someone is already booked here
    }
  });

  // Second pass: keep only slots that are really bookable
  return freeSlots
    .filter(event => {
      const start = event.getStartTime();
      return start > now && !blockedDays.has(dayKey(start)) && !bookedMinutes.has(minuteKey(start));
    })
    .sort((a, b) => a.getStartTime() - b.getStartTime())
    .map(event => {
      const start = event.getStartTime();
      const weekday = WEEKDAYS[start.getDay()];
      return {
        value: String(start.getTime()), // the backend re-checks this slot on submit
        date: `${weekday}, ${pad2(start.getDate())}.${pad2(start.getMonth() + 1)}.`,
        slotTime: formatTime(start),
        sendetagTime: BROADCAST_HOURS[weekday] || ''
      };
    });
}

/**
 * Returns the free placeholder event that starts exactly at startMs,
 * or null if that slot is no longer bookable (taken, blocked, deleted, out of range).
 */
function findBookableSlot(calendar, startMs) {
  const now = Date.now();
  if (!(startMs > now) || startMs > now + (DAYS_IN_FUTURE + 1) * DAY_MS) return null;

  const start = new Date(startMs);
  const dayStart = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  const dayEnd = new Date(start.getFullYear(), start.getMonth(), start.getDate() + 1);

  const blockedDays = new Set();
  let slot = null;
  let alreadyBooked = false;

  calendar.getEvents(dayStart, dayEnd).forEach(event => {
    const title = event.getTitle();
    if (event.isAllDayEvent()) {
      if (!isBroadcastDayTitle(title)) addBlockedDays(event, blockedDays);
      return;
    }
    if (isBroadcastDayTitle(title)) return;
    if (minuteKey(event.getStartTime()) !== minuteKey(start)) return;
    if (isFreeSlotEvent(event)) {
      if (!slot) slot = event;
    } else {
      alreadyBooked = true;
    }
  });

  if (alreadyBooked || blockedDays.has(dayKey(start))) return null;
  return slot;
}


// =========================================================================
// SECTION 4: FORM SUBMISSION (Writing Data & Sending Emails)
// =========================================================================

/**
 * Triggers when a DJ submits the application form.
 * Returns { ok: true } or { ok: false, error: '...', slotTaken?: true }.
 */
function submitApplication(formObject) {
  const check = validateApplication(formObject);
  if (check.error) return { ok: false, error: check.error };
  const app = check.data;

  const calendar = getCalendar();
  const token = Utilities.getUuid();
  const lock = LockService.getScriptLock();
  let start;
  let slotLabel;

  try {
    lock.waitLock(30000); // only one booking at a time -> no double bookings

    // 1. Make sure the slot is still free
    const slotEvent = findBookableSlot(calendar, app.slotStart);
    if (!slotEvent) {
      return {
        ok: false,
        slotTaken: true,
        error: 'Sorry, this slot has just been taken or is no longer available. Please choose another one.'
      };
    }
    start = slotEvent.getStartTime();
    const end = slotEvent.getEndTime();
    slotLabel = formatSlotLabel(start, end);

    // 2. Replace the placeholder with a yellow (pending) event
    const newEvent = calendar.createEvent(app.alias || app.name, start, end, {
      description: buildEventDescription(app)
    });
    newEvent.setColor(COLOR_YELLOW);
    slotEvent.deleteEvent();

    // 3. Log the application in the sheet
    const sheet = getApplicationsSheet();
    ensureScriptHeaders(sheet);
    sheet.appendRow([
      new Date(),
      slotLabel,
      app.name,
      app.alias,
      app.email,
      app.phone,
      app.instagram,
      app.concept,
      app.tech,
      app.other,
      app.howFound,
      'Yes',
      app.newsletter ? 'Yes' : 'No',
      app.sosList ? 'Yes' : 'No',
      STATUS_PENDING,
      newEvent.getId(),
      token
    ].map(toSafeCell));
    sheet.getRange(sheet.getLastRow(), COL.STATUS).setBackground('#fff2cc');
  } finally {
    lock.releaseLock();
  }

  // 4. Everything below does not need the lock
  updateSendetagStatus(calendar, start);
  if (app.newsletter) {
    addToMailingList({ email: app.email, name: app.name, alias: app.alias, source: 'Show Application' });
  }
  if (app.sosList) addToSosList(app);
  sendAdminNotification(app, slotLabel, token);

  return { ok: true };
}

/**
 * Server-side validation – never trust what the browser sends.
 */
function validateApplication(input) {
  const f = input || {};
  const data = {
    slotStart: Number(f.slot_raw_value),
    name: cleanLine(f.name, 200),
    alias: cleanLine(f.djalias, 200),
    email: cleanLine(f.email, 254),
    phone: cleanLine(f.telephone, 50),
    instagram: cleanLine(f.instagram, 100),
    concept: cleanText(f.concept, 5000),
    tech: cleanText(f.tech_support, 5000),
    other: cleanText(f.other_info, 5000),
    howFound: cleanLine(f.how_found, 300),
    coc: f.coc === true,
    newsletter: f.newsletter === true,
    sosList: f.sos_list === true
  };

  if (!Number.isFinite(data.slotStart)) return { error: 'Please select a time slot.' };
  if (!data.name && !data.alias) return { error: 'Please enter your real name or your DJ alias.' };
  if (!isValidEmail(data.email)) return { error: 'Please enter a valid email address.' };
  if (!data.concept) return { error: 'Please describe your show concept.' };
  if (!data.howFound) return { error: 'Please tell us how you found us.' };
  if (!data.coc) return { error: 'Please accept the Code of Conduct.' };
  return { data: data };
}

/**
 * Text stored in the calendar event. The "Email:" line is used by the DJ Schedule Checker.
 */
function buildEventDescription(app) {
  return `NEW SHOW APPLICATION

Name: ${app.name}
DJ Alias: ${app.alias || 'N/A'}
Email: ${app.email}
Phone: ${app.phone || 'N/A'}
Instagram: ${app.instagram || 'N/A'}

Concept:
${app.concept}

Tech Support:
${app.tech || 'None'}

Other:
${app.other || 'None'}`;
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
 * Adds the applicant to the local SOS list (no duplicates).
 */
function addToSosList(app) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SOS_TAB_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SOS_TAB_NAME);
      sheet.appendRow(['Registered on', 'Email', 'Name', 'DJ Alias']);
    }
    if (emailExistsInColumn(sheet, 1, app.email, 1)) return;
    sheet.appendRow([new Date(), app.email, app.name, app.alias].map(toSafeCell));
  } catch (err) {
    console.error('Could not add email to SOS list: ' + err);
  }
}

/**
 * Notification email to station management. Every applicant value is HTML-escaped.
 */
function sendAdminNotification(app, slotLabel, token) {
  const e = escapeHtml;
  const multiline = text => e(text).replace(/\n/g, '<br>');
  const displayName = app.alias || app.name;

  const acceptLink = `${ScriptApp.getService().getUrl()}?action=accept&token=${encodeURIComponent(token)}`;
  const declineSubject = encodeURIComponent(`Your ${STATION_NAME} Application - ${slotLabel}`);
  const declineBody = encodeURIComponent(`Hi ${displayName},\n\nThank you so much for applying for ${slotLabel}. Unfortunately, we cannot offer you this specific slot, but we'd love to see you apply for another date soon!\n\nBest,\n${STATION_NAME} Team`);
  const declineLink = `https://mail.google.com/mail/?view=cm&fs=1&to=${encodeURIComponent(app.email)}&su=${declineSubject}&body=${declineBody}`;

  const htmlBody = `
<div style="font-family: Arial, sans-serif; max-width: 600px; padding: 20px; border: 1px solid #eee; border-radius: 8px; color: #111;">
  <h2 style="color: #000; margin-top: 0;">New Show Application</h2>
  <p style="color: #555; margin: 0 0 20px 0;">A new request has been submitted for <strong>${e(slotLabel)}</strong>. The slot is currently blocked as yellow (pending) in your calendar.</p>

  <hr style="border: none; border-top: 1px solid #eee; margin: 20px 0;">

  <p style="margin: 0 0 8px 0;"><strong>Name:</strong> ${e(app.name || '–')} (${e(app.alias || 'No Alias')})</p>
  <p style="margin: 0 0 8px 0;"><strong>Email:</strong> ${e(app.email)}</p>
  <p style="margin: 0 0 8px 0;"><strong>Phone:</strong> ${e(app.phone || 'N/A')}</p>
  <p style="margin: 0 0 20px 0;"><strong>Instagram:</strong> ${e(app.instagram || 'N/A')}</p>

  <p style="margin: 0 0 5px 0;"><strong>Concept:</strong></p>
  <div style="color: #333; background: #f9f9f9; padding: 12px; border-radius: 6px; margin-bottom: 20px; font-size: 14px;">${multiline(app.concept)}</div>

  <p style="margin: 0 0 5px 0;"><strong>Tech Support:</strong></p>
  <div style="color: #333; background: #f9f9f9; padding: 12px; border-radius: 6px; margin-bottom: 20px; font-size: 14px;">${multiline(app.tech || 'None')}</div>

  <p style="margin: 0 0 8px 0;"><strong>Other:</strong> ${multiline(app.other || 'None')}</p>
  <p style="margin: 0 0 8px 0;"><strong>How found:</strong> ${e(app.howFound)}</p>
  <p style="margin: 0 0 8px 0;"><strong>Newsletter Opt-In:</strong> ${app.newsletter ? 'Yes' : 'No'}</p>
  <p style="margin: 0 0 20px 0;"><strong>SOS-List Opt-In:</strong> ${app.sosList ? 'Yes' : 'No'}</p>

  <hr style="border: none; border-top: 1px solid #eee; margin: 20px 0;">

  <p style="margin-bottom: 15px; font-weight: bold;">Quick Actions:</p>
  <a href="${e(acceptLink)}" target="_blank" style="background-color: #000; color: #fff; padding: 12px 20px; text-decoration: none; border-radius: 6px; display: inline-block; margin-right: 10px; font-weight: bold;">Accept & Send Confirmation</a>
  <a href="${e(declineLink)}" target="_blank" style="background-color: #f2f2f2; color: #333; padding: 12px 20px; text-decoration: none; border-radius: 6px; display: inline-block; font-weight: bold; border: 1px solid #ccc;">Decline Application</a>
</div>
  `;

  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    cc: Session.getEffectiveUser().getEmail(), // the script owner – see "Google Groups Email Trick" in the manual
    subject: `New Application: ${displayName} for ${slotLabel}`,
    htmlBody: htmlBody
  });
}


// =========================================================================
// SECTION 5: ACCEPT ACTION (Board clicked "Accept" in the admin email)
// =========================================================================

/**
 * The accept link only carries a random token. All data (email, name, slot,
 * calendar event) is read from the sheet, so nobody can forge a link that
 * sends emails to arbitrary addresses.
 */
function handleShowAccept(token) {
  if (!isValidToken(token)) {
    return renderMessagePage(
      'Link not valid',
      'This accept link is invalid or incomplete. Please use the button from the notification email.',
      true
    );
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getApplicationsSheet();
    const row = findRowByToken(sheet, COL.TOKEN, token);
    if (!row) {
      return renderMessagePage('Application not found', 'No application in the sheet matches this link. The row may have been deleted.', true);
    }

    const values = sheet.getRange(row, 1, 1, COL.TOKEN).getValues()[0];
    const get = col => values[col - 1];

    if (String(get(COL.STATUS)).trim() === STATUS_CONFIRMED) {
      return renderMessagePage('Already accepted', 'This application was already accepted earlier. No further email was sent.', false);
    }

    const email = String(get(COL.EMAIL)).trim();
    const djName = String(get(COL.ALIAS) || get(COL.NAME)).trim();
    const slotLabel = String(get(COL.SLOT)).trim();

    // 1. Calendar event -> green
    const calendar = getCalendar();
    const eventId = String(get(COL.EVENT_ID) || '').trim();
    const event = eventId ? calendar.getEventById(eventId) : null;
    let calendarNote = '';
    if (event) {
      event.setColor(COLOR_GREEN);
      updateSendetagStatus(calendar, event.getStartTime());
    } else {
      calendarNote = ' Note: the calendar event could not be found (maybe it was moved or deleted) – please check the calendar manually.';
    }

    // 2. Confirmation email to the DJ (before the status, so a failed mail can be retried)
    GmailApp.sendEmail(
      email,
      `Your ${STATION_NAME} Show: ${slotLabel}`,
      `Hi ${djName},\n\nAwesome news! We would love to have you play on ${slotLabel}.\n\nYour slot is confirmed. Please let us know if you have any questions before the show.\n\nCheers,\n${STATION_NAME} Team`,
      { from: SENDER_ALIAS }
    );

    // 3. Sheet status -> Confirmed (green)
    const statusCell = sheet.getRange(row, COL.STATUS);
    statusCell.setValue(STATUS_CONFIRMED);
    statusCell.setBackground('#d4edda');

    return renderMessagePage(
      'Success: Slot Accepted!',
      `The calendar was updated, the sheet was marked, and the confirmation email was sent to ${djName} from ${SENDER_ALIAS}.${calendarNote}`,
      false
    );
  } finally {
    lock.releaseLock();
  }
}


// =========================================================================
// SECTION 6: CALENDAR COLOR MANAGEMENT
// =========================================================================

/**
 * Updates the color of the all-day Broadcast Day event (Sendetag).
 * Yellow = something pending, Lavender = all slots confirmed (hosts needed).
 * A Sendetag that is already green (hosts confirmed manually) is never downgraded.
 */
function updateSendetagStatus(calendar, date) {
  try {
    const startOfDay = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    const endOfDay = new Date(date.getFullYear(), date.getMonth(), date.getDate() + 1);

    let sendetagEvent = null;
    let hasYellow = false;
    let hasEmptySlots = false;
    let allBookedAreGreen = true;
    let bookedCount = 0;

    calendar.getEvents(startOfDay, endOfDay).forEach(event => {
      const title = event.getTitle();
      if (event.isAllDayEvent()) {
        if (isBroadcastDayTitle(title)) sendetagEvent = event;
        return;
      }
      if (isBroadcastDayTitle(title)) return;
      if (isFreeSlotEvent(event)) {
        hasEmptySlots = true;
        return;
      }
      bookedCount++;
      const color = event.getColor();
      if (color === COLOR_YELLOW) {
        hasYellow = true;
        allBookedAreGreen = false;
      } else if (color !== COLOR_GREEN && color !== COLOR_GREEN_ALT) {
        allBookedAreGreen = false;
      }
    });

    if (!sendetagEvent) return;
    const currentColor = sendetagEvent.getColor();
    const isAlreadyGreen = currentColor === COLOR_GREEN || currentColor === COLOR_GREEN_ALT;

    if (hasYellow) {
      sendetagEvent.setColor(COLOR_YELLOW);
    } else if (bookedCount > 0 && allBookedAreGreen && !hasEmptySlots && !isAlreadyGreen) {
      sendetagEvent.setColor(COLOR_LAVENDER);
    }
  } catch (err) {
    console.error('Could not update Broadcast Day status: ' + err);
  }
}

function isBroadcastDayTitle(title) {
  const t = String(title || '').toLowerCase();
  return t.includes('sendetag') || t.includes('broadcast day');
}

/**
 * A free placeholder: the title contains "slot" AND it is not yellow/green.
 * The color check means a DJ alias like "Slotmachine" is never mistaken for a free slot.
 */
function isFreeSlotEvent(event) {
  const bookedColors = [COLOR_YELLOW, COLOR_GREEN, COLOR_GREEN_ALT];
  if (bookedColors.indexOf(event.getColor()) !== -1) return false;
  return String(event.getTitle() || '').toLowerCase().includes('slot');
}

function addBlockedDays(event, blockedDays) {
  const current = new Date(event.getStartTime());
  const end = event.getEndTime();
  while (current < end) {
    blockedDays.add(dayKey(current));
    current.setDate(current.getDate() + 1);
  }
}


// =========================================================================
// SECTION 7: HELPERS
// =========================================================================

function getCalendar() {
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) throw new Error('Calendar not found – check CALENDAR_ID.');
  return calendar;
}

function getApplicationsSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = APPLICATIONS_TAB_NAME ? ss.getSheetByName(APPLICATIONS_TAB_NAME) : ss.getSheets()[0];
  if (!sheet) throw new Error(`Applications tab "${APPLICATIONS_TAB_NAME}" not found.`);
  return sheet;
}

/**
 * Makes sure columns 15-17 exist and have a header (only fills empty header cells).
 */
function ensureScriptHeaders(sheet) {
  if (sheet.getMaxColumns() < COL.TOKEN) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), COL.TOKEN - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() < 1) return;
  const range = sheet.getRange(1, COL.STATUS, 1, SCRIPT_HEADERS.length);
  const current = range.getValues()[0];
  if (current.every(v => v !== '')) return;
  range.setValues([current.map((v, i) => (v !== '' ? v : SCRIPT_HEADERS[i]))]);
}

/**
 * Finds the row (>= 2) whose token column matches exactly. Returns 0 if not found.
 * Works even if rows were sorted or deleted in the meantime.
 */
function findRowByToken(sheet, column, token) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const cell = sheet.getRange(2, column, lastRow - 1, 1)
    .createTextFinder(token)
    .matchEntireCell(true)
    .findNext();
  return cell ? cell.getRow() : 0;
}

function emailExistsInColumn(sheet, columnIndex, email, firstDataRowIndex) {
  const target = String(email).toLowerCase().trim();
  const values = sheet.getDataRange().getValues();
  for (let i = firstDataRowIndex; i < values.length; i++) {
    if (String(values[i][columnIndex] || '').toLowerCase().trim() === target) return true;
  }
  return false;
}

function renderMessagePage(title, message, isError) {
  const color = isError ? '#991b1b' : '#166534';
  return HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; text-align: center; padding: 50px 20px; max-width: 600px; margin: 0 auto;">
      <h2 style="color: ${color};">${escapeHtml(title)}</h2>
      <p style="color: #444; font-size: 16px; line-height: 1.5;">${escapeHtml(message)}</p>
      <p style="margin-top: 30px; color: #666;">You can now close this tab.</p>
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

/**
 * Prevents formula injection: text starting with = + - @ is stored as plain text
 * (also keeps phone numbers like "+49 ..." from being read as a formula).
 */
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

function pad2(n) {
  return String(n).padStart(2, '0');
}

function formatTime(date) {
  return pad2(date.getHours()) + ':' + pad2(date.getMinutes());
}

/** e.g. "Wednesday, 04.11.2026 | 19:00 - 20:00" */
function formatSlotLabel(start, end) {
  return `${WEEKDAYS[start.getDay()]}, ${pad2(start.getDate())}.${pad2(start.getMonth() + 1)}.${start.getFullYear()} | ${formatTime(start)} - ${formatTime(end)}`;
}

function dayKey(date) {
  return date.getFullYear() + '-' + pad2(date.getMonth() + 1) + '-' + pad2(date.getDate());
}

function minuteKey(date) {
  return Math.floor(date.getTime() / 60000);
}
