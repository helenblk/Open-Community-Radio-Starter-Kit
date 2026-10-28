// =========================================================================
// OPEN COMMUNITY RADIO STARTER KIT
// Module 2: DJ Schedule Checker
//
// Runs as its own standalone Apps Script project (Drive -> New -> Google Apps Script).
// Released under the MIT License.
// =========================================================================


// =========================================================================
// CONFIGURATION – fill in your own values
// =========================================================================
const STATION_NAME = 'Your Station Name';
const CALENDAR_ID = 'YOUR_CALENDAR_ID_HERE'; // the same broadcast calendar as in Module 1
const MONTHS_AHEAD = 6; // How far ahead the checker looks

// Google Calendar color IDs used by the application tool
const COLOR_GREEN = '2';
const COLOR_YELLOW = '5';
const COLOR_GREEN_ALT = '10';

const WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];


// =========================================================================
// SECTION 1: WEB APP ROUTING (Loading the Interface)
// =========================================================================

/**
 * Triggers when someone visits the Web App URL and loads the search page.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(`${STATION_NAME} - Check My Dates`)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// =========================================================================
// SECTION 2: FETCHING THE DJ'S SCHEDULE
// =========================================================================

/**
 * Called by the frontend when a DJ submits their email address.
 *
 * The calendar search is only a pre-filter (it is a fuzzy full-text search –
 * "@gmail.com" would match everybody). An event is only returned if the
 * entered address appears as a COMPLETE email address in its description.
 */
function getMyShows(email) {
  const query = String(email || '').trim().toLowerCase();
  if (!isValidEmail(query)) {
    return { error: 'Please enter a valid email address.' };
  }

  const calendar = CalendarApp.getCalendarById(CALENDAR_ID);
  if (!calendar) {
    return { error: 'Calendar not found. Please contact the studio.' };
  }

  const now = new Date();
  const future = new Date(now.getTime());
  future.setMonth(future.getMonth() + MONTHS_AHEAD);

  const shows = [];
  calendar.getEvents(now, future, { search: query }).forEach(event => {
    if (event.isAllDayEvent()) return;
    const title = event.getTitle();
    if (isBroadcastDayTitle(title) || isFreeSlotEvent(event)) return;

    // Exact match – partial addresses or domains never match
    if (extractEmails(event.getDescription()).indexOf(query) === -1) return;

    const start = event.getStartTime();
    const end = event.getEndTime();
    shows.push({
      sortKey: start.getTime(),
      title: title,
      date: `${WEEKDAYS[start.getDay()]}, ${pad2(start.getDate())}.${pad2(start.getMonth() + 1)}.${start.getFullYear()}`,
      time: `${formatTime(start)} - ${formatTime(end)}`
    });
  });

  shows.sort((a, b) => a.sortKey - b.sortKey);
  return { shows: shows.map(s => ({ title: s.title, date: s.date, time: s.time })) };
}


// =========================================================================
// SECTION 3: HELPERS
// =========================================================================

/** All email addresses contained in a text (lower case). Works with plain text and HTML descriptions. */
function extractEmails(text) {
  const matches = String(text || '').match(/[A-Za-z0-9._%+'\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/g) || [];
  return matches.map(m => m.toLowerCase());
}

function isBroadcastDayTitle(title) {
  const t = String(title || '').toLowerCase();
  return t.includes('sendetag') || t.includes('broadcast day');
}

/** Free placeholder slot = title contains "slot" and it is not yellow/green (booked). */
function isFreeSlotEvent(event) {
  const bookedColors = [COLOR_YELLOW, COLOR_GREEN, COLOR_GREEN_ALT];
  if (bookedColors.indexOf(event.getColor()) !== -1) return false;
  return String(event.getTitle() || '').toLowerCase().includes('slot');
}

function isValidEmail(email) {
  const s = String(email || '');
  return s.length <= 254 && /^[^\s@<>"]+@[^\s@<>"]+\.[^\s@<>".]{2,}$/.test(s);
}

function pad2(n) {
  return String(n).padStart(2, '0');
}

function formatTime(date) {
  return pad2(date.getHours()) + ':' + pad2(date.getMinutes());
}
