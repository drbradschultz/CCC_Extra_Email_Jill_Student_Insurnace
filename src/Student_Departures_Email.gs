// ====== Student_Departures_Email.gs ======
// Weekly Monday email listing students who are leaving the clinic, sent to Dr. Schultz.
// Standalone module - all names prefixed with sd_ to avoid collisions with other scripts in this
// project (e.g. the insurance email).
//
// Source of truth: the workbook's roster tab ("Staff and Students 25-26"). Per the roster layout,
// student names are stacked in column C under section headers ("CCC Students", "PAC Students") and
// each student's end date sits beside the name in column D.
//
// Window (confirmed with the office): run Monday morning and report a two-week band centered on the
// send day -- any student who LEFT in the past 7 days, plus any student LEAVING in the next 7 days.
//
// End-date handling: dates may be written m/d/yy, mm/dd/yy, m/d/yyyy, etc. A date that is just a
// placeholder (contains "?", "X", or "TBD") means the real date is unknown -- those students are
// omitted entirely (we only list students with a confirmed date inside the window).
//
// Nothing is sent if no students fall in the window. Entry point: sd_runWeekly().

// ────────────────────────────────────────────────────────────────────────────
// CONFIG
// ────────────────────────────────────────────────────────────────────────────
const SD_CONFIG = {
  SPREADSHEET_ID: '1EaZDvarUiPMiEoHiLUuvOazo7TBPOCFWqOIczsZEJSs', // same workbook the other CCC scripts use
  ROSTER_TAB_FALLBACK: 'Staff and Students 25-26',
  NAME_COL: 'C',           // student names (stacked under "CCC Students" / "PAC Students" headers)
  END_DATE_COL: 'D',       // each student's end date, beside the name
  EMAIL_TO: 'DrSchultz@CollectiveCareClinic.com',
  EMAIL_CC: '',            // set to '' for none
  NOTIFY_ON_ERROR: 'DrSchultz@CollectiveCareClinic.com', // gets a heads-up if the roster tab can't be found
  CLINIC_NAME: 'Collective Care Clinic',
  WINDOW_BACK_DAYS: 7,     // students who left in the past 7 days
  WINDOW_FWD_DAYS: 7       // students leaving in the next 7 days
};

// Palette mirrored from the staffing summary / IOP email so this email matches them.
const SD_C = {
  navy:'#14304d', brass:'#b08d57',
  bg:'#eef1f4', card:'#ffffff', border:'#e3e8ee', line:'#eef1f4',
  text:'#2b2b2b', muted:'#8a94a3', subhead:'#f7f9fb',
  green:'#1f7a54', greenBg:'#e6f4ee',
  red:'#b3261e', redBg:'#fdeceb', redDark:'#7a1c16',
  amber:'#9a6a00', amberBg:'#fef7e6'
};

const SD_WEEKDAYS = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
const SD_MONTHS = ['January','February','March','April','May','June',
                   'July','August','September','October','November','December'];

// ────────────────────────────────────────────────────────────────────────────
// ENTRY POINT
// ────────────────────────────────────────────────────────────────────────────

// Builds and (only if there is at least one departing student) sends the weekly departures email.
function sd_runWeekly() {
  const ss = SpreadsheetApp.openById(SD_CONFIG.SPREADSHEET_ID);
  const tab = sd_findRosterSheet(ss);
  if (!tab) {
    if (SD_CONFIG.NOTIFY_ON_ERROR) {
      MailApp.sendEmail(SD_CONFIG.NOTIFY_ON_ERROR, 'Student departures email ERROR',
        'Could not find the roster tab ("' + SD_CONFIG.ROSTER_TAB_FALLBACK + '"). No email was sent.');
    }
    return;
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const entries = sd_collectDepartures(tab, today);
  if (!entries.length) return; // nobody leaving in the window -> no email at all

  // Split into "recently left" (before today) and "leaving soon" (today through +7 days).
  const todayYmd = sd_ymd(today);
  const recent = entries.filter(e => e.ymd < todayYmd).sort((a, b) => a.ymd - b.ymd);
  const upcoming = entries.filter(e => e.ymd >= todayYmd).sort((a, b) => a.ymd - b.ymd);

  const windowStart = sd_addDays(today, -SD_CONFIG.WINDOW_BACK_DAYS);
  const windowEnd = sd_addDays(today, SD_CONFIG.WINDOW_FWD_DAYS);
  const model = { today, windowStart, windowEnd, recent, upcoming, total: entries.length };

  const subject = 'Student Departures: ' + entries.length + ' student' + (entries.length === 1 ? '' : 's') +
    ' (' + sd_fmtShort(windowStart) + ' – ' + sd_fmtShort(windowEnd) + ')';

  MailApp.sendEmail({
    to: SD_CONFIG.EMAIL_TO,
    cc: SD_CONFIG.EMAIL_CC || undefined,
    subject: subject,
    body: sd_buildPlainText(model),
    htmlBody: sd_buildEmailHtml(model)
  });
}

// ────────────────────────────────────────────────────────────────────────────
// ROSTER READING
// ────────────────────────────────────────────────────────────────────────────

function sd_findRosterSheet(ss) {
  const exact = ss.getSheetByName(SD_CONFIG.ROSTER_TAB_FALLBACK);
  if (exact) return exact;
  const pattern = /^Staff and Students\s+\d{2}-\d{2}$/i;
  for (const sh of ss.getSheets()) if (pattern.test(sh.getName())) return sh;
  return null;
}

// Walk column C top to bottom, tracking the current section header. While inside a student section
// (CCC / PAC), read the name + its end date (column D) and keep only students whose CONFIRMED end
// date falls inside the window. Returns [{ name, clinic, date, ymd }].
function sd_collectDepartures(tab, today) {
  const lastRow = tab.getLastRow();
  if (lastRow < 1) return [];

  const nameCol = sd_colIdx(SD_CONFIG.NAME_COL) + 1;
  const dateCol = sd_colIdx(SD_CONFIG.END_DATE_COL) + 1;
  const names = tab.getRange(1, nameCol, lastRow, 1).getDisplayValues().flat();
  const dateRaw = tab.getRange(1, dateCol, lastRow, 1).getValues().flat();        // real Date if cell is a date
  const dateDisp = tab.getRange(1, dateCol, lastRow, 1).getDisplayValues().flat(); // text as shown (handles "TBD", "6/?/26")

  const windowStartYmd = sd_ymd(sd_addDays(today, -SD_CONFIG.WINDOW_BACK_DAYS));
  const windowEndYmd = sd_ymd(sd_addDays(today, SD_CONFIG.WINDOW_FWD_DAYS));

  const out = [];
  let section = null; // null | 'staff' | { clinic: 'CCC' | 'PAC' | 'Student' }

  for (let i = 0; i < names.length; i++) {
    const cell = String(names[i] || '').trim();
    if (!cell) continue;

    const header = sd_classifyHeader(cell);
    if (header) { section = header; continue; }

    // Only collect within a student section (skip the Staff section and anything above the first header).
    if (!section || section === 'staff') continue;
    if (!sd_isRealName(cell)) continue;

    const parsed = sd_parseEndDate(dateRaw[i], dateDisp[i], today);
    if (!parsed || parsed.placeholder) continue; // no usable / unknown date -> omit

    const ymd = sd_ymd(parsed);
    if (ymd < windowStartYmd || ymd > windowEndYmd) continue;

    out.push({ name: sd_cleanName(cell), clinic: section.clinic, date: parsed, ymd });
  }
  return out;
}

// Classify a column-C cell as a section header. Returns 'staff', { clinic } for a student section,
// or null when the cell is not a header.
function sd_classifyHeader(text) {
  const s = String(text || '').toLowerCase();
  if (/student/.test(s)) {
    const clinic = /\bpac\b/.test(s) ? 'PAC' : (/\bccc\b/.test(s) ? 'CCC' : 'Student');
    return { clinic };
  }
  if (/^staff\b/.test(s)) return 'staff';
  return null;
}

// Parse an end-date cell. Returns a Date (time stripped), { placeholder:true } for an unknown
// placeholder (?, X, TBD), or null when there is no usable date.
function sd_parseEndDate(rawVal, dispVal, today) {
  const s = String(dispVal == null ? '' : dispVal).trim();

  // Placeholder = date not actually known yet.
  if (s) {
    if (s.indexOf('?') >= 0) return { placeholder: true };
    if (/tbd/i.test(s)) return { placeholder: true };
    if (/(^|[^a-z])x([^a-z]|$)/i.test(s)) return { placeholder: true }; // standalone X (e.g. "X", "m/X/yy")
  }

  // A real spreadsheet date value is the most reliable source.
  if (rawVal instanceof Date && !isNaN(rawVal.getTime())) {
    return new Date(rawVal.getFullYear(), rawVal.getMonth(), rawVal.getDate());
  }

  if (!s) return null;

  // Text date: m/d, m/d/yy, mm/dd/yyyy, with / . or - separators.
  const m = s.match(/(\d{1,2})[\/.\-](\d{1,2})(?:[\/.\-](\d{2,4}))?/);
  if (!m) return null;
  const mm = parseInt(m[1], 10), dd = parseInt(m[2], 10);
  if (mm < 1 || mm > 12 || dd < 1 || dd > 31) return null;

  let year;
  if (m[3]) {
    year = m[3].length <= 2 ? 2000 + parseInt(m[3], 10) : parseInt(m[3], 10);
  } else {
    // No year written: pick the year (prev / this / next) that lands the date closest to today.
    year = sd_inferYear(mm, dd, today);
  }
  const d = new Date(year, mm - 1, dd);
  return isNaN(d.getTime()) ? null : d;
}

// Choose the year that places month/day nearest to `today` (handles a missing year near a boundary).
function sd_inferYear(mm, dd, today) {
  const base = today.getFullYear();
  let best = base, bestDiff = Infinity;
  [base - 1, base, base + 1].forEach(y => {
    const diff = Math.abs(new Date(y, mm - 1, dd).getTime() - today.getTime());
    if (diff < bestDiff) { bestDiff = diff; best = y; }
  });
  return best;
}

// Trim a roster name for display: collapse whitespace and drop a trailing credential
// (", PsyD" / "QMHP") so the email shows a clean name.
function sd_cleanName(name) {
  return String(name || '')
    .replace(/,?\s*(PhD|PsyD|MD|DO|LCSW|LSW|LPC|LPCC|LMFT|LMHC|LCPC|MA|MS|MSW|RN|APRN|PMHNP|NP|PA|BCBA|QMHP|CADC|Intern|Practicum|Extern|Fellow|Trainee)\b\.?/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
}

// ────────────────────────────────────────────────────────────────────────────
// RENDERING (mirrors the IOP email: navy letterhead, brass accent, card sections)
// ────────────────────────────────────────────────────────────────────────────

function sd_buildEmailHtml(model) {
  const greeting =
    '<div style="padding:20px 30px 4px 30px;">' +
      '<div style="font-size:17px;color:' + SD_C.navy + ';font-weight:700;letter-spacing:0.2px;">Student Departures</div>' +
      '<div style="font-size:13px;color:' + SD_C.muted + ';margin-top:4px;">' +
        'Students leaving the clinic between ' + sd_esc(sd_fmtLong(model.windowStart)) +
        ' and ' + sd_esc(sd_fmtLong(model.windowEnd)) + '.' +
      '</div>' +
    '</div>';

  let sections = '';
  if (model.upcoming.length) {
    sections += sd_card('Leaving This Week', sd_listBlock(model.upcoming, false));
  }
  if (model.recent.length) {
    sections += sd_card('Recently Left (Past Week)', sd_listBlock(model.recent, true));
  }

  const inner =
    sd_letterhead('Weekly Student Departures') +
    greeting +
    '<div style="padding:14px 30px 8px 30px;">' + sections + '</div>' +
    sd_footer();
  return sd_shell(inner);
}

// One student list. `past` tints the date label so a date already gone reads differently.
function sd_listBlock(entries, past) {
  const dateColor = past ? SD_C.redDark : SD_C.green;
  const rows = entries.map(e => {
    const tag = sd_clinicTag(e.clinic);
    return '<div style="padding:9px 0;border-bottom:1px solid ' + SD_C.line + ';">' +
        '<span style="font-size:14px;color:' + SD_C.text + ';font-weight:600;">' + sd_esc(e.name) + '</span>' +
        tag +
        '<div style="font-size:12.5px;color:' + dateColor + ';margin-top:2px;font-weight:600;">' +
          (past ? 'Left ' : 'Leaving ') + sd_esc(sd_fmtLong(e.date)) +
        '</div>' +
      '</div>';
  }).join('');
  return '<div>' + rows + '</div>';
}

// Per-clinic pill colors: a soft tint of the clinic color, with the deeper shade for text/border
// so the pills read as a gentle accent rather than a bold block.
const SD_CLINIC_COLORS = {
  CCC: { bg: '#daf0f2', text: '#1b8291', border: '#b6dfe4' }, // teal (brighter, so it reads distinct from navy)
  PAC: { bg: '#fef7e6', text: '#9a6a00', border: '#f0e2bf' }  // amber (matches palette amber/amberBg)
};

// Small clinic pill (CCC / PAC) shown beside each name.
function sd_clinicTag(clinic) {
  const label = clinic === 'Student' ? 'Student' : clinic;
  const c = SD_CLINIC_COLORS[clinic] || { bg: SD_C.subhead, text: SD_C.navy, border: SD_C.border };
  return '<span style="display:inline-block;margin-left:8px;font-size:10.5px;font-weight:700;' +
    'letter-spacing:0.06em;text-transform:uppercase;color:' + c.text + ';background:' + c.bg + ';' +
    'border:1px solid ' + c.border + ';border-radius:10px;padding:2px 8px;vertical-align:middle;">' +
    sd_esc(label) + '</span>';
}

function sd_letterhead(subtitle) {
  return '<div style="background:' + SD_C.navy + ';padding:26px 30px;">' +
    '<div style="font-family:Georgia,\'Times New Roman\',serif;font-size:22px;color:#ffffff;letter-spacing:0.4px;">' +
      sd_esc(SD_CONFIG.CLINIC_NAME) + '</div>' +
    '<div style="height:2px;width:48px;background:' + SD_C.brass + ';margin:11px 0 9px 0;"></div>' +
    '<div style="font-size:11px;letter-spacing:0.22em;text-transform:uppercase;color:#aab8c7;">' + sd_esc(subtitle) + '</div>' +
  '</div>';
}

function sd_shell(innerHtml) {
  return '<div style="background:' + SD_C.bg + ';padding:28px 12px;font-family:-apple-system,\'Segoe UI\',Roboto,Helvetica,Arial,sans-serif;">' +
    '<table align="center" width="680" cellpadding="0" cellspacing="0" role="presentation" style="max-width:680px;width:100%;margin:0 auto;">' +
    '<tr><td style="background:' + SD_C.card + ';border:1px solid ' + SD_C.border + ';border-radius:14px;overflow:hidden;">' +
    innerHtml + '</td></tr></table></div>';
}

function sd_footer() {
  return '<div style="padding:16px 30px 22px 30px;background:' + SD_C.subhead + ';border-top:1px solid ' + SD_C.line + ';">' +
    '<div style="font-size:11px;color:' + SD_C.muted + ';letter-spacing:0.02em;">' + sd_esc(SD_CONFIG.CLINIC_NAME) +
    ' &middot; Generated ' + sd_esc(sd_fmtLong(new Date())) + '</div></div>';
}

function sd_card(title, innerHtml) {
  return '<div style="border:1px solid ' + SD_C.border + ';border-radius:10px;overflow:hidden;margin:0 0 16px 0;">' +
    '<div style="background:' + SD_C.subhead + ';padding:12px 16px;border-bottom:1px solid ' + SD_C.border + ';' +
      'font-size:15px;font-weight:700;color:' + SD_C.navy + ';letter-spacing:0.2px;">' + sd_esc(title) + '</div>' +
    '<div style="padding:6px 16px 14px 16px;">' + innerHtml + '</div></div>';
}

function sd_buildPlainText(model) {
  let t = SD_CONFIG.CLINIC_NAME + ' — Weekly Student Departures\n';
  t += 'Window: ' + sd_fmtLong(model.windowStart) + ' – ' + sd_fmtLong(model.windowEnd) + '\n\n';
  if (model.upcoming.length) {
    t += 'LEAVING THIS WEEK\n';
    model.upcoming.forEach(e => { t += '- ' + e.name + ' (' + e.clinic + ') — leaving ' + sd_fmtLong(e.date) + '\n'; });
    t += '\n';
  }
  if (model.recent.length) {
    t += 'RECENTLY LEFT (PAST WEEK)\n';
    model.recent.forEach(e => { t += '- ' + e.name + ' (' + e.clinic + ') — left ' + sd_fmtLong(e.date) + '\n'; });
    t += '\n';
  }
  t += '(Generated ' + sd_fmtLong(new Date()) + ')\n';
  return t;
}

function sd_esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// ────────────────────────────────────────────────────────────────────────────
// HELPERS
// ────────────────────────────────────────────────────────────────────────────

function sd_colIdx(letter) {
  let col = 0;
  for (let i = 0; i < letter.length; i++) col = col * 26 + letter.charCodeAt(i) - 64;
  return col - 1;
}

function sd_ymd(d) { return d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate(); }

function sd_addDays(d, n) {
  const x = new Date(d);
  x.setDate(x.getDate() + n);
  x.setHours(0, 0, 0, 0);
  return x;
}

function sd_fmtLong(d) { return SD_WEEKDAYS[d.getDay()] + ', ' + SD_MONTHS[d.getMonth()] + ' ' + d.getDate(); }
function sd_fmtShort(d) { return (d.getMonth() + 1) + '/' + d.getDate(); }

// A roster cell that is a real person's name (not blank, not a section/label word).
function sd_isRealName(name) {
  const n = String(name || '').trim().toLowerCase().replace(/[.,:;!?]+$/, '');
  if (n.length <= 1) return false;
  if (/^\d+$/.test(n)) return false;
  if (/^(name|names|student|students|staff|supervisor|population|provider|providers|role|biller|notes|team|title|credential|credentials)$/.test(n)) return false;
  return true;
}

// ────────────────────────────────────────────────────────────────────────────
// PREVIEW + TRIGGER SETUP
// ────────────────────────────────────────────────────────────────────────────

// Run by hand: emails a sample with fabricated entries so you can eyeball the layout. Does not read
// the roster or affect the real weekly send.
function sd_previewSample() {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const model = {
    today,
    windowStart: sd_addDays(today, -SD_CONFIG.WINDOW_BACK_DAYS),
    windowEnd: sd_addDays(today, SD_CONFIG.WINDOW_FWD_DAYS),
    upcoming: [
      { name: 'Jordan Avery', clinic: 'CCC', date: sd_addDays(today, 2), ymd: sd_ymd(sd_addDays(today, 2)) },
      { name: 'Priya Nair', clinic: 'PAC', date: sd_addDays(today, 5), ymd: sd_ymd(sd_addDays(today, 5)) }
    ],
    recent: [
      { name: 'Sam Whitfield', clinic: 'CCC', date: sd_addDays(today, -3), ymd: sd_ymd(sd_addDays(today, -3)) }
    ],
    total: 3
  };
  MailApp.sendEmail({
    to: SD_CONFIG.NOTIFY_ON_ERROR || SD_CONFIG.EMAIL_TO,
    subject: '[PREVIEW] Weekly Student Departures',
    body: sd_buildPlainText(model),
    htmlBody: sd_buildEmailHtml(model)
  });
}

// Run once from the editor to schedule the Monday-morning send.
function sd_setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sd_runWeekly') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sd_runWeekly')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();
}
