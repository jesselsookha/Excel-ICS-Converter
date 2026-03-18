// Teaching Timetable Converter (Excel → ICS)// Teaching Timetable Converter before this file.

// =====================
// Constants & Helpers
// =====================

// Week, time and text extraction regexes
const WEEK_RE   = /^w\d+$/i;
const TIME_RE   = /^(\d{2})H(\d{2})\s*-\s*(\d{2})H(\d{2})$/;
const MODULE_RE = /\b[A-Z]{3,5}\d{4}\b/;
const GROUP_RE  = /\bGroup\s*\d+\b/i;
const COURSE_RE = /^(?<course>[A-Z0-9]+)\s+Group\b/i;
const LOC_RE    = /LR\s*\d+\s*-\s*CR|LR\s*\d+-CR|LR\s*\d+/i;

// Fallback tokens if the sheet contains explicit holiday words.
// (Kept only as a last resort; the date-based table is primary.)
const HOLIDAY_TOKENS = [
  'GOOD FRIDAY','FAMILY DAY','HUMAN RIGHTS','FREEDOM DAY',"WORKERS' DAY",'WORKERS DAY',
  'YOUTH DAY','WOMEN','HERITAGE','RECONCILIATION','CHRISTMAS','GOODWILL'
];

// Month names to numbers
const MONTHS = {
  Jan:1, Feb:2, Mar:3, Apr:4, May:5, Jun:6,
  Jul:7, Aug:8, Sep:9, Oct:10, Nov:11, Dec:12
};

// Public Holidays by year (expand for 2027+ as needed)
const HOLIDAYS_BY_YEAR = {
  2026: [
    { date: '2026-01-01', name: "New Year's Day" },
    { date: '2026-03-21', name: 'Human Rights Day' },
    { date: '2026-04-03', name: 'Good Friday' },                       // movable (2026)
    { date: '2026-04-06', name: 'Family Day' },                        // movable (2026)
    { date: '2026-04-27', name: 'Freedom Day' },
    { date: '2026-05-01', name: "Workers' Day" },
    { date: '2026-06-16', name: 'Youth Day' },
    { date: '2026-08-09', name: "National Women's Day" },
    { date: '2026-08-10', name: "Public holiday National Women's Day observed" }, // Monday observed
    { date: '2026-09-24', name: 'Heritage Day' },
    { date: '2026-12-16', name: 'Day of Reconciliation' },
    { date: '2026-12-25', name: 'Christmas Day' },
    { date: '2026-12-26', name: 'Day of Goodwill' }
  ]
  // 2027: [ ... ] // add here next year (or compute Easter dates dynamically later)
};

// App state
const STATE = {
  events: [],
  meta: { weeks: [], counts: { class: 0, holiday: 0, not_avail: 0 } }
};

// --- small utility helpers ---
function cellStr(v){ return (v === undefined || v === null) ? '' : String(v).trim(); }
function pad(n){ return String(n).padStart(2, '0'); }
function weekNum(w){ return parseInt(String(w||'').replace(/\D/g,''),10) || 0; }

function parseTimeRange(text){
  const s = cellStr(text);
  const m = s.match(TIME_RE);
  if(!m) return null;
  const [, sh, sm, eh, em ] = m;
  return { start: `${sh}:${sm}`, end: `${eh}:${em}` };
}

// Build YYYY-MM-DD as a string (avoid timezone conversion)
function parseDateLabel(label, year){
  const s = cellStr(label);
  if(!s) return { day:null, label:'', iso:null };
  const parts = s.split(/\s+/);
  if(parts.length < 3) return { day: parts[0] || null, label: s, iso: null };

  const [day, dTxt, monTxt] = parts;
  const num = parseInt((dTxt || '').replace(/\D/g,''), 10);
  const mon = MONTHS[(monTxt || '').slice(0,3)];
  if(!num || !mon) return { day, label: s, iso: null };

  const iso = `${year}-${pad(mon)}-${pad(num)}`;
  return { day, label: s, iso };
}

function sortEvents(a,b){
  if(weekNum(a.week) !== weekNum(b.week)) return weekNum(a.week) - weekNum(b.week);
  if((a.column_index||0) !== (b.column_index||0)) return (a.column_index||0) - (b.column_index||0);
  return (a.start_time||'').localeCompare(b.start_time||'');
}

// Stable UID
function simpleHash(s){
  let h = 0; for(let i=0;i<s.length;i++){ h = (h*31 + s.charCodeAt(i)) >>> 0; }
  return h.toString(16);
}
function makeUID(ev){
  const key = [
    ev.date || 'nodate',
    ev.start_time || '00:00',
    ev.module_code || ev.course_code || (ev.title || 'event'),
    ev.group || 'nogroup',
    ev.location || 'noloc'
  ].join('|').toLowerCase().replace(/[^a-z0-9|]+/g,'-');
  return `${key}-${simpleHash(key)}@timetable`;
}

// Holiday helpers
function getHolidaysForYear(year){
  const list = HOLIDAYS_BY_YEAR[year] || [];
  const map = new Map();
  for(const h of list){ map.set(h.date, h.name); }
  return map;
}
function isHolidayDate(iso, holidayMap){
  return iso && holidayMap.has(iso);
}

// =====================
// Parse Excel (SheetJS)
// =====================
async function parseExcel(file, year){
  if(!window.XLSX) throw new Error("SheetJS (XLSX) not loaded. Ensure ./js/xlsx.full.min.js is present.");

  const holidayMap = getHolidaysForYear(year);

  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array' });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });

  // find week rows
  const weekRows = [];
  for(let r=0; r<rows.length; r++){
    const a = cellStr(rows[r][0]);
    if(WEEK_RE.test(a)) weekRows.push(r);
  }

  const events = [];

  for(let i=0; i<weekRows.length; i++){
    const wRow = weekRows[i];
    const weekLabel = cellStr(rows[wRow][0]);

    // Day+date headers (B..G) on the same row
    const datesMap = {};
    for(let c=1;c<=6;c++) datesMap[c] = cellStr(rows[wRow][c]);

    const startR = wRow + 1;
    const endR = (i+1 < weekRows.length) ? weekRows[i+1] : rows.length;

    for(let r=startR; r<endR; r++){
      const timeCell = cellStr(rows[r][0]);
      const tr = parseTimeRange(timeCell);
      if(!tr) continue;

      for(let c=1;c<=6;c++){
        const text = cellStr(rows[r][c]);
        if(text === '') continue; // skip empty; keep NOT AVAIL + holidays

        const { day, label: date_label, iso: date } = parseDateLabel(datesMap[c], year);
        const upper = text.toUpperCase();

        let category = 'class';
        let holiday_name = null;

        // Primary: classify by date match to official holidays
        if (isHolidayDate(date, holidayMap)) {
          category = 'holiday';
          holiday_name = holidayMap.get(date);
        } else if (upper.includes('NOT AVAIL')) {
          category = 'not_avail';
        } else if (HOLIDAY_TOKENS.some(tok => upper.includes(tok))) {
          // Fallback if the spreadsheet uses explicit holiday text not tied to date headers
          category = 'holiday';
          holiday_name = text;
        }

        const module_code = (text.match(MODULE_RE)||[])[0] || null;
        const group = (text.match(GROUP_RE)||[])[0] || null;
        const locMatch = text.match(LOC_RE);
        const location = locMatch ? locMatch[0].replace(/\s{2,}/g,' ').trim() : null;
        const courseMatch = text.match(COURSE_RE);
        const course_code = courseMatch ? courseMatch.groups.course.toUpperCase() : null;

        const ev = {
          week: weekLabel,
          day: day,
          date_label,
          date: date || null,
          start_time: tr.start,
          end_time: tr.end,
          title: text,
          course_code,
          module_code,
          group,
          location,
          category,
          holiday_name,
          column_index: c
        };
        if(date){
          ev.datetime_start = `${date}T${tr.start}:00`;
          ev.datetime_end   = `${date}T${tr.end}:00`;
        }
        events.push(ev);
      }
    }
  }

  events.sort(sortEvents);
  return events;
}

// =====================
// ICS Export (enhanced)
// =====================
function escapeICS(s){ return String(s).replace(/\\/g,'\\\\').replace(/\n/g,'\\n').replace(/,/g,'\\,').replace(/;/g,'\\;'); }
function dtToICS(isoDateTime){ return isoDateTime.replace(/[-:]/g,'').replace('T','T'); }

// Build all-day ICS lines: DTSTART;VALUE=DATE:YYYYMMDD / DTEND;VALUE=DATE:YYYYMMDD(+1)
function dateToICSDate(iso){ return iso.replace(/-/g,''); }
function addOneDayISO(iso){
  const [y,m,d] = iso.split('-').map(n=>parseInt(n,10));
  const dt = new Date(y, m-1, d+1);
  const yyyy = dt.getFullYear(), mm = String(dt.getMonth()+1).padStart(2,'0'), dd = String(dt.getDate()).padStart(2,'0');
  return `${yyyy}-${mm}-${dd}`;
}

function toICS(events, { calName = 'Teaching Timetable', includeExtras = false } = {}){
  const lines = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Timetable Converter//EN',
    'CALSCALE:GREGORIAN',
    'METHOD:PUBLISH',
    `X-WR-CALNAME:${escapeICS(calName)}`,
    `X-WR-CALDESC:${escapeICS('Lecturer timetable generated from Excel')}`
  ];

  // 1) Classes (always include)
  const classEvents = events.filter(ev => ev.date && ev.category === 'class');

  for(const ev of classEvents){
    const uid = makeUID(ev);
    const dtStart = ev.datetime_start ? dtToICS(ev.datetime_start) : null;
    const dtEnd   = ev.datetime_end   ? dtToICS(ev.datetime_end)   : null;

    const summary =
      (ev.module_code ? `${ev.module_code}${ev.group? ' • '+ev.group:''}` :
      ev.course_code  ? `${ev.course_code}${ev.group? ' • '+ev.group:''}` :
      (ev.title || 'Event'));

    const descParts = [];
    if(ev.title) descParts.push(ev.title);
    if(ev.week || ev.day || ev.date_label) descParts.push(`Week/Day: ${ev.week || ''} ${ev.day || ''} ${ev.date_label || ''}`.trim());
    const description = descParts.join('\\n');
    const loc = ev.location || '';

    lines.push('BEGIN:VEVENT');
    if(dtStart) lines.push('DTSTART:' + dtStart);
    if(dtEnd)   lines.push('DTEND:'   + dtEnd);
    lines.push('UID:' + uid);
    lines.push('SUMMARY:' + escapeICS(summary));
    if(loc) lines.push('LOCATION:' + escapeICS(loc));
    if(description) lines.push('DESCRIPTION:' + escapeICS(description));
    lines.push('END:VEVENT');
  }

  // 2) Extras (Holidays & NOT AVAIL) — collapsed to ONE all-day event per date
  if(includeExtras){
    const extrasByDate = new Map(); // date -> { holiday_name?, hasNotAvail? }
    for(const ev of events){
      if(!ev.date) continue;
      if(ev.category !== 'holiday' && ev.category !== 'not_avail') continue;
      const entry = extrasByDate.get(ev.date) || { holiday_name: null, hasNotAvail: false };
      if(ev.category === 'holiday'){
        // prefer official holiday name if we have it
        entry.holiday_name = entry.holiday_name || ev.holiday_name || 'Public Holiday';
      } else if(ev.category === 'not_avail'){
        entry.hasNotAvail = true;
      }
      extrasByDate.set(ev.date, entry);
    }

    for(const [date, info] of extrasByDate.entries()){
      const yyyymmdd = dateToICSDate(date);
      const yyyymmddNext = dateToICSDate(addOneDayISO(date));

      if(info.holiday_name){
        const uid = `${yyyymmdd}-holiday-${simpleHash(info.holiday_name.toLowerCase())}@timetable`;
        lines.push('BEGIN:VEVENT');
        lines.push(`DTSTART;VALUE=DATE:${yyyymmdd}`);
        lines.push(`DTEND;VALUE=DATE:${yyyymmddNext}`);
        lines.push(`UID:${uid}`);
        lines.push(`SUMMARY:${escapeICS('Public Holiday — ' + info.holiday_name)}`);
        lines.push('CATEGORIES:HOLIDAY');
        lines.push('END:VEVENT');
      }
      if(info.hasNotAvail){
        const uid = `${yyyymmdd}-notavail-${simpleHash(date)}@timetable`;
        lines.push('BEGIN:VEVENT');
        lines.push(`DTSTART;VALUE=DATE:${yyyymmdd}`);
        lines.push(`DTEND;VALUE=DATE:${yyyymmddNext}`);
        lines.push(`UID:${uid}`);
        lines.push('SUMMARY:NOT AVAILABLE');
        lines.push('CATEGORIES:NOT_AVAIL');
        lines.push('DESCRIPTION:Lecturer not available (collapsed from multiple slots)');
        lines.push('END:VEVENT');
      }
    }
  }

  lines.push('END:VCALENDAR');
  return lines.join('\r\n');
}

// =====================
// UI rendering (Summary)
// =====================
function renderStats(){
  const el = document.getElementById('stats');
  if(!STATE.events.length){
    el.innerHTML = '<div class="row"><span>No data</span><span>—</span></div>';
    return;
  }
  const counts = {class:0,holiday:0,not_avail:0};
  const weeks = new Set();
  for(const e of STATE.events){ counts[e.category] = (counts[e.category]||0)+1; weeks.add(e.week); }
  STATE.meta = { weeks: Array.from(weeks).sort((a,b)=> parseInt(a.slice(1))-parseInt(b.slice(1))), counts };

  el.innerHTML = [
    `<div class="row"><span>Weeks detected</span><span>${STATE.meta.weeks.length}</span></div>`,
    `<div class="row"><span>Total entries</span><span>${STATE.events.length}</span></div>`,
    `<div class="row"><span>Classes</span><span>${counts.class||0}</span></div>`,
    `<div class="row"><span>Holidays</span><span>${counts.holiday||0}</span></div>`,
    `<div class="row"><span>NOT AVAIL</span><span>${counts.not_avail||0}</span></div>`
  ].join('');
}

function renderSubjects(){
  const root = document.getElementById('subjects');
  root.innerHTML = '';
  if(!STATE.events.length){ root.textContent = '—'; return; }

  // unique module/group combos
  const set = new Set();
  for(const e of STATE.events){
    if(e.category !== 'class') continue;
    const key = [e.module_code || e.course_code || 'Unknown', e.group || ''].join(' | ');
    set.add(key);
  }
  const list = Array.from(set).sort((a,b)=> a.localeCompare(b));
  if(list.length === 0){ root.textContent = '—'; return; }

  for(const item of list){
    const [mod, grp] = item.split(' | ');
    const chip = document.createElement('span');
    chip.className = 'chip mono';
    chip.textContent = grp ? `${mod} • ${grp}` : mod;
    root.appendChild(chip);
  }
}

// =====================
// Next 7 Days (compact)
// =====================
function withinDate(iso, start, end){
  const d = new Date(iso+'T00:00:00');
  const startMid = new Date(start.toDateString());
  return d >= startMid && d < end;
}

function renderNext7(){
  const root = document.getElementById('next7');
  root.innerHTML = '';
  if(!STATE.events.length){
    root.innerHTML = '<div class="day-item"><div class="day-head">No events</div></div>';
    return;
  }
  const today = new Date();
  const end = new Date(today); end.setDate(end.getDate()+7);
  const inRange = STATE.events.filter(ev => ev.date && ev.category === 'class' && withinDate(ev.date, today, end));

  const byDate = {};
  for(const ev of inRange){
    byDate[ev.date] = byDate[ev.date] || [];
    byDate[ev.date].push(ev);
  }
  const dates = Object.keys(byDate).sort();
  if(!dates.length){
    root.innerHTML = '<div class="day-item"><div class="day-head">No classes in the next 7 days.</div></div>';
    return;
  }

  for(const d of dates){
    const nice = new Date(d+'T00:00:00');
    const head = `${nice.toLocaleDateString(undefined,{weekday:'short'})} • ${nice.toLocaleDateString(undefined,{month:'short', day:'numeric'})}`;

    const day = document.createElement('div');
    day.className = 'day-item';

    const h = document.createElement('div');
    h.className = 'day-head';
    h.textContent = head;
    day.appendChild(h);

    const evs = byDate[d].sort((a,b)=> (a.start_time||'').localeCompare(b.start_time||''));
    for(const ev of evs){
      const row = document.createElement('div');
      row.className = 'event';

      // Neutral dot (no per-event colour coding)
      const dot = document.createElement('div');
      dot.className = 'dot';
      dot.style.background = '#a7c957'; // neutral accent

      const time = document.createElement('div');
      time.className = 'event-time mono';
      time.textContent = `${ev.start_time}–${ev.end_time}`;

      const info = document.createElement('div');
      const title = document.createElement('div');
      title.className = 'event-title';
      title.textContent = ev.module_code ? `${ev.module_code}${ev.group? ' • '+ev.group:''}` :
                        ev.course_code  ? `${ev.course_code}${ev.group? ' • '+ev.group:''}` :
                        ev.title;

      const meta = document.createElement('div');
      meta.className = 'event-meta';
      meta.textContent = `${ev.location ? ev.location : ''}${ev.course_code ? (ev.location ? ' • ' : '') + ev.course_code : ''}`;

      info.appendChild(title);
      info.appendChild(meta);

      row.appendChild(dot);
      row.appendChild(time);
      row.appendChild(info);

      day.appendChild(row);
    }

    root.appendChild(day);
  }
}

// =====================
// Wire up
// =====================
document.addEventListener('DOMContentLoaded', () => {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const fileName  = document.getElementById('fileName');
  const btnChoose = document.getElementById('btnChooseFile');
  const yearInput = document.getElementById('yearInput');
  const calNameInput = document.getElementById('calNameInput');
  const btnParse  = document.getElementById('btnParse');
  const btnICS    = document.getElementById('btnDownloadICS');
  const parseStatus = document.getElementById('parseStatus');
  const includeExtrasEl = document.getElementById('includeExtras');

  // show file name immediately on change
  fileInput.addEventListener('change', () => {
    fileName.textContent = fileInput.files && fileInput.files[0] ? fileInput.files[0].name : 'No file selected';
  });

  // browse button triggers real input
  btnChoose.addEventListener('click', () => fileInput.click());

  // drag & drop
  function prevent(e){ e.preventDefault(); e.stopPropagation(); }
  ['dragenter','dragover','dragleave','drop'].forEach(evt => { dropzone.addEventListener(evt, prevent); });
  dropzone.addEventListener('dragover', () => dropzone.classList.add('dragover'));
  dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
  dropzone.addEventListener('drop', (e) => {
    dropzone.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if(files && files[0]){
      fileInput.files = files;  // attach to input so our code path is consistent
      fileName.textContent = files[0].name; // update label immediately
    }
  });

  // parse
  btnParse.addEventListener('click', async () => {
    const f = fileInput.files && fileInput.files[0];
    const year = parseInt(yearInput.value,10) || 2026;
    if(!f){ alert('Please choose an Excel file first.'); return; }
    if(!window.XLSX){ alert('Excel parser (XLSX) not loaded. Ensure ./js/xlsx.full.min.js is present.'); return; }

    parseStatus.textContent = 'Parsing… (runs locally in your browser)';
    try {
      const events = await parseExcel(f, year);
      STATE.events = events;

      // compute & render summary
      renderStats();
      renderSubjects();
      renderNext7();

      // enable ICS download if any events exist
      btnICS.disabled = events.length === 0;

      const counts = STATE.meta.counts;
      parseStatus.textContent =
        `Parsed ${events.length} entries across ${STATE.meta.weeks.length} weeks · ` +
        `classes: ${counts.class||0} · holidays: ${counts.holiday||0} · NOT AVAIL: ${counts.not_avail||0}`;

    } catch(err){
      console.error(err);
      parseStatus.textContent = 'Error parsing file. Please check the layout (wN row + day headers in columns B–G + time rows).';
      alert('Parsing error: ' + err.message);
    }
  });

  // ICS download
  btnICS.addEventListener('click', () => {
    const includeExtras = includeExtrasEl.checked;
    const calName = calNameInput.value.trim() || 'Teaching Timetable';
    const ics = toICS(STATE.events, { calName, includeExtras });
    const blob = new Blob([ics], { type: 'text/calendar' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = (calName.replace(/[\\/:*?"<>|]+/g,'-') || 'Teaching Timetable') + '.ics';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(()=> URL.revokeObjectURL(a.href), 2000);
  });

  if(!window.XLSX){
    parseStatus.textContent = 'Excel parser not loaded. Ensure ./js/xlsx.full.min.js exists.';
  }
});
``
