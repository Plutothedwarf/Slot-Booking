/**
 * SLOT SIGNUP — Apps Script backend
 * ----------------------------------
 * This is a STANDALONE script (not bound to an existing Sheet). Run
 * setup() once and it creates the Spreadsheet, both tabs, and headers
 * for you, and sets your admin password. See SETUP.md.
 *
 * All requests are GET requests with a JSON payload in the "payload" query
 * param (this avoids CORS preflight, which Apps Script cannot handle).
 */

const ROOMS_SHEET = 'Rooms';
const BOOKINGS_SHEET = 'Bookings';

const ROOMS_HEADERS = ['room_id', 'room_name', 'date', 'start_time', 'end_time', 'interval_minutes', 'capacity', 'locked', 'created_at'];
const BOOKINGS_HEADERS = ['booking_id', 'room_id', 'slot_start', 'name', 'mobile', 'created_at', 'updated_at'];

/**
 * RUN THIS ONCE from the Apps Script editor (function dropdown -> setup -> Run).
 * Creates a spreadsheet named "Slot Signup Data" with both tabs and headers.
 * Safe to re-run: it will not create a second spreadsheet.
 *
 * ADMIN PASSWORD is NOT stored in this file. Set it in
 *   Project Settings -> Script Properties -> Add property
 *   name: ADMIN_PASSWORD   value: (your password)
 */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let sheetId = props.getProperty('SHEET_ID');
  let ss;
  if (sheetId) {
    try { ss = SpreadsheetApp.openById(sheetId); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('Slot Signup Data');
    props.setProperty('SHEET_ID', ss.getId());
  }

  if (!ss.getSheetByName(ROOMS_SHEET)) {
    const sh = ss.insertSheet(ROOMS_SHEET);
    sh.getRange(1, 1, 1, ROOMS_HEADERS.length).setValues([ROOMS_HEADERS]).setFontWeight('bold');
  }
  if (!ss.getSheetByName(BOOKINGS_SHEET)) {
    const sh = ss.insertSheet(BOOKINGS_SHEET);
    sh.getRange(1, 1, 1, BOOKINGS_HEADERS.length).setValues([BOOKINGS_HEADERS]).setFontWeight('bold');
  }
  // Keep date / time / mobile columns as plain text so Sheets never converts them.
  ss.getSheetByName(ROOMS_SHEET).getRange('C2:E').setNumberFormat('@');
  ss.getSheetByName(BOOKINGS_SHEET).getRange('C2:C').setNumberFormat('@');
  ss.getSheetByName(BOOKINGS_SHEET).getRange('E2:E').setNumberFormat('@');

  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  if (!props.getProperty('ADMIN_PASSWORD')) {
    Logger.log('WARNING: ADMIN_PASSWORD is not set. Add it under Project Settings -> Script Properties.');
  }
  Logger.log('Done. Spreadsheet URL: ' + ss.getUrl());
}

function getSS() {
  const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!sheetId) throw new Error('Not set up yet — run setup() first from the Apps Script editor.');
  return SpreadsheetApp.openById(sheetId);
}

function doGet(e) {
  try {
    const payload = e.parameter.payload ? JSON.parse(e.parameter.payload) : {};
    const action = e.parameter.action;
    let result;

    switch (action) {
      case 'getRooms':
        result = getRooms();
        break;
      case 'getRoom':
        result = getRoomPublic(payload.roomId);
        break;
      case 'bookSlot':
        result = withLock(() => bookSlot(payload));
        break;
      case 'cancelBooking':
        result = withLock(() => cancelBooking(payload));
        break;
      case 'myBooking':
        result = myBooking(payload);
        break;
      case 'adminLogin':
        result = adminLogin(payload);
        break;
      case 'adminGetRoom':
        requireAdmin(payload);
        result = adminGetRoom(payload.roomId);
        break;
      case 'adminCreateRoom':
        requireAdmin(payload);
        result = withLock(() => adminCreateRoom(payload));
        break;
      case 'adminUpdateRoom':
        requireAdmin(payload);
        result = withLock(() => adminUpdateRoom(payload));
        break;
      case 'adminDeleteRoom':
        requireAdmin(payload);
        result = withLock(() => adminDeleteRoom(payload));
        break;
      case 'adminToggleLock':
        requireAdmin(payload);
        result = withLock(() => adminToggleLock(payload));
        break;
      case 'adminMoveBooking':
        requireAdmin(payload);
        result = withLock(() => adminMoveBooking(payload));
        break;
      case 'adminDeleteBooking':
        requireAdmin(payload);
        result = withLock(() => adminDeleteBooking(payload));
        break;
      default:
        result = { error: 'Unknown action' };
    }
    return jsonOut(result);
  } catch (err) {
    return jsonOut({ error: err.message || String(err) });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Only one write runs at a time, so two people can never take the last spot
// together and row deletions can't shift under each other.
function withLock(fn) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch (e) { return { error: 'Server is busy, please try again in a moment' }; }
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------------- Sheet helpers ---------------- */

// Columns kept as plain text, by sheet (1-based column numbers).
const TEXT_COLS = { Rooms: [3, 4, 5], Bookings: [3, 5] };

function normTime(s) {
  const m = String(s || '').trim().match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*([AaPp][Mm])?$/);
  if (!m) return '';
  let h = Number(m[1]);
  if (m[3]) {
    const pm = m[3].toLowerCase() === 'pm';
    if (h === 12) h = 0;
    if (pm) h += 12;
  }
  return fmtTime(h, Number(m[2]));
}

function sheet(name) {
  const sh = getSS().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet tab: ' + name);
  return sh;
}

function readRows(sheetName, headers) {
  const sh = sheet(sheetName);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const range = sh.getRange(2, 1, lastRow - 1, headers.length);
  const values = range.getValues();
  const shown = range.getDisplayValues();
  return values.map((row, i) => {
    const obj = {};
    headers.forEach((h, idx) => {
      let v = row[idx];
      if (h === 'start_time' || h === 'end_time' || h === 'slot_start') {
        v = normTime(shown[i][idx]) || (v instanceof Date ? fmtTime(v.getHours(), v.getMinutes()) : String(v));
      } else if (h === 'date' || h === 'mobile') {
        v = shown[i][idx];
      }
      obj[h] = v;
    });
    obj._row = i + 2; // actual sheet row number
    return obj;
  }).filter(r => r[headers[0]] !== '' && r[headers[0]] !== null);
}

function setTextFormats(sh, sheetName, rowNumber) {
  (TEXT_COLS[sheetName] || []).forEach(col => sh.getRange(rowNumber, col).setNumberFormat('@'));
}

function appendRow(sheetName, headers, obj) {
  const sh = sheet(sheetName);
  const row = headers.map(h => obj[h] !== undefined ? obj[h] : '');
  const rowNumber = sh.getLastRow() + 1;
  setTextFormats(sh, sheetName, rowNumber);
  sh.getRange(rowNumber, 1, 1, row.length).setValues([row]);
}

function updateRow(sheetName, headers, rowNumber, obj) {
  const sh = sheet(sheetName);
  const row = headers.map(h => obj[h] !== undefined ? obj[h] : '');
  setTextFormats(sh, sheetName, rowNumber);
  sh.getRange(rowNumber, 1, 1, headers.length).setValues([row]);
}

function deleteRow(sheetName, rowNumber) {
  sheet(sheetName).deleteRow(rowNumber);
}

function uid() {
  return Utilities.getUuid();
}

/* ---------------- Admin auth ---------------- */

function requireAdmin(payload) {
  const stored = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!stored || payload.password !== stored) {
    throw new Error('Not authorized');
  }
}

function adminLogin(payload) {
  const stored = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  return { ok: !!stored && payload.password === stored };
}

/* ---------------- Rooms ---------------- */

function getRooms() {
  const rooms = readRows(ROOMS_SHEET, ROOMS_HEADERS);
  return rooms.map(r => ({
    room_id: r.room_id, room_name: r.room_name, date: r.date,
    locked: r.locked === true || r.locked === 'TRUE'
  }));
}

function buildSlots(room) {
  const slots = [];
  const startMin = timeToMinutes(room.start_time);
  const endMin = timeToMinutes(room.end_time);
  const interval = Number(room.interval_minutes);
  let cur = startMin;
  while (cur + interval <= endMin) {
    const startH = Math.floor(cur / 60), startM = cur % 60;
    slots.push(fmtTime(startH, startM));
    cur += interval;
  }
  return slots;
}

// Google Sheets silently converts text like "12:00" into a Date/time value
// when it lands in a cell, so reading it back can give either a Date object
// or a plain "HH:MM" string depending on how it was entered. Handle both.
function timeToMinutes(val) {
  if (val instanceof Date) {
    return val.getHours() * 60 + val.getMinutes();
  }
  const parts = String(val).split(':');
  return Number(parts[0]) * 60 + Number(parts[1]);
}

function fmtTime(h, m) {
  const pad = n => String(n).padStart(2, '0');
  return pad(h) + ':' + pad(m);
}

function getRoomPublic(roomId) {
  const rooms = readRows(ROOMS_SHEET, ROOMS_HEADERS);
  const room = rooms.find(r => r.room_id === roomId);
  if (!room) return { error: 'Room not found' };
  const slots = buildSlots(room);
  const bookings = readRows(BOOKINGS_SHEET, BOOKINGS_HEADERS).filter(b => b.room_id === roomId);
  const slotData = slots.map(s => {
    const taken = bookings.filter(b => b.slot_start === s);
    return { slot_start: s, taken: taken.length, capacity: Number(room.capacity), full: taken.length >= Number(room.capacity) };
  });
  return {
    room_id: room.room_id, room_name: room.room_name, date: room.date,
    locked: room.locked === true || room.locked === 'TRUE',
    slots: slotData
  };
}

function adminGetRoom(roomId) {
  const rooms = readRows(ROOMS_SHEET, ROOMS_HEADERS);
  const room = rooms.find(r => r.room_id === roomId);
  if (!room) return { error: 'Room not found' };
  const slots = buildSlots(room);
  const bookings = readRows(BOOKINGS_SHEET, BOOKINGS_HEADERS).filter(b => b.room_id === roomId);
  const slotData = slots.map(s => ({
    slot_start: s,
    capacity: Number(room.capacity),
    bookings: bookings.filter(b => b.slot_start === s).map(b => ({
      booking_id: b.booking_id, name: b.name, mobile: b.mobile
    }))
  }));
  // bookings whose time no longer matches any slot (e.g. after editing the event)
  const orphaned = bookings.filter(b => slots.indexOf(b.slot_start) === -1).map(b => ({
    booking_id: b.booking_id, name: b.name, mobile: b.mobile, slot_start: b.slot_start
  }));
  return {
    orphaned: orphaned,
    room_id: room.room_id, room_name: room.room_name, date: room.date,
    start_time: room.start_time, end_time: room.end_time,
    interval_minutes: room.interval_minutes, capacity: room.capacity,
    locked: room.locked === true || room.locked === 'TRUE',
    slots: slotData
  };
}

function validateRoomFields(p) {
  if (!p.room_name || !String(p.room_name).trim()) return 'Event name is required';
  const s = normTime(p.start_time), e = normTime(p.end_time);
  if (!s || !e) return 'Start and end time are required';
  if (timeToMinutes(e) <= timeToMinutes(s)) return 'End time must be after start time';
  if (!(Number(p.interval_minutes) >= 5)) return 'Slot length must be at least 5 minutes';
  if (!(Number(p.capacity) >= 1)) return 'Capacity must be at least 1';
  return null;
}

function adminCreateRoom(payload) {
  const err = validateRoomFields(payload);
  if (err) return { error: err };
  const room = {
    room_id: uid(),
    room_name: String(payload.room_name).trim(),
    date: payload.date || '',
    start_time: normTime(payload.start_time),
    end_time: normTime(payload.end_time),
    interval_minutes: Number(payload.interval_minutes),
    capacity: Number(payload.capacity),
    locked: false,
    created_at: new Date().toISOString()
  };
  appendRow(ROOMS_SHEET, ROOMS_HEADERS, room);
  return { ok: true, room_id: room.room_id };
}

function adminUpdateRoom(payload) {
  const rooms = readRows(ROOMS_SHEET, ROOMS_HEADERS);
  const room = rooms.find(r => r.room_id === payload.roomId);
  if (!room) return { error: 'Room not found' };
  const merged = Object.assign({}, room);
  ['room_name', 'date', 'start_time', 'end_time', 'interval_minutes', 'capacity'].forEach(k => {
    if (payload[k] !== undefined) merged[k] = payload[k];
  });
  const err = validateRoomFields(merged);
  if (err) return { error: err };
  room.room_name = String(merged.room_name).trim();
  room.date = merged.date || '';
  room.start_time = normTime(merged.start_time);
  room.end_time = normTime(merged.end_time);
  room.interval_minutes = Number(merged.interval_minutes);
  room.capacity = Number(merged.capacity);
  updateRow(ROOMS_SHEET, ROOMS_HEADERS, room._row, room);
  return { ok: true };
}

function adminDeleteRoom(payload) {
  const rooms = readRows(ROOMS_SHEET, ROOMS_HEADERS);
  const room = rooms.find(r => r.room_id === payload.roomId);
  if (!room) return { error: 'Room not found' };
  deleteRow(ROOMS_SHEET, room._row);
  // also delete its bookings
  let bookings = readRows(BOOKINGS_SHEET, BOOKINGS_HEADERS).filter(b => b.room_id === payload.roomId);
  bookings.sort((a, b) => b._row - a._row); // delete bottom-up
  bookings.forEach(b => deleteRow(BOOKINGS_SHEET, b._row));
  return { ok: true };
}

function adminToggleLock(payload) {
  const rooms = readRows(ROOMS_SHEET, ROOMS_HEADERS);
  const room = rooms.find(r => r.room_id === payload.roomId);
  if (!room) return { error: 'Room not found' };
  room.locked = payload.locked === true;
  updateRow(ROOMS_SHEET, ROOMS_HEADERS, room._row, room);
  return { ok: true, locked: room.locked };
}

/* ---------------- Bookings ---------------- */

function normName(s) {
  return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
}
// Digits only; a leading 91 or 0 on a longer number is dropped -> 10 digits.
function normMobile(s) {
  let d = String(s || '').replace(/\D/g, '');
  if (d.length > 10 && (d.indexOf('91') === 0 || d.indexOf('0') === 0)) d = d.slice(-10);
  return d;
}

// A person is identified by their mobile number only (names can repeat).
function findExistingBooking(roomId, mobile) {
  const m = normMobile(mobile);
  if (m.length !== 10) return null;
  const bookings = readRows(BOOKINGS_SHEET, BOOKINGS_HEADERS).filter(b => b.room_id === roomId);
  return bookings.find(b => normMobile(b.mobile) === m) || null;
}

function bookSlot(payload) {
  const rooms = readRows(ROOMS_SHEET, ROOMS_HEADERS);
  const room = rooms.find(r => r.room_id === payload.roomId);
  if (!room) return { error: 'Room not found' };
  if (room.locked === true || room.locked === 'TRUE') return { error: 'Bookings are closed for this event' };
  if (!payload.name || !String(payload.name).trim() || !payload.slot_start) return { error: 'Name and slot are required' };
  const mobile = normMobile(payload.mobile);
  if (mobile.length !== 10) return { error: 'Please enter a valid 10-digit mobile number' };
  if (buildSlots(room).indexOf(payload.slot_start) === -1) return { error: 'That time slot does not exist' };

  const existing = findExistingBooking(payload.roomId, mobile);

  const bookings = readRows(BOOKINGS_SHEET, BOOKINGS_HEADERS).filter(b => b.room_id === payload.roomId);
  const takenForSlot = bookings.filter(b => b.slot_start === payload.slot_start && (!existing || b.booking_id !== existing.booking_id));
  if (takenForSlot.length >= Number(room.capacity)) return { error: 'That slot just filled up, please pick another' };

  const name = String(payload.name).trim();
  if (existing) {
    existing.slot_start = payload.slot_start;
    existing.name = name;
    existing.mobile = mobile;
    existing.updated_at = new Date().toISOString();
    updateRow(BOOKINGS_SHEET, BOOKINGS_HEADERS, existing._row, existing);
    return { ok: true, booking_id: existing.booking_id, updated: true };
  }
  const booking = {
    booking_id: uid(), room_id: payload.roomId, slot_start: payload.slot_start,
    name: name, mobile: mobile,
    created_at: new Date().toISOString(), updated_at: new Date().toISOString()
  };
  appendRow(BOOKINGS_SHEET, BOOKINGS_HEADERS, booking);
  return { ok: true, booking_id: booking.booking_id, updated: false };
}

function myBooking(payload) {
  const existing = findExistingBooking(payload.roomId, payload.mobile);
  if (!existing) return { found: false };
  return { found: true, slot_start: existing.slot_start, booking_id: existing.booking_id };
}

function cancelBooking(payload) {
  const rooms = readRows(ROOMS_SHEET, ROOMS_HEADERS);
  const room = rooms.find(r => r.room_id === payload.roomId);
  if (room && (room.locked === true || room.locked === 'TRUE')) return { error: 'Bookings are locked for this event' };
  const bookings = readRows(BOOKINGS_SHEET, BOOKINGS_HEADERS);
  const b = bookings.find(x => x.booking_id === payload.bookingId);
  if (!b) return { error: 'Booking not found' };
  deleteRow(BOOKINGS_SHEET, b._row);
  return { ok: true };
}

function adminMoveBooking(payload) {
  const bookings = readRows(BOOKINGS_SHEET, BOOKINGS_HEADERS);
  const b = bookings.find(x => x.booking_id === payload.bookingId);
  if (!b) return { error: 'Booking not found' };
  const rooms = readRows(ROOMS_SHEET, ROOMS_HEADERS);
  const room = rooms.find(r => r.room_id === b.room_id);
  const takenForSlot = bookings.filter(x => x.room_id === b.room_id && x.slot_start === payload.newSlot && x.booking_id !== b.booking_id);
  if (room && takenForSlot.length >= Number(room.capacity)) return { error: 'Target slot is full' };
  b.slot_start = payload.newSlot;
  b.updated_at = new Date().toISOString();
  updateRow(BOOKINGS_SHEET, BOOKINGS_HEADERS, b._row, b);
  return { ok: true };
}

function adminDeleteBooking(payload) {
  const bookings = readRows(BOOKINGS_SHEET, BOOKINGS_HEADERS);
  const b = bookings.find(x => x.booking_id === payload.bookingId);
  if (!b) return { error: 'Booking not found' };
  deleteRow(BOOKINGS_SHEET, b._row);
  return { ok: true };
}
