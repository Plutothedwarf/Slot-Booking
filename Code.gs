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
 * RUN THIS ONCE from the Apps Script editor (function dropdown → setup →
 * Run). Change ADMIN_PASSWORD below first. Creates a new spreadsheet named
 * "Slot Signup Data" with both tabs and headers already set up, and stores
 * your admin password. Safe to re-run later just to change the password —
 * it won't create a second spreadsheet if one already exists.
 */
function setup() {
  const ADMIN_PASSWORD = 'CHANGE_ME';

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
  // Force start_time/end_time columns to plain text so Sheets never
  // silently converts "12:00" into a time value on us.
  ss.getSheetByName(ROOMS_SHEET).getRange('D2:E').setNumberFormat('@');
  if (!ss.getSheetByName(BOOKINGS_SHEET)) {
    const sh = ss.insertSheet(BOOKINGS_SHEET);
    sh.getRange(1, 1, 1, BOOKINGS_HEADERS.length).setValues([BOOKINGS_HEADERS]).setFontWeight('bold');
  }
  // Remove the default "Sheet1" tab if it's still there and empty
  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  props.setProperty('ADMIN_PASSWORD', ADMIN_PASSWORD);

  Logger.log('Done. Spreadsheet URL: ' + ss.getUrl());
  Logger.log('Now deploy this project as a Web App (see SETUP.md step 3).');
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
        result = bookSlot(payload);
        break;
      case 'cancelBooking':
        result = cancelBooking(payload);
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
        result = adminCreateRoom(payload);
        break;
      case 'adminUpdateRoom':
        requireAdmin(payload);
        result = adminUpdateRoom(payload);
        break;
      case 'adminDeleteRoom':
        requireAdmin(payload);
        result = adminDeleteRoom(payload);
        break;
      case 'adminToggleLock':
        requireAdmin(payload);
        result = adminToggleLock(payload);
        break;
      case 'adminMoveBooking':
        requireAdmin(payload);
        result = adminMoveBooking(payload);
        break;
      case 'adminDeleteBooking':
        requireAdmin(payload);
        result = adminDeleteBooking(payload);
        break;
      case 'adminExportCsv':
        requireAdmin(payload);
        result = adminExportCsv(payload);
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

/* ---------------- Sheet helpers ---------------- */

function sheet(name) {
  const sh = getSS().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet tab: ' + name);
  return sh;
}

function readRows(sheetName, headers) {
  const sh = sheet(sheetName);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const values = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values.map((row, i) => {
    const obj = {};
    headers.forEach((h, idx) => obj[h] = row[idx]);
    obj._row = i + 2; // actual sheet row number
    return obj;
  }).filter(r => r[headers[0]] !== '' && r[headers[0]] !== null);
}

function appendRow(sheetName, headers, obj) {
  const sh = sheet(sheetName);
  const row = headers.map(h => obj[h] !== undefined ? obj[h] : '');
  sh.appendRow(row);
}

function updateRow(sheetName, headers, rowNumber, obj) {
  const sh = sheet(sheetName);
  const row = headers.map(h => obj[h] !== undefined ? obj[h] : '');
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
  return {
    room_id: room.room_id, room_name: room.room_name, date: room.date,
    start_time: room.start_time, end_time: room.end_time,
    interval_minutes: room.interval_minutes, capacity: room.capacity,
    locked: room.locked === true || room.locked === 'TRUE',
    slots: slotData
  };
}

function adminCreateRoom(payload) {
  const room = {
    room_id: uid(),
    room_name: payload.room_name,
    date: payload.date || '',
    start_time: payload.start_time,
    end_time: payload.end_time,
    interval_minutes: payload.interval_minutes,
    capacity: payload.capacity,
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
  ['room_name', 'date', 'start_time', 'end_time', 'interval_minutes', 'capacity'].forEach(k => {
    if (payload[k] !== undefined) room[k] = payload[k];
  });
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
function normMobile(s) {
  return String(s || '').replace(/\D/g, '');
}

function findExistingBooking(roomId, name, mobile) {
  const bookings = readRows(BOOKINGS_SHEET, BOOKINGS_HEADERS).filter(b => b.room_id === roomId);
  return bookings.find(b => normMobile(b.mobile) === normMobile(mobile) && normMobile(mobile) !== '')
      || bookings.find(b => normName(b.name) === normName(name));
}

function bookSlot(payload) {
  const rooms = readRows(ROOMS_SHEET, ROOMS_HEADERS);
  const room = rooms.find(r => r.room_id === payload.roomId);
  if (!room) return { error: 'Room not found' };
  if (room.locked === true || room.locked === 'TRUE') return { error: 'Bookings are closed for this event' };
  if (!payload.name || !payload.slot_start) return { error: 'Name and slot are required' };
  if (!payload.mobile) return { error: 'Mobile number is required' };

  const existing = findExistingBooking(payload.roomId, payload.name, payload.mobile);

  const bookings = readRows(BOOKINGS_SHEET, BOOKINGS_HEADERS).filter(b => b.room_id === payload.roomId);
  const takenForSlot = bookings.filter(b => b.slot_start === payload.slot_start && (!existing || b.booking_id !== existing.booking_id));
  if (takenForSlot.length >= Number(room.capacity)) return { error: 'That slot just filled up, please pick another' };

  if (existing) {
    existing.slot_start = payload.slot_start;
    existing.name = payload.name;
    existing.mobile = payload.mobile || existing.mobile;
    existing.updated_at = new Date().toISOString();
    updateRow(BOOKINGS_SHEET, BOOKINGS_HEADERS, existing._row, existing);
    return { ok: true, booking_id: existing.booking_id, updated: true };
  } else {
    const booking = {
      booking_id: uid(), room_id: payload.roomId, slot_start: payload.slot_start,
      name: payload.name, mobile: payload.mobile || '',
      created_at: new Date().toISOString(), updated_at: new Date().toISOString()
    };
    appendRow(BOOKINGS_SHEET, BOOKINGS_HEADERS, booking);
    return { ok: true, booking_id: booking.booking_id, updated: false };
  }
}

function myBooking(payload) {
  const existing = findExistingBooking(payload.roomId, payload.name, payload.mobile);
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

/* ---------------- Export ---------------- */

function adminExportCsv(payload) {
  const data = adminGetRoom(payload.roomId);
  if (data.error) return data;
  let csv = 'Time Slot,Name,Mobile\n';
  data.slots.forEach(s => {
    if (s.bookings.length === 0) {
      csv += `${s.slot_start},,\n`;
    } else {
      s.bookings.forEach(b => {
        csv += `${s.slot_start},"${(b.name || '').replace(/"/g, '""')}","${b.mobile || ''}"\n`;
      });
    }
  });
  return { ok: true, csv: csv, filename: (data.room_name || 'room').replace(/[^a-z0-9]+/gi, '_') + '_bookings.csv' };
}
