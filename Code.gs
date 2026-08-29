/*
 * ShiftLog Google Apps Script backend.
 *
 * Before deploying, set Script Properties:
 *   SHIFTLOG_PIN          required, the shared team PIN
 *   SESSION_TTL_SECONDS   optional, from 60 to 21600 (defaults to 14400)
 */

const SESSION_PREFIX = 'shiftlog.session.';
const DEFAULT_SESSION_TTL_SECONDS = 14400;
const MAX_SESSION_TTL_SECONDS = 21600;

function doGet(e) {
  const params = (e && e.parameter) || {};
  const action = String(params.action || '');

  try {
    if (action === 'login') {
      return jsonOut(login(params.pin));
    }

    if (!isValidSession(params.token)) {
      return jsonOut({ error: 'Authentication required', authRequired: true });
    }

    switch (action) {
      case 'getStaff': return jsonOut(getStaff());
      case 'getAbsent': return jsonOut(getAbsent());
      case 'getOpenBreaks': return jsonOut(getOpenBreaks());
      case 'toggleAbsent': return jsonOut(toggleAbsent(params.staff));
      case 'startBreak': return jsonOut(startBreak(params.staff));
      case 'endBreak': return jsonOut(endBreak(params.staff));
      case 'addStaff': return jsonOut(addStaff(params.name));
      case 'removeStaff': return jsonOut(removeStaff(params.name));
      case 'renameStaff': return jsonOut(renameStaff(params.oldName, params.newName));
      default: return jsonOut({ error: 'Unknown action' });
    }
  } catch (err) {
    return jsonOut({ error: err.message || String(err) });
  }
}

function login(pin) {
  const expectedPin = PropertiesService.getScriptProperties().getProperty('SHIFTLOG_PIN');
  if (!expectedPin) throw new Error('SHIFTLOG_PIN has not been configured');
  if (!safeEqual(String(pin || ''), expectedPin)) return { error: 'Invalid PIN' };

  const token = createSessionToken();
  const ttl = sessionTtlSeconds();
  CacheService.getScriptCache().put(SESSION_PREFIX + token, '1', ttl);
  return { token: token, expiresIn: ttl };
}

function isValidSession(token) {
  token = String(token || '');
  return token.length === 64 && CacheService.getScriptCache().get(SESSION_PREFIX + token) === '1';
}

function createSessionToken() {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    Utilities.getUuid() + ':' + new Date().getTime() + ':' + Math.random()
  );
  return bytes.map(function(byte) {
    return ('0' + ((byte + 256) % 256).toString(16)).slice(-2);
  }).join('');
}

function sessionTtlSeconds() {
  const configured = Number(PropertiesService.getScriptProperties().getProperty('SESSION_TTL_SECONDS'));
  if (!configured) return DEFAULT_SESSION_TTL_SECONDS;
  return Math.max(60, Math.min(MAX_SESSION_TTL_SECONDS, Math.floor(configured)));
}

function safeEqual(left, right) {
  if (left.length !== right.length) return false;
  let difference = 0;
  for (let i = 0; i < left.length; i++) difference |= left.charCodeAt(i) ^ right.charCodeAt(i);
  return difference === 0;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Sheet not found: ' + name);
  return sheet;
}

function todayStr() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function normDate(value) {
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return value === null || value === undefined ? '' : String(value).trim();
}

function getStaff() {
  const rows = getSheet('Staff').getDataRange().getValues();
  return rows.slice(1).filter(function(row) {
    return String(row[0] || '').trim() && (row[1] === true || String(row[1]).toUpperCase() === 'TRUE');
  }).map(function(row) { return String(row[0]).trim(); });
}

function getAbsent() {
  const today = todayStr();
  return getSheet('Status').getDataRange().getValues().slice(1).filter(function(row) {
    return normDate(row[0]) === today && String(row[2] || '').trim() === 'Absent';
  }).map(function(row) { return String(row[1] || '').trim(); }).filter(Boolean);
}

function getOpenBreaks() {
  const today = todayStr();
  const rows = getSheet('Log').getDataRange().getValues();
  const result = {};
  rows.slice(1).forEach(function(row) {
    const name = String(row[1] || '').trim();
    const isOpen = row[3] === '' || row[3] === null || row[3] === undefined;
    if (name && normDate(row[0]) === today && isOpen) result[name] = String(row[2] || '').trim();
  });
  return result;
}

function addStaff(name) {
  name = requiredText(name, 'Missing name');
  return withWriteLock(function() {
    getSheet('Staff').appendRow([name, true]);
    return { ok: true };
  });
}

function removeStaff(name) {
  name = requiredText(name, 'Missing name');
  return withWriteLock(function() {
    const sheet = getSheet('Staff');
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0] || '').trim() === name) {
        sheet.getRange(i + 1, 2).setValue(false);
        break;
      }
    }
    return { ok: true };
  });
}

function renameStaff(oldName, newName) {
  oldName = requiredText(oldName, 'Missing name');
  newName = requiredText(newName, 'Missing name');
  return withWriteLock(function() {
    const sheet = getSheet('Staff');
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0] || '').trim() === oldName) {
        sheet.getRange(i + 1, 1).setValue(newName);
        return { ok: true };
      }
    }
    throw new Error('Staff member not found');
  });
}

function toggleAbsent(staff) {
  staff = requiredText(staff, 'Missing staff');
  return withWriteLock(function() {
    const sheet = getSheet('Status');
    const rows = sheet.getDataRange().getValues();
    const today = todayStr();
    for (let i = 1; i < rows.length; i++) {
      if (normDate(rows[i][0]) === today && String(rows[i][1] || '').trim() === staff) {
        sheet.deleteRow(i + 1);
        return { ok: true, absent: false };
      }
    }
    sheet.appendRow([today, staff, 'Absent']);
    return { ok: true, absent: true };
  });
}

function startBreak(staff) {
  staff = requiredText(staff, 'Missing staff');
  return withWriteLock(function() {
    const now = new Date();
    const startTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
    getSheet('Log').appendRow([todayStr(), staff, startTime, '', '']);
    return { ok: true, startTime: startTime };
  });
}

function endBreak(staff) {
  staff = requiredText(staff, 'Missing staff');
  return withWriteLock(function() {
    const sheet = getSheet('Log');
    const rows = sheet.getDataRange().getValues();
    const today = todayStr();
    for (let i = rows.length - 1; i >= 1; i--) {
      const open = rows[i][3] === '' || rows[i][3] === null || rows[i][3] === undefined;
      if (normDate(rows[i][0]) !== today || String(rows[i][1] || '').trim() !== staff || !open) continue;
      const start = parseBreakTime(rows[i][2]);
      if (!start) throw new Error('Invalid break start time for ' + staff);
      const now = new Date();
      const duration = Math.max(0, Math.round((now.getTime() - start.getTime()) / 60000));
      const endTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
      sheet.getRange(i + 1, 4, 1, 2).setValues([[endTime, duration]]);
      return { ok: true, durationMinutes: duration, startTime: String(rows[i][2]) };
    }
    throw new Error('No open break found for ' + staff);
  });
}

function parseBreakTime(value) {
  const match = String(value || '').trim().match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!match) return null;
  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  const seconds = Number(match[3] || 0);
  if (hours > 23 || minutes > 59 || seconds > 59) return null;
  const date = new Date();
  date.setHours(hours, minutes, seconds, 0);
  return isNaN(date.getTime()) ? null : date;
}

function requiredText(value, message) {
  const text = String(value || '').trim();
  if (!text) throw new Error(message);
  return text;
}

function withWriteLock(work) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { return work(); }
  finally { lock.releaseLock(); }
}
