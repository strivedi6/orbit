/* =====================================================================
   ORBIT — Google Apps Script backend
   1. Open the "Orbit Data" sheet (Drive › Orbit) → Extensions → Apps Script.
   2. Paste this whole file into Code.gs and save.
   3. Edit ADMIN below (your login), then Run ▸ setup  (approve permissions once).
   4. Deploy ▸ New deployment ▸ Web app — Execute as: Me, Who has access: Anyone.
   5. Copy the web app URL into CONFIG.API_URL at the top of index.html.
   ===================================================================== */

var ADMIN = {
  propertyCode: 'DXBAI',
  propertyName: 'Embassy Suites by Hilton Dubai Business Bay',
  teamName: 'Marketing',
  username: 'strivedi6',
  name: 'Shrikant Trivedi',
  jobTitle: 'Cluster Marketing Manager',
  pin: '0000'            // put your own PIN here before running setup(); never commit the real one
};
// Extra properties and teams created by setup(). Add or remove freely; you can also add them later in the app.
var SEED = {
  DXBAI: { name: 'Embassy Suites by Hilton Dubai Business Bay', teams: ['Marketing', 'Sales', 'Food & Beverage', 'Front Office', 'Human Resources', 'Finance', 'Engineering', 'Housekeeping'] },
  DXBPI: { name: 'Hilton Garden Inn Dubai Business Bay', teams: ['Marketing', 'Sales', 'Food & Beverage', 'Front Office', 'Human Resources', 'Finance', 'Engineering', 'Housekeeping'] }
};
// Your Drive: Orbit › Orbit Data (database) and Orbit › Orbit Uploads (task files)
var SPREADSHEET_ID = '1GxBNh_12qWoX-00D-M0Ibu-wAq8Hqr4AkUiWRrOO_ME';
var UPLOAD_FOLDER_ID = '1_yy9_fzZVAOIixQxLBDJosERv94NAX1k';
var TEAM_COLORS = ['blue', 'orange', 'red', 'teal', 'purple', 'green', 'indigo', 'brown'];

/* ---------------- web app entry points ---------------- */
function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Bad request.' }); }
  var lock = LockService.getScriptLock();
  var readOnly = req.action === 'sync' || req.action === 'teams';
  if (!readOnly) lock.waitLock(25000);
  try { SheetDB.reset(); return json_(coreHandle(SheetDB, req)); }
  finally { if (!readOnly) lock.releaseLock(); }
}
function doGet() {
  return ContentService.createTextOutput('Orbit backend is running. The app talks to this URL with POST requests.');
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------------- one-time setup ---------------- */
function setup() {
  SheetDB.reset();
  Object.keys(TABLES).forEach(function (t) { SheetDB.sheet(t); });
  Object.keys(SEED).forEach(function (code) {
    var p = findProperty(SheetDB, code);
    if (!p) SheetDB.insert('properties', { id: uid('h'), code: code, name: SEED[code].name, createdAt: nowISO() });
    SEED[code].teams.forEach(function (name, i) {
      var ex = SheetDB.find('teams', function (t) { return t.propertyCode === code && t.name === name; });
      if (!ex) SheetDB.insert('teams', { id: uid('g'), propertyCode: code, name: name, color: TEAM_COLORS[i % TEAM_COLORS.length], createdAt: nowISO() });
    });
  });
  var u = seedAdmin(SheetDB, ADMIN);
  var book = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = book.getSheetByName('Sheet1'); if (sh && book.getSheets().length > 1) book.deleteSheet(sh);
  Logger.log('Ready. Main admin: ' + ADMIN.propertyCode + ' / ' + u.username + '. Now deploy as a web app.');
}

/* ---------------- Google Sheets storage adapter ---------------- */
var SheetDB = (function () {
  var cache = {};
  var book = null;
  function ss() { return book || (book = SpreadsheetApp.openById(SPREADSHEET_ID)); }
  function sheet(t) {
    var sh = ss().getSheetByName(t);
    if (!sh) {
      sh = ss().insertSheet(t);
      sh.getRange(1, 1, sh.getMaxRows(), TABLES[t].length).setNumberFormat('@');
      sh.getRange(1, 1, 1, TABLES[t].length).setValues([TABLES[t]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    return sh;
  }
  function load(t) {
    if (cache[t]) return cache[t];
    var sh = sheet(t), vals = sh.getDataRange().getValues(), head = vals[0], rows = [];
    // Add any columns a newer version of the app needs, at the end of the header row.
    var missing = TABLES[t].filter(function (k) { return head.indexOf(k) < 0; });
    if (missing.length) {
      var start = head.filter(function (h) { return h !== ''; }).length;
      head = head.slice(0, start).concat(missing);
      sh.getRange(1, 1, sh.getMaxRows(), head.length).setNumberFormat('@');
      sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
      vals = sh.getDataRange().getValues();
    }
    for (var i = 1; i < vals.length; i++) {
      if (vals[i].join('') === '') continue;
      var o = {};
      for (var j = 0; j < head.length; j++) {
        var k = head[j], v = vals[i][j];
        if (JSON_COLS[k]) { try { v = v ? JSON.parse(v) : (k === 'reactions' ? {} : []); } catch (e) { v = k === 'reactions' ? {} : []; } }
        else if (v instanceof Date) v = k === 'due' ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : v.toISOString();
        else if (v === 'true') v = true; else if (v === 'false') v = false;
        else if (k === 'order' || k === 'failed') v = Number(v) || 0;
        else v = v === null || v === undefined ? '' : String(v);
        o[k] = v;
      }
      o.__row = i + 1;
      rows.push(o);
    }
    cache[t] = { head: head, rows: rows, sh: sh };
    return cache[t];
  }
  function strip(o) { if (!o) return null; var c = JSON.parse(JSON.stringify(o)); delete c.__row; return c; }
  function toRow(head, o) {
    return head.map(function (k) { var v = o[k]; if (JSON_COLS[k]) return JSON.stringify(v || (k === 'reactions' ? {} : [])); if (v === null || v === undefined) return ''; return typeof v === 'boolean' ? String(v) : v; });
  }
  return {
    reset: function () { cache = {}; },
    sheet: sheet,
    all: function (t) { return load(t).rows.map(strip); },
    get: function (t, id) { var r = load(t).rows.filter(function (x) { return x.id === id; })[0]; return strip(r); },
    find: function (t, fn) { var rows = load(t).rows; for (var i = 0; i < rows.length; i++) if (fn(rows[i])) return strip(rows[i]); return null; },
    insert: function (t, o) {
      var c = load(t); var row = toRow(c.head, o);
      c.sh.appendRow(row);
      var n = c.sh.getLastRow(); c.sh.getRange(n, 1, 1, row.length).setNumberFormat('@').setValues([row]);
      var copy = JSON.parse(JSON.stringify(o)); copy.__row = n; c.rows.push(copy);
    },
    update: function (t, id, patch) {
      var c = load(t), r = c.rows.filter(function (x) { return x.id === id; })[0]; if (!r) return;
      Object.keys(patch).forEach(function (k) { r[k] = patch[k]; });
      c.sh.getRange(r.__row, 1, 1, c.head.length).setValues([toRow(c.head, r)]);
    },
    remove: function (t, id) {
      var c = load(t), r = c.rows.filter(function (x) { return x.id === id; })[0]; if (!r) return;
      c.sh.deleteRow(r.__row); delete cache[t];
    },
    saveFile: function (name, type, b64) {
      var folder = DriveApp.getFolderById(UPLOAD_FOLDER_ID);
      var file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), type, name));
      try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* some Workspace domains block link sharing */ }
      return { url: 'https://drive.google.com/file/d/' + file.getId() + '/view', fileId: file.getId() };
    }
  };
})();

/* =====================================================================
   ORBIT CORE — shared business logic.
   The same code runs in the browser (demo mode, data on this device)
   and inside Google Apps Script (live mode, data in Google Sheets).
   A storage adapter "DB" is passed in with:
     all(t) get(t,id) find(t,fn) insert(t,o) update(t,id,patch) remove(t,id)
     saveFile(name,type,base64) -> {url,fileId}
   ===================================================================== */

var TABLES = {
  properties: ['id','code','name','createdAt','color','logo'],
  teams:      ['id','propertyCode','name','color','createdAt'],
  users:      ['id','propertyCode','teamId','username','name','jobTitle','role','status','pinHash','salt','failed','lockedUntil','avatarColor','createdAt','approvedBy','approvedAt','photo','photoV','teamIds'],
  sessions:   ['id','userId','createdAt','lastSeen'],
  projects:   ['id','teamId','name','color','icon','description','ownerId','status','due','createdBy','createdAt','updatedAt','archived','stages','shareTeams','access','members','managers'],
  tasks:      ['id','teamId','projectId','parentId','title','notes','status','priority','assignees','due','createdBy','createdAt','updatedAt','completedAt','order','attachments','repeat','nextId','repeatOn','subStages'],
  comments:   ['id','teamId','taskId','userId','text','createdAt','parentId','reactions','editedAt'],
  activity:   ['id','teamId','userId','type','refId','text','targets','createdAt']
};
var JSON_COLS = { assignees: 1, attachments: 1, targets: 1, reactions: 1, teamIds: 1, stages: 1, subStages: 1, shareTeams: 1, mentions: 1, members: 1, managers: 1 };
var REACTIONS = ['👍', '❤️', '😂', '🎉', '👀', '✅', '🙏', '🔥'];
var DEFAULT_STAGES = [
  { id: 'todo', name: 'To Do', color: 'gray' }, { id: 'doing', name: 'In Progress', color: 'blue' },
  { id: 'review', name: 'In Review', color: 'purple' }, { id: 'followup', name: 'Follow Up', color: 'orange' },
  { id: 'done', name: 'Done', color: 'green' }];
/* Repeating tasks: when one is completed, the next one is created with the next due date. */
var REPEATS = ['daily', 'weekdays', 'weekly', 'biweekly', 'monthly', 'quarterly', 'yearly'];
function ymd(d) { return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
function addMonths(d, n) { var day = d.getDate(), r = new Date(d.getFullYear(), d.getMonth() + n, 1); r.setDate(Math.min(day, new Date(r.getFullYear(), r.getMonth() + 1, 0).getDate())); return r; }
function stepDate(d, rep) {
  var r = new Date(d.getTime());
  if (rep === 'daily') r.setDate(r.getDate() + 1);
  else if (rep === 'weekdays') { do { r.setDate(r.getDate() + 1); } while (r.getDay() === 0 || r.getDay() === 6); }
  else if (rep === 'weekly') r.setDate(r.getDate() + 7);
  else if (rep === 'biweekly') r.setDate(r.getDate() + 14);
  else if (rep === 'monthly') r = addMonths(d, 1);
  else if (rep === 'quarterly') r = addMonths(d, 3);
  else if (rep === 'yearly') r = addMonths(d, 12);
  return r;
}
/* "on" refines the frequency: weekdays for weekly ("1,3" = Mon, Wed), a day of the month for monthly/quarterly ("15" or "last"). */
function monthDay(y, m, on) { var last = new Date(y, m + 1, 0).getDate(); return new Date(y, m, on === 'last' ? last : Math.min(parseInt(on, 10), last)); }
function weekStart(d) { var r = new Date(d.getTime()); r.setDate(r.getDate() - ((r.getDay() + 6) % 7)); r.setHours(0, 0, 0, 0); return r.getTime(); }
function stepOn(d, rep, on) {
  if ((rep === 'weekly' || rep === 'biweekly') && on) {
    var days = String(on).split(',').map(Number).filter(function (x) { return x >= 0 && x <= 6; });
    if (days.length) {
      var r = new Date(d.getTime());
      for (var i = 0; i < 7; i++) { r.setDate(r.getDate() + 1); if (days.indexOf(r.getDay()) >= 0) break; }
      if (rep === 'biweekly' && weekStart(r) !== weekStart(d)) r.setDate(r.getDate() + 7);
      return r;
    }
  }
  if ((rep === 'monthly' || rep === 'quarterly') && on && (on === 'last' || /^\d{1,2}$/.test(String(on)))) {
    var step = rep === 'monthly' ? 1 : 3;
    var same = monthDay(d.getFullYear(), d.getMonth(), on);
    if (same > d) return same;
    return monthDay(d.getFullYear(), d.getMonth() + step, on);
  }
  return stepDate(d, rep);
}
function nextDue(due, rep, today, on) {
  var t = today ? new Date(today + 'T00:00:00') : new Date(); t.setHours(0, 0, 0, 0);
  var d = due ? new Date(due + 'T00:00:00') : new Date(t.getTime());
  var n = stepOn(d, rep, on), guard = 0;
  while (n < t && guard++ < 1000) n = stepOn(n, rep, on);
  if (n <= t && !due) n = stepOn(t, rep, on);
  return ymd(n);
}
/* First date on or after today for a new repeating task. */
function firstDue(rep, today, on) { var t = new Date(today + 'T00:00:00'); t.setDate(t.getDate() - 1); return nextDue(ymd(t), rep, today, on); }
function stagesOf(p) { return p && Array.isArray(p.stages) && p.stages.length ? p.stages : DEFAULT_STAGES; }
/* A stage list either ends in a Done column (completing moves the task there), or has none: then the stages are
   groups such as outlets, and completing only ticks the task, which stays where it is. */
function hasDoneStage(list) { return (list || []).some(function (s) { return s.id === 'done'; }); }
function taskDoneIn(t, list) { return hasDoneStage(list) ? t.status === 'done' : !!t.completedAt; }
var PRIORITIES = ['none','low','medium','high','urgent'];
var PROJECT_STATUSES = ['on_track','at_risk','off_track','on_hold','complete'];
var AVATAR_COLORS = ['blue','indigo','purple','pink','red','orange','green','teal','brown'];
var MAX_UPLOAD_BYTES = 10 * 1024 * 1024;
var MAX_IMAGE_CHARS = 48000; // one Google Sheets cell holds 50,000 characters
var ID_RE = /^[a-z][a-z0-9]{6,40}$/;
function cleanId(id, prefix) { return ID_RE.test(String(id || '')) ? String(id) : uid(prefix); }
function cleanImage(v) {
  v = String(v || '');
  if (!v) return '';
  if (!/^data:image\/(png|jpeg|webp|gif);base64,/.test(v)) fail('Choose a PNG, JPG or WebP image.');
  if (v.length > MAX_IMAGE_CHARS) fail('That image is too large. Try a smaller one.');
  return v;
}
function publicProperty(p) { return { id: p.id, code: p.code, name: p.name, color: p.color || 'blue', logo: p.logo || '' }; }

function sha256(ascii) {
  function rr(v, a) { return (v >>> a) | (v << (32 - a)); }
  var mp = Math.pow, maxWord = mp(2, 32), i, j, result = '';
  var words = [], bitLen = ascii.length * 8;
  var hash = [], k = [], pc = 0, isComposite = {};
  for (var c = 2; pc < 64; c++) {
    if (!isComposite[c]) {
      for (i = 0; i < 313; i += c) isComposite[i] = c;
      hash[pc] = (mp(c, .5) * maxWord) | 0;
      k[pc++] = (mp(c, 1 / 3) * maxWord) | 0;
    }
  }
  hash = hash.slice(0, 8);
  ascii += '\x80';
  while (ascii.length % 64 - 56) ascii += '\x00';
  for (i = 0; i < ascii.length; i++) {
    j = ascii.charCodeAt(i);
    if (j >> 8) return '';
    words[i >> 2] |= j << ((3 - i) % 4) * 8;
  }
  words[words.length] = ((bitLen / maxWord) | 0);
  words[words.length] = (bitLen);
  for (j = 0; j < words.length;) {
    var w = words.slice(j, j += 16), old = hash;
    hash = hash.slice(0, 8);
    for (i = 0; i < 64; i++) {
      var w15 = w[i - 15], w2 = w[i - 2], a = hash[0], e = hash[4];
      var t1 = hash[7] + (rr(e, 6) ^ rr(e, 11) ^ rr(e, 25)) + ((e & hash[5]) ^ ((~e) & hash[6])) + k[i] +
        (w[i] = (i < 16) ? w[i] : (w[i - 16] + (rr(w15, 7) ^ rr(w15, 18) ^ (w15 >>> 3)) + w[i - 7] + (rr(w2, 17) ^ rr(w2, 19) ^ (w2 >>> 10))) | 0);
      var t2 = (rr(a, 2) ^ rr(a, 13) ^ rr(a, 22)) + ((a & hash[1]) ^ (a & hash[2]) ^ (hash[1] & hash[2]));
      hash = [(t1 + t2) | 0].concat(hash);
      hash[4] = (hash[4] + t1) | 0;
    }
    for (i = 0; i < 8; i++) hash[i] = (hash[i] + old[i]) | 0;
  }
  for (i = 0; i < 8; i++) for (j = 3; j + 1; j--) { var b = (hash[i] >> (j * 8)) & 255; result += ((b < 16) ? 0 : '') + b.toString(16); }
  return result;
}

function uid(p) { return (p || '') + Date.now().toString(36) + Math.random().toString(36).slice(2, 9); }
function nowISO() { return new Date().toISOString(); }
function fail(m) { throw new Error(m); }
function str(v, max) { v = (v == null ? '' : String(v)).trim(); return max ? v.slice(0, max) : v; }
function lc(v) { return str(v).toLowerCase(); }
function truthy(v) { return v === true || v === 'true' || v === 1 || v === '1'; }
function pinHash(salt, username, pin) { return sha256(String(salt) + '|' + lc(username) + '|' + String(pin)); }
function makeToken() { return 't' + sha256(uid() + Math.random() + Date.now()).slice(0, 40); }

function publicUser(u) {
  return { id: u.id, propertyCode: u.propertyCode, teamId: u.teamId, username: u.username, name: u.name,
    jobTitle: u.jobTitle, role: u.role, status: u.status, avatarColor: u.avatarColor, createdAt: u.createdAt,
    approvedBy: u.approvedBy || '', approvedAt: u.approvedAt || '', teamIds: Array.isArray(u.teamIds) ? u.teamIds : [], photoV: u.photo ? String(u.photoV || '1') : '' };
}
function isAdmin(u) { return u.role === 'admin'; }
/* For the main admin: account state that isn't otherwise shared (locks, signed-in devices, last active). */
function adminInfo(DB) {
  var info = {}, now = Date.now(), sessions = DB.all('sessions'), cut = now - 90 * 864e5, stale = 0;
  DB.all('users').forEach(function (x) { info[x.id] = { locked: !!(x.lockedUntil && new Date(x.lockedUntil).getTime() > now), failed: Number(x.failed) || 0, devices: 0, lastSeen: '' }; });
  sessions.forEach(function (s) {
    var i = info[s.userId]; if (!i) { stale++; return; }
    var seen = s.lastSeen || s.createdAt; if (new Date(seen).getTime() < cut) stale++;
    i.devices++; if (!i.lastSeen || seen > i.lastSeen) i.lastSeen = seen;
  });
  return { users: info, sessions: sessions.length, staleSessions: stale };
}
/* A person's teams: their home team plus any extra teams (cluster colleagues working across properties). */
function userTeams(u) {
  var l = [u.teamId];
  (Array.isArray(u.teamIds) ? u.teamIds : []).forEach(function (t) { if (t && l.indexOf(t) < 0) l.push(t); });
  return l;
}
function canSeeTeam(u, teamId) { return isAdmin(u) || userTeams(u).indexOf(teamId) >= 0; }
function canManageTeam(u, teamId) { return isAdmin(u) || (u.role === 'head' && userTeams(u).indexOf(teamId) >= 0); }
/* A project belongs to one team and can be shared with teams at other properties. */
function projTeams(p) { return [p.teamId].concat(Array.isArray(p.shareTeams) ? p.shareTeams : []); }
/* Projects are open to their teams, or limited to chosen people. Managers edit the project, its stages and can archive it. */
function listOf(v) { return Array.isArray(v) ? v : []; }
function isProjectManager(u, p) { return p.createdBy === u.id || listOf(p.managers).indexOf(u.id) >= 0; }
function canSeeProject(u, p) {
  if (isAdmin(u) || isProjectManager(u, p)) return true;
  if (p.access === 'private') return listOf(p.members).indexOf(u.id) >= 0;
  var ts = userTeams(u); return projTeams(p).some(function (t) { return ts.indexOf(t) >= 0; });
}
function canManageProject(u, p) { return isAdmin(u) || isProjectManager(u, p) || projTeams(p).some(function (t) { return canManageTeam(u, t); }); }
function projectPeople(DB, p) {
  if (p.access === 'private') return [p.createdBy].concat(listOf(p.members), listOf(p.managers));
  var allowed = projTeams(p);
  return DB.all('users').filter(function (x) { return x.status === 'active' && userTeams(x).some(function (t) { return allowed.indexOf(t) >= 0; }); })
    .map(function (x) { return x.id; }).concat(listOf(p.managers));
}
function canSeeTask(DB, u, t) {
  var p = t.projectId ? DB.get('projects', t.projectId) : null;
  return p ? canSeeProject(u, p) : canSeeTeam(u, t.teamId);
}
function findProperty(DB, code) { code = str(code).toUpperCase(); return DB.find('properties', function (p) { return String(p.code).toUpperCase() === code; }); }

function auth(DB, token) {
  if (!token) fail('AUTH');
  var s = DB.get('sessions', token);
  if (!s) fail('AUTH');
  var u = DB.get('users', s.userId);
  if (!u || u.status !== 'active') fail('AUTH');
  // remember when each device was last used (at most every 30 minutes, to keep writes low)
  if (!s.lastSeen || Date.now() - new Date(s.lastSeen).getTime() > 30 * 60000) DB.update('sessions', s.id, { lastSeen: nowISO() });
  return u;
}
function log(DB, u, teamId, type, refId, text, targets) {
  DB.insert('activity', { id: uid('a'), teamId: teamId, userId: u.id, type: type, refId: refId || '', text: text,
    targets: (targets || []).filter(function (t) { return t && t !== u.id; }), createdAt: nowISO() });
}
function headsOf(DB, teamId) {
  return DB.all('users').filter(function (x) { return x.status === 'active' && (x.role === 'admin' || (x.role === 'head' && x.teamId === teamId)); })
    .map(function (x) { return x.id; });
}
function taskOrFail(DB, u, id) {
  var t = DB.get('tasks', id);
  if (!t || !canSeeTask(DB, u, t)) fail('That task no longer exists.');
  return t;
}

var ACTIONS = {
  /* ---------- public: no login needed ---------- */
  teams: function (DB, r) {
    var p = findProperty(DB, r.propertyCode);
    if (!p) fail('We could not find property code "' + str(r.propertyCode).toUpperCase() + '". Check the code with your department head.');
    var teams = DB.all('teams').filter(function (t) { return t.propertyCode === p.code; })
      .map(function (t) { return { id: t.id, name: t.name, color: t.color }; })
      .sort(function (a, b) { return a.name < b.name ? -1 : 1; });
    return { property: publicProperty(p), teams: teams };
  },

  signup: function (DB, r) {
    var p = findProperty(DB, r.propertyCode);
    if (!p) fail('Unknown property code.');
    var team = DB.get('teams', r.teamId);
    if (!team || team.propertyCode !== p.code) fail('Choose a team for this property.');
    var name = str(r.name, 60), username = lc(r.username), jobTitle = str(r.jobTitle, 60), pin = str(r.pin);
    if (name.length < 2) fail('Enter your full name.');
    if (!/^[a-z0-9._-]{3,20}$/.test(username)) fail('Username must be 3–20 letters, numbers, dots or dashes.');
    if (!/^\d{4}$/.test(pin)) fail('Your PIN must be exactly 4 digits.');
    var taken = DB.find('users', function (x) { return x.propertyCode === p.code && lc(x.username) === username; });
    if (taken) fail('That username is already taken at ' + p.code + '. Try another.');
    var salt = uid('s');
    var u = { id: uid('u'), propertyCode: p.code, teamId: team.id, username: username, name: name, jobTitle: jobTitle,
      role: 'member', status: 'pending', pinHash: pinHash(salt, username, pin), salt: salt, failed: 0, lockedUntil: '',
      avatarColor: AVATAR_COLORS[Math.floor(Math.random() * AVATAR_COLORS.length)], createdAt: nowISO(), approvedBy: '', approvedAt: '' };
    DB.insert('users', u);
    DB.insert('activity', { id: uid('a'), teamId: team.id, userId: u.id, type: 'signup', refId: u.id,
      text: 'asked to join ' + team.name, targets: headsOf(DB, team.id), createdAt: nowISO() });
    return { status: 'pending', team: team.name, property: p.code };
  },

  login: function (DB, r) {
    var p = findProperty(DB, r.propertyCode);
    var username = lc(r.username), pin = str(r.pin);
    if (!p) fail('Unknown property code.');
    var u = DB.find('users', function (x) { return x.propertyCode === p.code && lc(x.username) === username; });
    if (!u) fail('Username or PIN is incorrect.');
    if (u.lockedUntil && new Date(u.lockedUntil) > new Date()) fail('Too many wrong PINs. Try again after 15 minutes, or ask your department head to reset your PIN.');
    if (pinHash(u.salt, u.username, pin) !== u.pinHash) {
      var f = (Number(u.failed) || 0) + 1;
      DB.update('users', u.id, { failed: f, lockedUntil: f >= 5 ? new Date(Date.now() + 15 * 60000).toISOString() : '' });
      fail(f >= 5 ? 'Too many wrong PINs. Your account is locked for 15 minutes.' : 'Username or PIN is incorrect.');
    }
    if (u.status === 'pending') fail('PENDING');
    if (u.status === 'disabled') fail('Your account has been turned off. Contact your department head.');
    DB.update('users', u.id, { failed: 0, lockedUntil: '' });
    var token = makeToken();
    DB.insert('sessions', { id: token, userId: u.id, createdAt: nowISO(), lastSeen: nowISO() });
    return { token: token, user: publicUser(u), photo: u.photo || '', data: ACTIONS.sync(DB, { token: token }) };
  },

  logout: function (DB, r) {
    if (r.token && DB.get('sessions', r.token)) DB.remove('sessions', r.token);
    return { ok: true };
  },

  /* ---------- everything the signed-in person can see ---------- */
  sync: function (DB, r) {
    var u = auth(DB, r.token), admin = isAdmin(u), mine = userTeams(u);
    var projects = DB.all('projects').filter(function (p) { return canSeeProject(u, p); });
    var pIds = {}, tset = {};
    mine.forEach(function (t) { tset[t] = 1; });
    projects.forEach(function (p) { pIds[p.id] = 1; projTeams(p).forEach(function (t) { tset[t] = 1; }); });
    var tasks = DB.all('tasks').filter(function (t) { return admin || (t.projectId ? pIds[t.projectId] : mine.indexOf(t.teamId) >= 0); });
    var tIds = {}; tasks.forEach(function (t) { tIds[t.id] = 1; });
    var users = DB.all('users').filter(function (x) {
      if (admin) return true;
      var xt = userTeams(x);
      if (!xt.some(function (t) { return tset[t]; })) return false;
      if (x.status === 'active') return true;
      return u.role === 'head' && xt.some(function (t) { return mine.indexOf(t) >= 0; });
    }).map(publicUser);
    var activity = DB.all('activity').filter(function (a) {
      return admin || mine.indexOf(a.teamId) >= 0 || tIds[a.refId] || (a.targets || []).indexOf(u.id) >= 0;
    }).sort(function (a, b) { return a.createdAt < b.createdAt ? 1 : -1; }).slice(0, 250);
    return {
      me: publicUser(u),
      properties: DB.all('properties').map(publicProperty),
      teams: DB.all('teams'), users: users,
      projects: projects, tasks: tasks,
      comments: DB.all('comments').filter(function (c) { return tIds[c.taskId]; }),
      activity: activity,
      admin: admin ? adminInfo(DB) : null,
      serverTime: nowISO()
    };
  },

  /* ---------- projects ---------- */
  saveProject: function (DB, r) {
    var u = auth(DB, r.token), p = r.project || {};
    var fields = { name: str(p.name, 80), color: str(p.color, 20) || 'blue', icon: str(p.icon, 8), description: str(p.description, 2000),
      ownerId: str(p.ownerId), status: PROJECT_STATUSES.indexOf(p.status) >= 0 ? p.status : 'on_track', due: str(p.due, 10),
      archived: truthy(p.archived), updatedAt: nowISO() };
    if (!fields.name) fail('Give the project a name.');
    // Only the main admin can share with any property; everyone else only with properties they belong to
    // (a share that is already on the project is kept).
    var myProps = userTeams(u).map(function (id) { return (DB.get('teams', id) || {}).propertyCode; });
    var shares = function (home, before) {
      var seen = {};
      return (Array.isArray(p.shareTeams) ? p.shareTeams : []).filter(function (id) {
        var t = id && DB.get('teams', id);
        if (!t || id === home || seen[id]) return false;
        if (!isAdmin(u) && myProps.indexOf(t.propertyCode) < 0 && (before || []).indexOf(id) < 0) return false;
        seen[id] = 1; return true;
      });
    };
    var people = function (list) { return listOf(list).filter(function (id, i, a) { return a.indexOf(id) === i && DB.get('users', id); }); };
    if (p.access !== undefined) fields.access = p.access === 'private' ? 'private' : 'team';
    if (p.members !== undefined) fields.members = people(p.members);
    if (p.managers !== undefined) fields.managers = people(p.managers);
    var existing = p.id ? DB.get('projects', p.id) : null;
    if (existing) {
      if (!canSeeProject(u, existing)) fail('That project no longer exists.');
      if (!canManageProject(u, existing)) fail('Only this project’s managers or a department head can change it.');
      if (p.shareTeams !== undefined) fields.shareTeams = shares(existing.teamId, existing.shareTeams);
      DB.update('projects', existing.id, fields);
      return DB.get('projects', existing.id);
    }
    var teamId = p.teamId && canSeeTeam(u, p.teamId) ? p.teamId : u.teamId;
    if (!DB.get('teams', teamId)) fail('Choose a team for this project.');
    fields.shareTeams = shares(teamId);
    var obj = Object.assign({ id: cleanId(p.id, 'p'), teamId: teamId, createdBy: u.id, createdAt: nowISO(), stages: [], access: 'team', members: [], managers: [] }, fields);
    if (!obj.ownerId) obj.ownerId = u.id;
    DB.insert('projects', obj);
    log(DB, u, teamId, 'project', obj.id, 'created the project ' + obj.name, []);
    return obj;
  },

  deleteProject: function (DB, r) {
    var u = auth(DB, r.token), p = DB.get('projects', r.id);
    if (!p) return { ok: true };
    if (!canSeeProject(u, p)) fail('That project no longer exists.');
    if (!canManageProject(u, p)) fail('Only this project’s managers or a department head can delete it.');
    DB.all('tasks').filter(function (t) { return t.projectId === p.id; }).forEach(function (t) {
      DB.all('comments').filter(function (c) { return c.taskId === t.id; }).forEach(function (c) { DB.remove('comments', c.id); });
      DB.remove('tasks', t.id);
    });
    DB.remove('projects', p.id);
    log(DB, u, p.teamId, 'project', '', 'deleted the project ' + p.name, []);
    return { ok: true };
  },

  /* ---------- tasks & subtasks ---------- */
  saveTask: function (DB, r) {
    var u = auth(DB, r.token), t = r.task || {};
    var existing = t.id ? DB.get('tasks', t.id) : null;
    if (existing) { if (!canSeeTask(DB, u, existing)) fail('That task no longer exists.'); existing = JSON.parse(JSON.stringify(existing)); }
    var base = existing || {};
    var projectId = t.projectId !== undefined ? str(t.projectId) : base.projectId;
    var project = projectId ? DB.get('projects', projectId) : null;
    if (projectId && (!project || !canSeeProject(u, project))) fail('Choose a project you belong to.');
    var parent = t.parentId || base.parentId ? DB.get('tasks', t.parentId || base.parentId) : null;
    var teamId = project ? project.teamId : (parent ? parent.teamId : (existing ? existing.teamId : u.teamId));
    if (!project && !parent && !existing && t.teamId && canSeeTeam(u, t.teamId)) teamId = t.teamId;
    if (project ? !canSeeProject(u, project) : parent ? !canSeeTask(DB, u, parent) : !canSeeTeam(u, teamId)) fail('You can’t add tasks to that team.');
    var team = project ? projectPeople(DB, project) : DB.all('users').filter(function (x) { return x.status === 'active' && userTeams(x).indexOf(teamId) >= 0; }).map(function (x) { return x.id; });
    var stageList = parent ? stagesOf({ stages: parent.subStages }) : stagesOf(project);
    var stageIds = stageList.map(function (s) { return s.id; }), doneCol = hasDoneStage(stageList);
    var admins = DB.all('users').filter(function (x) { return x.role === 'admin'; }).map(function (x) { return x.id; });

    var f = {};
    if (t.title !== undefined) f.title = str(t.title, 300);
    if (t.notes !== undefined) f.notes = str(t.notes, 10000);
    if (t.status !== undefined) f.status = stageIds.indexOf(t.status) >= 0 ? t.status : stageIds[0];
    if (t.priority !== undefined) f.priority = PRIORITIES.indexOf(t.priority) >= 0 ? t.priority : 'none';
    if (t.due !== undefined) f.due = str(t.due, 10);
    if (t.repeat !== undefined) f.repeat = REPEATS.indexOf(t.repeat) >= 0 ? t.repeat : '';
    if (t.repeatOn !== undefined) f.repeatOn = /^(last|[0-9,]{1,20})$/.test(String(t.repeatOn)) ? String(t.repeatOn) : '';
    if (t.order !== undefined) f.order = Number(t.order) || 0;
    if (t.projectId !== undefined) f.projectId = projectId || '';
    if (t.assignees !== undefined) f.assignees = (t.assignees || []).filter(function (id) { return team.indexOf(id) >= 0 || admins.indexOf(id) >= 0; });
    f.teamId = teamId; f.updatedAt = nowISO();
    if (doneCol) { if (f.status) f.completedAt = f.status === 'done' ? (base.status === 'done' ? base.completedAt : nowISO()) : ''; }
    else if (t.done !== undefined) f.completedAt = t.done ? (base.completedAt || nowISO()) : '';
    var wasDone = existing ? taskDoneIn(existing, stageList) : false;

    if (existing) {
      if (f.title === '') fail('A task needs a title.');
      DB.update('tasks', existing.id, f);
      var after = DB.get('tasks', existing.id);
      var added = (f.assignees || []).filter(function (id) { return (existing.assignees || []).indexOf(id) < 0; });
      if (added.length) log(DB, u, teamId, 'assign', existing.id, 'assigned “' + after.title + '”', added);
      if (!wasDone && taskDoneIn(after, stageList)) {
        log(DB, u, teamId, 'complete', existing.id, 'completed “' + after.title + '”', [after.createdBy].concat(after.assignees || []));
        if (after.repeat && !after.nextId) {
          var nid = cleanId(t.spawnId, 't');
          if (!DB.get('tasks', nid)) {
            var nx = { id: nid, teamId: after.teamId, projectId: after.projectId || '', parentId: after.parentId || '', title: after.title, notes: after.notes || '',
              status: doneCol ? stageIds[0] : after.status, priority: after.priority || 'none', assignees: after.assignees || [], due: nextDue(after.due, after.repeat, t.today, after.repeatOn), repeatOn: after.repeatOn || '', subStages: after.subStages || [],
              createdBy: after.createdBy, createdAt: nowISO(), updatedAt: nowISO(), completedAt: '', order: Date.now(), attachments: [], repeat: after.repeat, nextId: '' };
            DB.insert('tasks', nx);
            DB.all('tasks').filter(function (x) { return x.parentId === after.id; }).forEach(function (s, i) {
              DB.insert('tasks', { id: uid('t'), teamId: s.teamId, projectId: s.projectId || '', parentId: nid, title: s.title, notes: s.notes || '', status: stagesOf({ stages: after.subStages })[0].id,
                priority: s.priority || 'none', assignees: s.assignees || [], due: '', createdBy: s.createdBy, createdAt: nowISO(), updatedAt: nowISO(), completedAt: '',
                order: Date.now() + i, attachments: [], repeat: '', nextId: '' });
            });
          }
          DB.update('tasks', after.id, { nextId: nid });
          after = DB.get('tasks', after.id);
        }
      }
      return after;
    }
    if (!f.title) fail('A task needs a title.');
    var obj = Object.assign({ id: cleanId(t.id, 't'), parentId: parent ? parent.id : '', projectId: projectId || '', notes: '', status: 'todo',
      priority: 'none', assignees: [], due: '', order: Date.now(), attachments: [], completedAt: '',
      createdBy: u.id, createdAt: nowISO() }, f);
    DB.insert('tasks', obj);
    if ((obj.assignees || []).length) log(DB, u, teamId, 'assign', obj.id, 'assigned “' + obj.title + '”', obj.assignees);
    else if (!parent) log(DB, u, teamId, 'task', obj.id, 'added “' + obj.title + '”', []);
    return obj;
  },

  deleteTask: function (DB, r) {
    var u = auth(DB, r.token), t = DB.get('tasks', r.id);
    if (!t) return { ok: true };
    if (!canSeeTask(DB, u, t)) fail('That task no longer exists.');
    if (!(canManageTeam(u, t.teamId) || t.createdBy === u.id)) fail('Only the person who created this task or a department head can delete it.');
    var kill = function (id) {
      DB.all('tasks').filter(function (x) { return x.parentId === id; }).forEach(function (x) { kill(x.id); });
      DB.all('comments').filter(function (c) { return c.taskId === id; }).forEach(function (c) { DB.remove('comments', c.id); });
      DB.remove('tasks', id);
    };
    kill(t.id);
    return { ok: true };
  },

  addComment: function (DB, r) {
    var u = auth(DB, r.token), t = taskOrFail(DB, u, r.taskId), text = str(r.text, 4000);
    if (!text) fail('Write a comment first.');
    if (r.id && DB.get('comments', r.id)) return DB.get('comments', r.id);
    var parent = r.parentId ? DB.get('comments', r.parentId) : null;
    if (parent && parent.parentId) parent = DB.get('comments', parent.parentId) || parent;
    var c = { id: cleanId(r.id, 'c'), teamId: t.teamId, taskId: t.id, userId: u.id, text: text, createdAt: nowISO(), parentId: parent ? parent.id : '', reactions: {} };
    DB.insert('comments', c);
    var mentioned = (Array.isArray(r.mentions) ? r.mentions : []).filter(function (id) {
      var x = DB.get('users', id); return x && x.id !== u.id && x.status === 'active' && canSeeTask(DB, x, t);
    });
    if (mentioned.length) log(DB, u, t.teamId, 'mention', t.id, 'mentioned you on “' + t.title + '”', mentioned);
    var others = function (l) { return l.filter(function (id) { return mentioned.indexOf(id) < 0; }); };
    if (parent) log(DB, u, t.teamId, 'comment', t.id, 'replied to a comment on “' + t.title + '”', others([parent.userId]));
    else log(DB, u, t.teamId, 'comment', t.id, 'commented on “' + t.title + '”', others([t.createdBy].concat(t.assignees || [])));
    return c;
  },

  /* Tap a reaction to add it; tap it again to take it back. */
  reactComment: function (DB, r) {
    var u = auth(DB, r.token), c = DB.get('comments', r.id);
    var ct = c ? DB.get('tasks', c.taskId) : null;
    if (!c || !ct || !canSeeTask(DB, u, ct)) fail('That comment no longer exists.');
    if (REACTIONS.indexOf(r.emoji) < 0) fail('Choose one of the reactions.');
    var re = c.reactions && !Array.isArray(c.reactions) ? c.reactions : {};
    var list = re[r.emoji] || [], i = list.indexOf(u.id);
    if (i >= 0) list.splice(i, 1); else list.push(u.id);
    if (list.length) re[r.emoji] = list; else delete re[r.emoji];
    DB.update('comments', c.id, { reactions: re });
    return DB.get('comments', c.id);
  },

  editComment: function (DB, r) {
    var u = auth(DB, r.token), c = DB.get('comments', r.id), text = str(r.text, 4000);
    if (!c) fail('That comment no longer exists.');
    if (c.userId !== u.id) fail('You can only edit your own comments.');
    if (!text) fail('A comment can’t be empty. Delete it instead.');
    DB.update('comments', c.id, { text: text, editedAt: nowISO() });
    return DB.get('comments', c.id);
  },

  deleteComment: function (DB, r) {
    var u = auth(DB, r.token), c = DB.get('comments', r.id);
    if (!c) return { ok: true };
    if (!(c.userId === u.id || isAdmin(u))) fail('Only the person who wrote this comment or the main admin can delete it.');
    DB.all('comments').filter(function (x) { return x.parentId === c.id; }).forEach(function (x) { DB.remove('comments', x.id); });
    DB.remove('comments', c.id);
    return { ok: true };
  },

  upload: function (DB, r) {
    var u = auth(DB, r.token), t = taskOrFail(DB, u, r.taskId);
    var data = String(r.data || '');
    if (!data) fail('No file received.');
    if (data.length * 0.75 > MAX_UPLOAD_BYTES) fail('Files can be up to 10 MB.');
    var name = str(r.name, 120) || 'file';
    var saved = DB.saveFile(name, str(r.type, 100) || 'application/octet-stream', data);
    var att = { id: uid('f'), name: name, type: str(r.type, 100), size: Number(r.size) || Math.round(data.length * 0.75),
      url: saved.url, fileId: saved.fileId || '', by: u.id, at: nowISO() };
    var list = (t.attachments || []).concat([att]);
    DB.update('tasks', t.id, { attachments: list, updatedAt: nowISO() });
    log(DB, u, t.teamId, 'file', t.id, 'attached ' + name + ' to “' + t.title + '”', [t.createdBy].concat(t.assignees || []));
    return att;
  },

  removeAttachment: function (DB, r) {
    var u = auth(DB, r.token), t = taskOrFail(DB, u, r.taskId);
    var att = (t.attachments || []).filter(function (a) { return a.id === r.attId; })[0];
    if (!att) return { ok: true };
    if (!(att.by === u.id || canManageTeam(u, t.teamId))) fail('Only the person who added this file or a department head can remove it.');
    DB.update('tasks', t.id, { attachments: (t.attachments || []).filter(function (a) { return a.id !== r.attId; }) });
    return { ok: true };
  },

  /* ---------- people & approvals ---------- */
  updateUser: function (DB, r) {
    var u = auth(DB, r.token), target = DB.get('users', r.userId), patch = r.patch || {};
    if (!target) fail('That person no longer exists.');
    if (!userTeams(target).some(function (t) { return canManageTeam(u, t); })) fail('Only the main admin or this team’s department head can change this.');
    if (target.role === 'admin' && !isAdmin(u)) fail('You can’t change the main admin.');
    var f = {};
    if (patch.name !== undefined) f.name = str(patch.name, 60) || target.name;
    if (patch.jobTitle !== undefined) f.jobTitle = str(patch.jobTitle, 60);
    if (patch.photo !== undefined) { f.photo = cleanImage(patch.photo); f.photoV = String(Date.now()); }
    if (patch.role !== undefined) {
      var allowed = isAdmin(u) ? ['member', 'head', 'admin'] : ['member', 'head'];
      if (allowed.indexOf(patch.role) < 0) fail('You can’t give that role.');
      if (target.id === u.id && patch.role !== u.role) fail('You can’t change your own role.');
      f.role = patch.role;
    }
    if (patch.teamId !== undefined && patch.teamId !== target.teamId) {
      if (!isAdmin(u)) fail('Only the main admin can move people between teams.');
      var nt = DB.get('teams', patch.teamId); if (!nt) fail('Choose a team.');
      f.teamId = nt.id; f.propertyCode = nt.propertyCode;
    }
    if (patch.teamIds !== undefined) {
      var home = f.teamId || target.teamId, cur = Array.isArray(target.teamIds) ? target.teamIds : [], mineT = userTeams(u), next = [];
      (Array.isArray(patch.teamIds) ? patch.teamIds : []).forEach(function (id) {
        if (!id || id === home || next.indexOf(id) >= 0 || !DB.get('teams', id)) return;
        if (!isAdmin(u) && cur.indexOf(id) < 0 && mineT.indexOf(id) < 0) fail('You can only give access to teams you manage.');
        next.push(id);
      });
      if (!isAdmin(u)) cur.forEach(function (id) { if (next.indexOf(id) < 0 && mineT.indexOf(id) < 0) next.push(id); });
      f.teamIds = next;
    }
    if (patch.status !== undefined) {
      if (['active', 'disabled'].indexOf(patch.status) < 0) fail('Unknown status.');
      if (target.id === u.id) fail('You can’t turn off your own account.');
      f.status = patch.status;
      if (patch.status === 'active' && target.status === 'pending') { f.approvedBy = u.id; f.approvedAt = nowISO(); }
    }
    DB.update('users', target.id, f);
    if (f.status === 'active' && target.status === 'pending')
      log(DB, u, f.teamId || target.teamId, 'approved', target.id, 'approved ' + target.name + ' to join the team', [target.id]);
    if (f.status === 'disabled') DB.all('sessions').filter(function (s) { return s.userId === target.id; }).forEach(function (s) { DB.remove('sessions', s.id); });
    return publicUser(DB.get('users', target.id));
  },

  /* Main admin or a department head adds someone directly — no approval step. */
  addUser: function (DB, r) {
    var u = auth(DB, r.token), n = r.user || {};
    var team = DB.get('teams', n.teamId);
    if (!team) fail('Choose a team.');
    if (!canManageTeam(u, team.id)) fail('Only the main admin or this team’s department head can add people here.');
    var name = str(n.name, 60), username = lc(n.username), pin = str(n.pin), role = n.role || 'member';
    if (name.length < 2) fail('Enter their full name.');
    if (!/^[a-z0-9._-]{3,20}$/.test(username)) fail('Usernames are 3–20 letters, numbers, dots or dashes.');
    if (!/^\d{4}$/.test(pin)) fail('The PIN must be exactly 4 digits.');
    if ((isAdmin(u) ? ['member', 'head', 'admin'] : ['member', 'head']).indexOf(role) < 0) fail('You can’t give that role.');
    if (DB.find('users', function (x) { return x.propertyCode === team.propertyCode && lc(x.username) === username; })) fail('That username is already taken at ' + team.propertyCode + '.');
    var salt = uid('s');
    var nu = { id: uid('u'), propertyCode: team.propertyCode, teamId: team.id, username: username, name: name, jobTitle: str(n.jobTitle, 60), role: role, status: 'active',
      pinHash: pinHash(salt, username, pin), salt: salt, failed: 0, lockedUntil: '', avatarColor: AVATAR_COLORS[Math.floor(Math.random() * AVATAR_COLORS.length)],
      createdAt: nowISO(), approvedBy: u.id, approvedAt: nowISO(), photo: '', photoV: '' };
    if (n.photo) { nu.photo = cleanImage(n.photo); nu.photoV = String(Date.now()); }
    DB.insert('users', nu);
    log(DB, u, team.id, 'approved', nu.id, 'added ' + name + ' to ' + team.name, [nu.id]);
    return publicUser(nu);
  },

  /* ---------- main admin tools ---------- */
  unlockUser: function (DB, r) {
    var u = auth(DB, r.token), target = DB.get('users', r.userId);
    if (!target) fail('That person no longer exists.');
    if (!isAdmin(u) && !canManageTeam(u, target.teamId)) fail('You can’t unlock this account.');
    DB.update('users', target.id, { failed: 0, lockedUntil: '' });
    return { ok: true };
  },
  signOutUser: function (DB, r) {
    var u = auth(DB, r.token), target = DB.get('users', r.userId);
    if (!target) fail('That person no longer exists.');
    if (!isAdmin(u) && !canManageTeam(u, target.teamId)) fail('You can’t sign this person out.');
    var n = 0;
    DB.all('sessions').filter(function (s) { return s.userId === target.id && s.id !== r.token; }).forEach(function (s) { DB.remove('sessions', s.id); n++; });
    return { ok: true, count: n };
  },
  /* Repairs for data that went out of shape (deleted teams, renamed stages, removed projects). Main admin only. */
  fixData: function (DB, r) {
    var u = auth(DB, r.token); if (!isAdmin(u)) fail('Only the main admin can do this.');
    var n = 0, team = r.teamId ? DB.get('teams', r.teamId) : null;
    if (r.kind === 'orphanProjects') {
      if (!team) fail('Choose a team.');
      DB.all('projects').filter(function (p) { return !DB.get('teams', p.teamId); }).forEach(function (p) { DB.update('projects', p.id, { teamId: team.id, updatedAt: nowISO() }); n++; });
    } else if (r.kind === 'homelessUsers') {
      if (!team) fail('Choose a team.');
      DB.all('users').filter(function (x) { return !DB.get('teams', x.teamId); }).forEach(function (x) { DB.update('users', x.id, { teamId: team.id, propertyCode: team.propertyCode }); n++; });
    } else if (r.kind === 'orphanTasks') {
      DB.all('tasks').filter(function (t) { return t.projectId && !DB.get('projects', t.projectId); }).forEach(function (t) { DB.update('tasks', t.id, { projectId: '', updatedAt: nowISO() }); n++; });
      DB.all('tasks').filter(function (t) { return t.parentId && !DB.get('tasks', t.parentId); }).forEach(function (t) { DB.update('tasks', t.id, { parentId: '', updatedAt: nowISO() }); n++; });
    } else if (r.kind === 'badStatus') {
      DB.all('tasks').forEach(function (t) {
        var parent = t.parentId ? DB.get('tasks', t.parentId) : null, p = t.projectId ? DB.get('projects', t.projectId) : null;
        var list = parent ? stagesOf({ stages: parent.subStages }) : stagesOf(p), ids = list.map(function (s) { return s.id; });
        if (ids.indexOf(t.status) < 0) { DB.update('tasks', t.id, { status: ids[0], updatedAt: nowISO() }); n++; }
      });
    } else if (r.kind === 'staleSessions') {
      var cut = Date.now() - 90 * 864e5;
      DB.all('sessions').filter(function (s) { return !DB.get('users', s.userId) || new Date(s.lastSeen || s.createdAt).getTime() < cut; }).forEach(function (s) { DB.remove('sessions', s.id); n++; });
    } else fail('Unknown repair.');
    return { ok: true, count: n };
  },

  removeUser: function (DB, r) {
    var u = auth(DB, r.token), target = DB.get('users', r.userId);
    if (!target) return { ok: true };
    if (target.id === u.id) fail('You can’t remove yourself.');
    if (!canManageTeam(u, target.teamId) || (target.role === 'admin' && !isAdmin(u))) fail('You can’t remove this person.');
    DB.all('sessions').filter(function (s) { return s.userId === target.id; }).forEach(function (s) { DB.remove('sessions', s.id); });
    DB.remove('users', target.id);
    return { ok: true };
  },

  resetPin: function (DB, r) {
    var u = auth(DB, r.token), target = DB.get('users', r.userId), pin = str(r.pin);
    if (!target) fail('That person no longer exists.');
    if (!/^\d{4}$/.test(pin)) fail('The PIN must be exactly 4 digits.');
    if (target.id !== u.id && (!canManageTeam(u, target.teamId) || (target.role === 'admin' && !isAdmin(u)))) fail('You can’t reset this PIN.');
    if (target.id === u.id && pinHash(u.salt, u.username, str(r.oldPin)) !== u.pinHash) fail('Your current PIN is incorrect.');
    var salt = uid('s');
    DB.update('users', target.id, { salt: salt, pinHash: pinHash(salt, target.username, pin), failed: 0, lockedUntil: '' });
    return { ok: true };
  },

  /* Main admin or the project team's department head edits the stages (columns) of a project. Done always stays last. */
  saveStages: function (DB, r) {
    var u = auth(DB, r.token), task = r.taskId ? DB.get('tasks', r.taskId) : null;
    if (r.taskId && (!task || !canSeeTask(DB, u, task))) fail('That task no longer exists.');
    var p = task ? (task.projectId ? DB.get('projects', task.projectId) : null) : DB.get('projects', r.projectId);
    if (!task && (!p || !canSeeProject(u, p))) fail('That project no longer exists.');
    if (task) {
      if (!(task.createdBy === u.id || listOf(task.assignees).indexOf(u.id) >= 0 || (p ? canManageProject(u, p) : canManageTeam(u, task.teamId))))
        fail('Only the person who created this task, the people assigned to it or a manager can change its subtask stages.');
    } else if (!canManageProject(u, p)) fail('Only this project’s managers, a department head or the main admin can change stages.');
    var out = [], seen = {}, done = null;
    (Array.isArray(r.stages) ? r.stages : []).slice(0, 12).forEach(function (s) {
      var name = str(s && s.name, 30); if (!name) return;
      var id = /^[a-z][a-z0-9]{1,30}$/.test(String(s.id || '')) ? String(s.id) : uid('s');
      if (seen[id]) return; seen[id] = 1;
      var st = { id: id, name: name, color: str(s.color, 20) || 'gray' };
      if (id === 'done') done = st; else out.push(st);
    });
    if (!out.length) fail('Keep at least one stage before Done.');
    if (!r.noDone) out.push(done || { id: 'done', name: 'Done', color: 'green' });
    var ids = out.map(function (s) { return s.id; });
    var oldList = task ? stagesOf({ stages: task.subStages }) : stagesOf(p);
    var inScope = function (t) { return task ? t.parentId === task.id : t.projectId === p.id && !t.parentId; };
    var keepAs = r.noDone && ids.indexOf(String(r.doneAs || '')) >= 0 ? String(r.doneAs) : '';
    if (r.noDone && hasDoneStage(oldList))
      DB.all('tasks').filter(function (t) { return inScope(t) && t.status === 'done'; })   // stay completed, just leave the old Done column
        .forEach(function (t) { DB.update('tasks', t.id, { status: keepAs || ids[0], completedAt: t.completedAt || nowISO(), updatedAt: nowISO() }); });
    if (!r.noDone && !hasDoneStage(oldList))
      DB.all('tasks').filter(function (t) { return inScope(t) && t.completedAt && t.status !== 'done'; })
        .forEach(function (t) { DB.update('tasks', t.id, { status: 'done', updatedAt: nowISO() }); });
    if (task) {
      DB.all('tasks').filter(function (t) { return t.parentId === task.id && ids.indexOf(t.status) < 0; })
        .forEach(function (t) { DB.update('tasks', t.id, { status: ids[0], updatedAt: nowISO() }); });
      DB.update('tasks', task.id, { subStages: out, updatedAt: nowISO() });
      return DB.get('tasks', task.id);
    }
    DB.all('tasks').filter(function (t) { return t.projectId === p.id && !t.parentId && ids.indexOf(t.status) < 0; })
      .forEach(function (t) { DB.update('tasks', t.id, { status: ids[0], updatedAt: nowISO() }); });
    DB.update('projects', p.id, { stages: out, updatedAt: nowISO() });
    return DB.get('projects', p.id);
  },

  updateProfile: function (DB, r) {
    var u = auth(DB, r.token), p = r.profile || {}, f = {};
    if (p.name !== undefined) { f.name = str(p.name, 60); if (f.name.length < 2) fail('Enter your full name.'); }
    if (p.jobTitle !== undefined) f.jobTitle = str(p.jobTitle, 60);
    if (p.photo !== undefined) { f.photo = cleanImage(p.photo); f.photoV = String(Date.now()); }
    if (p.username !== undefined && lc(p.username) !== lc(u.username)) {
      var nu = lc(p.username);
      if (!/^[a-z0-9._-]{3,20}$/.test(nu)) fail('Usernames are 3–20 letters, numbers, dots or dashes.');
      if (DB.find('users', function (x) { return x.id !== u.id && x.propertyCode === u.propertyCode && lc(x.username) === nu; })) fail('That username is already taken at ' + u.propertyCode + '.');
      if (pinHash(u.salt, u.username, str(p.pin)) !== u.pinHash) fail('Enter your current PIN to change your username.');
      var salt = uid('s');
      f.username = nu; f.salt = salt; f.pinHash = pinHash(salt, nu, str(p.pin));
    }
    DB.update('users', u.id, f);
    return publicUser(DB.get('users', u.id));
  },

  photos: function (DB, r) {
    var u = auth(DB, r.token), ids = r.ids || [], out = {};
    DB.all('users').forEach(function (x) {
      if (ids.indexOf(x.id) < 0) return;
      if (!(isAdmin(u) || x.teamId === u.teamId)) return;
      out[x.id] = { v: String(x.photoV || ''), data: x.photo || '' };
    });
    return out;
  },

  /* Several changes in one request, applied in order. Each change succeeds or fails on its own. */
  batch: function (DB, r) {
    auth(DB, r.token);
    return (r.ops || []).slice(0, 100).map(function (op) {
      if (!op || op.action === 'batch' || op.action === 'login' || op.action === 'signup') return { ok: false, error: 'Not allowed in a batch.' };
      op.token = r.token;
      return coreHandle(DB, op);
    });
  },

  /* ---------- properties & teams (main admin) ---------- */
  saveProperty: function (DB, r) {
    var u = auth(DB, r.token); if (!isAdmin(u)) fail('Only the main admin can add properties.');
    var code = str(r.property && r.property.code, 12).toUpperCase().replace(/[^A-Z0-9]/g, '');
    var name = str(r.property && r.property.name, 100);
    if (code.length < 3) fail('Property codes need at least 3 letters or numbers.');
    if (!name) fail('Enter the property name.');
    var look = {};
    if (r.property.color !== undefined) look.color = str(r.property.color, 20) || 'blue';
    if (r.property.logo !== undefined) look.logo = cleanImage(r.property.logo);
    var ex = findProperty(DB, code);
    if (ex) { DB.update('properties', ex.id, Object.assign({ name: name }, look)); return publicProperty(DB.get('properties', ex.id)); }
    var p = Object.assign({ id: uid('h'), code: code, name: name, color: 'blue', logo: '', createdAt: nowISO() }, look);
    DB.insert('properties', p);
    return publicProperty(p);
  },

  saveTeam: function (DB, r) {
    var u = auth(DB, r.token); if (!isAdmin(u)) fail('Only the main admin can add teams.');
    var t = r.team || {}, name = str(t.name, 60);
    if (!name) fail('Enter a team name.');
    if (t.id) { DB.update('teams', t.id, { name: name, color: str(t.color, 20) || 'blue' }); return DB.get('teams', t.id); }
    var p = findProperty(DB, t.propertyCode); if (!p) fail('Choose a property.');
    var obj = { id: uid('g'), propertyCode: p.code, name: name, color: str(t.color, 20) || 'blue', createdAt: nowISO() };
    DB.insert('teams', obj);
    return obj;
  },

  deleteTeam: function (DB, r) {
    var u = auth(DB, r.token); if (!isAdmin(u)) fail('Only the main admin can delete teams.');
    if (DB.find('users', function (x) { return x.teamId === r.id; })) fail('Move or remove everyone in this team first.');
    if (DB.find('projects', function (x) { return x.teamId === r.id; })) fail('Delete this team’s projects first.');
    DB.remove('teams', r.id);
    return { ok: true };
  }
};

function coreHandle(DB, req) {
  try {
    var fn = ACTIONS[req && req.action];
    if (!fn) fail('Unknown request.');
    return { ok: true, data: fn(DB, req) };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}

function seedAdmin(DB, cfg) {
  var p = findProperty(DB, cfg.propertyCode);
  if (!p) { p = { id: uid('h'), code: cfg.propertyCode, name: cfg.propertyName, createdAt: nowISO() }; DB.insert('properties', p); }
  var team = DB.find('teams', function (t) { return t.propertyCode === p.code && t.name === cfg.teamName; });
  if (!team) { team = { id: uid('g'), propertyCode: p.code, name: cfg.teamName, color: 'blue', createdAt: nowISO() }; DB.insert('teams', team); }
  var ex = DB.find('users', function (x) { return x.propertyCode === p.code && lc(x.username) === lc(cfg.username); });
  if (ex) return ex;
  var salt = uid('s');
  var u = { id: uid('u'), propertyCode: p.code, teamId: team.id, username: lc(cfg.username), name: cfg.name, jobTitle: cfg.jobTitle,
    role: 'admin', status: 'active', pinHash: pinHash(salt, cfg.username, cfg.pin), salt: salt, failed: 0, lockedUntil: '',
    avatarColor: 'indigo', createdAt: nowISO(), approvedBy: '', approvedAt: '' };
  DB.insert('users', u);
  return u;
}
