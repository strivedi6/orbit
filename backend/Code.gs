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
        if (JSON_COLS[k]) { try { v = v ? JSON.parse(v) : []; } catch (e) { v = []; } }
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
    return head.map(function (k) { var v = o[k]; if (JSON_COLS[k]) return JSON.stringify(v || []); if (v === null || v === undefined) return ''; return typeof v === 'boolean' ? String(v) : v; });
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
  users:      ['id','propertyCode','teamId','username','name','jobTitle','role','status','pinHash','salt','failed','lockedUntil','avatarColor','createdAt','approvedBy','approvedAt','photo','photoV'],
  sessions:   ['id','userId','createdAt','lastSeen'],
  projects:   ['id','teamId','name','color','icon','description','ownerId','status','due','createdBy','createdAt','updatedAt','archived'],
  tasks:      ['id','teamId','projectId','parentId','title','notes','status','priority','assignees','due','createdBy','createdAt','updatedAt','completedAt','order','attachments'],
  comments:   ['id','teamId','taskId','userId','text','createdAt'],
  activity:   ['id','teamId','userId','type','refId','text','targets','createdAt']
};
var JSON_COLS = { assignees: 1, attachments: 1, targets: 1 };
var STATUSES = ['todo','doing','review','done'];
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
    approvedBy: u.approvedBy || '', approvedAt: u.approvedAt || '', photoV: u.photo ? String(u.photoV || '1') : '' };
}
function isAdmin(u) { return u.role === 'admin'; }
function canSeeTeam(u, teamId) { return isAdmin(u) || u.teamId === teamId; }
function canManageTeam(u, teamId) { return isAdmin(u) || (u.role === 'head' && u.teamId === teamId); }
function findProperty(DB, code) { code = str(code).toUpperCase(); return DB.find('properties', function (p) { return String(p.code).toUpperCase() === code; }); }

function auth(DB, token) {
  if (!token) fail('AUTH');
  var s = DB.get('sessions', token);
  if (!s) fail('AUTH');
  var u = DB.get('users', s.userId);
  if (!u || u.status !== 'active') fail('AUTH');
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
  if (!t || !canSeeTeam(u, t.teamId)) fail('That task no longer exists.');
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
    var u = auth(DB, r.token);
    var admin = isAdmin(u);
    var teams = DB.all('teams').filter(function (t) { return admin || t.id === u.teamId; });
    var teamIds = {}; teams.forEach(function (t) { teamIds[t.id] = 1; });
    var inScope = function (x) { return !!teamIds[x.teamId]; };
    var users = DB.all('users').filter(function (x) {
      if (admin) return true;
      if (x.teamId !== u.teamId) return false;
      return x.status === 'active' || u.role === 'head';
    }).map(publicUser);
    var activity = DB.all('activity').filter(inScope)
      .sort(function (a, b) { return a.createdAt < b.createdAt ? 1 : -1; }).slice(0, 250);
    return {
      me: publicUser(u),
      properties: DB.all('properties').filter(function (p) { return admin || p.code === u.propertyCode; }).map(publicProperty),
      teams: teams, users: users,
      projects: DB.all('projects').filter(inScope),
      tasks: DB.all('tasks').filter(inScope),
      comments: DB.all('comments').filter(inScope),
      activity: activity,
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
    var existing = p.id ? DB.get('projects', p.id) : null;
    if (existing) {
      if (!canSeeTeam(u, existing.teamId)) fail('You can’t edit this project.');
      DB.update('projects', existing.id, fields);
      return DB.get('projects', existing.id);
    }
    var teamId = isAdmin(u) && p.teamId ? p.teamId : u.teamId;
    if (!DB.get('teams', teamId)) fail('Choose a team for this project.');
    var obj = Object.assign({ id: cleanId(p.id, 'p'), teamId: teamId, createdBy: u.id, createdAt: nowISO() }, fields);
    if (!obj.ownerId) obj.ownerId = u.id;
    DB.insert('projects', obj);
    log(DB, u, teamId, 'project', obj.id, 'created the project ' + obj.name, []);
    return obj;
  },

  deleteProject: function (DB, r) {
    var u = auth(DB, r.token), p = DB.get('projects', r.id);
    if (!p) return { ok: true };
    if (!canSeeTeam(u, p.teamId)) fail('That project no longer exists.');
    if (!(canManageTeam(u, p.teamId) || p.createdBy === u.id)) fail('Only the person who created this project or a department head can delete it.');
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
    if (existing) { if (!canSeeTeam(u, existing.teamId)) fail('That task no longer exists.'); existing = JSON.parse(JSON.stringify(existing)); }
    var base = existing || {};
    var projectId = t.projectId !== undefined ? str(t.projectId) : base.projectId;
    var project = projectId ? DB.get('projects', projectId) : null;
    if (projectId && (!project || !canSeeTeam(u, project.teamId))) fail('Choose a project you belong to.');
    var parent = t.parentId || base.parentId ? DB.get('tasks', t.parentId || base.parentId) : null;
    var teamId = project ? project.teamId : (parent ? parent.teamId : (existing ? existing.teamId : u.teamId));
    if (isAdmin(u) && !project && !parent && !existing && t.teamId) teamId = t.teamId;
    if (!canSeeTeam(u, teamId)) fail('You can’t add tasks to that team.');
    var team = DB.all('users').filter(function (x) { return x.teamId === teamId && x.status === 'active'; }).map(function (x) { return x.id; });
    var admins = DB.all('users').filter(function (x) { return x.role === 'admin'; }).map(function (x) { return x.id; });

    var f = {};
    if (t.title !== undefined) f.title = str(t.title, 300);
    if (t.notes !== undefined) f.notes = str(t.notes, 10000);
    if (t.status !== undefined) f.status = STATUSES.indexOf(t.status) >= 0 ? t.status : 'todo';
    if (t.priority !== undefined) f.priority = PRIORITIES.indexOf(t.priority) >= 0 ? t.priority : 'none';
    if (t.due !== undefined) f.due = str(t.due, 10);
    if (t.order !== undefined) f.order = Number(t.order) || 0;
    if (t.projectId !== undefined) f.projectId = projectId || '';
    if (t.assignees !== undefined) f.assignees = (t.assignees || []).filter(function (id) { return team.indexOf(id) >= 0 || admins.indexOf(id) >= 0; });
    f.teamId = teamId; f.updatedAt = nowISO();
    if (f.status) f.completedAt = f.status === 'done' ? (base.status === 'done' ? base.completedAt : nowISO()) : '';

    if (existing) {
      if (f.title === '') fail('A task needs a title.');
      DB.update('tasks', existing.id, f);
      var after = DB.get('tasks', existing.id);
      var added = (f.assignees || []).filter(function (id) { return (existing.assignees || []).indexOf(id) < 0; });
      if (added.length) log(DB, u, teamId, 'assign', existing.id, 'assigned “' + after.title + '”', added);
      if (f.status && f.status !== existing.status && f.status === 'done')
        log(DB, u, teamId, 'complete', existing.id, 'completed “' + after.title + '”', [after.createdBy].concat(after.assignees || []));
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
    if (!canSeeTeam(u, t.teamId)) fail('That task no longer exists.');
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
    var c = { id: cleanId(r.id, 'c'), teamId: t.teamId, taskId: t.id, userId: u.id, text: text, createdAt: nowISO() };
    DB.insert('comments', c);
    log(DB, u, t.teamId, 'comment', t.id, 'commented on “' + t.title + '”', [t.createdBy].concat(t.assignees || []));
    return c;
  },

  deleteComment: function (DB, r) {
    var u = auth(DB, r.token), c = DB.get('comments', r.id);
    if (!c) return { ok: true };
    if (!(c.userId === u.id || canManageTeam(u, c.teamId))) fail('You can only delete your own comments.');
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
    if (!canManageTeam(u, target.teamId)) fail('Only the main admin or this team’s department head can change this.');
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

  updateProfile: function (DB, r) {
    var u = auth(DB, r.token), p = r.profile || {}, f = {};
    if (p.name !== undefined) { f.name = str(p.name, 60); if (f.name.length < 2) fail('Enter your full name.'); }
    if (p.jobTitle !== undefined) f.jobTitle = str(p.jobTitle, 60);
    if (p.photo !== undefined) { f.photo = cleanImage(p.photo); f.photoV = String(Date.now()); }
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
