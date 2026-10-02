/**
 * Google Apps Script - Central Storage for MTCIT PPM Platform
 * يخزن البيانات في جداول منفصلة (ورقة لكل كيان) بدلاً من JSON واحد
 *
 * Deployment: انشر كـ Web App مع إعدادات:
 *   Execute as: Me
 *   Who has access: Anyone (أو Anyone with link)
 */

const SHEET_NAMES = {
  PROJECTS: 'Projects',
  OBJECTIVES: 'Objectives',
  KPIS: 'KPIs',
  TASKS: 'Tasks',
  RISKS: 'Risks',
  BUDGETS: 'Budgets',
  SERVICES: 'Services',
  TEAM_MEMBERS: 'TeamMembers',
  RESOURCES: 'Resources',
  FIVE_YEAR_PLANS: 'FiveYearPlans',
  SECTORS: 'Sectors',
  DIRECTORATES: 'Directorates',
  SYNC_META: 'SyncMeta', // لتخزين معلومات المزامنة (آخر تحديث، إصدار، إلخ)
  USERS: 'Users',
  ROLE_PERMISSIONS: 'RolePermissions',
  AUDIT_LOG: 'AuditLog',
  CHANGE_LOG: 'ChangeLog',
  CONFLICTS: 'Conflicts'
};

// تعريف الأعمدة لكل ورقة
const SHEET_SCHEMAS = {
  [SHEET_NAMES.PROJECTS]: [
    'id', 'name', 'type', 'sector', 'directorate', 'manager', 'objectiveId',
    'status', 'start', 'end', 'plannedBudget', 'allocatedBudget', 'actualCost',
    'progress', 'description', 'createdAt', 'updatedAt'
  ],
  [SHEET_NAMES.OBJECTIVES]: [
    'id', 'title', 'desc', 'perspective', 'parentId', 'weight', 'target',
    'progress', 'startYear', 'endYear', 'yearWeights', 'createdAt', 'updatedAt'
  ],
  [SHEET_NAMES.KPIS]: [
    'id', 'objectiveId', 'name', 'baseline', 'q1', 'q2', 'q3', 'q4', 'target',
    'actual', 'unit', 'createdAt', 'updatedAt'
  ],
  [SHEET_NAMES.TASKS]: [
    'id', 'projectId', 'name', 'predecessor', 'successor', 'planStart', 'planEnd',
    'actualStart', 'actualEnd', 'responsible', 'progress', 'order', 'sortOrder',
    'createdAt', 'updatedAt'
  ],
  [SHEET_NAMES.RISKS]: [
    'id', 'projectId', 'type', 'level', 'desc', 'action', 'owner', 'status',
    'probability', 'impact', 'createdAt', 'updatedAt'
  ],
  [SHEET_NAMES.BUDGETS]: [
    'id', 'projectId', 'item', 'planned', 'spent', 'date', 'notes',
    'createdAt', 'updatedAt'
  ],
  [SHEET_NAMES.SERVICES]: [
    'id', 'type', 'name', 'projectId', 'beneficiaries', 'launchDate', 'status',
    'progress', 'description', 'createdAt', 'updatedAt'
  ],
  [SHEET_NAMES.TEAM_MEMBERS]: [
    'id', 'projectId', 'name', 'role', 'specialty', 'email', 'phone',
    'createdAt', 'updatedAt'
  ],
  [SHEET_NAMES.RESOURCES]: [
    'id', 'projectId', 'manager', 'email', 'phone', 'department', 'stakeholders',
    'createdAt', 'updatedAt'
  ],
  [SHEET_NAMES.FIVE_YEAR_PLANS]: [
    'id', 'year', 'focus', 'goals', 'target', 'actual', 'budget', 'createdAt', 'updatedAt'
  ],
  [SHEET_NAMES.SECTORS]: [
    'id', 'name', 'createdAt', 'updatedAt'
  ],
  [SHEET_NAMES.DIRECTORATES]: [
    'id', 'name', 'createdAt', 'updatedAt'
  ],
  [SHEET_NAMES.SYNC_META]: [
    'key', 'value', 'updatedAt'
  ],
  [SHEET_NAMES.USERS]: [
    'id', 'username', 'name', 'email', 'role', 'directorate', 'active',
    'passwordSalt', 'passwordHash', 'mustChangePassword', 'updatedAt'
  ],
  [SHEET_NAMES.ROLE_PERMISSIONS]: [
    'role', 'pageId', 'pageName', 'access'
  ],
  [SHEET_NAMES.AUDIT_LOG]: [
    'timestamp', 'username', 'action', 'status', 'details'
  ],
  [SHEET_NAMES.CHANGE_LOG]: [
    'sequence', 'entity', 'recordId', 'operation', 'changedFields', 'revision',
    'updatedBy', 'timestamp', 'snapshot'
  ],
  [SHEET_NAMES.CONFLICTS]: [
    'id', 'entity', 'recordId', 'field', 'centralValue', 'localValue',
    'centralRevision', 'baseRevision', 'username', 'timestamp', 'status'
  ]
};

const SECURITY_TABS = [
  ['tab-dashboard','لوحة المؤشرات'], ['tab-fiveyear','الخطة الخمسية'],
  ['tab-strategy','الأهداف الاستراتيجية BSC'], ['tab-projects','المشاريع'],
  ['tab-resources','الموارد'], ['tab-gantt','خطة المشروع'],
  ['tab-risks','المخاطر'], ['tab-budget','الموازنة'],
  ['tab-services','الخدمات والمبادرات'], ['tab-reports','التقارير'],
  ['tab-export','تصدير'], ['tab-security','صلاحيات المستخدمين']
];
const ROLE_LABELS = { admin:'مدير النظام PMO', manager:'مدير مشروع', viewer:'عارض تنفيذي' };

// المفتاح المستخدم في appData/JSON -> اسم الورقة الفعلي في Google Sheets.
// إبقاء هذه الخريطة في مكان واحد يمنع اختلاف مفاتيح القراءة والكتابة.
const ENTITY_SHEETS = {
  projects: SHEET_NAMES.PROJECTS,
  objectives: SHEET_NAMES.OBJECTIVES,
  kpis: SHEET_NAMES.KPIS,
  tasks: SHEET_NAMES.TASKS,
  risks: SHEET_NAMES.RISKS,
  budgets: SHEET_NAMES.BUDGETS,
  services: SHEET_NAMES.SERVICES,
  teamMembers: SHEET_NAMES.TEAM_MEMBERS,
  resources: SHEET_NAMES.RESOURCES,
  fiveYearPlans: SHEET_NAMES.FIVE_YEAR_PLANS,
  sectors: SHEET_NAMES.SECTORS,
  directorates: SHEET_NAMES.DIRECTORATES
};

// حقول التحكم بالمزامنة التعاونية. تضاف إلى جميع أوراق البيانات دون تغيير حقول الأعمال.
Object.values(ENTITY_SHEETS).forEach(sheetName => {
  ['revision', 'fieldVersions', 'deleted', 'updatedBy'].forEach(field => {
    if (SHEET_SCHEMAS[sheetName].indexOf(field) < 0) SHEET_SCHEMAS[sheetName].push(field);
  });
});

// حقول تمثل يوماً تقويمياً فقط وليست طابعاً زمنياً.
const DATE_ONLY_FIELDS = {
  [SHEET_NAMES.PROJECTS]: ['start', 'end'],
  [SHEET_NAMES.TASKS]: ['planStart', 'planEnd', 'actualStart', 'actualEnd'],
  [SHEET_NAMES.BUDGETS]: ['date'],
  [SHEET_NAMES.SERVICES]: ['launchDate']
};

function isDateOnlyField(sheetName, field) {
  return (DATE_ONLY_FIELDS[sheetName] || []).indexOf(field) !== -1;
}

function normalizeDateOnly(value, timeZone) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, timeZone || Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const text = String(value).trim();
  const match = text.match(/^(\d{4}-\d{2}-\d{2})/);
  if (match) return match[1];
  return text;
}

function cleanExistingIsoDateCells(sheet, sheetName, headers, rows, timeZone) {
  const dateFields = DATE_ONLY_FIELDS[sheetName] || [];
  if (!dateFields.length || !rows.length) return;
  dateFields.forEach(field => {
    const columnIndex = headers.indexOf(field);
    if (columnIndex < 0) return;
    rows.forEach((row, rowIndex) => {
      const value = row[columnIndex];
      // تنظيف النصوص الزمنية القديمة فقط، وعدم لمس الصيغ أو القيم الأخرى.
      if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}T/.test(value.trim())) {
        const cell = sheet.getRange(rowIndex + 2, columnIndex + 1);
        if (!cell.getFormula()) {
          cell.setNumberFormat('@');
          cell.setValue(normalizeDateOnly(value, timeZone));
        }
      }
    });
  });
}

function securityPepper_() {
  const props = PropertiesService.getScriptProperties();
  let pepper = props.getProperty('PPM_SECURITY_PEPPER');
  if (!pepper) {
    pepper = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('PPM_SECURITY_PEPPER', pepper);
  }
  return pepper;
}

function centralPasswordHash_(password, salt) {
  const bytes = Utilities.computeHmacSha256Signature(String(password) + '|' + String(salt), securityPepper_());
  return Utilities.base64Encode(bytes);
}

function safeEqual_(a, b) {
  a = String(a || ''); b = String(b || '');
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function audit_(username, action, status, details) {
  const sheet = ensureSheet(SHEET_NAMES.AUDIT_LOG, SHEET_SCHEMAS[SHEET_NAMES.AUDIT_LOG]);
  sheet.appendRow([new Date().toISOString(), username || '', action || '', status || '', details || '']);
}

function getPermissionsForRole_(role) {
  // مدير النظام لا يعتمد على اكتمال ورقة الصلاحيات؛ يجب أن يرى جميع الصفحات دائماً.
  if (role === 'admin') {
    const adminPages = {};
    SECURITY_TABS.forEach(t => { adminPages[t[0]] = 'write'; });
    return adminPages;
  }
  const rows = getSheetData(SHEET_NAMES.ROLE_PERMISSIONS);
  const pages = {};
  rows.filter(r => String(r.role) === String(role)).forEach(r => {
    const access = ['write','read','none'].indexOf(String(r.access)) >= 0 ? String(r.access) : 'none';
    pages[String(r.pageId)] = access;
  });
  SECURITY_TABS.forEach(t => { if (!pages[t[0]]) pages[t[0]] = 'none'; });
  return pages;
}

function defaultPermission_(role, pageId) {
  if (role === 'admin') return 'write';
  if (role === 'manager') {
    if (pageId === 'tab-security') return 'none';
    return ['tab-dashboard','tab-fiveyear','tab-strategy','tab-reports','tab-export'].indexOf(pageId) >= 0 ? 'read' : 'write';
  }
  return ['tab-export','tab-security'].indexOf(pageId) >= 0 ? 'none' : 'read';
}

function authenticate_(username, password) {
  username = String(username || '').trim().toLowerCase();
  const user = getSheetData(SHEET_NAMES.USERS).find(u => String(u.username || '').trim().toLowerCase() === username);
  if (!user || user.active === false || String(user.active).toLowerCase() === 'false') {
    audit_(username, 'login', 'denied', 'unknown or inactive user');
    throw new Error('بيانات الدخول غير صحيحة أو الحساب غير مفعّل');
  }
  const actual = centralPasswordHash_(password, user.passwordSalt);
  if (!safeEqual_(actual, user.passwordHash)) {
    audit_(username, 'login', 'denied', 'invalid password');
    throw new Error('بيانات الدخول غير صحيحة');
  }
  const pages = getPermissionsForRole_(user.role);
  const canWrite = Object.keys(pages).some(k => pages[k] === 'write' && k !== 'tab-security');
  const session = {
    username: username, name: user.name || username, email: user.email || '', role: user.role || 'viewer',
    roleLabel: ROLE_LABELS[user.role] || user.role || 'viewer', directorate: user.directorate || '',
    scope: user.role === 'manager' ? 'directorate' : 'all', pages: pages, canWrite: canWrite,
    mustChange: user.mustChangePassword === true || String(user.mustChangePassword).toLowerCase() === 'true',
    at: new Date().toISOString()
  };
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('session:' + token, JSON.stringify(session), 21600);
  audit_(username, 'login', 'success', 'central authentication');
  return { token: token, session: session, expiresIn: 21600 };
}

function requireSession_(token, requireWrite) {
  const raw = CacheService.getScriptCache().get('session:' + String(token || ''));
  if (!raw) throw new Error('انتهت جلسة الدخول أو أنها غير صالحة');
  const session = JSON.parse(raw);
  const liveUser = getSheetData(SHEET_NAMES.USERS).find(u => String(u.username || '').trim().toLowerCase() === String(session.username || '').trim().toLowerCase());
  if (!liveUser || liveUser.active === false || String(liveUser.active).toLowerCase() === 'false') {
    CacheService.getScriptCache().remove('session:' + String(token || ''));
    throw new Error('تم تعطيل الحساب أو حذفه. سجل الدخول من جديد.');
  }
  session.role = liveUser.role || 'viewer';
  session.name = liveUser.name || session.username;
  session.email = liveUser.email || '';
  session.directorate = liveUser.directorate || '';
  session.pages = getPermissionsForRole_(session.role);
  session.canWrite = Object.keys(session.pages).some(k => session.pages[k] === 'write' && k !== 'tab-security');
  if (requireWrite && !session.canWrite) throw new Error('لا تملك صلاحية رفع التعديلات');
  CacheService.getScriptCache().put('session:' + String(token || ''), JSON.stringify(session), 21600);
  return session;
}

function htmlBridgeResponse_(token, result) {
  const payload = JSON.stringify({ type:'T21_CENTRAL_AUTH_RESPONSE', token:String(token || ''), result:result })
    .replace(/</g, '\\u003c');
  return HtmlService.createHtmlOutput('<!doctype html><meta charset="utf-8"><script>try{top.postMessage(' + payload + ',"*")}catch(e){};try{parent.postMessage(' + payload + ',"*")}catch(e){};<\/script>')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function htmlWriteBridgeResponse_(token, result) {
  const payload = JSON.stringify({ type:'T21_CENTRAL_WRITE_RESPONSE', token:String(token || ''), result:result })
    .replace(/</g, '\\u003c');
  return HtmlService.createHtmlOutput('<!doctype html><meta charset="utf-8"><script>try{top.postMessage(' + payload + ',"*")}catch(e){};try{parent.postMessage(' + payload + ',"*")}catch(e){};<\/script>')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function initializeSecuritySheets() {
  const users = ensureSheet(SHEET_NAMES.USERS, SHEET_SCHEMAS[SHEET_NAMES.USERS]);
  const perms = ensureSheet(SHEET_NAMES.ROLE_PERMISSIONS, SHEET_SCHEMAS[SHEET_NAMES.ROLE_PERMISSIONS]);
  ensureSheet(SHEET_NAMES.AUDIT_LOG, SHEET_SCHEMAS[SHEET_NAMES.AUDIT_LOG]);
  ensureSheet(SHEET_NAMES.CHANGE_LOG, SHEET_SCHEMAS[SHEET_NAMES.CHANGE_LOG]);
  ensureSheet(SHEET_NAMES.CONFLICTS, SHEET_SCHEMAS[SHEET_NAMES.CONFLICTS]);
  if (users.getLastRow() < 2) {
    const salt = Utilities.getUuid();
    users.appendRow(['u_admin','admin','مدير النظام','', 'admin','',true,salt,centralPasswordHash_('admin123',salt),true,new Date().toISOString()]);
  }
  // استكمال أي صفوف ناقصة حتى لو كانت الورقة تحتوي على بعض الصلاحيات مسبقاً.
  const existing = perms.getLastRow() >= 2
    ? perms.getRange(2,1,perms.getLastRow()-1,4).getValues()
    : [];
  const keys = {};
  existing.forEach(row => { keys[String(row[0]) + '|' + String(row[1])] = true; });
  const missing = [];
  ['admin','manager','viewer'].forEach(role => {
    SECURITY_TABS.forEach(t => {
      const key = role + '|' + t[0];
      if (!keys[key]) missing.push([role,t[0],t[1],defaultPermission_(role,t[0])]);
    });
  });
  if (missing.length) perms.getRange(perms.getLastRow()+1,1,missing.length,4).setValues(missing);
  SpreadsheetApp.getUi().alert('تم تهيئة واستكمال Users وRolePermissions وسجلات التدقيق. تمت إضافة ' + missing.length + ' صلاحية مفقودة.');
}

function setCentralUserPassword(username, newPassword) {
  if (!username || !newPassword) throw new Error('أدخل اسم المستخدم وكلمة المرور الجديدة');
  const sheet = ensureSheet(SHEET_NAMES.USERS, SHEET_SCHEMAS[SHEET_NAMES.USERS]);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const userCol = headers.indexOf('username'), saltCol = headers.indexOf('passwordSalt');
  const hashCol = headers.indexOf('passwordHash'), changeCol = headers.indexOf('mustChangePassword'), updatedCol = headers.indexOf('updatedAt');
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][userCol]).trim().toLowerCase() === String(username).trim().toLowerCase()) {
      const salt = Utilities.getUuid();
      sheet.getRange(r+1,saltCol+1).setValue(salt);
      sheet.getRange(r+1,hashCol+1).setValue(centralPasswordHash_(newPassword,salt));
      sheet.getRange(r+1,changeCol+1).setValue(false);
      sheet.getRange(r+1,updatedCol+1).setValue(new Date().toISOString());
      return true;
    }
  }
  throw new Error('اسم المستخدم غير موجود');
}

function validatePassword_(password, username) {
  password = String(password || '');
  if (password.length < 8) throw new Error('يجب ألا تقل كلمة المرور عن 8 أحرف');
  if (password.length > 128) throw new Error('كلمة المرور طويلة جداً');
  if (String(username || '').toLowerCase() === password.toLowerCase()) throw new Error('لا يمكن استخدام اسم المستخدم ككلمة مرور');
  return password;
}

function userRowInfo_(username) {
  const sheet = ensureSheet(SHEET_NAMES.USERS, SHEET_SCHEMAS[SHEET_NAMES.USERS]);
  const values = sheet.getDataRange().getValues();
  const headers = values[0] || SHEET_SCHEMAS[SHEET_NAMES.USERS];
  const usernameCol = headers.indexOf('username');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][usernameCol] || '').trim().toLowerCase() === String(username || '').trim().toLowerCase()) {
      return { sheet:sheet, headers:headers, values:values[i], row:i + 1 };
    }
  }
  return null;
}

function setUserCell_(info, field, value) {
  const col = info.headers.indexOf(field);
  if (col < 0) throw new Error('حقل المستخدم غير موجود: ' + field);
  info.sheet.getRange(info.row, col + 1).setValue(value);
}

function setPasswordForUser_(info, password, mustChange) {
  const username = info.values[info.headers.indexOf('username')];
  validatePassword_(password, username);
  const salt = Utilities.getUuid();
  setUserCell_(info, 'passwordSalt', salt);
  setUserCell_(info, 'passwordHash', centralPasswordHash_(password, salt));
  setUserCell_(info, 'mustChangePassword', !!mustChange);
  setUserCell_(info, 'updatedAt', new Date().toISOString());
}

function securityAdminAction_(action, data, actor) {
  data = data || {};
  const now = new Date().toISOString();

  if (action === 'changeOwnPassword') {
    const info = userRowInfo_(actor.username);
    if (!info) throw new Error('الحساب غير موجود');
    const salt = info.values[info.headers.indexOf('passwordSalt')];
    const hash = info.values[info.headers.indexOf('passwordHash')];
    if (!safeEqual_(centralPasswordHash_(data.currentPassword, salt), hash)) throw new Error('كلمة المرور الحالية غير صحيحة');
    if (String(data.currentPassword || '') === String(data.newPassword || '')) throw new Error('يجب أن تختلف كلمة المرور الجديدة عن الحالية');
    setPasswordForUser_(info, data.newPassword, false);
    audit_(actor.username, 'changePassword', 'success', 'self-service from browser');
    return { ok:true, message:'تم تغيير كلمة المرور بنجاح', mustChange:false };
  }

  if (actor.role !== 'admin') throw new Error('هذه العملية متاحة لمدير النظام فقط');
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (action === 'createUser') {
      const username = String(data.username || '').trim().toLowerCase();
      if (!/^[a-z0-9._-]{3,40}$/.test(username)) throw new Error('اسم المستخدم يجب أن يكون 3 إلى 40 حرفاً إنجليزياً أو رقماً');
      if (userRowInfo_(username)) throw new Error('اسم المستخدم موجود بالفعل');
      const role = ['admin','manager','viewer'].indexOf(String(data.role)) >= 0 ? String(data.role) : 'viewer';
      validatePassword_(data.password, username);
      const salt = Utilities.getUuid();
      const sheet = ensureSheet(SHEET_NAMES.USERS, SHEET_SCHEMAS[SHEET_NAMES.USERS]);
      sheet.appendRow(['u_' + Utilities.getUuid(), username, String(data.name || username).trim(), String(data.email || '').trim(), role,
        String(data.directorate || '').trim(), true, salt, centralPasswordHash_(data.password, salt), true, now]);
      audit_(actor.username, 'createUser', 'success', username + ', role=' + role);
    } else if (action === 'updateUser') {
      const info = userRowInfo_(data.username);
      if (!info) throw new Error('المستخدم غير موجود');
      const role = ['admin','manager','viewer'].indexOf(String(data.role)) >= 0 ? String(data.role) : 'viewer';
      if (String(data.username).toLowerCase() === String(actor.username).toLowerCase() && role !== 'admin') throw new Error('لا يمكن لمدير النظام تخفيض صلاحية حسابه الحالي');
      ['name','email','directorate'].forEach(field => setUserCell_(info, field, String(data[field] || '').trim()));
      setUserCell_(info, 'role', role);
      setUserCell_(info, 'updatedAt', now);
      audit_(actor.username, 'updateUser', 'success', String(data.username) + ', role=' + role);
    } else if (action === 'setUserActive') {
      const info = userRowInfo_(data.username);
      if (!info) throw new Error('المستخدم غير موجود');
      if (String(data.username).toLowerCase() === String(actor.username).toLowerCase() && data.active === false) throw new Error('لا يمكنك تعطيل حسابك الحالي');
      setUserCell_(info, 'active', !!data.active);
      setUserCell_(info, 'updatedAt', now);
      audit_(actor.username, data.active ? 'activateUser' : 'deactivateUser', 'success', String(data.username));
    } else if (action === 'deleteUser') {
      const info = userRowInfo_(data.username);
      if (!info) throw new Error('المستخدم غير موجود');
      if (String(data.username).toLowerCase() === String(actor.username).toLowerCase()) throw new Error('لا يمكنك حذف حسابك الحالي');
      audit_(actor.username, 'deleteUser', 'success', String(data.username));
      info.sheet.deleteRow(info.row);
    } else if (action === 'resetUserPassword') {
      const info = userRowInfo_(data.username);
      if (!info) throw new Error('المستخدم غير موجود');
      setPasswordForUser_(info, data.password, true);
      audit_(actor.username, 'resetUserPassword', 'success', String(data.username));
    } else if (action === 'updatePermission') {
      const role = String(data.role || '');
      const pageId = String(data.pageId || '');
      const access = String(data.access || 'none');
      if (role === 'admin') throw new Error('صلاحيات مدير النظام ثابتة وكاملة');
      if (['manager','viewer'].indexOf(role) < 0 || !SECURITY_TABS.some(t => t[0] === pageId) || ['write','read','none'].indexOf(access) < 0) throw new Error('قيمة الصلاحية غير صحيحة');
      const sheet = ensureSheet(SHEET_NAMES.ROLE_PERMISSIONS, SHEET_SCHEMAS[SHEET_NAMES.ROLE_PERMISSIONS]);
      const values = sheet.getDataRange().getValues();
      let row = 0;
      for (let i = 1; i < values.length; i++) if (String(values[i][0]) === role && String(values[i][1]) === pageId) { row = i + 1; break; }
      if (row) sheet.getRange(row, 1, 1, 4).setValues([[role,pageId,SECURITY_TABS.find(t => t[0] === pageId)[1],access]]);
      else sheet.appendRow([role,pageId,SECURITY_TABS.find(t => t[0] === pageId)[1],access]);
      audit_(actor.username, 'updatePermission', 'success', role + '/' + pageId + '=' + access);
    } else {
      throw new Error('عملية إدارة غير معروفة');
    }
    return { ok:true, message:'تم حفظ التغيير بنجاح' };
  } finally {
    lock.releaseLock();
  }
}

function nextChangeSequence_() {
  const props = PropertiesService.getScriptProperties();
  const next = Number(props.getProperty('PPM_CHANGE_SEQUENCE') || 0) + 1;
  props.setProperty('PPM_CHANGE_SEQUENCE', String(next));
  return next;
}

function latestChangeSequence_() {
  return Number(PropertiesService.getScriptProperties().getProperty('PPM_CHANGE_SEQUENCE') || 0);
}

function parseFieldVersions_(value) {
  if (value && typeof value === 'object') return value;
  try { return JSON.parse(String(value || '{}')); } catch (e) { return {}; }
}

function appendChange_(entity, record, operation, changedFields, username) {
  const sequence = nextChangeSequence_();
  const sheet = ensureSheet(SHEET_NAMES.CHANGE_LOG, SHEET_SCHEMAS[SHEET_NAMES.CHANGE_LOG]);
  sheet.appendRow([
    sequence, entity, record.id, operation, (changedFields || []).join(','), Number(record.revision || 0),
    username || '', new Date().toISOString(), JSON.stringify(record)
  ]);
  return sequence;
}

function appendConflict_(entity, recordId, field, centralValue, localValue, centralRevision, baseRevision, username) {
  const sheet = ensureSheet(SHEET_NAMES.CONFLICTS, SHEET_SCHEMAS[SHEET_NAMES.CONFLICTS]);
  const id = 'cf_' + Utilities.getUuid();
  sheet.appendRow([id, entity, recordId, field, JSON.stringify(centralValue), JSON.stringify(localValue),
    centralRevision, baseRevision, username, new Date().toISOString(), 'open']);
  return id;
}

function entityConfig_(entity) {
  const sheetName = ENTITY_SHEETS[entity];
  if (!sheetName) throw new Error('كيان مزامنة غير معروف: ' + entity);
  return { sheetName:sheetName, schema:SHEET_SCHEMAS[sheetName] };
}

function findRecordRow_(sheet, schema, id) {
  if (sheet.getLastRow() < 2) return null;
  const idCol = schema.indexOf('id');
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, schema.length).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][idCol]) === String(id)) {
      const record = {};
      schema.forEach((field, c) => { record[field] = values[i][c]; });
      return { row:i + 2, record:record };
    }
  }
  return null;
}

function writeRecordRow_(sheet, schema, row, record) {
  const values = schema.map(field => {
    let value = record[field];
    if (field === 'fieldVersions' && value && typeof value === 'object') value = JSON.stringify(value);
    return value === undefined || value === null ? '' : value;
  });
  sheet.getRange(row, 1, 1, schema.length).setValues([values]);
}

function applyDeltaOperation_(op, actor) {
  const entity = String(op.entity || '');
  const cfg = entityConfig_(entity);
  const sheet = ensureSheet(cfg.sheetName, cfg.schema);
  const found = findRecordRow_(sheet, cfg.schema, op.id);
  const baseRevision = Number(op.baseRevision || 0);
  const changes = op.changes && typeof op.changes === 'object' ? op.changes : {};
  const changedFields = Object.keys(changes).filter(f => cfg.schema.indexOf(f) >= 0 && ['revision','fieldVersions','updatedBy','updatedAt','createdAt'].indexOf(f) < 0);

  if (op.operation === 'create') {
    if (found && found.record.deleted !== true && String(found.record.deleted).toLowerCase() !== 'true') {
      appendConflict_(entity, op.id, '*', found.record, changes, found.record.revision || 0, baseRevision, actor.username);
      return { status:'conflict', entity:entity, id:op.id, fields:['*'] };
    }
    const revision = found ? Number(found.record.revision || 0) + 1 : 1;
    const record = Object.assign({}, changes, { id:op.id, revision:revision, deleted:false, updatedBy:actor.username, updatedAt:new Date().toISOString() });
    record.createdAt = record.createdAt || new Date().toISOString();
    record.fieldVersions = {};
    changedFields.forEach(f => { record.fieldVersions[f] = revision; });
    writeRecordRow_(sheet, cfg.schema, found ? found.row : sheet.getLastRow() + 1, record);
    const seq = appendChange_(entity, record, 'create', changedFields, actor.username);
    return { status:'applied', entity:entity, id:op.id, record:record, sequence:seq };
  }

  if (!found) {
    appendConflict_(entity, op.id, '*', null, changes, 0, baseRevision, actor.username);
    return { status:'conflict', entity:entity, id:op.id, fields:['*'] };
  }

  const record = found.record;
  const fieldVersions = parseFieldVersions_(record.fieldVersions);
  const conflicting = [];
  if (op.operation === 'delete') {
    if (Number(record.revision || 0) > baseRevision) conflicting.push('*');
  } else {
    changedFields.forEach(field => { if (Number(fieldVersions[field] || 0) > baseRevision) conflicting.push(field); });
  }
  if (conflicting.length) {
    conflicting.forEach(field => appendConflict_(entity, op.id, field,
      field === '*' ? record : record[field], field === '*' ? changes : changes[field],
      Number(record.revision || 0), baseRevision, actor.username));
    return { status:'conflict', entity:entity, id:op.id, fields:conflicting, central:record };
  }

  const revision = Number(record.revision || 0) + 1;
  if (op.operation === 'delete') {
    record.deleted = true;
    fieldVersions.deleted = revision;
  } else {
    changedFields.forEach(field => { record[field] = changes[field]; fieldVersions[field] = revision; });
  }
  record.revision = revision;
  record.fieldVersions = fieldVersions;
  record.updatedBy = actor.username;
  record.updatedAt = new Date().toISOString();
  writeRecordRow_(sheet, cfg.schema, found.row, record);
  const seq = appendChange_(entity, record, op.operation === 'delete' ? 'delete' : 'update', op.operation === 'delete' ? ['deleted'] : changedFields, actor.username);
  return { status:'applied', entity:entity, id:op.id, record:record, sequence:seq };
}

function changesSince_(sinceSequence) {
  const rows = getSheetData(SHEET_NAMES.CHANGE_LOG);
  return rows.filter(r => Number(r.sequence || 0) > Number(sinceSequence || 0)).map(r => {
    let snapshot = {};
    try { snapshot = JSON.parse(r.snapshot || '{}'); } catch (e) {}
    return { sequence:Number(r.sequence || 0), entity:r.entity, recordId:r.recordId, operation:r.operation,
      changedFields:String(r.changedFields || '').split(',').filter(Boolean), revision:Number(r.revision || 0),
      updatedBy:r.updatedBy, timestamp:r.timestamp, record:snapshot };
  });
}

function collaborativeSync_(payload, actor) {
  const operations = Array.isArray(payload.operations) ? payload.operations : [];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const results = operations.map(op => applyDeltaOperation_(op, actor));
    const changes = changesSince_(Number(payload.sinceSequence || 0));
    audit_(actor.username, 'deltaSync', 'success', 'operations=' + operations.length + ', conflicts=' + results.filter(r => r.status === 'conflict').length);
    return { ok:true, results:results, changes:changes, latestSequence:latestChangeSequence_() };
  } finally { lock.releaseLock(); }
}

function securityActivity_(actor) {
  if (actor.role !== 'admin') throw new Error('عرض سجل التدقيق متاح لمدير النظام فقط');
  const users = getSheetData(SHEET_NAMES.USERS).map(u => ({ id:u.id, username:u.username, name:u.name, email:u.email, role:u.role, directorate:u.directorate, active:u.active, mustChangePassword:u.mustChangePassword, updatedAt:u.updatedAt }));
  const permissions = getSheetData(SHEET_NAMES.ROLE_PERMISSIONS);
  const audit = getSheetData(SHEET_NAMES.AUDIT_LOG).slice(-250).reverse();
  const changes = getSheetData(SHEET_NAMES.CHANGE_LOG).slice(-250).reverse();
  const conflicts = getSheetData(SHEET_NAMES.CONFLICTS).slice(-250).reverse();
  return { users:users, permissions:permissions, audit:audit, changes:changes, conflicts:conflicts, latestSequence:latestChangeSequence_() };
}

function installCollaborativeEditTrigger() {
  const ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'handleCollaborativeSheetEdit').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('handleCollaborativeSheetEdit').forSpreadsheet(ss).onEdit().create();
  SpreadsheetApp.getUi().alert('تم تفعيل تتبع التعديلات المباشرة في Google Sheet.');
}

function handleCollaborativeSheetEdit(e) {
  if (!e || !e.range || e.range.getRow() < 2) return;
  const sheet = e.range.getSheet();
  const entity = Object.keys(ENTITY_SHEETS).find(key => ENTITY_SHEETS[key] === sheet.getName());
  if (!entity) return;
  const schema = SHEET_SCHEMAS[sheet.getName()];
  const editedField = schema[e.range.getColumn() - 1];
  if (!editedField || ['revision','fieldVersions','updatedBy','updatedAt'].indexOf(editedField) >= 0) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const values = sheet.getRange(e.range.getRow(), 1, 1, schema.length).getValues()[0];
    const record = {}; schema.forEach((field, i) => { record[field] = values[i]; });
    if (!record.id) return;
    const revision = Number(record.revision || 0) + 1;
    const versions = parseFieldVersions_(record.fieldVersions);
    versions[editedField] = revision;
    record.revision = revision; record.fieldVersions = versions;
    record.updatedBy = Session.getActiveUser().getEmail() || 'sheet-editor';
    record.updatedAt = new Date().toISOString();
    writeRecordRow_(sheet, schema, e.range.getRow(), record);
    appendChange_(entity, record, 'update', [editedField], record.updatedBy);
    audit_(record.updatedBy, 'sheetEdit', 'success', entity + '/' + record.id + '/' + editedField);
  } finally { lock.releaseLock(); }
}

/**
 * نقطة الدخول للطلبات POST (من الواجهة الأمامية)
 */
function doPost(e) {
  try {
    let payload = {};
    try { payload = JSON.parse((e.postData && e.postData.contents) || '{}'); }
    catch (parseErr) { payload = (e && e.parameter) || {}; }
    const action = payload.action;

    if (action === 'bridgeAuth') {
      const bridgeToken = payload.bridgeToken || '';
      try {
        return htmlBridgeResponse_(bridgeToken, { ok:true, auth:authenticate_(payload.username, payload.password) });
      } catch (authErr) {
        return htmlBridgeResponse_(bridgeToken, { ok:false, error:String(authErr.message || authErr) });
      }
    }

    if (action === 'authenticate') {
      return jsonResponse({ ok:true, auth:authenticate_(payload.username, payload.password) });
    }

    if (action === 'bridgeSetDB') {
      const bridgeToken = payload.bridgeToken || '';
      try {
        const actor = requireSession_(payload.token, true);
        const lock = LockService.getScriptLock();
        lock.waitLock(15000);
        try {
          const current = getSyncMeta();
          const baseVersion = Number(payload.baseVersion || 0);
          const currentVersion = Number(current.version || 0);
          if (baseVersion !== currentVersion) {
            audit_(actor.username, 'setDB', 'conflict', 'base=' + baseVersion + ', current=' + currentVersion);
            return htmlWriteBridgeResponse_(bridgeToken, { ok:false, code:'SYNC_CONFLICT', error:'توجد نسخة مركزية أحدث. حمّل النسخة المركزية وراجع تغييراتك قبل إعادة الرفع.', currentVersion:currentVersion });
          }
          setAllData(JSON.parse(payload.dbJson || '{}'));
          const updatedMeta = getSyncMeta();
          audit_(actor.username, 'setDB', 'success', 'version=' + updatedMeta.version);
          return htmlWriteBridgeResponse_(bridgeToken, { ok:true, version:updatedMeta.version, updatedAt:updatedMeta.lastSyncAt });
        } finally { lock.releaseLock(); }
      } catch (writeErr) {
        return htmlWriteBridgeResponse_(bridgeToken, { ok:false, error:String(writeErr.message || writeErr) });
      }
    }

    if (action === 'bridgeDeltaSync') {
      const bridgeToken = payload.bridgeToken || '';
      try {
        const actor = requireSession_(payload.token, true);
        const result = collaborativeSync_({
          operations:JSON.parse(payload.operationsJson || '[]'),
          sinceSequence:Number(payload.sinceSequence || 0)
        }, actor);
        return htmlWriteBridgeResponse_(bridgeToken, result);
      } catch (syncErr) {
        return htmlWriteBridgeResponse_(bridgeToken, { ok:false, error:String(syncErr.message || syncErr) });
      }
    }

    if (action === 'bridgeSecurityAction') {
      const bridgeToken = payload.bridgeToken || '';
      try {
        const actor = requireSession_(payload.token, false);
        const data = JSON.parse(payload.dataJson || '{}');
        const result = securityAdminAction_(String(payload.securityAction || ''), data, actor);
        if (result.mustChange === false) {
          actor.mustChange = false;
          CacheService.getScriptCache().put('session:' + String(payload.token || ''), JSON.stringify(actor), 21600);
        }
        return htmlWriteBridgeResponse_(bridgeToken, result);
      } catch (securityErr) {
        return htmlWriteBridgeResponse_(bridgeToken, { ok:false, error:String(securityErr.message || securityErr) });
      }
    }

    if (action === 'getDB') {
      requireSession_(payload.token, false);
      return jsonResponse({ ok: true, db: getAllData() });
    }

    if (action === 'setDB') {
      const actor = requireSession_(payload.token, true);
      const lock = LockService.getScriptLock();
      lock.waitLock(15000);
      try {
        const current = getSyncMeta();
        const baseVersion = Number(payload.baseVersion || 0);
        const currentVersion = Number(current.version || 0);
        if (baseVersion !== currentVersion) {
          audit_(actor.username, 'setDB', 'conflict', 'base=' + baseVersion + ', current=' + currentVersion);
          return jsonResponse({ ok:false, code:'SYNC_CONFLICT', error:'توجد نسخة مركزية أحدث. حمّل النسخة المركزية وراجع تغييراتك قبل إعادة الرفع.', currentVersion:currentVersion });
        }
      const db = payload.db || {};
      setAllData(db);
      // إرجاع version و updatedAt لتحديث meta في الواجهة الأمامية
      const updatedMeta = getSyncMeta();
      audit_(actor.username, 'setDB', 'success', 'version=' + updatedMeta.version);
      return jsonResponse({
        ok: true,
        message: 'Data saved successfully',
        version: updatedMeta.version,
        updatedAt: updatedMeta.lastSyncAt
      });
      } finally {
        lock.releaseLock();
      }
    }

    if (action === 'getSheet') {
      const sheetReader = requireSession_(payload.token, false);
      // جلب ورقة محددة فقط
      const sheetName = payload.sheetName;
      if ([SHEET_NAMES.USERS, SHEET_NAMES.ROLE_PERMISSIONS, SHEET_NAMES.AUDIT_LOG].indexOf(sheetName) >= 0 && sheetReader.role !== 'admin') {
        return jsonResponse({ ok:false, error:'هذه الورقة متاحة لمدير النظام فقط' });
      }
      if (sheetName && SHEET_SCHEMAS[sheetName]) {
        return jsonResponse({ ok: true, data: getSheetData(sheetName) });
      }
      return jsonResponse({ ok: false, error: 'Invalid sheet name' });
    }

    if (action === 'setSheet') {
      const sheetWriter = requireSession_(payload.token, true);
      // حفظ ورقة محددة فقط
      const sheetName = payload.sheetName;
      const data = payload.data || [];
      if ([SHEET_NAMES.USERS, SHEET_NAMES.ROLE_PERMISSIONS, SHEET_NAMES.AUDIT_LOG].indexOf(sheetName) >= 0 && sheetWriter.role !== 'admin') {
        return jsonResponse({ ok:false, error:'تعديل هذه الورقة متاح لمدير النظام فقط' });
      }
      if (sheetName && SHEET_SCHEMAS[sheetName]) {
        setSheetData(sheetName, data);
        return jsonResponse({ ok: true });
      }
      return jsonResponse({ ok: false, error: 'Invalid sheet name' });
    }

    return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    console.error('doPost error:', err);
    return jsonResponse({ ok: false, error: String(err) });
  }
}

/**
 * نقطة الدخول للطلبات GET (للتحقق من الخدمة)
 */
function doGet(e) {
  const params = (e && e.parameter) || {};
  try {
    if (params.action === 'bridgeGetDB') {
      requireSession_(params.token, false);
      return bridgeResponse(params.bridgeToken, { ok: true, db: getAllData() });
    }

    if (params.action === 'getSecurityActivity') {
      const actor = requireSession_(params.token, false);
      const activityResult = { ok:true, activity:securityActivity_(actor) };
      return params.callback ? jsonpResponse(params.callback, activityResult) : jsonResponse(activityResult);
    }

    const result = params.action === 'getDB'
      ? (requireSession_(params.token, false), { ok: true, db: getAllData() })
      : {
          ok: true,
          message: 'MTCIT PPM Central Storage is running.',
          version: '3.2.0',
          structure: 'tabular',
          sheets: Object.values(SHEET_NAMES)
        };

    // JSONP يسمح للواجهة المحلية file:// بقراءة البيانات دون حظر CORS.
    if (params.callback) {
      return jsonpResponse(params.callback, result);
    }
    return jsonResponse(result);
  } catch (err) {
    const errorResult = { ok: false, error: String(err) };
    if (params.action === 'bridgeGetDB') {
      return bridgeResponse(params.bridgeToken, errorResult);
    }
    // لا نرجع JSON عاديًا لطلب JSONP، وإلا سيظهر في المتصفح كأنه خطأ شبكة.
    return params.callback
      ? jsonpResponse(params.callback, errorResult)
      : jsonResponse(errorResult);
  }
}

/**
 * جسر قراءة للصفحات المحلية file:// عبر iframe + postMessage.
 * HtmlService يشغّل المحتوى داخل iframe داخلي؛ لذلك يجب الإرسال إلى window.top
 * للوصول إلى صفحة البرنامج، مع parent كمسار احتياطي.
 */
function bridgeResponse(token, result) {
  const message = {
    type: 'T20_CENTRAL_DB_RESPONSE',
    token: String(token || ''),
    result: result
  };
  const safeMessage = JSON.stringify(message)
    .replace(/</g, '\\u003c')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
  const html = '<!doctype html><html><head><meta charset="utf-8"></head><body>' +
    '<script>(function(){' +
      'var message=' + safeMessage + ';' +
      'function send(){' +
        'try{window.top.postMessage(message,"*");}catch(e){}' +
        'try{window.parent.postMessage(message,"*");}catch(e){}' +
      '}' +
      'send();setTimeout(send,250);setTimeout(send,1000);' +
    '})();<\/script>' +
    '</body></html>';
  return HtmlService
    .createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function jsonpResponse(callbackName, obj) {
  const callback = String(callbackName || '');
  if (!/^[A-Za-z_$][\w$\.]*$/.test(callback)) {
    return ContentService
      .createTextOutput('/* Invalid JSONP callback */')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(obj) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/**
 * جلب جميع البيانات من جميع الأوراق وتركيب كائن موحد
 */
function getAllData() {
  const db = {};

  // أسماء مفاتيح SHEET_SCHEMAS هي أسماء الأوراق نفسها وليست مفاتيح SHEET_NAMES.
  Object.entries(ENTITY_SHEETS).forEach(([entityKey, sheetName]) => {
    const rows = getSheetData(sheetName);
    // الواجهة تتوقع القوائم المرجعية كنصوص، بينما الورقة تحفظها كصفوف id/name.
    db[entityKey] = (entityKey === 'sectors' || entityKey === 'directorates')
      ? rows.map(row => row.name).filter(Boolean)
      : rows;
  });

  // إضافة معلومات المزامنة
  const meta = getSyncMeta();
  meta.changeSequence = latestChangeSequence_();
  db._meta = meta;

  return db;
}

/**
 * حفظ جميع البيانات في الأوراق المقابلة
 */
function applyProjectBudgetTotals(projects, budgets) {
  const totals = (Array.isArray(budgets) ? budgets : []).reduce((acc, item) => {
    const projectId = String((item && item.projectId) || '');
    if (!projectId) return acc;
    const planned = Number(item.planned);
    const spent = Number(item.spent);
    if (!acc[projectId]) acc[projectId] = { allocatedBudget: 0, actualCost: 0 };
    acc[projectId].allocatedBudget += Number.isFinite(planned) ? planned : 0;
    acc[projectId].actualCost += Number.isFinite(spent) ? spent : 0;
    return acc;
  }, {});
  return (Array.isArray(projects) ? projects : []).map(project => {
    const total = totals[String((project && project.id) || '')] || { allocatedBudget: 0, actualCost: 0 };
    return { ...project, allocatedBudget: total.allocatedBudget, actualCost: total.actualCost };
  });
}

function setAllData(db) {
  const now = new Date();

  // الترتيب المحدد للمراجع
  const entityOrder = Object.keys(ENTITY_SHEETS);

  entityOrder.forEach(entityKey => {
    const sheetName = ENTITY_SHEETS[entityKey];
    // دعم المفتاح القديم team_members عند مزامنة نسخة محلية سابقة.
    let data = entityKey === 'teamMembers'
      ? (db.teamMembers || db.team_members || [])
      : (db[entityKey] || []);
    if (entityKey === 'projects') {
      data = applyProjectBudgetTotals(data, db.budgets || []);
    }
    // تحويل القيم النصية في القوائم المرجعية إلى صفوف قابلة للحفظ والاسترجاع.
    if ((entityKey === 'sectors' || entityKey === 'directorates') && Array.isArray(data)) {
      data = data.map((item, index) => {
        if (item && typeof item === 'object') return item;
        const name = String(item || '').trim();
        return { id: entityKey + '_' + (index + 1), name: name };
      }).filter(item => item.name);
    }
    if (Array.isArray(data) && data.length > 0) {
      // إضافة timestamps
      const dataWithTimestamps = data.map(item => ({
        ...item,
        updatedAt: now.toISOString(),
        createdAt: item.createdAt || now.toISOString()
      }));
      setSheetData(sheetName, dataWithTimestamps);
    } else {
      // حتى لو فارغة، نضمن وجود الورقة مع الهيدر
      ensureSheet(sheetName, SHEET_SCHEMAS[sheetName]);
    }
  });

  // جلب معلومات المزامنة الحالية أولاً
  const currentMeta = getSyncMeta();

  // تحديث معلومات المزامنة
  setSyncMeta({
    lastSyncAt: now.toISOString(),
    version: (currentMeta.version || 0) + 1,
    recordCounts: entityOrder.reduce((acc, key) => {
      acc[key] = (db[key] || []).length;
      return acc;
    }, {})
  });
}

/**
 * جلب بيانات ورقة واحدة وتحويلها لمصفوفة كائنات
 */
function getSheetData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const timeZone = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();

  if (!sheet) {
    return [];
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return []; // فقط الهيدر موجود
  }

  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(1, 1, lastRow, lastCol);
  const values = range.getValues();

  const headers = values[0];
  const rows = values.slice(1);
  cleanExistingIsoDateCells(sheet, sheetName, headers, rows, timeZone);

  return rows.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      if (header) {
        let value = row[index];
        // تحويل التواريخ والأرقام
        if (isDateOnlyField(sheetName, header)) {
          value = normalizeDateOnly(value, timeZone);
        } else if (value instanceof Date) {
          value = value.toISOString();
        } else if (typeof value === 'number' && !Number.isInteger(value)) {
          // الاحتفاظ بالأرقام العشرية كما هي
        } else if (typeof value === 'string' && value === '') {
          value = null;
        }
        obj[header] = value;
      }
    });
    return obj;
  }).filter(obj => {
    if (!Object.keys(obj).length) return false;
    if (sheetName === SHEET_NAMES.ROLE_PERMISSIONS) return !!(obj.role && obj.pageId);
    if (sheetName === SHEET_NAMES.AUDIT_LOG) return !!obj.timestamp;
    if (sheetName === SHEET_NAMES.CHANGE_LOG) return obj.sequence !== null && obj.sequence !== undefined && obj.sequence !== '';
    if (sheetName === SHEET_NAMES.SYNC_META) return !!obj.key;
    return !!obj.id;
  }); // لكل ورقة مفتاح منطقي مختلف؛ لا تشترط id على سجلات التدقيق والصلاحيات
}

/**
 * كتابة بيانات في ورقة محددة (استبدال كامل)
 */
function setSheetData(sheetName, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const schema = SHEET_SCHEMAS[sheetName];

  if (!schema) {
    throw new Error('Unknown sheet: ' + String(sheetName) + '. Check ENTITY_SHEETS mapping.');
  }

  if (!Array.isArray(data)) {
    throw new Error('Invalid data for sheet ' + sheetName + ': expected an array');
  }

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  // مسح المحتوى الحالي (عدا الهيدر)
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, schema.length).clearContent();
  }

  // كتابة الهيدر إذا لم يكن موجوداً
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, schema.length).setValues([schema]);
    // تنسيق الهيدر
    const headerRange = sheet.getRange(1, 1, 1, schema.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1e3a5f');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  if (data.length === 0) {
    return; // لا توجد بيانات للكتابة
  }

  // إعداد البيانات للكتابة
  const rowsData = data.map(item => {
    return schema.map(field => {
      let value = item[field];
      if (isDateOnlyField(sheetName, field)) {
        return normalizeDateOnly(value, ss.getSpreadsheetTimeZone());
      }
      if (value instanceof Date) {
        return value;
      }
      if (value === null || value === undefined) {
        return '';
      }
      return value;
    });
  });

  // كتابة البيانات دفعة واحدة
  const startRow = 2;
  // إبقاء حقول اليوم التقويمي كنص YYYY-MM-DD لمنع إضافة الوقت والمنطقة الزمنية.
  (DATE_ONLY_FIELDS[sheetName] || []).forEach(field => {
    const column = schema.indexOf(field) + 1;
    if (column > 0) sheet.getRange(startRow, column, rowsData.length, 1).setNumberFormat('@');
  });
  sheet.getRange(startRow, 1, rowsData.length, schema.length).setValues(rowsData);

  // ضبط عرض الأعمدة تلقائياً
  sheet.autoResizeColumns(1, schema.length);
}

/**
 * التأكد من وجود ورقة مع الهيدر الصحيح
 */
function ensureSheet(sheetName, schema) {
  if (!sheetName || !Array.isArray(schema)) {
    throw new Error('Invalid sheet configuration: ' + String(sheetName));
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, schema.length).setValues([schema]);
    const headerRange = sheet.getRange(1, 1, 1, schema.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1e3a5f');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * جلب معلومات المزامنة
 */
function getSyncMeta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SYNC_META);

  if (!sheet || sheet.getLastRow() < 2) {
    return { version: 0, lastSyncAt: null, recordCounts: {} };
  }

  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3);
  const values = range.getValues();

  const meta = { version: 0, lastSyncAt: null, recordCounts: {} };
  values.forEach(row => {
    const key = row[0];
    const value = row[1];
    if (key === 'version') meta.version = parseInt(value) || 0;
    else if (key === 'lastSyncAt') meta.lastSyncAt = value;
    else if (key === 'recordCounts') {
      try { meta.recordCounts = JSON.parse(value); } catch(e) {}
    }
  });

  return meta;
}

/**
 * حفظ معلومات المزامنة
 */
function setSyncMeta(meta) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.SYNC_META);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.SYNC_META);
    sheet.getRange(1, 1, 1, 3).setValues([['key', 'value', 'updatedAt']]);
    const headerRange = sheet.getRange(1, 1, 1, 3);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#1e3a5f');
    headerRange.setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  const now = new Date().toISOString();
  const rows = [
    ['version', meta.version || 0, now],
    ['lastSyncAt', meta.lastSyncAt || now, now],
    ['recordCounts', JSON.stringify(meta.recordCounts || {}), now]
  ];

  // مسح القديم وكتابة الجديد
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
  }
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
}

/**
 * إنشاء الاستجابة JSON
 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * دالة مساعدة للتهيئة اليدوية (يمكن تشغيلها من محرر Apps Script)
 * تنشئ جميع الأوراق مع الهيدر الصحيح
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.entries(SHEET_SCHEMAS).forEach(([sheetName, schema]) => {
    ensureSheet(sheetName, schema);
  });

  // ورقة المزامنة
  ensureSheet(SHEET_NAMES.SYNC_META, SHEET_SCHEMAS[SHEET_NAMES.SYNC_META]);

  // حذف الورقة القديمة إذا وجدت
  const oldSheet = ss.getSheetByName('CentralDB');
  if (oldSheet) {
    // نقل البيانات القديمة إن وجدت
    migrateFromOldFormat(oldSheet);
    // يمكن حذفها أو الاحتفاظ بها للنسخ الاحتياطي
    // ss.deleteSheet(oldSheet);
  }

  SpreadsheetApp.getUi().alert('تم تهيئة جميع الأوراق بنجاح!');
}

/**
 * ترحيل البيانات من التنسيق القديم (JSON في خلية واحدة) للتنسيق الجديد
 */
function migrateFromOldFormat(oldSheet) {
  try {
    const jsonCell = oldSheet.getRange(2, 2).getValue();
    if (!jsonCell) return;

    const oldData = JSON.parse(jsonCell);
    if (!oldData || typeof oldData !== 'object') return;

    // تعيين البيانات في الأوراق الجديدة
    const mapping = {
      'projects': SHEET_NAMES.PROJECTS,
      'objectives': SHEET_NAMES.OBJECTIVES,
      'kpis': SHEET_NAMES.KPIS,
      'tasks': SHEET_NAMES.TASKS,
      'risks': SHEET_NAMES.RISKS,
      'budgets': SHEET_NAMES.BUDGETS,
      'services': SHEET_NAMES.SERVICES,
      'team_members': SHEET_NAMES.TEAM_MEMBERS,
      'resources': SHEET_NAMES.RESOURCES,
      'fiveYearPlans': SHEET_NAMES.FIVE_YEAR_PLANS,
      'sectors': SHEET_NAMES.SECTORS,
      'directorates': SHEET_NAMES.DIRECTORATES
    };

    Object.entries(mapping).forEach(([oldKey, newSheetName]) => {
      if (oldData[oldKey] && Array.isArray(oldData[oldKey])) {
        const dataWithTimestamps = oldData[oldKey].map(item => ({
          ...item,
          createdAt: item.createdAt || new Date().toISOString(),
          updatedAt: new Date().toISOString()
        }));
        setSheetData(newSheetName, dataWithTimestamps);
      }
    });

    console.log('Migration completed successfully');
  } catch (err) {
    console.error('Migration failed:', err);
  }
}

/**
 * دالة اختبار: طباعة هيكل البيانات الحالي
 */
function debugPrintStructure() {
  const db = getAllData();
  Object.keys(db).forEach(key => {
    if (key !== '_meta') {
      console.log(`${key}: ${db[key].length} records`);
      if (db[key].length > 0) {
        console.log('  Sample:', JSON.stringify(db[key][0], null, 2));
      }
    }
  });
  console.log('Meta:', JSON.stringify(db._meta, null, 2));
}




function setUserPasswordFromPrompt() {
  const ui = SpreadsheetApp.getUi();

  const usernameResponse = ui.prompt(
    'تعيين كلمة مرور',
    'أدخل اسم المستخدم:',
    ui.ButtonSet.OK_CANCEL
  );

  if (usernameResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const passwordResponse = ui.prompt(
    'تعيين كلمة مرور',
    'أدخل كلمة المرور الجديدة:',
    ui.ButtonSet.OK_CANCEL
  );

  if (passwordResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const username = usernameResponse.getResponseText().trim();
  const password = passwordResponse.getResponseText();

  setCentralUserPassword(username, password);

  ui.alert('تم تعيين كلمة المرور للمستخدم بنجاح.');
}
