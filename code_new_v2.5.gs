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
  SYNC_META: 'SyncMeta' // لتخزين معلومات المزامنة (آخر تحديث، إصدار، إلخ)
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
  ]
};

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

/**
 * نقطة الدخول للطلبات POST (من الواجهة الأمامية)
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    const action = payload.action;

    if (action === 'getDB') {
      return jsonResponse({ ok: true, db: getAllData() });
    }

    if (action === 'setDB') {
      const db = payload.db || {};
      setAllData(db);
      // إرجاع version و updatedAt لتحديث meta في الواجهة الأمامية
      const updatedMeta = getSyncMeta();
      return jsonResponse({
        ok: true,
        message: 'Data saved successfully',
        version: updatedMeta.version,
        updatedAt: updatedMeta.lastSyncAt
      });
    }

    if (action === 'getSheet') {
      // جلب ورقة محددة فقط
      const sheetName = payload.sheetName;
      if (sheetName && SHEET_SCHEMAS[sheetName]) {
        return jsonResponse({ ok: true, data: getSheetData(sheetName) });
      }
      return jsonResponse({ ok: false, error: 'Invalid sheet name' });
    }

    if (action === 'setSheet') {
      // حفظ ورقة محددة فقط
      const sheetName = payload.sheetName;
      const data = payload.data || [];
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
      return bridgeResponse(params.token, { ok: true, db: getAllData() });
    }

    const result = params.action === 'getDB'
      ? { ok: true, db: getAllData() }
      : {
          ok: true,
          message: 'MTCIT PPM Central Storage is running.',
          version: '2.5.0',
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
      return bridgeResponse(params.token, errorResult);
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
  db._meta = meta;

  return db;
}

/**
 * حفظ جميع البيانات في الأوراق المقابلة
 */
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
  }).filter(obj => Object.keys(obj).length > 0 && obj.id); // استبعاد الصفوف الفارغة
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
