function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Attendance') || ss.getSheets()[0];
  var action = e.parameter.action || 'save';

  if (action === 'report') {
    return handleReport(ss, sheet, e);
  }
  if (action === 'requestSubmit') {
    return handleRequestSubmit(ss, e);
  }
  if (action === 'requestsReport') {
    return handleRequestsReport(ss, e);
  }
  if (action === 'requestSetStatus') {
    return handleRequestSetStatus(ss, e);
  }
  if (action === 'registerPush') {
    return handleRegisterPush(ss, e);
  }
  if (action === 'dashboardData') {
    return handleDashboardData(ss, sheet, e);
  }
  if (action === 'nextOrderNumber') {
    return handleNextOrderNumber(ss, e);
  }
  if (action === 'orderSubmit') {
    return handleOrderSubmit(ss, e);
  }
  if (action === 'ordersSearch') {
    return handleOrdersSearch(ss, e);
  }

  // ---- default action: save a new attendance record ----
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['الاسم', 'النوع', 'التاريخ', 'الوقت', 'خط العرض', 'خط الطول', 'رابط الموقع', 'الجهاز']);
  }

  var name   = e.parameter.name   || '';
  var type   = e.parameter.type   || '';
  var date   = e.parameter.date   || '';
  var time   = e.parameter.time   || '';
  var lat    = e.parameter.lat    || '';
  var lng    = e.parameter.lng    || '';
  var map    = e.parameter.map    || '';
  var device = e.parameter.device || '';

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    // امنع التكرار: لو نفس الموظف ونفس النوع (حضور/انصراف) اتسجل بالفعل في
    // نفس اليوم ده (حتى لو من جهاز تاني في نفس اللحظة تقريباً)، ما تضيفش صف
    // جديد، وارجع بيانات الصف الأصلي بدل منه.
    var tz = ss.getSpreadsheetTimeZone();
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
      for (var i = 0; i < values.length; i++) {
        var r = values[i];
        if (r[0] === name && r[1] === type && normalizeDate_(r[2], tz) === date) {
          return ContentService.createTextOutput(JSON.stringify({
            ok: true, alreadyRecorded: true, time: normalizeTime_(r[3], tz)
          })).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

    sheet.appendRow([name, type, date, time, lat, lng, map, device]);

    // بعت إشعار فوري لموبايل الإداري بتسجيل الحضور/الانصراف، بنفس فكرة
    // إشعارات طلبات الإذن/المأمورية/الإجازة بالظبط — لو حصل أي خطأ في
    // الإرسال، ده مبيأثرش على تسجيل الحضور نفسه (اتسجل فعلاً فوق).
    sendPushToAll_(
      name + ' — ' + (type === 'حضور' ? 'سجّل حضور' : 'سجّل انصراف'),
      name + ' سجّل ' + type + ' الساعة ' + time + (device ? (' — ' + device) : '')
    );

    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ---- read-only: return the day's attendance records for the dashboard ----
function handleReport(ss, sheet, e) {
  var rows = getAttendanceRows_(ss, sheet, e.parameter.date || '');
  return ContentService.createTextOutput(JSON.stringify({ ok: true, rows: rows }))
    .setMimeType(ContentService.MimeType.JSON);
}

// بيبني قايمة صفوف الحضور/الانصراف (مستخدمة من handleReport وكمان من
// handleDashboardData عشان نتجنب تكرار نفس الكود).
function getAttendanceRows_(ss, sheet, dateFilter) {
  var tz = ss.getSpreadsheetTimeZone();
  var lastRow = sheet.getLastRow();
  var rows = [];

  if (lastRow > 1) {
    var values = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    for (var i = 0; i < values.length; i++) {
      var r = values[i];
      var rowDate = normalizeDate_(r[2], tz);
      if (dateFilter && rowDate !== dateFilter) continue;
      rows.push({
        name: r[0],
        type: r[1],
        date: rowDate,
        time: normalizeTime_(r[3], tz),
        lat:  r[4],
        lng:  r[5],
        map:  r[6],
        device: r[7]
      });
    }
  }
  return rows;
}

// ---- طلبات الإذن والمأمورية والإجازة: تاب منفصل اسمه "Requests" في نفس
// الشيت ----
// الأعمدة: المعرف | الاسم | النوع (إذن/مأمورية/إجازة) | التاريخ | من الساعة |
// إلى الساعة | السبب | الحالة | وقت الطلب | من تاريخ الإجازة | إلى تاريخ الإجازة
// آخر عمودين (من/إلى تاريخ الإجازة) مستخدمين لطلب "إجازة" بس (فترة أيام)،
// وفاضلين فاضيين لطلبات الإذن/المأمورية العاديين (بتاعتهم "من الساعة/إلى
// الساعة" زي ما هو). ضفتهم في آخر الأعمدة (مش في النص) عشان الصفوف القديمة
// المسجلة قبل كده تفضل صح من غير أي تعديل يدوي.
// الحالة بتبدأ "قيد المراجعة"، وبتتغيّر لـ "موافق" أو "مرفوض" من التطبيق
// نفسه (شاشة "اعتماد الإذن والمأموريات") — مفيش تعديل يدوي مطلوب في الشيت.
function getRequestsSheet_(ss) {
  var headers = ['المعرف', 'الاسم', 'النوع', 'التاريخ', 'من الساعة', 'إلى الساعة', 'السبب', 'الحالة', 'وقت الطلب', 'من تاريخ الإجازة', 'إلى تاريخ الإجازة'];
  var sheet = ss.getSheetByName('Requests');
  if (!sheet) {
    sheet = ss.insertSheet('Requests');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    // لو تاب "Requests" كان موجود من قبل بترويسة قديمة (من نسخة أقدم من
    // الكود، فيها أعمدة أقل أو مرتبة غير كده)، نظبط صف العناوين تلقائي
    // عشان يتطابق مع البيانات اللي فعلاً بتتسجل تحته (id/من/إلى الساعة..).
    var currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    var matches = true;
    for (var i = 0; i < headers.length; i++) {
      if (currentHeaders[i] !== headers[i]) { matches = false; break; }
    }
    if (!matches) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  return sheet;
}

function handleRequestSubmit(ss, e) {
  var sheet = getRequestsSheet_(ss);
  var id     = e.parameter.id     || Utilities.getUuid();
  var name   = e.parameter.name   || '';
  var type   = e.parameter.type   || ''; // 'إذن' أو 'مأمورية'
  var date   = e.parameter.date   || '';
  var from   = e.parameter.from   || '';
  var to     = e.parameter.to     || '';
  var reason = e.parameter.reason || '';
  var leaveFrom = e.parameter.leaveFrom || '';
  var leaveTo   = e.parameter.leaveTo   || '';
  var tz = ss.getSpreadsheetTimeZone();
  var submittedAt = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    sheet.appendRow([id, name, type, date, from, to, reason, 'قيد المراجعة', submittedAt, leaveFrom, leaveTo]);
  } finally {
    lock.releaseLock();
  }

  // بعت إشعار فوري لموبايل الإداري (لو ميزة الإشعارات متظبطة) — لو حصل أي
  // خطأ في الإرسال، ده مبيأثرش على تسجيل الطلب في الشيت أصلاً (اتسجل
  // فعلاً فوق قبل ما نوصل للسطر ده).
  sendPushToAll_(
    'طلب ' + type + ' جديد',
    name + ' طلب ' + type + (from ? (' من ' + from) : '') + (to ? (' إلى ' + to) : '') + (reason ? (' — ' + reason) : '')
  );

  return ContentService.createTextOutput(JSON.stringify({ ok: true, id: id }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleRequestsReport(ss, e) {
  var rows = getRequestRows_(ss, e.parameter.date || '');
  return ContentService.createTextOutput(JSON.stringify({ ok: true, rows: rows }))
    .setMimeType(ContentService.MimeType.JSON);
}

// بيبني قايمة صفوف طلبات الإذن/المأمورية/الإجازة (مستخدمة من
// handleRequestsReport وكمان من handleDashboardData).
function getRequestRows_(ss, dateFilter) {
  var sheet = ss.getSheetByName('Requests');
  var tz = ss.getSpreadsheetTimeZone();
  var rows = [];

  if (sheet && sheet.getLastRow() > 1) {
    var lastRow = sheet.getLastRow();
    var values = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
    for (var i = 0; i < values.length; i++) {
      var r = values[i];
      var rowDate = normalizeDate_(r[3], tz);
      var leaveFrom = normalizeDate_(r[9], tz);
      var leaveTo = normalizeDate_(r[10], tz);
      // طلبات "إجازة" بتاخد أكتر من يوم — لازم نتأكد إن التاريخ المطلوب
      // واقع جوه مدى [leaveFrom, leaveTo] بدل ما نقارنه بتاريخ أول يوم بس
      // (اللي كان بيخلي الطلب يختفي من اليوم التاني للإجازة وطالع).
      if (dateFilter) {
        if (r[2] === 'إجازة' && leaveFrom) {
          var leaveEndDay = leaveTo || leaveFrom;
          if (dateFilter < leaveFrom || dateFilter > leaveEndDay) continue;
        } else if (rowDate !== dateFilter) {
          continue;
        }
      }
      rows.push({
        id: r[0],
        name: r[1],
        type: r[2],
        date: rowDate,
        from: normalizeTime_(r[4], tz),
        to: normalizeTime_(r[5], tz),
        reason: r[6],
        status: r[7],
        submittedAt: r[8],
        leaveFrom: leaveFrom,
        leaveTo: leaveTo
      });
    }
  }
  return rows;
}

// ---- بيرجع الحضور/الانصراف وطلبات الإذن/المأمورية سوا في طلب واحد بس،
// عشان لوحة الحضور تحمّل بسرعة (بدل ما تعمل نداءين متوازيين لنفس
// السكريبت، اللي كان بيسبب "جاري التحميل" يفضل شغال لفترة طويلة على
// الموبايل لو الطلبين اتصادفوا في نفس اللحظة). ----
function handleDashboardData(ss, sheet, e) {
  var date = e.parameter.date || '';
  var rows = getAttendanceRows_(ss, sheet, date);
  var requestRows = getRequestRows_(ss, date);
  return ContentService.createTextOutput(JSON.stringify({ ok: true, rows: rows, requestRows: requestRows }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------
// سجل أوامر الشراء (تاب "Orders") + رقم الأمر التلقائي + البحث من التطبيق
// ---------------------------------------------------------------------
// آخر رقم بيتاخد من تاب "Orders" نفسه (أكبر رقم موجود + 1) عشان منحتاجش
// تاب إضافي منفصل بس عشان عداد. عشان نضمن عدم تكرار الرقم حتى لو كذا
// شخص من كذا منطقة فتحوا النموذج في نفس اللحظة، أول ما حد ياخد رقم
// جديد بنحجزه فورًا بصف مؤقت في نفس الشيت (جوه LockService)، فأي طلب
// تاني للرقم التالي هيلاقي الرقم ده محجوز ويطلع اللي بعده مباشرة. وبعدين
// لما الأمر يتحفظ فعليًا (عند التحميل/المشاركة) الصف/الصفوف المحجوزة
// بتتمسح ويتحط بدالها الصفوف الحقيقية.
//
// كل صنف (قطر) في الأمر بياخد سطر لوحده في الشيت، وكل سطوره بتحمل نفس
// رقم الأمر وبيانات العميل/المورد (مكررة على كل سطر) — يعني أمر فيه
// قطرين بيطلع سطرين، مش سطر واحد فيه كل حاجة مجمّعة. عمود JSON خام
// (آخر عمود قبل وقت الحفظ) بيحمل كل أصناف الأمر مع بعض في كل سطوره،
// للاستخدام الداخلي بس (عشان التطبيق يقدر يرجّع الأمر بكل أصنافه لو حد
// فتحه تاني للتعديل أو لإعادة تنزيل نفس الـ PDF).
var ORDERS_HEADERS_ = ['رقم الأمر', 'التاريخ', 'اسم المورد', 'اسم العميل', 'عناية', 'عنوان التوصيل', 'مسئول التواصل', 'رقم التواصل', 'اسم المندوب', 'ملاحظات', 'القطر (مم)', 'الكمية (طن)', 'المواصفة', 'النوع', 'إجمالي الكمية (طن)', 'الأصناف (بيانات النظام)', 'وقت الحفظ'];
var ORD_COL_NUM_ = 1;
var ORD_COL_DIAMETER_ = 11;
var ORD_COL_QTY_ = 12;
var ORD_COL_SPEC_ = 13;
var ORD_COL_TYPE_ = 14;
var ORD_COL_TOTAL_ = 15;
var ORD_COL_ITEMS_JSON_ = 16;

function getOrdersSheet_(ss) {
  var sheet = ss.getSheetByName('Orders');
  if (!sheet) {
    sheet = ss.insertSheet('Orders');
    sheet.appendRow(ORDERS_HEADERS_);
    sheet.getRange(1, 1, 1, ORDERS_HEADERS_.length).setFontWeight('bold');
    return sheet;
  }
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headerVals = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headerVals[ORD_COL_DIAMETER_ - 1] === ORDERS_HEADERS_[ORD_COL_DIAMETER_ - 1]) {
    return sheet; // التصميم بالفعل محدّث (سطر لكل صنف)
  }
  // ترقية تلقائية من أي نسخة قديمة من شيت "Orders" (سواء بعمود JSON خام
  // واحد، أو بأعمدة أقطار/كميات مجمّعة في نفس السطر) لتصميم "سطر لكل
  // صنف". بنلاقي الأعمدة الثابتة بالاسم (زي "رقم الأمر"، "وقت الحفظ"...)
  // في السطر القديم، وعمود بيانات الأصناف (JSON) عن طريق شكل محتواه
  // (بيبدأ بـ "[") مش بالاسم، عشان الترقية تشتغل مهما كان اسم/مكان
  // العمود ده في أي نسخة سابقة — وبعدين بنفكّ كل أصناف السطر القديم
  // لسطر مستقل لكل صنف.
  var lastRow = sheet.getLastRow();
  var oldVals = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  var idxOf = function (name) {
    for (var c = 0; c < headerVals.length; c++) { if (headerVals[c] === name) return c; }
    return -1;
  };
  var coreIdx = {
    num: idxOf('رقم الأمر'), date: idxOf('التاريخ'), supplier: idxOf('اسم المورد'),
    client: idxOf('اسم العميل'), attention: idxOf('عناية'), address: idxOf('عنوان التوصيل'),
    contactName: idxOf('مسئول التواصل'), contactPhone: idxOf('رقم التواصل'),
    rep: idxOf('اسم المندوب'), notes: idxOf('ملاحظات'), savedAt: idxOf('وقت الحفظ')
  };
  var g = function (row, idx) { return idx > -1 ? row[idx] : ''; };
  var newRows = [];
  for (var r = 0; r < oldVals.length; r++) {
    var row = oldVals[r];
    var jsonIdx = -1;
    for (var c = 0; c < row.length; c++) {
      if (String(row[c] || '').trim().charAt(0) === '[') { jsonIdx = c; break; }
    }
    var itemsJson = jsonIdx > -1 ? String(row[jsonIdx] || '[]') : '[]';
    var numRaw = g(row, coreIdx.num);
    var numInt = extractOrderNumberInt_(numRaw);
    var numFixed = numInt ? ('0000' + numInt).slice(-4) : numRaw;
    var core = [
      numFixed, g(row, coreIdx.date), g(row, coreIdx.supplier), g(row, coreIdx.client), g(row, coreIdx.attention),
      g(row, coreIdx.address), g(row, coreIdx.contactName), g(row, coreIdx.contactPhone), g(row, coreIdx.rep), g(row, coreIdx.notes)
    ];
    var exploded = explodeOrderItems_(itemsJson);
    var savedAt = g(row, coreIdx.savedAt);
    if (exploded.items.length === 0) {
      newRows.push(core.concat(['', '', '', '', exploded.total, itemsJson, savedAt]));
    } else {
      for (var v = 0; v < exploded.items.length; v++) {
        var it = exploded.items[v];
        newRows.push(core.concat([it.diameter, it.qty, it.spec, it.type, exploded.total, itemsJson, savedAt]));
      }
    }
  }
  sheet.clear();
  sheet.getRange(1, ORD_COL_NUM_, Math.max(sheet.getMaxRows(), newRows.length + 1), 1).setNumberFormat('@');
  sheet.appendRow(ORDERS_HEADERS_);
  sheet.getRange(1, 1, 1, ORDERS_HEADERS_.length).setFontWeight('bold');
  if (newRows.length) {
    sheet.getRange(2, 1, newRows.length, ORDERS_HEADERS_.length).setValues(newRows);
  }
  return sheet;
}

// بيقرا أي رقم أمر (نص أو رقم، بأصفار على الشمال أو من غيرها) ويطلّع منه
// الرقم الصحيح بس — بيرجع 0 لو مفيش أرقام خالص (صف فاضي مثلاً).
function extractOrderNumberInt_(v) {
  var m = String(v || '').match(/\d+/);
  return m ? parseInt(m[0], 10) : 0;
}

// بيحوّل مصفوفة الأصناف (JSON) لقائمة أصناف صالحة (كل واحد بقيمه
// الأربعة) + إجمالي الكمية — عشان كل صنف يتحط في سطر مستقل بعد كده.
function explodeOrderItems_(itemsJsonStr) {
  var raw = [];
  try { raw = JSON.parse(itemsJsonStr || '[]'); } catch (err) { raw = []; }
  var items = [];
  var total = 0;
  for (var i = 0; i < raw.length; i++) {
    var it = raw[i] || {};
    var diameter = String(it.diameter || '').trim();
    var qty = String(it.qty || '').trim();
    var spec = String(it.spec || '').trim();
    var type = String(it.type || '').trim();
    if (!diameter && !qty && !spec && !type) continue;
    items.push({ diameter: diameter, qty: qty, spec: spec, type: type });
    var qtyNum = parseFloat(qty);
    if (!isNaN(qtyNum)) total += qtyNum;
  }
  return { items: items, total: total ? String(total) : '' };
}

// نص مختصر (سطر واحد) بيجمع القطر مع الكمية لكل صنف — يُستخدم بس في
// عرض نتائج البحث داخل التطبيق (مش في الشيت نفسه).
function summarizeOrderItemsForSearch_(itemsJsonStr) {
  var exploded = explodeOrderItems_(itemsJsonStr);
  var lines = exploded.items.map(function (it) {
    var parts = [];
    if (it.diameter) parts.push('قطر ' + it.diameter + ' مم');
    if (it.qty) parts.push(it.qty + ' طن');
    return parts.join(' × ');
  });
  return lines.join('، ');
}

// بيرجع رقم تسلسلي جديد كل مرة، بمنتهى الأمان حتى لو أكتر من شخص من
// أكتر من منطقة طلبوا رقم في نفس اللحظة بالظبط — LockService بيقفل
// تنفيذ السكريبت لحظيًا، وبنحجز الرقم فورًا بصف مؤقت في شيت "Orders"
// نفسه (مفيش تاب تاني) عشان محدش تاني ياخد نفس الرقم قبل ما الأمر يتحفظ.
function handleNextOrderNumber(ss, e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = getOrdersSheet_(ss);
    var lastRow = sheet.getLastRow();
    var max = 0;
    if (lastRow > 1) {
      var nums = sheet.getRange(2, ORD_COL_NUM_, lastRow - 1, 1).getValues();
      for (var i = 0; i < nums.length; i++) {
        var n = extractOrderNumberInt_(nums[i][0]);
        if (n > max) max = n;
      }
    }
    var next = max + 1;
    var formatted = ('0000' + next).slice(-4);

    var newRowIndex = lastRow + 1;
    var placeholder = new Array(ORDERS_HEADERS_.length).fill('');
    placeholder[ORD_COL_NUM_ - 1] = formatted;
    placeholder[9] = 'محجوز تلقائيًا لحد ما يتحفظ الأمر فعليًا';
    placeholder[ORD_COL_ITEMS_JSON_ - 1] = '[]';
    placeholder[ORDERS_HEADERS_.length - 1] = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sheet.getRange(newRowIndex, ORD_COL_NUM_).setNumberFormat('@');
    sheet.getRange(newRowIndex, 1, 1, placeholder.length).setValues([placeholder]);

    return ContentService.createTextOutput(JSON.stringify({ ok: true, number: formatted }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// كل أمر بيتحفظ لحظة ما حد ينزّل أو يشارك ملف الـ PDF بتاعه — بيمسح أي
// سطور قديمة لنفس رقم الأمر (زي الصف المحجوز من nextOrderNumber، أو
// سطور أمر سابق بعدد أصناف مختلف لو الأمر ده اتعدّل)، وبعدين بيحط سطر
// مستقل لكل صنف من أصناف الأمر الحالي.
function handleOrderSubmit(ss, e) {
  var num = (e.parameter.num || '').trim();
  if (!num) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'missing order number' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var sheet = getOrdersSheet_(ss);
  var tz = ss.getSpreadsheetTimeZone();
  var itemsJson = e.parameter.items || '[]';
  var exploded = explodeOrderItems_(itemsJson);
  var savedAt = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  var core = [
    num,
    e.parameter.date || '',
    e.parameter.supplier || '',
    e.parameter.client || '',
    e.parameter.attention || '',
    e.parameter.address || '',
    e.parameter.contactName || '',
    e.parameter.contactPhone || '',
    e.parameter.rep || '',
    e.parameter.notes || ''
  ];
  var newRows = exploded.items.length
    ? exploded.items.map(function (it) {
        return core.concat([it.diameter, it.qty, it.spec, it.type, exploded.total, itemsJson, savedAt]);
      })
    : [core.concat(['', '', '', '', exploded.total, itemsJson, savedAt])];

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var numInt = extractOrderNumberInt_(num);
    var lastRow = sheet.getLastRow();
    if (lastRow > 1 && numInt) {
      var ids = sheet.getRange(2, ORD_COL_NUM_, lastRow - 1, 1).getValues();
      // بنمسح كل السطور القديمة لنفس رقم الأمر من تحت لفوق (عشان مسح
      // سطر ميغيّرش ترقيم السطور اللي لسه هنمسحها بعده)، وبعدين بنضيف
      // سطر مستقل لكل صنف من الأصناف الحالية.
      for (var r = ids.length - 1; r >= 0; r--) {
        if (extractOrderNumberInt_(ids[r][0]) === numInt) {
          sheet.deleteRow(r + 2);
        }
      }
    }
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, ORD_COL_NUM_, newRows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, 1, newRows.length, ORDERS_HEADERS_.length).setValues(newRows);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// بيدوّر في رقم الأمر + اسم العميل + اسم المورد + التاريخ عن أي جزء يطابق
// النص المكتوب (من غير حساسية لحالة الأحرف)، وبيرجع أحدث 30 نتيجة —
// نتيجة واحدة لكل رقم أمر (حتى لو ليه أكتر من سطر/صنف في الشيت)، وبيتجاهل
// أي صف لسه محجوز بس ماتحفظش فيه أمر فعلي لحد دلوقتي.
function handleOrdersSearch(ss, e) {
  var q = (e.parameter.q || '').trim().toLowerCase();
  var sheet = ss.getSheetByName('Orders');
  var rows = [];
  var seen = {};
  if (sheet && q && sheet.getLastRow() > 1) {
    var lastRow = sheet.getLastRow();
    var vals = sheet.getRange(2, 1, lastRow - 1, ORDERS_HEADERS_.length).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      var r = vals[i];
      var num = String(r[0] || '');
      var date = String(r[1] || '');
      var supplier = String(r[2] || '');
      var client = String(r[3] || '');
      if (!date && !supplier && !client) continue; // صف محجوز لسه من غير أمر فعلي
      if (seen[num]) continue; // نفس الأمر ظهر قبل كده (سطر صنف تاني بس)
      var haystack = (num + ' ' + date + ' ' + supplier + ' ' + client).toLowerCase();
      if (haystack.indexOf(q) === -1) continue;
      seen[num] = true;
      var itemsJsonVal = String(r[ORD_COL_ITEMS_JSON_ - 1] || '[]');
      rows.push({
        num: num,
        date: date,
        supplier: supplier,
        client: client,
        attention: r[4],
        address: r[5],
        contactName: r[6],
        contactPhone: r[7],
        rep: r[8],
        notes: r[9],
        itemsSummary: summarizeOrderItemsForSearch_(itemsJsonVal),
        items: itemsJsonVal
      });
      if (rows.length >= 30) break;
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true, rows: rows }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---- بيغيّر عمود "الحالة" لطلب معين (بمعرّفه الفريد) لـ "موافق" أو
// "مرفوض" — ده اللي بيخلي الاعتماد يتم من داخل التطبيق نفسه من غير ما
// حد يفتح الشيت يدوياً. ----
function handleRequestSetStatus(ss, e) {
  var sheet = ss.getSheetByName('Requests');
  var id = e.parameter.id || '';
  var status = e.parameter.status || '';

  if (!sheet || !id || !status) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'missing id or status' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(id)) {
          sheet.getRange(i + 2, 8).setValue(status); // العمود التامن = الحالة
          return ContentService.createTextOutput(JSON.stringify({ ok: true }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'not found' }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------
// إشعارات الموبايل الفورية (Firebase Cloud Messaging)
// ---------------------------------------------------------------------
// إعداد لازم يتعمل مرة واحدة (اتفعل بعد ما يكون عندك مشروع Firebase):
// 1) Extensions > Libraries في محرر Apps Script > أضف مكتبة بالـ Script ID:
//    1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF (مكتبة OAuth2
//    الرسمية من جوجل) — اختار آخر إصدار واحفظ باسم "OAuth2".
// 2) Project Settings (⚙️) > Script Properties > أضف خاصية واحدة بس:
//    FCM_SERVICE_ACCOUNT_JSON = محتوى ملف الـ JSON بتاع حساب الخدمة كامل
//    زي ما هو (افتح الملف بـ Notepad، Ctrl+A ثم Ctrl+C، والصقه بالكامل —
//    من { الأولى لحد } الأخيرة — كقيمة واحدة). كده مفيش احتمال إن جزء من
//    المفتاح يتقطع زي ما بيحصل لو حاولت تنسخ private_key لوحده.
// لو الخاصية دي لسه مش متظبطة أو الـ JSON فيه غلطة، الدالة دي هترجع من
// غير ما تعمل حاجة ومن غير ما توقف تسجيل الطلب في الشيت.
function getServiceAccountInfo_() {
  var raw = PropertiesService.getScriptProperties().getProperty('FCM_SERVICE_ACCOUNT_JSON');
  if (!raw) return null;
  try {
    var info = JSON.parse(raw);
    console.log('getServiceAccountInfo_: تم تحليل الملف بنجاح — project_id=' + info.project_id +
      ' | طول المفتاح=' + (info.private_key ? info.private_key.length : 0));
    return info;
  } catch (e) {
    console.log('getServiceAccountInfo_: تعذر تحليل JSON_ACCOUNT_SERVICE_FCM - ' + e);
    return null;
  }
}

function getFcmService_(info) {
  return OAuth2.createService('FCM')
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setPrivateKey(info.private_key)
    .setIssuer(info.client_email)
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setScope('https://www.googleapis.com/auth/firebase.messaging');
}

// بيسجل كل خطوة في تاب "PushLog" في نفس الشيت — أسهل بكتير من فتح "سجل
// التنفيذ" في محرر Apps Script، لأنك بتفتحه زي أي تاب عادي في نفس الشيت
// اللي بتشتغل عليه أصلاً.
function logPush_(message) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('PushLog');
    if (!sheet) {
      sheet = ss.insertSheet('PushLog');
      sheet.appendRow(['الوقت', 'الرسالة']);
    }
    var tz = ss.getSpreadsheetTimeZone();
    sheet.appendRow([Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss'), message]);
  } catch (e) {
    // متعمل حاجة — التسجيل نفسه ميقفش أي وظيفة تانية.
  }
}

function getPushTokensSheet_(ss) {
  var sheet = ss.getSheetByName('PushTokens');
  if (!sheet) {
    sheet = ss.insertSheet('PushTokens');
    sheet.appendRow(['التوكن', 'تاريخ التسجيل']);
  }
  return sheet;
}

// بيسجل توكن جهاز جديد (من زرار "فعّل إشعارات الجوال" في التطبيق) —
// بيتجاهل التكرار لو نفس التوكن مسجل قبل كده.
function handleRegisterPush(ss, e) {
  var token = e.parameter.token || '';
  if (!token) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: 'no token' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  var sheet = getPushTokensSheet_(ss);
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var existing = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < existing.length; i++) {
        if (existing[i][0] === token) {
          return ContentService.createTextOutput(JSON.stringify({ ok: true, alreadyRegistered: true }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
    }
    var tz = ss.getSpreadsheetTimeZone();
    sheet.appendRow([token, Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss')]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// بيبعت إشعار Push لكل الأجهزة المسجلة في تاب "PushTokens". لو ميزة
// الإشعارات لسه مش متظبطة (Script Properties فاضية) أو حصل أي خطأ، بترجع
// بهدوء من غير ما توقف باقي الكود.
function sendPushToAll_(title, body) {
  logPush_('sendPushToAll_: بدأ التشغيل — العنوان: ' + title);
  try {
    var info = getServiceAccountInfo_();
    if (!info) { logPush_('sendPushToAll_: FCM_SERVICE_ACCOUNT_JSON مش متظبط أو فيه غلطة في Script Properties'); return; }
    var projectId = info.project_id;
    logPush_('sendPushToAll_: تم قراءة بيانات الحساب بنجاح — project_id=' + projectId);

    var service = getFcmService_(info);
    if (!service.hasAccess()) {
      logPush_('sendPushToAll_: تعذر الحصول على توكن الدخول - ' + service.getLastError());
      return;
    }
    var accessToken = service.getAccessToken();
    logPush_('sendPushToAll_: تم الحصول على توكن الدخول بنجاح');

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('PushTokens');
    if (!sheet || sheet.getLastRow() < 2) { logPush_('sendPushToAll_: مفيش أي أجهزة مسجلة في تاب PushTokens'); return; }
    var tokens = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    logPush_('sendPushToAll_: عدد الأجهزة المسجلة = ' + tokens.length);

    tokens.forEach(function (row) {
      var token = row[0];
      if (!token) return;
      // بنبعت الإشعار كـ "data" مش "notification" عمداً — لما الـ payload
      // فيه notification، بعض الأجهزة (خصوصًا أندرويد) بتعرض الإشعار
      // مرتين: مرة تلقائي من نظام FCM نفسه ومرة تانية من الكود بتاعنا في
      // firebase-messaging-sw.js، فتحس إنه "متكرر". بـ data-only، إحنا
      // اللي بنتحكم في عرض الإشعار بنفسنا مرة واحدة بس (من firebase-
      // messaging-sw.js)، فمفيش تكرار تاني.
      var payload = {
        message: {
          token: token,
          data: { title: String(title), body: String(body) }
        }
      };
      var response = UrlFetchApp.fetch('https://fcm.googleapis.com/v1/projects/' + projectId + '/messages:send', {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + accessToken },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      var code = response.getResponseCode();
      if (code !== 200) {
        logPush_('sendPushToAll_: فشل الإرسال (كود ' + code + ') لتوكن ينتهي بـ ...' + String(token).slice(-8) + ' - ' + response.getContentText());
      } else {
        logPush_('sendPushToAll_: تم الإرسال بنجاح لتوكن ينتهي بـ ...' + String(token).slice(-8));
      }
    });
  } catch (e) {
    logPush_('sendPushToAll_ exception: ' + e);
  }
}

// دالة اختبار: شغّلها يدوياً من محرر Apps Script (اختار اسمها من قائمة
// الدوال جنب زرار "تشغيل" في الأعلى، وبعدين دوس "تشغيل") عشان تختبر
// الإشعارات على طول من غير ما تحتاج تعمل طلب إذن حقيقي من التطبيق. لو
// دي أول مرة، ممكن يطلب منك "تفويض" الصلاحيات — وافق عليها.
// بعد التشغيل (أو بعد أي طلب إذن حقيقي من التطبيق)، افتح تاب "PushLog" في
// نفس الشيت (هيتعمل تلقائي أول مرة) وهتلاقي فيه كل خطوة بالترتيب وبالوقت —
// أسهل بكتير من "سجل التنفيذ" في محرر الكود.
function testSendPush() {
  sendPushToAll_('إشعار تجريبي 🔔', 'لو وصلك ده، الإشعارات شغالة تمام!');
}

function normalizeDate_(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  }
  return String(v || '');
}

function normalizeTime_(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, tz, 'HH:mm:ss');
  }
  return String(v || '');
}

// ---------------------------------------------------------------------
// تقرير التأخير والأذونات الشهري
// ---------------------------------------------------------------------
// بيحسب لكل موظف، عن شهر معيّن:
//  - عدد أيام التأخير + إجمالي دقائق التأخير (بمقارنة وقت "حضور" الفعلي
//    بميعاد الحضور الرسمي بتاعه، من تاب "مواعيد الحضور")
//  - عدد طلبات "إذن" المعتمدة، وعدد طلبات "مأمورية" المعتمدة
//  - عدد أيام "إجازة" المعتمدة (بيحسب كل الأيام من "من تاريخ" لـ "إلى
//    تاريخ" اللي واقعة جوه الشهر المطلوب)
// النتيجة بتتكتب في تاب "التقرير الشهري". التاب ده وتاب "مواعيد الحضور"
// بيتعملوا تلقائي أول مرة، ومفيش أي حاجة تانية مطلوبة غير إنك تملي ميعاد
// حضور كل موظف في تاب "مواعيد الحضور" مرة واحدة.

var ATTENDANCE_EMPLOYEES_SEED_ = ['أحمد مفرح', 'محمود كارم', 'مني محمد', 'شهد خالد', 'سيد كارم', 'محمود رجب', 'صبري يحي'];

// عمود "ب" = ميعاد الحضور الرسمي، عمود "ج" = ميعاد الانصراف الرسمي.
// لو التاب موجود بالفعل من قبل (نسخة أقدم فيها عمود الحضور بس)، بيضيف
// عمود الانصراف تلقائي من غير ما يلمس أي مواعيد اتكتبت بالفعل، وبيضيف كمان
// أي موظف جديد اتضاف في ATTENDANCE_EMPLOYEES_SEED_ ولسه مش موجود صف ليه.
function getOfficialTimesSheet_(ss) {
  var headers = ['الاسم', 'ميعاد الحضور الرسمي', 'ميعاد الانصراف الرسمي'];
  var sheet = ss.getSheetByName('مواعيد الحضور');
  if (!sheet) {
    sheet = ss.insertSheet('مواعيد الحضور');
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    ATTENDANCE_EMPLOYEES_SEED_.forEach(function (name) {
      sheet.appendRow([name, '', '']);
    });
    sheet.getRange('B2:C').setNumberFormat('HH:mm');
    sheet.autoResizeColumns(1, headers.length);
    return sheet;
  }

  // تحديث الترويسة لو ناقصها عمود "ميعاد الانصراف الرسمي" (نسخة أقدم).
  var currentHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  var headerMismatch = false;
  for (var i = 0; i < headers.length; i++) {
    if (currentHeaders[i] !== headers[i]) { headerMismatch = true; break; }
  }
  if (headerMismatch) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }

  // إضافة أي موظف جديد لسه مالوش صف، من غير ما نلمس صفوف/مواعيد الموظفين
  // الموجودين بالفعل.
  var lastRow = sheet.getLastRow();
  var existingNames = {};
  if (lastRow > 1) {
    var col = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    col.forEach(function (r) { if (r[0]) existingNames[r[0]] = true; });
  }
  ATTENDANCE_EMPLOYEES_SEED_.forEach(function (name) {
    if (!existingNames[name]) sheet.appendRow([name, '', '']);
  });
  return sheet;
}

function getMonthlyReportSheet_(ss) {
  var sheet = ss.getSheetByName('التقرير الشهري');
  if (!sheet) {
    sheet = ss.insertSheet('التقرير الشهري');
  }
  return sheet;
}

// بيرجع "HH:mm" من قيمة الخلية سواء كانت وقت (Date) أو نص مكتوب يدوي.
function normalizeTimeHHMM_(v, tz) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, tz, 'HH:mm');
  }
  var s = String(v || '').trim();
  var m = s.match(/^(\d{1,2}):(\d{2})/);
  return m ? (('0' + m[1]).slice(-2) + ':' + m[2]) : '';
}

// الفرق بالدقايق بين الوقت الفعلي والميعاد الرسمي (موجب = متأخر بكام دقيقة).
function timeDiffMinutes_(officialHHMM, actualTimeStr) {
  var o = String(officialHHMM || '').match(/^(\d{1,2}):(\d{2})/);
  var a = String(actualTimeStr || '').match(/^(\d{1,2}):(\d{2})/);
  if (!o || !a) return 0;
  var oMin = parseInt(o[1], 10) * 60 + parseInt(o[2], 10);
  var aMin = parseInt(a[1], 10) * 60 + parseInt(a[2], 10);
  return aMin - oMin;
}

// عدد أيام إجازة (من leaveFrom لـ leaveTo شاملين) اللي واقعة جوه شهر
// monthStr ('yyyy-MM') بالظبط — لو الإجازة بتمتد لشهر تاني، بيتحسب بس
// الجزء اللي جوه الشهر المطلوب.
function countLeaveDaysInMonth_(fromStr, toStr, monthStr) {
  if (!fromStr) return 0;
  toStr = toStr || fromStr;
  var from = new Date(fromStr + 'T00:00:00');
  var to = new Date(toStr + 'T00:00:00');
  if (isNaN(from.getTime()) || isNaN(to.getTime())) return 0;
  var monthStart = new Date(monthStr + '-01T00:00:00');
  if (isNaN(monthStart.getTime())) return 0;
  var monthEnd = new Date(monthStart.getFullYear(), monthStart.getMonth() + 1, 0);
  var start = from > monthStart ? from : monthStart;
  var end = to < monthEnd ? to : monthEnd;
  if (start > end) return 0;
  var msPerDay = 24 * 60 * 60 * 1000;
  return Math.round((end - start) / msPerDay) + 1;
}

// بيحسب ويكتب تقرير شهر معيّن ('yyyy-MM') — لو مبعتش، بياخد الشهر الحالي.
function refreshMonthlyReport_(monthStr) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  monthStr = (monthStr && /^\d{4}-\d{2}$/.test(monthStr)) ? monthStr : Utilities.formatDate(new Date(), tz, 'yyyy-MM');

  // 1) مواعيد الحضور والانصراف الرسمية لكل موظف
  var timesSheet = getOfficialTimesSheet_(ss);
  var officialIn = {}, officialOut = {};
  if (timesSheet.getLastRow() > 1) {
    var timesData = timesSheet.getRange(2, 1, timesSheet.getLastRow() - 1, 3).getValues();
    timesData.forEach(function (r) {
      var name = r[0];
      if (!name) return;
      var tIn = normalizeTimeHHMM_(r[1], tz);
      var tOut = normalizeTimeHHMM_(r[2], tz);
      if (tIn) officialIn[name] = tIn;
      if (tOut) officialOut[name] = tOut;
    });
  }

  // 2) التأخير (حضور بعد الميعاد) والانصراف المبكر (انصراف قبل الميعاد)
  // من تاب "Attendance"
  var lateCount = {}, lateMinutes = {};
  var earlyCount = {}, earlyMinutes = {};
  var attSheet = ss.getSheetByName('Attendance') || ss.getSheets()[0];
  if (attSheet && attSheet.getLastRow() > 1) {
    var attVals = attSheet.getRange(2, 1, attSheet.getLastRow() - 1, 8).getValues();
    attVals.forEach(function (r) {
      var name = r[0], type = r[1];
      if (!name) return;
      var dateStr = normalizeDate_(r[2], tz);
      if (dateStr.slice(0, 7) !== monthStr) return;
      var timeStr = normalizeTime_(r[3], tz);
      if (type === 'حضور') {
        var official = officialIn[name];
        if (!official) return; // مفيش ميعاد حضور رسمي متظبط لموظف ده لسه
        var diff = timeDiffMinutes_(official, timeStr);
        if (diff > 0) {
          lateCount[name] = (lateCount[name] || 0) + 1;
          lateMinutes[name] = (lateMinutes[name] || 0) + diff;
        }
      } else if (type === 'انصراف') {
        var officialO = officialOut[name];
        if (!officialO) return; // مفيش ميعاد انصراف رسمي متظبط لموظف ده لسه
        var diffOut = timeDiffMinutes_(officialO, timeStr);
        if (diffOut < 0) {
          earlyCount[name] = (earlyCount[name] || 0) + 1;
          earlyMinutes[name] = (earlyMinutes[name] || 0) + (-diffOut);
        }
      }
    });
  }

  // 3) الأذونات/المأموريات/الإجازات المعتمدة من تاب "Requests"
  var permitCount = {}, missionCount = {}, leaveDays = {};
  var reqSheet = ss.getSheetByName('Requests');
  if (reqSheet && reqSheet.getLastRow() > 1) {
    var reqVals = reqSheet.getRange(2, 1, reqSheet.getLastRow() - 1, 11).getValues();
    reqVals.forEach(function (r) {
      var name = r[1], type = r[2], status = r[7];
      if (!name || status !== 'موافق') return;
      if (type === 'إذن') {
        var d1 = normalizeDate_(r[3], tz);
        if (d1.slice(0, 7) === monthStr) permitCount[name] = (permitCount[name] || 0) + 1;
      } else if (type === 'مأمورية') {
        var d2 = normalizeDate_(r[3], tz);
        if (d2.slice(0, 7) === monthStr) missionCount[name] = (missionCount[name] || 0) + 1;
      } else if (type === 'إجازة') {
        var lf = normalizeDate_(r[9], tz), lt = normalizeDate_(r[10], tz);
        var days = countLeaveDaysInMonth_(lf, lt, monthStr);
        if (days > 0) leaveDays[name] = (leaveDays[name] || 0) + days;
      }
    });
  }

  // 4) قايمة كل الموظفين اللي ليهم أي بيانات (مواعيد متظبطة، تأخير، أو طلبات)
  var names = {};
  ATTENDANCE_EMPLOYEES_SEED_.forEach(function (n) { names[n] = true; });
  Object.keys(officialIn).forEach(function (n) { names[n] = true; });
  Object.keys(officialOut).forEach(function (n) { names[n] = true; });
  Object.keys(lateCount).forEach(function (n) { names[n] = true; });
  Object.keys(earlyCount).forEach(function (n) { names[n] = true; });
  Object.keys(permitCount).forEach(function (n) { names[n] = true; });
  Object.keys(missionCount).forEach(function (n) { names[n] = true; });
  Object.keys(leaveDays).forEach(function (n) { names[n] = true; });

  // 5) كتابة التقرير
  var reportHeaders = ['الاسم', 'عدد أيام التأخير', 'إجمالي دقائق التأخير', 'عدد أيام الانصراف المبكر', 'إجمالي دقائق الانصراف المبكر', 'عدد الأذونات المعتمدة', 'عدد المأموريات المعتمدة', 'عدد أيام الإجازة المعتمدة'];
  var reportSheet = getMonthlyReportSheet_(ss);
  reportSheet.clear();
  reportSheet.appendRow(['تقرير شهر: ' + monthStr]);
  reportSheet.getRange(1, 1).setFontWeight('bold').setFontSize(13);
  reportSheet.appendRow(reportHeaders);
  reportSheet.getRange(2, 1, 1, reportHeaders.length).setFontWeight('bold').setBackground('#f1eee9');

  Object.keys(names).sort().forEach(function (name) {
    reportSheet.appendRow([
      name,
      lateCount[name] || 0,
      lateMinutes[name] || 0,
      earlyCount[name] || 0,
      earlyMinutes[name] || 0,
      permitCount[name] || 0,
      missionCount[name] || 0,
      leaveDays[name] || 0
    ]);
  });

  reportSheet.autoResizeColumns(1, reportHeaders.length);
  reportSheet.setFrozenRows(2);
}

// ---------------------------------------------------------------------
// أعمدة إضافية في نفس شيت الحضور الأساسي (اللي فيه تسجيلات حضور/انصراف —
// "Attendance" أو "ورقة1" لو التاب لسه مسمّى كده) — بتضيف 3 أعمدة بعد
// الأعمدة الأصلية التمنية مباشرة:
//   التأخير | نوع طلب اليوم | حالة الطلب
// "التأخير" بتظهر "متأخر" أو "لا" لصفوف "حضور" بس (بالمقارنة بميعاد
// الحضور الرسمي من تاب "مواعيد الحضور")، وبتفضل فاضية لصفوف "انصراف" ولو
// مفيش ميعاد رسمي متظبط للموظف ده. "نوع طلب اليوم"/"حالة الطلب" بيدوروا
// في تاب "Requests" على أي طلب إذن/مأمورية/إجازة يخص نفس الموظف ونفس
// اليوم (الإجازة بتتحسب لو اليوم واقع بين تاريخ بدايتها ونهايتها).
// ---------------------------------------------------------------------

function getAttendanceMainSheet_(ss) {
  return ss.getSheetByName('Attendance') || ss.getSheets()[0];
}

var ATTENDANCE_EXTRA_HEADERS_ = ['التأخير', 'نوع طلب اليوم', 'حالة الطلب'];
var ATTENDANCE_EXTRA_START_COL_ = 9; // بعد الأعمدة التمنية الأصلية (الاسم..الجهاز)

function ensureAttendanceExtraHeaders_(sheet) {
  var current = sheet.getRange(1, ATTENDANCE_EXTRA_START_COL_, 1, ATTENDANCE_EXTRA_HEADERS_.length).getValues()[0];
  var mismatch = false;
  for (var i = 0; i < ATTENDANCE_EXTRA_HEADERS_.length; i++) {
    if (current[i] !== ATTENDANCE_EXTRA_HEADERS_[i]) { mismatch = true; break; }
  }
  if (mismatch) {
    sheet.getRange(1, ATTENDANCE_EXTRA_START_COL_, 1, ATTENDANCE_EXTRA_HEADERS_.length)
      .setValues([ATTENDANCE_EXTRA_HEADERS_]).setFontWeight('bold');
  }
}

// قايمة مبسطة من تاب "Requests" (اسم، من يوم، إلى يوم، نوع، حالة) — بنبنيها
// مرة واحدة ونستخدمها لكل صفوف الحضور، بدل ما ندوّر في الشيت لكل صف لوحده.
function buildRequestsIndex_(ss, tz) {
  var sheet = ss.getSheetByName('Requests');
  var index = [];
  if (!sheet || sheet.getLastRow() < 2) return index;
  var vals = sheet.getRange(2, 1, sheet.getLastRow() - 1, 11).getValues();
  vals.forEach(function (r) {
    var name = r[1], type = r[2], status = r[7];
    if (!name) return;
    var dateStr = normalizeDate_(r[3], tz);
    var leaveFrom = normalizeDate_(r[9], tz);
    var leaveTo = normalizeDate_(r[10], tz);
    var fromDay = (type === 'إجازة' && leaveFrom) ? leaveFrom : dateStr;
    var toDay = (type === 'إجازة' && leaveTo) ? leaveTo : dateStr;
    if (!fromDay) return;
    index.push({ name: name, from: fromDay, to: (toDay || fromDay), type: type, status: status });
  });
  return index;
}

function findRequestForDay_(index, name, dateStr) {
  if (!dateStr) return null;
  for (var i = 0; i < index.length; i++) {
    var it = index[i];
    if (it.name === name && dateStr >= it.from && dateStr <= it.to) return it;
  }
  return null;
}

// بيحسب ويكتب الأعمدة الإضافية التلاتة لكل صفوف شيت الحضور الحالية.
function refreshAttendanceExtraColumns_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var sheet = getAttendanceMainSheet_(ss);
  if (!sheet) return;

  // الترويسة لازم تتظبط حتى لو الشيت لسه فاضي من البيانات (أول مرة قبل ما
  // أي حد يسجل حضور خالص).
  if (sheet.getLastRow() >= 1) ensureAttendanceExtraHeaders_(sheet);
  if (sheet.getLastRow() < 2) return; // مفيش صفوف بيانات لسه نحسبلها حاجة

  var timesSheet = getOfficialTimesSheet_(ss);
  var officialIn = {};
  if (timesSheet.getLastRow() > 1) {
    var timesData = timesSheet.getRange(2, 1, timesSheet.getLastRow() - 1, 3).getValues();
    timesData.forEach(function (r) {
      var name = r[0];
      if (!name) return;
      var tIn = normalizeTimeHHMM_(r[1], tz);
      if (tIn) officialIn[name] = tIn;
    });
  }

  var reqIndex = buildRequestsIndex_(ss, tz);

  var lastRow = sheet.getLastRow();
  var vals = sheet.getRange(2, 1, lastRow - 1, 4).getValues(); // الاسم|النوع|التاريخ|الوقت
  var out = vals.map(function (r) {
    var name = r[0], type = r[1];
    var dateStr = normalizeDate_(r[2], tz);
    var timeStr = normalizeTime_(r[3], tz);

    var lateText = '';
    if (type === 'حضور') {
      var official = officialIn[name];
      if (official) {
        lateText = (timeDiffMinutes_(official, timeStr) > 0) ? 'متأخر' : 'لا';
      }
    }

    var req = findRequestForDay_(reqIndex, name, dateStr);
    return [lateText, req ? req.type : '', req ? req.status : ''];
  });

  sheet.getRange(2, ATTENDANCE_EXTRA_START_COL_, out.length, 3).setValues(out);
}

// شغّلها يدوياً (أو من قايمة "📊 تقرير الحضور" فوق) لتحديث تقرير الشهر
// الحالي وأعمدة التأخير/الحالة في شيت الحضور مع بعض.
function refreshCurrentMonthReport() {
  refreshAttendanceExtraColumns_();
  refreshMonthlyReport_();
}

// بتظهر بوكس صغير تكتب فيه أي شهر عايز تقريره (بصيغة yyyy-MM، مثلاً
// 2026-08) — مفيدة لو عايز تراجع شهر فات.
function promptRefreshReportForMonth() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('تحديث تقرير شهر محدد', 'اكتب الشهر بصيغة yyyy-MM (مثال: 2026-08)', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var val = resp.getResponseText().trim();
  if (!/^\d{4}-\d{2}$/.test(val)) {
    ui.alert('صيغة غير صحيحة. لازم تكتبه بالشكل ده بالظبط: 2026-08');
    return;
  }
  refreshAttendanceExtraColumns_();
  refreshMonthlyReport_(val);
  ui.alert('تم تحديث تقرير شهر ' + val + '.');
}

// بيضيف قايمة "📊 تقرير الحضور" فوق في الشيت، وبيحدّث تقرير الشهر الحالي
// تلقائي أول ما تفتح الملف — عشان التقرير يفضل حديث من غير ما حد يحتاج
// يعمل حاجة يدوي.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 تقرير الحضور')
    .addItem('تحديث كل شيء (التأخير + التقرير الشهري)', 'refreshCurrentMonthReport')
    .addItem('تحديث تقرير شهر محدد...', 'promptRefreshReportForMonth')
    .addToUi();
  try { refreshCurrentMonthReport(); } catch (e) { /* أول مرة قبل ما البيانات تتظبط - متعمل حاجة */ }
}

// ============================================================
// نسخ احتياطي تلقائي يومي
// ============================================================
// بتعمل نسخة كاملة من الشيت (بكل التابات) وتحطها في مجلد خاص في Drive،
// وكمان تبعت نسخة Excel على الإيميل — عشان لو حصل أي حظر أو مشكلة في
// المشروع، تكون في إيدك نسخة كاملة جاهزة ترفعها تاني من غير ما تخسر
// أي بيانات. لازم نضيف مشغّل (Trigger) وقتي يشغّلها مرة كل يوم —
// من "الساعة ⏰ Triggers" في القائمة الجانبية اليسار.
function backupNow() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var file = DriveApp.getFileById(ss.getId());
    var folderName = 'نسخ احتياطية - حضور وانصراف شيل اوت';
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    var tz = ss.getSpreadsheetTimeZone();
    var stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd_HH-mm');
    file.makeCopy('نسخة احتياطية - ' + stamp, folder);

    // تنظيف: خلي آخر 30 نسخة بس عشان المجلد ميمتلئش
    var iter = folder.getFiles();
    var list = [];
    while (iter.hasNext()) list.push(iter.next());
    list.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
    for (var i = 30; i < list.length; i++) { list[i].setTrashed(true); }

    // نسخة Excel كمان تتبعت بالإيميل، عشان تكون متاحة حتى لو حصل حظر
    // على المشروع أو الشيت نفسه.
    var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
    var token = ScriptApp.getOAuthToken();
    var resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    if (resp.getResponseCode() === 200) {
      var blob = resp.getBlob().setName('حضور_وانصراف_' + stamp + '.xlsx');
      MailApp.sendEmail({
        to: Session.getActiveUser().getEmail(),
        subject: 'نسخة احتياطية - حضور وانصراف شيل اوت - ' + stamp,
        body: 'نسخة احتياطية تلقائية يومية من شيت الحضور والانصراف.\n\nلو حصلت أي مشكلة في الشيت الأساسي، تقدر تفتح المرفق ده مباشرة أو تلاقي نسخة منه محفوظة في مجلد "' + folderName + '" في Drive بتاعك.',
        attachments: [blob]
      });
    }

    logPush_('backupNow_: تم عمل نسخة احتياطية بنجاح - ' + stamp);
  } catch (err) {
    logPush_('backupNow_ exception: ' + err);
  }
}
