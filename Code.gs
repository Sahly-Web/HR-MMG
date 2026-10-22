/**
 * ============================================================================
 * نظام إدارة الموارد البشرية المخصص لمساحات العمل المشتركة (Coworking Space HRMS)
 * Google Apps Script Backend + Automatic Sheets Architect & Formatter
 * الملف: Code.gs
 * ============================================================================
 */

const SYSTEM_CONFIG = {
  APP_TITLE: 'Coworking Space Enterprise HRMS',
  TIMEZONE: 'Africa/Cairo',
  CURRENCY_FORMAT: '#,##0.00 "ج.م"',
  PERCENT_FORMAT: '0.0"%"',
  DATE_FORMAT: 'yyyy-mm-dd',
  DATETIME_FORMAT: 'yyyy-mm-dd hh:mm:ss',
  HEADER_BG: '#0f172a',
  HEADER_FG: '#ffffff',
  ACCENT_COLOR: '#0d9488'
};

/**
 * تعريف هيكل الجداول التسعة وأعمدتها وألوانها المميزة للتنسيق التلقائي
 */
const SHEETS_SCHEMA = {
  Branches: {
    tabColor: '#0d9488',
    headers: [
      'معرف الفرع (ID)',
      'كود الفرع',
      'اسم الفرع',
      'المدينة',
      'خط العرض (Lat)',
      'خط الطول (Lng)',
      'نطاق البصمة (متر)',
      'تشغيل 24/7',
      'الحد الأدنى للاستقبال',
      'سعة المكاتب الكلية',
      'نسبة الإشغال الحالية %',
      'مدير الفرع',
      'الحالة'
    ],
    widths: [130, 100, 190, 110, 115, 115, 125, 100, 140, 130, 140, 160, 100],
    currencyCols: [],
    percentCols: [11]
  },
  Employees: {
    tabColor: '#2563eb',
    headers: [
      'معرف الموظف (ID)',
      'الرقم الوظيفي',
      'الاسم الكامل',
      'الرقم القومي',
      'الهاتف',
      'البريد الإلكتروني',
      'الفرع الأساسي',
      'القسم',
      'الدور الوظيفي',
      'نوع العقد',
      'الراتب الأساسي',
      'أجر الساعة (للجزئي)',
      'مسموح بالتنقل (Roaming)',
      'توزيع التكلفة بين الفروع',
      'خاضع للتأمينات',
      'تاريخ التعيين',
      'الحالة'
    ],
    widths: [130, 110, 180, 140, 130, 190, 160, 130, 160, 140, 130, 130, 145, 200, 115, 120, 110],
    currencyCols: [11, 12],
    percentCols: []
  },
  Shifts_Attendance: {
    tabColor: '#7c3aed',
    headers: [
      'معرف السجل (ID)',
      'التاريخ',
      'معرف الموظف',
      'اسم الموظف',
      'الفرع المسجل به',
      'نوع الوردية',
      'وقت الحضور',
      'وقت الانصراف',
      'ساعات العمل الفعلية',
      'ساعات إضافية (Overtime)',
      'إحداثيات الموظف (Lat,Lng)',
      'المسافة عن الفرع (متر)',
      'حالة الـ Geofence',
      'نوع الحضور (أساسي/متنقل)'
    ],
    widths: [130, 115, 120, 175, 160, 140, 110, 110, 135, 145, 180, 145, 140, 150],
    currencyCols: [],
    percentCols: []
  },
  Sales_Commissions: {
    tabColor: '#059669',
    headers: [
      'معرف العملية (ID)',
      'تاريخ الفاتورة',
      'رقم فاتورة BMS',
      'اسم العضو / الشركة',
      'الفرع',
      'نوع العملية',
      'فئة المنتج (Coworking)',
      'عدد المكاتب/الساعات',
      'مدة العقد (شهور)',
      'صافي الفاتورة (بدون ضريبة)',
      'الموظف الأساسي (Closer)',
      'نسبة الأساسي %',
      'الموظف المساعد (Tour Host)',
      'نسبة المساعد %',
      'نسبة العمولة المطبقة %',
      'إجمالي العمولة المحتسبة',
      'عمولة الأساسي',
      'عمولة المساعد',
      'الحالة'
    ],
    widths: [130, 115, 140, 185, 150, 135, 165, 130, 125, 165, 170, 115, 175, 115, 140, 155, 130, 130, 130],
    currencyCols: [10, 16, 17, 18],
    percentCols: [12, 14, 15]
  },
  Payroll_Slips: {
    tabColor: '#d97706',
    headers: [
      'معرف القسيمة (ID)',
      'دورة الرواتب (الشهر)',
      'معرف الموظف',
      'اسم الموظف',
      'الفرع الأساسي',
      'نوع العقد',
      'أيام/ساعات العمل',
      'الأجر الأساسي المستحق',
      'أجر العمل الإضافي',
      'عمولات المبيعات والتجديد',
      'حافز إشغال الفرع و KPIs',
      'إجمالي الاستحقاقات (Gross)',
      'حصة الموظف تأمينات (11%)',
      'ضريبة كسب العمل',
      'خصم استرداد عمولة (Clawback) / جزاءات',
      'إجمالي الاستقطاعات',
      'صافي الراتب المستحق (Net)',
      'حالة الصرف'
    ],
    widths: [130, 130, 120, 175, 155, 135, 130, 150, 135, 160, 155, 165, 155, 135, 180, 145, 165, 120],
    currencyCols: [8, 9, 10, 11, 12, 13, 14, 15, 16, 17],
    percentCols: []
  },
  Leaves_Coverage: {
    tabColor: '#db2777',
    headers: [
      'معرف الطلب (ID)',
      'تاريخ التقديم',
      'معرف الموظف',
      'اسم الموظف',
      'الفرع',
      'الدور الوظيفي',
      'نوع الإجازة',
      'من تاريخ',
      'إلى تاريخ',
      'عدد الأيام',
      'الموظف البديل للتغطية',
      'فحص تغطية الاستقبال الآلي',
      'حالة الاعتماد'
    ],
    widths: [130, 120, 120, 175, 155, 150, 130, 115, 115, 100, 175, 195, 130],
    currencyCols: [],
    percentCols: []
  },
  Shift_Swaps_ESS: {
    tabColor: '#4f46e5',
    headers: [
      'معرف التبديل (ID)',
      'تاريخ الطلب',
      'الموظف الطالب',
      'وردية الطالب وتاريخها',
      'الزميل المستهدف',
      'وردية الزميل وتاريخها',
      'الفرع',
      'فحص منع الوردية المزدوجة (Double Shift)',
      'موافقة الزميل',
      'اعتماد مدير الفرع'
    ],
    widths: [130, 120, 170, 185, 170, 185, 150, 210, 130, 140],
    currencyCols: [],
    percentCols: []
  },
  ATS_KPIs: {
    tabColor: '#0284c7',
    headers: [
      'معرف السجل (ID)',
      'النوع (توظيف / تقييم KPI)',
      'الاسم (المرشح / الموظف)',
      'الفرع',
      'الوظيفة / الدور',
      'تاريخ يوم المعايشة (Trial Shift) / الشهر',
      'تقييم رضا الأعضاء (NPS)',
      'معدل الاحتفاظ بالأعضاء (Retention %)',
      'تحقيق تارجت المبيعات %',
      'النتيجة النهائية من 100',
      'القرار / التوصية'
    ],
    widths: [130, 150, 180, 155, 155, 185, 150, 185, 155, 150, 160],
    currencyCols: [],
    percentCols: [8, 9]
  },
  Weekly_Roster: {
    tabColor: '#0d9488',
    headers: ['معرف الموظف', 'اسم الموظف', 'الفرع', 'السبت', 'الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'إجمالي الساعات الأسبوعية'],
    widths: [120, 175, 160, 145, 145, 145, 145, 145, 145, 145, 150],
    currencyCols: [],
    percentCols: []
  },
  Offboarding_Assets: {
    tabColor: '#dc2626',
    headers: [
      'معرف الإخلاء (ID)',
      'تاريخ الإجراء',
      'معرف الموظف',
      'اسم الموظف',
      'الفرع',
      'نوع الإجراء (إخلاء طرف / جزاء)',
      'العهدة المستلمة (Laptop / Keys)',
      'إيقاف بطاقة Smart Lock RFID',
      'إيقاف صلاحية نظام BMS والواي فاي',
      'تسوية المستحقات النهائية',
      'الحالة'
    ],
    widths: [130, 120, 120, 175, 150, 170, 195, 175, 195, 155, 120],
    currencyCols: [10],
    percentCols: []
  }
};

/**
 * إضافة قائمة مخصصة داخل Google Sheets عند فتح الملف
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 نظام HR المساحة المشتركة')
    .addItem('1️⃣ تهيئة وتنسيق جميع الشيتات تلقائياً (Setup & Auto-Format)', 'setupAndFormatDatabase')
    .addItem('2️⃣ إدراج بيانات تجريبية واقعية (Seed Demo Data)', 'seedInitialCoworkingData')
    .addSeparator()
    .addItem('🖥️ فتح واجهة النظام التفاعلية داخل الشيت', 'openAppModal')
    .addToUi();
}

/**
 * فتح واجهة index.html كنافذة كاملة داخل Google Sheets
 */
function openAppModal() {
  const html = HtmlService.createHtmlOutputFromFile('index')
    .setWidth(1280)
    .setHeight(820)
    .setTitle(SYSTEM_CONFIG.APP_TITLE);
  SpreadsheetApp.getUi().showModalDialog(html, 'نظام إدارة الموارد البشرية لمساحات العمل المشتركة');
}

/**
 * نقطة الدخول لتطبيق الويب (Web App Entry Point)
 * تدعم عرض واجهة index.html أو الرد بـ JSON عند تمرير ?action=...
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    const result = handleApiRequest(e.parameter.action, e.parameter);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(SYSTEM_CONFIG.APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * استقبال طلبات POST من واجهة الويب (Fetch API)
 */
function doPost(e) {
  try {
    const body = e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    const action = body.action || (e.parameter && e.parameter.action);
    const payload = body.payload || body;
    const result = handleApiRequest(action, payload);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      ok: false,
      error: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * الدالة المركزية لتوجيه الطلبات (تعمل مع google.script.run ومع doGet/doPost)
 */
function handleApiRequest(action, payload) {
  try {
    switch (action) {
      case 'INIT_AND_FORMAT':
        return setupAndFormatDatabase(true);
      case 'GET_ALL_DATA':
        return getAllDatabaseState();
      case 'ADD_BRANCH':
        return appendRowToSheet('Branches', payload.rowArray);
      case 'ADD_EMPLOYEE':
        return appendRowToSheet('Employees', payload.rowArray);
      case 'LOG_ATTENDANCE':
        return appendRowToSheet('Shifts_Attendance', payload.rowArray);
      case 'RECORD_REVENUE_COMMISSION':
        return appendRowToSheet('Sales_Commissions', payload.rowArray);
      case 'GENERATE_PAYSLIP':
        return appendRowToSheet('Payroll_Slips', payload.rowArray);
      case 'SUBMIT_LEAVE':
        return appendRowToSheet('Leaves_Coverage', payload.rowArray);
      case 'SUBMIT_SHIFT_SWAP':
        return appendRowToSheet('Shift_Swaps_ESS', payload.rowArray);
      case 'ADD_ATS_KPI':
        return appendRowToSheet('ATS_KPIs', payload.rowArray);
      case 'SAVE_WEEKLY_ROSTER':
        return appendRowToSheet('Weekly_Roster', payload.rowArray);
      case 'INITIATE_OFFBOARDING':
        return appendRowToSheet('Offboarding_Assets', payload.rowArray);
      default:
        return { ok: false, message: 'إجراء غير معروف: ' + action };
    }
  } catch (error) {
    return { ok: false, message: error.toString() };
  }
}

/**
 * ============================================================================
 * محرك الإنشاء والتنسيق التلقائي الشامل لجميع شيتات النظام
 * (Automatic Sheet Builder & RTL Visual Formatter)
 * ============================================================================
 */
function setupAndFormatDatabase(includeSeed) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(SYSTEM_CONFIG.TIMEZONE);
  ss.setSpreadsheetLocale('ar_EG');

  const sheetNames = Object.keys(SHEETS_SCHEMA);

  sheetNames.forEach(function(sheetName) {
    const schema = SHEETS_SCHEMA[sheetName];
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    // 1. ضبط اتجاه الشيت من اليمين إلى اليسار ولون التبويب
    sheet.setRightToLeft(true);
    sheet.setTabColor(schema.tabColor);

    // 2. كتابة رؤوس الأعمدة وتجميد الصف الأول
    const numCols = schema.headers.length;
    const headerRange = sheet.getRange(1, 1, 1, numCols);
    headerRange.setValues([schema.headers]);
    sheet.setFrozenRows(1);
    sheet.setRowHeight(1, 40);

    // 3. تنسيق ترويسة الجدول (Header Styling)
    headerRange
      .setBackground(SYSTEM_CONFIG.HEADER_BG)
      .setFontColor(SYSTEM_CONFIG.HEADER_FG)
      .setFontFamily('Cairo')
      .setFontSize(11)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setWrap(true);

    // 4. ضبط عرض الأعمدة تلقائياً حسب المخطط الهندسي
    schema.widths.forEach(function(width, idx) {
      sheet.setColumnWidth(idx + 1, width);
    });

    // 5. إزالة التناوب اللوني القديم وإضافة Banding احترافي
    const existingBandings = sheet.getBandings();
    existingBandings.forEach(function(b) { b.remove(); });
    const maxRows = Math.max(sheet.getLastRow(), 50);
    const dataRange = sheet.getRange(1, 1, maxRows, numCols);
    dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.TEAL, true, false);

    // 6. تنسيق خلايا البيانات والمحاذاة والخطوط
    const bodyRange = sheet.getRange(2, 1, maxRows - 1, numCols);
    bodyRange
      .setFontFamily('Cairo')
      .setFontSize(10)
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center');

    // 7. تطبيق تنسيق العملات والنسب المئوية
    schema.currencyCols.forEach(function(colIndex1Based) {
      sheet.getRange(2, colIndex1Based, maxRows - 1, 1).setNumberFormat(SYSTEM_CONFIG.CURRENCY_FORMAT);
    });
    schema.percentCols.forEach(function(colIndex1Based) {
      sheet.getRange(2, colIndex1Based, maxRows - 1, 1).setNumberFormat(SYSTEM_CONFIG.PERCENT_FORMAT);
    });

    // 8. إضافة قواعد التنسيق الشرطي للحالات (Conditional Formatting)
    applyStatusConditionalFormatting(sheet, numCols, maxRows);
  });

  if (includeSeed || ss.getSheetByName('Branches').getLastRow() <= 1) {
    seedInitialCoworkingData();
  }

  SpreadsheetApp.flush();
  return {
    ok: true,
    message: 'تم إنشاء وتنسيق جميع الجداول التسعة تلقائياً بنجاح بتصميم RTL احترافي!'
  };
}

/**
 * تطبيق قواعد تلوين تلقائي للحالات (نشط، معتمد، داخل النطاق، استرداد عمولة، إلخ)
 */
function applyStatusConditionalFormatting(sheet, numCols, maxRows) {
  sheet.clearConditionalFormatRules();
  const fullRange = sheet.getRange(2, 1, maxRows - 1, numCols);
  const rules = [];

  const positiveWords = ['ACTIVE', 'نشط', 'معتمد', 'INSIDE_GEOFENCE', 'داخل النطاق', 'تم الصرف', 'مكتمل', 'آمن'];
  positiveWords.forEach(function(word) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains(word)
        .setBackground('#dcfce7')
        .setFontColor('#166534')
        .setBold(true)
        .setRanges([fullRange])
        .build()
    );
  });

  const warningWords = ['PENDING', 'قيد المراجعة', 'PROBATION', 'تحت الاختبار', 'ROAMING', 'متنقل'];
  warningWords.forEach(function(word) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains(word)
        .setBackground('#fef3c7')
        .setFontColor('#92400e')
        .setBold(true)
        .setRanges([fullRange])
        .build()
    );
  });

  const dangerWords = ['CLAWBACK', 'استرداد', 'OUTSIDE_GEOFENCE', 'خارج النطاق', 'مرفوض', 'عجز تغطية'];
  dangerWords.forEach(function(word) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains(word)
        .setBackground('#fee2e2')
        .setFontColor('#991b1b')
        .setBold(true)
        .setRanges([fullRange])
        .build()
    );
  });

  sheet.setConditionalFormatRules(rules);
}

/**
 * إضافة صف جديد لأي شيت مع الحفاظ على ارتفاع الصف والتنسيق
 */
function appendRowToSheet(sheetName, rowArray) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    setupAndFormatDatabase(false);
    sheet = ss.getSheetByName(sheetName);
  }
  sheet.appendRow(rowArray);
  const lastRow = sheet.getLastRow();
  sheet.setRowHeight(lastRow, 32);
  return {
    ok: true,
    sheet: sheetName,
    rowIndex: lastRow,
    message: 'تم الحفظ في شيت ' + sheetName + ' بنجاح'
  };
}

/**
 * قراءة جميع البيانات من الشيتات التسعة وإرجاعها ككائن JSON متكامل للواجهة
 */
function getAllDatabaseState() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const data = {};
  Object.keys(SHEETS_SCHEMA).forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) {
      data[sheetName] = [];
    } else {
      const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, SHEETS_SCHEMA[sheetName].headers.length).getValues();
      data[sheetName] = values;
    }
  });
  return { ok: true, data: data };
}

/**
 * تغذية الشيتات ببيانات أولية واقعية لمساحة عمل مشتركة متعددة الفروع
 */
function seedInitialCoworkingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const bSheet = ss.getSheetByName('Branches');
  if (bSheet && bSheet.getLastRow() <= 1) {
    const branches = [
      ['BR-01', 'ALEX-KAFR', 'فرع كفر عبده - الإسكندرية', 'الإسكندرية', 31.2156, 29.9553, 120, 'نعم (24/7)', 2, 140, 91.5, 'سارة المنشاوي', 'نشط (ACTIVE)'],
      ['BR-02', 'ALEX-SMOUHA', 'فرع سموحة هب - الإسكندرية', 'الإسكندرية', 31.2112, 29.9441, 100, 'لا (8ص-11م)', 2, 95, 84.0, 'عمر عبد العزيز', 'نشط (ACTIVE)'],
      ['BR-03', 'CAI-NEWCAIRO', 'فرع التجمع الخامس - القاهرة', 'القاهرة', 30.0131, 31.4285, 150, 'نعم (24/7)', 3, 220, 88.0, 'أحمد الكيلاني', 'نشط (ACTIVE)']
    ];
    bSheet.getRange(2, 1, branches.length, branches[0].length).setValues(branches);
  }

  const eSheet = ss.getSheetByName('Employees');
  if (eSheet && eSheet.getLastRow() <= 1) {
    const employees = [
      ['EMP-101', 'CW-2026-01', 'سارة المنشاوي', '29504120201124', '01092233441', 'sara@coworkhub.eg', 'فرع كفر عبده - الإسكندرية', 'إدارة المجتمع (Community)', 'Community Manager', 'دوام كامل (FULL_TIME)', 16500, 0, 'نعم (Roaming)', '80% كفر عبده | 20% سموحة', 'نعم', '2024-03-01', 'نشط (ACTIVE)'],
      ['EMP-102', 'CW-2026-02', 'كريم طارق الجمل', '29708190204412', '01283344552', 'karim@coworkhub.eg', 'فرع كفر عبده - الإسكندرية', 'المبيعات والتأجير (Sales)', 'Sales & Leasing Executive', 'دوام كامل (FULL_TIME)', 12000, 0, 'نعم (Roaming)', '50% كفر عبده | 50% التجمع', 'نعم', '2024-06-15', 'نشط (ACTIVE)'],
      ['EMP-103', 'CW-2026-03', 'مريم وائل نصار', '30201150208831', '01145566773', 'mariam@coworkhub.eg', 'فرع كفر عبده - الإسكندرية', 'الاستقبال (Front Desk)', 'Front Desk Host', 'دوام جزئي بالساعة (PART_TIME)', 0, 85, 'لا', '100% كفر عبده', 'لا', '2025-02-10', 'نشط (ACTIVE)'],
      ['EMP-104', 'CW-2026-04', 'يوسف محمود زكي', '30109220203319', '01556677884', 'youssef@coworkhub.eg', 'فرع سموحة هب - الإسكندرية', 'الاستقبال (Front Desk)', 'Night Shift Receptionist', 'دوام كامل (FULL_TIME)', 9500, 0, 'نعم (Roaming)', '100% سموحة هب', 'نعم', '2025-01-01', 'نشط (ACTIVE)'],
      ['EMP-105', 'CW-2026-05', 'حسن إبراهيم الشاذلي', '29111040209911', '01019988775', 'hassan@coworkhub.eg', 'فرع التجمع الخامس - القاهرة', 'الصيانة والمرافق (Facilities)', 'Facilities & IT Specialist', 'عقد مرن بالوردية (FREELANCE)', 11000, 0, 'نعم (Roaming)', '40% كفر عبده | 60% التجمع', 'لا', '2024-09-01', 'نشط (ACTIVE)']
    ];
    eSheet.getRange(2, 1, employees.length, employees[0].length).setValues(employees);
  }
}
