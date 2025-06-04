/**
 * هذا السكريبت يقوم بإنشاء الأوراق المطلوبة في جدول بيانات Google
 * وإضافة رؤوس الأعمدة لكل ورقة تلقائيًا.
 *
 * الخطوات:
 * 1. افتح جدول بيانات Google الذي تريد العمل عليه (أو قم بإنشاء واحد جديد).
 * 2. اذهب إلى Extensions > Apps Script لفتح محرر Apps Script.
 * 3. امسح أي كود موجود في ملف Code.gs.
 * 4. انسخ والصق هذا الكود كاملاً في ملف Code.gs.
 * 5. احفظ المشروع (Ctrl + S أو File > Save project).
 * 6. في شريط الأدوات العلوي في محرر Apps Script، اختر الدالة "setupSheetsAndHeaders" من القائمة المنسدلة.
 * 7. انقر على زر "Run" (تشغيل).
 * 8. سيطلب منك Google التصريح للسكريبت بالوصول إلى جداول البيانات الخاصة بك. وافق على ذلك.
 * 9. بعد انتهاء التشغيل، عد إلى جدول بيانات Google الخاص بك. ستجد الأوراق والرؤوس قد تم إنشاؤها.
 */

function setupSheetsAndHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // الحصول على جدول البيانات النشط

  // تعريف أسماء الأوراق ورؤوس الأعمدة لكل ورقة
  const sheetsConfig = {
    'Users': [
      'Username',
      'PasswordHash',
      'Role',
      'TeacherProfileId',
      'CreatedAt',
      'UpdatedAt'
    ],
    'Students': [
      'StudentID',
      'Name',
      'Age',
      'Phone',
      'Gender',
      'GuardianName',
      'GuardianPhone',
      'GuardianRelation',
      'TeacherID',
      'SubscriptionType',
      'Duration',
      'PaymentStatus',
      'PaymentAmount',
      'PaymentDate',
      'SessionsCompletedThisPeriod',
      'AbsencesThisPeriod',
      'IsRenewalNeeded',
      'LastRenewalDate',
      'ScheduledAppointments', // سيتم تخزينها كـ JSON String
      'IsArchived',
      'ArchivedReason',
      'ArchivedAt',
      'IsTrial',
      'TrialStatus',
      'TrialNotes',
      'CreatedAt',
      'UpdatedAt'
    ],
    'Teachers': [
      'TeacherID',
      'Name',
      'Age',
      'ContactNumber',
      'ZoomLink',
      'CurrentMonthSessions',
      'CurrentMonthAbsences',
      'CurrentMonthTrialSessions',
      'EstimatedMonthlyEarnings',
      'Specialization',
      'Bio',
      'Active',
      'HireDate',
      'Rating',
      'AvailableTimeSlots', // سيتم تخزينها كـ JSON String
      'LastPaymentDate',
      'CreatedAt',
      'UpdatedAt'
    ],
    'Sessions': [
      'SessionID',
      'StudentID',
      'TeacherID',
      'TeacherTimeSlotID',
      'Date',
      'TimeSlot',
      'DayOfWeek',
      'Status',
      'Report',
      'IsTrial',
      'CountsTowardsBalance',
      'CreatedAt',
      'UpdatedAt'
    ],
    'Transactions': [
      'TransactionID',
      'EntityType',
      'EntityID',
      'Amount',
      'Type',
      'Description',
      'Date',
      'Status',
      'RelatedSessionID',
      'CreatedAt',
      'UpdatedAt'
    ],
    'AccountingSummary': [
      'MonthYear',
      'TotalRevenue',
      'TotalExpenses',
      'TotalSalariesPaid',
      'CharityExpenses',
      'NetProfit',
      'CreatedAt',
      'UpdatedAt'
    ]
  };

  // حلقة لإنشاء أو إعادة تسمية الأوراق وإضافة الرؤوس
  for (const sheetName in sheetsConfig) {
    if (sheetsConfig.hasOwnProperty(sheetName)) {
      let sheet = ss.getSheetByName(sheetName);

      // إذا كانت الورقة موجودة، قم بمسح محتواها
      if (sheet) {
        sheet.clearContents();
        Logger.log('تم مسح محتوى ورقة موجودة: ' + sheetName);
      } else {
        // إذا لم تكن الورقة موجودة، قم بإنشائها
        sheet = ss.insertSheet(sheetName);
        Logger.log('تم إنشاء ورقة جديدة: ' + sheetName);
      }

      // **نقطة التحقق:** تأكد من أن 'sheet' كائن صالح قبل محاولة استخدام الدوال عليه
      if (sheet === null) { // إذا فشل insertSheet أو getSheetByName بطريقة ما
        Logger.log('خطأ: لم يتم الحصول على مرجع لورقة ' + sheetName + '. تخطي التحديث لهذه الورقة.');
        continue; // تخطي هذه الورقة والانتقال إلى التالية
      }

      // إضافة رؤوس الأعمدة
      const headers = sheetsConfig[sheetName];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // تجميد الصف الأول (صف الرؤوس) - هذا هو السطر الذي كان يسبب المشكلة
      try {
        sheet.freezeRows(1);
        Logger.log('تم تجميد الصف الأول في ورقة: ' + sheetName);
      } catch (e) {
        Logger.log('خطأ في تجميد الصفوف لورقة ' + sheetName + ': ' + e.message);
        // لا توقف السكريبت، فقط سجل الخطأ
      }
      
      Logger.log('تمت إضافة الرؤوس إلى ورقة: ' + sheetName);
    }
  }

  // (اختياري) حذف ورقة "Sheet1" الافتراضية إذا كانت لا تزال موجودة وفارغة
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && defaultSheet.getLastRow() === 0 && defaultSheet.getLastColumn() === 0) {
    // تحقق من أن الورقة فارغة تمامًا قبل الحذف
    ss.deleteSheet(defaultSheet);
    Logger.log('تم حذف ورقة "Sheet1" الافتراضية.');
  }

  SpreadsheetApp.getUi().alert('إعداد جدول البيانات اكتمل بنجاح!', 'تم إنشاء جميع الأوراق وإضافة رؤوس الأعمدة.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * دالة نقطة الدخول لتطبيق الويب.
 * تُقدم دائماً ملف index.html.
 * @param {GoogleAppsScript.Events.DoGet} e حدث doGet.
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet(e) {
  Logger.log("doGet received parameters (always loads index.html): " + JSON.stringify(e.parameter));
  const template = HtmlService.createTemplateFromFile('index');

  // تمرير معلمات URL إلى القالب، لتكون متاحة في JavaScript
  // هذه المعلمات ستُستخدم لتحديد الصفحة التي يجب تحميل محتواها.
  for (const key in e.parameter) {
    template[key] = e.parameter[key];
  }
  return template.evaluate().setTitle('أكاديمية غيث');
}




/**
 * تُرجع المحتوى HTML الخام لملف معين.
 * تُستخدم لتحميل أقسام الصفحات ديناميكياً.
 * @param {string} pageName اسم ملف HTML (بدون امتداد .html)
 * @returns {string} محتوى HTML للملف.
 */
function getHtmlContent(pageName) {
  Logger.log("Getting HTML content for: " + pageName);
  try {
    const template = HtmlService.createTemplateFromFile(pageName);
    // لا نُمرر هنا e.parameter، بل سنعتمد على localStorage أو طلبات منفصلة للبيانات
    return template.evaluate().getContent();
  } catch (error) {
    Logger.log("Error getting HTML content for " + pageName + ": " + error.message);
    return '<div><h1>خطأ في تحميل الصفحة</h1><p>عذرًا، حدث خطأ أثناء تحميل المحتوى.</p></div>';
  }
}


/**
 * دالة لتحديث حالة المتصفح (URL) باستخدام google.script.history.
 * لا تُرجع HtmlOutput.
 * @param {string} pageName اسم ملف HTML للصفحة المراد الانتقال إليها (للاستخدام في `google.script.history`).
 * @param {Object} [params={}] معلمات اختيارية لتمريرها مع الحالة.
 */
function redirectToPage(pageName, params = {}) {
  Logger.log("Redirecting via History API to: " + pageName + " with params: " + JSON.stringify(params));
  // لا تفعل شيئاً هنا، لأن google.script.history.replace سيتم استدعاؤها في الواجهة الأمامية
  // this function is just a placeholder to align client-side calls
  // The client-side will call google.script.history.replace() directly now
}



/**
 * دالة مساعدة لتضمين محتوى ملف HTML آخر (مثل CSS أو JavaScript) داخل ملف HTML الرئيسي.
 * تُستخدم بواسطة <اسم السكريبت>.html` (مثل `<?!= include('Style') ?>`)
 * @param {string} filename اسم الملف المراد تضمينه (بدون امتداد .html).
 * @returns {string} محتوى الملف.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

/**
 * دالة بسيطة لتوليد معرف فريد (UUIDv4)
 * يمكن استخدامها لـ StudentID, TeacherID, SessionID, TransactionID
 * @returns {string} معرف UUID فريد.
 */
function generateUuid() {
  return Utilities.getUuid();
}

/**
 * دالة عامة لجلب البيانات من ورقة معينة.
 * @param {string} sheetName اسم الورقة.
 * @returns {Array<Array<any>>} جميع بيانات الورقة (بما في ذلك الرأس).
 */
function getSheetData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('الورقة "' + sheetName + '" غير موجودة.');
  }
  return sheet.getDataRange().getValues();
}

/**
 * دالة عامة لكتابة صف جديد إلى ورقة معينة.
 * @param {string} sheetName اسم الورقة.
 * @param {Array<any>} rowData مصفوفة بالبيانات المراد إضافتها كصف.
 */
function appendRowToSheet(sheetName, rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('الورقة "' + sheetName + '" غير موجودة.');
  }
  sheet.appendRow(rowData);
}

/**
 * دالة عامة لتحديث صف موجود في ورقة.
 * تتطلب معرفًا فريدًا للبحث عن الصف.
 * @param {string} sheetName اسم الورقة.
 * @param {number} idColumnIndex فهرس العمود الذي يحتوي على المعرف الفريد (عادةً 0 لأول عمود).
 * @param {any} idValue قيمة المعرف للبحث عنها.
 * @param {Array<any>} newRowData مصفوفة بالبيانات الجديدة للصف.
 * @returns {boolean} True إذا تم التحديث، False إذا لم يتم العثور على الصف.
 */
function updateRowInSheet(sheetName, idColumnIndex, idValue, newRowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('الورقة "' + sheetName + '" غير موجودة.');
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { // ابدأ من 1 لتخطي الرأس
    if (data[i][idColumnIndex] === idValue) {
      sheet.getRange(i + 1, 1, 1, newRowData.length).setValues([newRowData]);
      return true;
    }
  }
  return false;
}

/**
 * دالة عامة لحذف صف موجود في ورقة.
 * @param {string} sheetName اسم الورقة.
 * @param {number} idColumnIndex فهرس العمود الذي يحتوي على المعرف الفريد.
 * @param {any} idValue قيمة المعرف للبحث عنها.
 * @returns {boolean} True إذا تم الحذف، False إذا لم يتم العثور على الصف.
 */
function deleteRowFromSheet(sheetName, idColumnIndex, idValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('الورقة "' + sheetName + '" غير موجودة.');
  }
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][idColumnIndex] === idValue) {
      sheet.deleteRow(i + 1); // +1 لأن الصفوف في Apps Script تبدأ من 1
      return true;
    }
  }
  return false;
}

/**
 * دالة بسيطة لتجزئة كلمة المرور (لأغراض العرض فقط، ليست آمنة مثل Bcrypt).
 * @param {string} password كلمة المرور.
 * @returns {string} تجزئة SHA-256.
 */
function hashPasswordSimple(password) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return Utilities.base64Encode(digest);
}

// ===========================================
//  الخدمات الأساسية (التي تتفاعل مع الواجهة الأمامية)
// ===========================================

// Auth Service (Simplified)
function processLogin(username, password) {
  Logger.log("processLogin function started for username: " + username);

  try {
    const users = getSheetData('Users');
    Logger.log("Users data fetched. Number of rows: " + users.length);

    if (users.length === 0) {
      Logger.log("Users sheet is empty or only has headers.");
      return { success: false, message: 'لا يوجد مستخدمون مسجلون.' };
    }

    const headers = users[0];
    const userData = users.slice(1);

    Logger.log("Headers: " + headers.join(', '));
    Logger.log("First user row: " + (userData.length > 0 ? userData[0].join(', ') : "N/A"));

    const foundUserRow = userData.find(row => {
      const userObj = {};
      headers.forEach((header, i) => userObj[header] = row[i]);

      const sheetUsername = String(userObj.Username).trim();
      const sheetPasswordHash = String(userObj.PasswordHash).trim();
      const inputUsername = String(username).trim();
      const inputPasswordHash = hashPasswordSimple(String(password).trim());

      Logger.log("Comparing: Input User='" + inputUsername + "', Sheet User='" + sheetUsername + "'");
      Logger.log("Comparing: Input Hash='" + inputPasswordHash + "', Sheet Hash='" + sheetPasswordHash + "'");

      return sheetUsername === inputUsername && sheetPasswordHash === inputPasswordHash;
    });

    if (foundUserRow) {
      const userObj = {};
      headers.forEach((header, i) => userObj[header] = foundUserRow[i]);
      Logger.log("Login successful for user: " + userObj.Username + ", Role: " + userObj.Role);

      // عند النجاح، نُرجع كائن JSON بسيط.
      // الواجهة الأمامية في index.html ستستخدم هذا لتخزين بيانات المستخدم
      // ثم تستدعي navigateTo لتغيير المحتوى.
      return {
        success: true,
        username: userObj.Username,
        role: userObj.Role,
        teacherProfileId: userObj.TeacherProfileId || null,
        message: 'تم تسجيل الدخول بنجاح!'
      };

    } else { // عند الفشل، نُرجع كائن JSON مع success: false
      Logger.log("Login failed: Username or password incorrect.");
      return { success: false, message: 'اسم المستخدم أو كلمة المرور غير صحيحة.' };
    }
  } catch (e) {
    Logger.log("Error in processLogin: " + e.message + " Stack: " + e.stack);
    return { success: false, message: 'حدث خطأ غير متوقع: ' + e.message };
  }
}

// Student Service (Simplified)
function getAllStudentsAndTeachers(filterOptions) {
  try {
    const students = getAllStudents(filterOptions.isArchived, filterOptions.teacherId);
    const teachers = getAllTeachers(); // جلب جميع المعلمين
    return { success: true, students: students, teachers: teachers };
  } catch (e) {
    Logger.log("Error in getAllStudentsAndTeachers: " + e.message);
    return { success: false, message: e.message };
  }
}

function getAllStudents(isArchivedFilter = 'false', teacherIdFilter = null) {
  let studentsData = getSheetData('Students');
  const headers = studentsData[0];
  let students = studentsData.slice(1); // تخطي الرأس

  let filteredStudents = students.map(row => {
    const student = {};
    headers.forEach((header, i) => {
      // التعامل مع الحقول المتداخلة مثل guardianDetails و paymentDetails
      if (header.includes('.')) {
        const parts = header.split('.');
        if (!student[parts[0]]) {
          student[parts[0]] = {};
        }
        student[parts[0]][parts[1]] = row[i];
      } else if (header === 'ScheduledAppointments' && row[i]) {
        try {
          student[header] = JSON.parse(row[i]);
        } catch (e) {
          Logger.log('Error parsing ScheduledAppointments for student ' + row[0] + ': ' + e.message);
          student[header] = []; // تعيين مصفوفة فارغة في حالة الخطأ
        }
      } else {
        student[header] = row[i];
      }
    });
    return student;
  });

  // تطبيق الفلاتر
  if (isArchivedFilter === 'true') {
    filteredStudents = filteredStudents.filter(s => s.IsArchived === true);
  } else if (isArchivedFilter === 'false') {
    filteredStudents = filteredStudents.filter(s => s.IsArchived === false);
  }
  // إذا كان 'all'، لا يوجد فلتر للأرشفة

  if (teacherIdFilter && teacherIdFilter !== 'all') {
    filteredStudents = filteredStudents.filter(s => s.TeacherID === teacherIdFilter);
  }

  return filteredStudents;
}

function getStudentAndTeachersData(studentId) {
  try {
    const student = getStudentById(studentId);
    const teachers = getAllTeachers();
    if (!student) {
      return { success: false, message: 'لم يتم العثور على الطالب.' };
    }
    return { success: true, student: student, teachers: teachers };
  } catch (e) {
    Logger.log("Error in getStudentAndTeachersData: " + e.message);
    return { success: false, message: e.message };
  }
}

function getStudentById(studentId) {
  const studentsData = getSheetData('Students');
  const headers = studentsData[0];
  const data = studentsData.slice(1);

  const studentRow = data.find(row => String(row[0]) === String(studentId)); // افترض أن StudentID في العمود الأول
  if (!studentRow) return null;

  const student = {};
  headers.forEach((header, i) => {
    if (header.includes('.')) {
      const parts = header.split('.');
      if (!student[parts[0]]) student[parts[0]] = {};
      student[parts[0]][parts[1]] = studentRow[i];
    } else if (header === 'ScheduledAppointments' && studentRow[i]) {
      try {
        student[header] = JSON.parse(studentRow[i]);
      } catch (e) {
        Logger.log('Error parsing ScheduledAppointments for student ' + studentRow[0] + ': ' + e.message);
        student[header] = [];
      }
    } else {
      student[header] = studentRow[i];
    }
  });
  return student;
}

function addStudent(studentData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  const headers = getSheetData('Students')[0];
  const newId = generateUuid();
  const now = new Date();

  // تحويل الكائن إلى صف بناءً على الرؤوس
  const newRow = headers.map(header => {
    if (header === 'StudentID') return newId;
    if (header === 'CreatedAt' || header === 'UpdatedAt') return now;
    if (header === 'IsArchived') return false; // طالب جديد ليس مؤرشفًا
    if (header === 'IsRenewalNeeded') return false; // طالب جديد لا يحتاج لتجديد
    if (header === 'IsTrial') return studentData.SubscriptionType === 'حلقة تجريبية';
    
    if (header.includes('.')) {
      const parts = header.split('.');
      return studentData[parts[0]] ? studentData[parts[0]][parts[1]] : null;
    }
    // التعامل مع المواعيد المجدولة كـ JSON string
    if (header === 'ScheduledAppointments' && studentData.ScheduledAppointments) {
      return JSON.stringify(studentData.ScheduledAppointments);
    }
    // التعامل مع 'PaymentDate' كـ Date Object
    if (header === 'PaymentDate' && studentData.PaymentDate) {
      return new Date(studentData.PaymentDate);
    }
    // للحقول الأخرى، استخدم القيمة المقدمة أو null
    return studentData[header] !== undefined ? studentData[header] : null;
  });

  appendRowToSheet('Students', newRow);

  // تحديث المعلم والمواعيد (نفس منطق Node.js، لكن باستخدام Sheets API)
  updateTeacherSlotsAndCreateSessions(studentData.TeacherID, newId, studentData.ScheduledAppointments, studentData.SubscriptionType === 'حلقة تجريبية');

  // إضافة حركة مالية (إذا لم تكن حلقة تجريبية ومبلغ الدفع > 0)
  if (studentData.SubscriptionType !== 'حلقة تجريبية' && studentData.PaymentAmount > 0) {
    addTransaction({
      EntityType: 'Student',
      EntityID: newId,
      Amount: studentData.PaymentAmount,
      Type: 'subscription_payment',
      Description: `دفعة تسجيل اشتراك جديد للطالب ${studentData.Name} (${studentData.SubscriptionType})`,
      Date: studentData.PaymentDate ? new Date(studentData.PaymentDate) : now,
      Status: studentData.PaymentStatus
    });
  }

  return { success: true, message: 'تم إضافة الطالب بنجاح!', studentId: newId };
}

function updateStudent(studentId, updatedData) {
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  const studentsData = studentsSheet.getDataRange().getValues();
  const headers = studentsData[0];
  const data = studentsData.slice(1);

  const rowIndex = data.findIndex(row => String(row[0]) === String(studentId)); // StudentID في العمود الأول
  if (rowIndex === -1) {
    throw new Error('لم يتم العثور على الطالب.');
  }

  const currentRow = data[rowIndex];
  const oldStudent = {};
  headers.forEach((header, i) => {
    if (header.includes('.')) {
      const parts = header.split('.');
      if (!oldStudent[parts[0]]) oldStudent[parts[0]] = {};
      oldStudent[parts[0]][parts[1]] = currentRow[i];
    } else if (header === 'ScheduledAppointments' && currentRow[i]) {
      try {
        oldStudent[header] = JSON.parse(currentRow[i]);
      } catch (e) {
        Logger.log('Error parsing ScheduledAppointments for old student ' + studentId + ': ' + e.message);
        oldStudent[header] = [];
      }
    } else {
      oldStudent[header] = currentRow[i];
    }
  });

  const oldTeacherId = oldStudent.TeacherID;
  const oldScheduledAppointments = oldStudent.ScheduledAppointments;
  const newTeacherId = updatedData.TeacherID;
  const newScheduledAppointments = updatedData.ScheduledAppointments || [];

  // 1. تحرير المواعيد القديمة من المعلم القديم وحذف الجلسات
  if (oldTeacherId) { // فقط إذا كان هناك معلم قديم
    releaseTeacherSlotsAndClearSessions(oldTeacherId, studentId, oldScheduledAppointments);
  }

  // تحديث بيانات الطالب
  const updatedRow = headers.map((header, i) => {
    if (header === 'StudentID') return studentId;
    if (header === 'CreatedAt') return currentRow[i]; // لا نغير تاريخ الإنشاء
    if (header === 'UpdatedAt') return new Date();
    if (header === 'IsArchived') return !!updatedData.IsArchived; // تأكد أنها boolean
    if (header === 'IsRenewalNeeded') return !!updatedData.IsRenewalNeeded; // تأكد أنها boolean
    if (header === 'IsTrial') return updatedData.SubscriptionType === 'حلقة تجريبية';
    
    if (header.includes('.')) {
      const parts = header.split('.');
      return updatedData[parts[0]] ? updatedData[parts[0]][parts[1]] : currentRow[i];
    }
    if (header === 'ScheduledAppointments') {
      return JSON.stringify(newScheduledAppointments);
    }
    if (header === 'PaymentDate' && updatedData.PaymentDate) {
      return new Date(updatedData.PaymentDate);
    }
    return updatedData[header] !== undefined ? updatedData[header] : currentRow[i];
  });

  studentsSheet.getRange(rowIndex + 2, 1, 1, updatedRow.length).setValues([updatedRow]); // +2 لتخطي الرأس


  // 2. حجز المواعيد الجديدة للمعلم الجديد وإنشاء Sessions لها
  if (newTeacherId && newScheduledAppointments.length > 0) {
    updateTeacherSlotsAndCreateSessions(newTeacherId, studentId, newScheduledAppointments, updatedData.SubscriptionType === 'حلقة تجريبية');
  } else if (newTeacherId && newScheduledAppointments.length === 0) {
    // إذا تم اختيار معلم ولكن بدون مواعيد (قد يكون حالة خاصة أو خطأ في الواجهة)
    // لا تفعل شيئاً سوى تحذير
    Logger.log(`المعلم ${newTeacherId} تم تحديده للطالب ${studentId} ولكن لا توجد مواعيد مجدولة.`);
  }


  // معالجة الأرشفة إذا تغيرت حالة الاشتراك إلى 'لم يشترك'
  if (updatedData.SubscriptionType === 'لم يشترك' && oldStudent.SubscriptionType !== 'لم يشترك') {
    archiveStudent(studentId, 'تم الأرشفة تلقائياً بسبب عدم الاشتراك بعد الفترة التجريبية.');
  }

  return { success: true, message: 'تم تحديث بيانات الطالب بنجاح!' };
}

function archiveStudent(studentId, reason) {
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  const studentsData = studentsSheet.getDataRange().getValues();
  const headers = studentsData[0];
  const data = studentsData.slice(1);

  const rowIndex = data.findIndex(row => String(row[0]) === String(studentId));
  if (rowIndex === -1) {
    throw new Error('الطالب غير موجود.');
  }

  const studentRow = studentsData[rowIndex + 1]; // +1 لأننا نعمل على البيانات الأصلية مع الرأس
  if (studentRow[headers.indexOf('IsArchived')] === true) {
    throw new Error('الطالب مؤرشف بالفعل.');
  }

  studentRow[headers.indexOf('IsArchived')] = true;
  studentRow[headers.indexOf('ArchivedAt')] = new Date();
  studentRow[headers.indexOf('ArchivedReason')] = reason;
  studentRow[headers.indexOf('SessionsCompletedThisPeriod')] = 0;
  studentRow[headers.indexOf('AbsencesThisPeriod')] = 0;
  studentRow[headers.indexOf('IsRenewalNeeded')] = false;
  studentRow[headers.indexOf('UpdatedAt')] = new Date();

  studentsSheet.getRange(rowIndex + 2, 1, 1, studentRow.length).setValues([studentRow]);

  // تحرير المواعيد الأسبوعية من المعلم عند الأرشفة
  const teacherId = studentRow[headers.indexOf('TeacherID')];
  let scheduledAppointments = [];
  try {
    scheduledAppointments = JSON.parse(studentRow[headers.indexOf('ScheduledAppointments')] || '[]');
  } catch(e) {
    Logger.log('Error parsing ScheduledAppointments during archive: ' + e.message);
  }
  
  releaseTeacherSlotsAndClearSessions(teacherId, studentId, scheduledAppointments);

  // مسح ارتباط المعلم والمواعيد المجدولة من الطالب بعد الأرشفة
  studentRow[headers.indexOf('TeacherID')] = null;
  studentRow[headers.indexOf('ScheduledAppointments')] = JSON.stringify([]); // مسح المواعيد
  studentsSheet.getRange(rowIndex + 2, 1, 1, studentRow.length).setValues([studentRow]); // حفظ مرة أخرى

  return { success: true, message: 'تم أرشفة الطالب بنجاح.' };
}

function unarchiveStudent(studentId) {
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  const studentsData = studentsSheet.getDataRange().getValues();
  const headers = studentsData[0];
  const data = studentsData.slice(1);

  const rowIndex = data.findIndex(row => String(row[0]) === String(studentId));
  if (rowIndex === -1) {
    throw new Error('الطالب غير موجود.');
  }

  const studentRow = studentsData[rowIndex + 1]; // +1 لأننا نعمل على البيانات الأصلية مع الرأس
  if (studentRow[headers.indexOf('IsArchived')] === false) {
    throw new Error('الطالب ليس مؤرشفًا بالفعل.');
  }

  studentRow[headers.indexOf('IsArchived')] = false;
  studentRow[headers.indexOf('ArchivedAt')] = null;
  studentRow[headers.indexOf('ArchivedReason')] = null;
  studentRow[headers.indexOf('UpdatedAt')] = new Date();

  studentsSheet.getRange(rowIndex + 2, 1, 1, studentRow.length).setValues([studentRow]);

  return { success: true, message: 'تمت إعادة تنشيط الطالب بنجاح!' };
}

// Teacher Service (Simplified)
function getAllTeachers() {
  const teachersData = getSheetData('Teachers');
  const headers = teachersData[0];
  const teachers = teachersData.slice(1);
  return teachers.map(row => {
    const teacher = {};
    headers.forEach((header, i) => {
      if (header === 'AvailableTimeSlots' && row[i]) {
        try {
          teacher[header] = JSON.parse(row[i]);
        } catch (e) {
          Logger.log('Error parsing AvailableTimeSlots for teacher ' + row[0] + ': ' + e.message);
          teacher[header] = [];
        }
      } else if (header.includes('.')) {
        const parts = header.split('.');
        if (!teacher[parts[0]]) teacher[parts[0]] = {};
        teacher[parts[0]][parts[1]] = row[i];
      } else {
        teacher[header] = row[i];
      }
    });
    return teacher;
  });
}

function getTeacherById(teacherId) {
  const teachersData = getSheetData('Teachers');
  const headers = teachersData[0];
  const data = teachersData.slice(1);

  const teacherRow = data.find(row => String(row[0]) === String(teacherId));
  if (!teacherRow) return null;

  const teacher = {};
  headers.forEach((header, i) => {
    if (header === 'AvailableTimeSlots' && teacherRow[i]) {
      try {
        teacher[header] = JSON.parse(teacherRow[i]);
      } catch (e) {
        Logger.log('Error parsing AvailableTimeSlots for teacher ' + teacherRow[0] + ': ' + e.message);
        teacher[header] = [];
      }
    } else if (header.includes('.')) {
      const parts = header.split('.');
      if (!teacher[parts[0]]) teacher[parts[0]] = {};
      teacher[parts[0]][parts[1]] = teacherRow[i];
    } else {
      teacher[header] = teacherRow[i];
    }
  });
  return teacher;
}

function getTeacherAvailableSlots(teacherId) {
  try {
    const teacher = getTeacherById(teacherId);
    if (!teacher) {
      return { success: false, message: 'المعلم غير موجود.' };
    }
    return { success: true, slots: teacher.AvailableTimeSlots || [] };
  } catch (e) {
    Logger.log("Error in getTeacherAvailableSlots: " + e.message);
    return { success: false, message: e.message };
  }
}

function addTeacher(teacherData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teachers');
  const headers = getSheetData('Teachers')[0];
  const newId = generateUuid();
  const now = new Date();

  const newRow = headers.map(header => {
    if (header === 'TeacherID') return newId;
    if (header === 'CreatedAt' || header === 'UpdatedAt' || header === 'HireDate') return now;
    if (header === 'Active') return true; // الافتراضي نشط
    if (header === 'AvailableTimeSlots') return JSON.stringify(teacherData.AvailableTimeSlots || []);
    if (header.includes('financialDetails.lastPaymentDate')) return null; // Default to null for new teachers
    return teacherData[header] !== undefined ? teacherData[header] : null;
  });

  appendRowToSheet('Teachers', newRow);
  return { success: true, message: 'تم إضافة المعلم بنجاح!', teacherId: newId };
}

function updateTeacher(teacherId, updatedData) {
  const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teachers');
  const teachersData = teachersSheet.getDataRange().getValues();
  const headers = teachersData[0];
  const data = teachersData.slice(1);

  const rowIndex = data.findIndex(row => String(row[0]) === String(teacherId));
  if (rowIndex === -1) {
    throw new Error('لم يتم العثور على المعلم.');
  }

  const currentRow = data[rowIndex];
  const oldTeacher = {};
  headers.forEach((header, i) => {
    if (header === 'AvailableTimeSlots' && currentRow[i]) {
      try {
        oldTeacher[header] = JSON.parse(currentRow[i]);
      } catch (e) {
        Logger.log('Error parsing AvailableTimeSlots for old teacher ' + teacherId + ': ' + e.message);
        oldTeacher[header] = [];
      }
    } else if (header.includes('.')) {
      const parts = header.split('.');
      if (!oldTeacher[parts[0]]) oldTeacher[parts[0]] = {};
      oldTeacher[parts[0]][parts[1]] = currentRow[i];
    } else {
      oldTeacher[header] = currentRow[i];
    }
  });

  const newAvailableTimeSlots = updatedData.AvailableTimeSlots || [];
  const updatedFinalSlots = [];

  // دمج المواعيد الجديدة مع الحفاظ على حالة المحجوزة للقديمة
  oldTeacher.AvailableTimeSlots.forEach(oldSlot => {
    const foundInNew = newAvailableTimeSlots.find(newSlot =>
      newSlot.dayOfWeek === oldSlot.dayOfWeek && newSlot.timeSlot === oldSlot.timeSlot
    );
    if (foundInNew) {
      // إذا كان الموعد موجودًا في القائمة الجديدة، خذ حالته المحجوزة من القديم
      updatedFinalSlots.push({ ...foundInNew, isBooked: oldSlot.isBooked, bookedBy: oldSlot.bookedBy });
    } else {
      // إذا كان الموعد القديم محجوزًا وغير موجود في الجديد، احتفظ به
      if (oldSlot.isBooked) {
        updatedFinalSlots.push(oldSlot);
      }
    }
  });
  // إضافة المواعيد الجديدة تمامًا التي لم تكن موجودة في القديم
  newAvailableTimeSlots.forEach(newSlot => {
    const foundInOld = oldTeacher.AvailableTimeSlots.some(oldSlot =>
      newSlot.dayOfWeek === oldSlot.dayOfWeek && newSlot.timeSlot === newSlot.timeSlot
    );
    if (!foundInOld) {
      updatedFinalSlots.push({ dayOfWeek: newSlot.dayOfWeek, timeSlot: newSlot.timeSlot, isBooked: false, bookedBy: null });
    }
  });


  // تحديث الصف
  const updatedRow = headers.map((header, i) => {
    if (header === 'TeacherID') return teacherId;
    if (header === 'CreatedAt' || header === 'HireDate') return currentRow[i];
    if (header === 'UpdatedAt') return new Date();
    if (header === 'AvailableTimeSlots') return JSON.stringify(updatedFinalSlots);
    if (header.includes('.')) { // للتعامل مع financialDetails
        const parts = header.split('.');
        return updatedData[parts[0]] ? updatedData[parts[0]][parts[1]] : currentRow[i];
    }
    return updatedData[header] !== undefined ? updatedData[header] : currentRow[i];
  });

  teachersSheet.getRange(rowIndex + 2, 1, 1, updatedRow.length).setValues([updatedRow]);
  return { success: true, message: 'تم تحديث بيانات المعلم بنجاح!' };
}

// Session Management (Simplified)
function getTeacherTodaySessions(teacherId, dayOfWeek) {
  const sessionsData = getSheetData('Sessions').slice(1);
  const sessionsHeaders = getSheetData('Sessions')[0];
  const studentsData = getSheetData('Students').slice(1);
  const studentsHeaders = getSheetData('Students')[0];

  const todaySessions = sessionsData.filter(row => String(row[sessionsHeaders.indexOf('TeacherID')]) === String(teacherId) && String(row[sessionsHeaders.indexOf('DayOfWeek')]) === String(dayOfWeek)); // teacherId in col C, DayOfWeek in col G

  return todaySessions.map(sessionRow => {
    const session = {};
    sessionsHeaders.forEach((header, i) => session[header] = sessionRow[i]);

    // جلب بيانات الطالب المرتبطة
    const studentRow = studentsData.find(sRow => String(sRow[studentsHeaders.indexOf('StudentID')]) === String(session.StudentID)); // StudentID in col A
    if (studentRow) {
      session.Student = {};
      studentsHeaders.forEach((h, i) => session.Student[h] = studentRow[i]);
    }
    return session;
  });
}

function updateSessionStatus(sessionId, status, report) {
  const sessionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sessions');
  const sessionsData = sessionsSheet.getDataRange().getValues();
  const sessionsHeaders = sessionsData[0];
  const data = sessionsData.slice(1);

  const rowIndex = data.findIndex(row => String(row[0]) === String(sessionId)); // SessionID in col A
  if (rowIndex === -1) {
    throw new Error('الجلسة غير موجودة.');
  }

  const currentRow = data[rowIndex];
  const oldSession = {};
  sessionsHeaders.forEach((h, i) => oldSession[h] = currentRow[i]);

  const studentId = oldSession.StudentID;
  const teacherId = oldSession.TeacherID;
  const oldStatus = oldSession.Status;
  const isTrial = oldSession.IsTrial;
  const sessionDate = new Date(oldSession.Date);

  // تحديث صف الجلسة
  const newRow = [...currentRow];
  newRow[sessionsHeaders.indexOf('Status')] = status;
  newRow[sessionsHeaders.indexOf('Report')] = (status === 'حضَر') ? report : '';
  newRow[sessionsHeaders.indexOf('UpdatedAt')] = new Date();
  newRow[sessionsHeaders.indexOf('CountsTowardsBalance')] = (status !== 'طلب تأجيل'); // تحديث حقل countsTowardsBalance

  sessionsSheet.getRange(rowIndex + 2, 1, 1, newRow.length).setValues([newRow]);

  // تحديث عدادات الطالب والمعلم
  updateStudentAndTeacherCounters(studentId, teacherId, oldStatus, status, isTrial, sessionDate);

  return { success: true, message: 'تم تحديث حالة الحصة بنجاح.' };
}

// Financial Management (Simplified)
function addTransaction(transactionData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  const headers = getSheetData('Transactions')[0];
  const newId = generateUuid();
  const now = new Date();

  const newRow = headers.map(header => {
    if (header === 'TransactionID') return newId;
    if (header === 'CreatedAt' || header === 'UpdatedAt') return now;
    if (header === 'Date') return new Date(transactionData.Date || now);
    return transactionData[header] !== undefined ? transactionData[header] : null;
  });

  appendRowToSheet('Transactions', newRow);

  // تحديث تفاصيل الدفع للطالب/المعلم إذا كانت حركة دفع
  if (transactionData.EntityType === 'Student' && transactionData.Type === 'subscription_payment') {
    const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
    const studentData = studentsSheet.getDataRange().getValues();
    const studentHeaders = studentData[0];
    const studentRowIndex = studentData.findIndex(row => String(row[0]) === String(transactionData.EntityID)); // StudentID in col A
    if (studentRowIndex !== -1) {
      const studentRow = studentData[studentRowIndex];
      studentRow[studentHeaders.indexOf('PaymentStatus')] = transactionData.Status;
      studentRow[studentHeaders.indexOf('PaymentAmount')] = transactionData.Amount;
      studentRow[studentHeaders.indexOf('PaymentDate')] = transactionData.Date;
      studentsSheet.getRange(studentRowIndex + 1, 1, 1, studentRow.length).setValues([studentRow]);
    }
  } else if (transactionData.EntityType === 'Teacher' && transactionData.Type === 'salary_payment') {
    const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teachers');
    const teacherData = teachersSheet.getDataRange().getValues();
    const teacherHeaders = teacherData[0];
    const teacherRowIndex = teacherData.findIndex(row => String(row[0]) === String(transactionData.EntityID)); // TeacherID in col A
    if (teacherRowIndex !== -1) {
      const teacherRow = teacherData[teacherRowIndex];
      teacherRow[teacherHeaders.indexOf('LastPaymentDate')] = transactionData.Date;
      teachersSheet.getRange(teacherRowIndex + 1, 1, 1, teacherRow.length).setValues([teacherRow]);
    }
  }

  return { success: true, message: 'تم إضافة الحركة المالية بنجاح!', transactionId: newId };
}

function getFinancialTransactions(entityType, entityId, type, startDate, endDate) {
  const transactionsData = getSheetData('Transactions');
  const headers = transactionsData[0];
  let transactions = transactionsData.slice(1);
  const students = getAllStudents(); // لجلب أسماء الطلاب والمعلمين
  const teachers = getAllTeachers();

  let filtered = transactions.filter(row => {
    const transaction = {};
    headers.forEach((h, i) => transaction[h] = row[i]);

    let match = true;
    if (entityType && entityType !== 'all' && transaction.EntityType !== entityType) match = false;
    if (entityId && String(transaction.EntityID) !== String(entityId)) match = false; // تحويل إلى String للمقارنة
    if (type && type !== 'all' && transaction.Type !== type) match = false;

    const transactionDate = new Date(transaction.Date);
    if (startDate && transactionDate < new Date(startDate)) match = false;
    if (endDate) {
      const end = new Date(endDate);
      end.setHours(23, 59, 59, 999);
      if (transactionDate > end) match = false;
    }
    return match;
  }).map(row => {
    const transaction = {};
    headers.forEach((h, i) => transaction[h] = row[i]);

    // إضافة اسم الكيان المرتبط (الطالب/المعلم)
    if (transaction.EntityType === 'Student') {
      const student = students.find(s => String(s.StudentID) === String(transaction.EntityID));
      transaction.EntityName = student ? student.Name : 'طالب محذوف';
    } else if (transaction.EntityType === 'Teacher') {
      const teacher = teachers.find(t => String(t.TeacherID) === String(transaction.EntityID));
      transaction.EntityName = teacher ? teacher.Name : 'معلم محذوف';
    } else if (transaction.EntityType === 'SystemExpense') {
      transaction.EntityName = 'مصروفات عامة';
    }
    return transaction;
  });

  // الفرز حسب التاريخ الأحدث
  filtered.sort((a, b) => new Date(b.Date).getTime() - new Date(a.Date).getTime());

  return filtered;
}


// Accounting Scheduler (Simplified)
function calculateAndSaveMonthlySummary(year, month) {
  const monthString = `${year}-${String(month).padStart(2, '0')}`;
  const startDate = new Date(year, month - 1, 1);
  const endDate = new Date(year, month, 0, 23, 59, 59);

  const transactionsData = getSheetData('Transactions');
  const headers = transactionsData[0];
  const transactions = transactionsData.slice(1);

  let totalRevenue = 0;
  let totalExpenses = 0;
  let totalSalariesPaid = 0;
  let charityExpenses = 0;

  transactions.forEach(row => {
    const transaction = {};
    headers.forEach((h, i) => transaction[h] = row[i]);

    const transactionDate = new Date(transaction.Date);
    if (transactionDate >= startDate && transactionDate <= endDate) {
      if (transaction.Type === 'subscription_payment' || transaction.Type === 'other_income') {
        totalRevenue += transaction.Amount;
      } else if (transaction.Type === 'salary_payment') {
        totalSalariesPaid += transaction.Amount;
      } else if (transaction.Type === 'charity_expense') {
        charityExpenses += transaction.Amount;
      } else if (['system_expense', 'advertisement_expense', 'other_expense'].includes(transaction.Type)) {
        totalExpenses += transaction.Amount;
      }
    }
  });

  const netProfit = totalRevenue - totalExpenses - totalSalariesPaid - charityExpenses;

  const accountingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AccountingSummary');
  const summaryData = accountingSheet.getDataRange().getValues();
  const summaryHeaders = summaryData[0];
  const existingRowIndex = summaryData.findIndex(row => String(row[0]) === String(monthString));

  const newSummaryRow = summaryHeaders.map(header => {
    switch (header) {
      case 'MonthYear': return monthString;
      case 'TotalRevenue': return totalRevenue;
      case 'TotalExpenses': return totalExpenses;
      case 'TotalSalariesPaid': return totalSalariesPaid;
      case 'CharityExpenses': return charityExpenses;
      case 'NetProfit': return netProfit;
      case 'CreatedAt': return existingRowIndex === -1 ? new Date() : summaryData[existingRowIndex][summaryHeaders.indexOf('CreatedAt')];
      case 'UpdatedAt': return new Date();
      default: return null;
    }
  });

  if (existingRowIndex === -1) {
    appendRowToSheet('AccountingSummary', newSummaryRow);
  } else {
    accountingSheet.getRange(existingRowIndex + 1, 1, 1, newSummaryRow.length).setValues([newSummaryRow]);
  }
}

// دالة لجلب الملخصات الشهرية (لصفحة التقارير المالية)
function getMonthlyFinancialSummary(year, month) {
  const monthString = `${year}-${String(month).padStart(2, '0')}`;
  const accountingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('AccountingSummary');
  if (!accountingSheet) throw new Error('ورقة ملخص الحسابات غير موجودة.');

  const data = accountingSheet.getDataRange().getValues();
  const headers = data[0];
  const summaryRow = data.slice(1).find(row => String(row[0]) === String(monthString));

  if (!summaryRow) return null; // لا توجد بيانات لهذا الشهر

  const summary = {};
  headers.forEach((header, i) => summary[header] = summaryRow[i]);
  return summary;
}


// ===========================================
//  المنطق الداخلي ومساعدات Cron Jobs
// ===========================================

// دالة مساعده: لإفراغ مواعيد المعلم وحذف الجلسات المرتبطة
function releaseTeacherSlotsAndClearSessions(teacherId, studentId, scheduledAppointments) {
  if (!teacherId || !Array.isArray(scheduledAppointments) || scheduledAppointments.length === 0) return;

  const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teachers');
  const teachersData = teachersSheet.getDataRange().getValues();
  const teacherHeaders = teachersData[0];
  const teacherRowIndex = teachersData.findIndex(row => String(row[0]) === String(teacherId));

  if (teacherRowIndex === -1) return; // المعلم غير موجود

  const teacherRow = teachersData[teacherRowIndex];
  let availableTimeSlots = [];
  try {
    availableTimeSlots = JSON.parse(teacherRow[teacherHeaders.indexOf('AvailableTimeSlots')] || '[]');
  } catch (e) {
    Logger.log('Error parsing AvailableTimeSlots for teacher ' + teacherId + ' in release: ' + e.message);
    availableTimeSlots = [];
  }

  // إفراغ الخانات المحجوزة لهذا الطالب
  scheduledAppointments.forEach(appt => {
    const slotIndex = availableTimeSlots.findIndex(slot =>
      String(slot.dayOfWeek) === String(appt.DayOfWeek) && // تأكد من المقارنة الصحيحة
      String(slot.timeSlot) === String(appt.TimeSlot) &&
      slot.isBooked === true &&
      String(slot.bookedBy) === String(studentId) // تأكد أنه محجوز بواسطة هذا الطالب
    );
    if (slotIndex !== -1) {
      availableTimeSlots[slotIndex].isBooked = false;
      availableTimeSlots[slotIndex].bookedBy = null;
    }
  });

  // تحديث مواعيد المعلم
  teacherRow[teacherHeaders.indexOf('AvailableTimeSlots')] = JSON.stringify(availableTimeSlots);
  teachersSheet.getRange(teacherRowIndex + 1, 1, 1, teacherRow.length).setValues([teacherRow]);

  // حذف الجلسات المجدولة (المستقبلية وغير المكتملة)
  const sessionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sessions');
  const sessionsData = sessionsSheet.getDataRange().getValues();
  const sessionsHeaders = sessionsData[0];
  const sessionsToDelete = [];

  for (let i = 1; i < sessionsData.length; i++) {
    const session = {};
    sessionsHeaders.forEach((h, idx) => session[h] = sessionsData[i][idx]);

    const isMatch = String(session.StudentID) === String(studentId) &&
                    String(session.TeacherID) === String(teacherId) &&
                    (String(session.Status) === 'مجدولة' || String(session.Status) === 'طلب تأجيل') &&
                    scheduledAppointments.some(appt => String(appt.DayOfWeek) === String(session.DayOfWeek) && String(appt.TimeSlot) === String(session.TimeSlot));

    if (isMatch) {
      sessionsToDelete.push(i + 1); // أرقام الصفوف للحذف (رقم الصف في الشيت + 1)
    }
  }

  // حذف الصفوف من الأسفل للأعلى لتجنب مشاكل الفهرسة
  sessionsToDelete.sort((a, b) => b - a).forEach(rowNum => {
    sessionsSheet.deleteRow(rowNum);
  });
}


// دالة مساعده: لحجز مواعيد المعلم وإنشاء الجلسات
function updateTeacherSlotsAndCreateSessions(teacherId, studentId, scheduledAppointments, isTrialSession = false) {
  if (!teacherId || !Array.isArray(scheduledAppointments) || scheduledAppointments.length === 0) return;

  const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teachers');
  const teachersData = teachersSheet.getDataRange().getValues();
  const teacherHeaders = teachersData[0];
  const teacherRowIndex = teachersData.findIndex(row => String(row[0]) === String(teacherId));

  if (teacherRowIndex === -1) throw new Error('المعلم الجديد غير موجود لجدولة المواعيد.');

  const teacherRow = teachersData[teacherRowIndex];
  let availableTimeSlots = [];
  try {
    availableTimeSlots = JSON.parse(teacherRow[teacherHeaders.indexOf('AvailableTimeSlots')] || '[]');
  } catch (e) {
    Logger.log('Error parsing AvailableTimeSlots for teacher ' + teacherId + ' in update: ' + e.message);
    availableTimeSlots = [];
  }

  const sessionsToCreate = [];
  scheduledAppointments.forEach(appt => {
    const slotIndex = availableTimeSlots.findIndex(slot =>
      String(slot.dayOfWeek) === String(appt.DayOfWeek) && String(slot.timeSlot) === String(appt.TimeSlot)
    );

    if (slotIndex !== -1) {
      const targetSlot = availableTimeSlots[slotIndex];

      // تحقق إذا كانت الخانة محجوزة بالفعل من قبل طالب آخر
      // (وإذا كان الطالب الحالي هو نفسه، فلا بأس، هذا يعني أنه يختار نفس الخانة القديمة)
      if (targetSlot.isBooked && String(targetSlot.bookedBy) !== String(studentId)) {
        throw new Error(`الخانة الزمنية ${appt.TimeSlot} في يوم ${appt.DayOfWeek} محجوزة بالفعل من قبل طالب آخر (${targetSlot.bookedBy}).`);
      }

      targetSlot.isBooked = true;
      targetSlot.bookedBy = studentId;

      // أضف الجلسة لإنشائها لاحقًا
      sessionsToCreate.push({
        StudentID: studentId,
        TeacherID: teacherId,
        TeacherTimeSlotID: `${appt.DayOfWeek}-${appt.TimeSlot}`, // معرف مؤقت للخانة
        Date: new Date(), // تاريخ اليوم الذي تم فيه الجدولة
        TimeSlot: appt.TimeSlot,
        DayOfWeek: appt.DayOfWeek,
        Status: 'مجدولة',
        Report: null,
        IsTrial: isTrialSession,
        CountsTowardsBalance: true,
        CreatedAt: new Date(),
        UpdatedAt: new Date()
      });
    } else {
      throw new Error(`الخانة الزمنية ${appt.TimeSlot} في يوم ${appt.DayOfWeek} غير متاحة للمعلم.`);
    }
  });

  // حفظ مواعيد المعلم المحدثة
  teacherRow[teacherHeaders.indexOf('AvailableTimeSlots')] = JSON.stringify(availableTimeSlots);
  teachersSheet.getRange(teacherRowIndex + 1, 1, 1, teacherRow.length).setValues([teacherRow]);

  // إنشاء الجلسات في ورقة Sessions
  const sessionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sessions');
  const sessionsHeaders = getSheetData('Sessions')[0];

  const existingSessionData = sessionsSheet.getDataRange().getValues(); // جلب جميع الجلسات الموجودة
  const existingSessionHeaders = existingSessionData[0];

  sessionsToCreate.forEach(sessionObj => {
    // تحقق مما إذا كانت الجلسة موجودة بالفعل (نفس الطالب والمعلم واليوم والوقت والحالة مجدولة)
    const sessionExists = existingSessionData.slice(1).some(existingRow => {
      const existingSession = {};
      existingSessionHeaders.forEach((h, idx) => existingSession[h] = existingRow[idx]);
      
      return String(existingSession.StudentID) === String(sessionObj.StudentID) &&
             String(existingSession.TeacherID) === String(sessionObj.TeacherID) &&
             String(existingSession.DayOfWeek) === String(sessionObj.DayOfWeek) &&
             String(existingSession.TimeSlot) === String(sessionObj.TimeSlot) &&
             String(existingSession.Status) === 'مجدولة';
    });

    if (!sessionExists) {
      const newSessionRow = sessionsHeaders.map(header => {
        if (header === 'SessionID') return generateUuid();
        return sessionObj[header] !== undefined ? sessionObj[header] : null;
      });
      sessionsSheet.appendRow(newSessionRow);
    } else {
      Logger.log(`الجلسة موجودة بالفعل للطالب ${sessionObj.StudentID} والمعلم ${sessionObj.TeacherID} في ${sessionObj.DayOfWeek} ${sessionObj.TimeSlot}. لم يتم الإنشاء.`);
    }
  });
}


// دالة مساعده: تحديث عدادات الطالب والمعلم بعد تغيير حالة الحصة
function updateStudentAndTeacherCounters(studentId, teacherId, oldStatus, newStatus, isTrial, sessionDate) {
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  const studentsData = studentsSheet.getDataRange().getValues();
  const studentHeaders = studentsData[0];
  const studentRowIndex = studentsData.findIndex(row => String(row[0]) === String(studentId)); // StudentID is col A

  const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teachers');
  const teachersData = teachersSheet.getDataRange().getValues();
  const teacherHeaders = teachersData[0];
  const teacherRowIndex = teachersData.findIndex(row => String(row[0]) === String(teacherId)); // TeacherID is col A

  if (studentRowIndex === -1 || teacherRowIndex === -1) return;

  // نسخ الصفوف لتجنب التعديل المباشر على البيانات المجلوبة قبل setValues
  const studentRow = [...studentsData[studentRowIndex]];
  const teacherRow = [...teachersData[teacherRowIndex]];

  // ==================== تحديث الطالب ====================
  let sessionsChange = 0;
  let absencesChange = 0;

  // حساب التغيير في الحصص المكتملة
  if (String(oldStatus) === 'حضَر' && String(newStatus) !== 'حضَر') {
      sessionsChange = -1;
  } else if (String(oldStatus) !== 'حضَر' && String(newStatus) === 'حضَر') {
      sessionsChange = 1;
  }

  // حساب التغيير في الغيابات
  if (String(oldStatus) === 'غاب' && String(newStatus) !== 'غاب') {
      absencesChange = -1;
  } else if (String(oldStatus) !== 'غاب' && String(newStatus) === 'غاب') {
      absencesChange = 1;
  }

  // تحديث عدادات الطالب
  const currentSessionsCompleted = studentRow[studentHeaders.indexOf('SessionsCompletedThisPeriod')] || 0;
  const currentAbsences = studentRow[studentHeaders.indexOf('AbsencesThisPeriod')] || 0;
  
  studentRow[studentHeaders.indexOf('SessionsCompletedThisPeriod')] = Math.max(0, currentSessionsCompleted + sessionsChange);
  studentRow[studentHeaders.indexOf('AbsencesThisPeriod')] = Math.max(0, currentAbsences + absencesChange);

  // تحديث `isRenewalNeeded` بناءً على نوع الاشتراك وعدد الحصص المكتملة
  const subscriptionType = studentRow[studentHeaders.indexOf('SubscriptionType')];
  const sessionsCompleted = studentRow[studentHeaders.indexOf('SessionsCompletedThisPeriod')];
  
  const SUBSCRIPTION_SLOTS_MAP = {
    'نصف ساعة / 4 حصص': 4,
    'نصف ساعة / 8 حصص': 8,
    'ساعة / 4 حصص': 8,
    'ساعة / 8 حصص': 16,
    'مخصص': 12, // افتراضي لـ 'مخصص'
    'حلقة تجريبية': 1,
    'أخرى': 0
  };
  const requiredSlots = SUBSCRIPTION_SLOTS_MAP[subscriptionType] || 0;
  
  if (requiredSlots > 0 && subscriptionType !== 'مخصص') {
    studentRow[studentHeaders.indexOf('IsRenewalNeeded')] = (sessionsCompleted >= requiredSlots);
  } else {
    studentRow[studentHeaders.indexOf('IsRenewalNeeded')] = false;
  }
  
  // يتم استخدام getRange(rowIndex + 1, ...) لأن getValues تعيد مصفوفة تبدأ من 0
  // ولكن setValues تحتاج إلى أرقام صفوف حقيقية في الشيت (تبدأ من 1)
  // والـ rowIndex الذي وجدناه هو فهرس المصفوفة (بعد تخطي الرأس).
  // لذلك، إذا كان rowIndex هو 0 (أول صف بيانات بعد الرأس)، فالصف في الشيت هو 2.
  studentsSheet.getRange(studentRowIndex + 2, 1, 1, studentRow.length).setValues([studentRow]);


  // ==================== تحديث المعلم ====================
  let teacherSessionsChange = 0;
  let teacherAbsencesChange = 0;
  let teacherTrialSessionsChange = 0;
  let earningsChange = 0;

  const getEarningValue = (status, isTrialSession) => {
      if (isTrialSession) return 0;
      if (String(status) === 'حضَر') return 20;
      if (String(status) === 'غاب') return 10;
      return 0;
  };

  const oldEarning = getEarningValue(oldStatus, isTrial);
  const newEarning = getEarningValue(newStatus, isTrial);
  earningsChange = newEarning - oldEarning;

  if (isTrial) {
      if (String(oldStatus) === 'حضَر' && String(newStatus) !== 'حضَر') {
          teacherTrialSessionsChange = -1;
      } else if (String(oldStatus) !== 'حضَر' && String(newStatus) === 'حضَر') {
          teacherTrialSessionsChange = 1;
      }
      const currentTrialSessions = teacherRow[teacherHeaders.indexOf('CurrentMonthTrialSessions')] || 0;
      teacherRow[teacherHeaders.indexOf('CurrentMonthTrialSessions')] = Math.max(0, currentTrialSessions + teacherTrialSessionsChange);
  } else {
      if (String(oldStatus) === 'حضَر' && String(newStatus) !== 'حضَر') {
          teacherSessionsChange = -1;
      } else if (String(oldStatus) !== 'حضَر' && String(newStatus) === 'حضَر') {
          teacherSessionsChange = 1;
      }
      if (String(oldStatus) === 'غاب' && String(newStatus) !== 'غاب') {
          teacherAbsencesChange = -1;
      } else if (String(oldStatus) !== 'غاب' && String(newStatus) === 'غاب') {
          teacherAbsencesChange = 1;
      }
      const currentMonthSessions = teacherRow[teacherHeaders.indexOf('CurrentMonthSessions')] || 0;
      const currentMonthAbsences = teacherRow[teacherHeaders.indexOf('CurrentMonthAbsences')] || 0;
      teacherRow[teacherHeaders.indexOf('CurrentMonthSessions')] = Math.max(0, currentMonthSessions + teacherSessionsChange);
      teacherRow[teacherHeaders.indexOf('CurrentMonthAbsences')] = Math.max(0, currentMonthAbsences + teacherAbsencesChange);
  }

  const currentEstimatedEarnings = teacherRow[teacherHeaders.indexOf('EstimatedMonthlyEarnings')] || 0;
  teacherRow[teacherHeaders.indexOf('EstimatedMonthlyEarnings')] = Math.max(0, currentEstimatedEarnings + earningsChange);

  teachersSheet.getRange(teacherRowIndex + 2, 1, 1, teacherRow.length).setValues([teacherRow]);
}

// ===========================================
//  CRON Jobs (Time-driven Triggers)
// ===========================================

/**
 * وظيفة لإعادة تعيين رصيد الحصص الشهري والتنبيه بالحاجة للتجديد.
 * تُجدول للتشغيل في بداية كل شهر.
 */
function resetMonthlySessionsCronJob() {
  Logger.log('بدء مهمة إعادة تعيين الحصص الشهرية.');
  const studentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
  const teachersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Teachers');

  // إعادة تعيين الطلاب
  const studentsData = studentsSheet.getDataRange().getValues();
  const studentHeaders = studentsData[0];
  const studentsToUpdate = studentsData.slice(1);
  for (let i = 0; i < studentsToUpdate.length; i++) {
    const row = studentsToUpdate[i];
    row[studentHeaders.indexOf('SessionsCompletedThisPeriod')] = 0;
    row[studentHeaders.indexOf('AbsencesThisPeriod')] = 0;
    row[studentHeaders.indexOf('IsRenewalNeeded')] = false; // تُعاد التقييم لاحقًا أو يدوياً
    row[studentHeaders.indexOf('UpdatedAt')] = new Date(); // تحديث وقت التحديث
  }
  // قم بتحديث النطاق الكامل للبيانات (تخطي الرأس)
  studentsSheet.getRange(2, 1, studentsToUpdate.length, studentsToUpdate[0].length).setValues(studentsToUpdate);


  // إعادة تعيين المعلمين
  const teachersData = teachersSheet.getDataRange().getValues();
  const teacherHeaders = teachersData[0];
  const teachersToUpdate = teachersData.slice(1);
  for (let i = 0; i < teachersToUpdate.length; i++) {
    const row = teachersToUpdate[i];
    row[teacherHeaders.indexOf('CurrentMonthSessions')] = 0;
    row[teacherHeaders.indexOf('CurrentMonthAbsences')] = 0;
    row[teacherHeaders.indexOf('CurrentMonthTrialSessions')] = 0;
    row[teacherHeaders.indexOf('EstimatedMonthlyEarnings')] = 0;
    row[teacherHeaders.indexOf('UpdatedAt')] = new Date(); // تحديث وقت التحديث
  }
  // قم بتحديث النطاق الكامل للبيانات (تخطي الرأس)
  teachersSheet.getRange(2, 1, teachersToUpdate.length, teachersToUpdate[0].length).setValues(teachersToUpdate);

  Logger.log('اكتملت مهمة إعادة تعيين الحصص الشهرية.');
}

/**
 * وظيفة لتجميع وتحديث الملخص المحاسبي الشهري.
 * تُجدول للتشغيل في بداية كل شهر (بعد إعادة تعيين الحصص).
 */
function updateMonthlyAccountingSummaryCronJob() {
  Logger.log('بدء مهمة تحديث الملخص المحاسبي الشهري.');
  const now = new Date();
  let targetMonth = now.getMonth(); // الشهر الحالي (0-11)
  let targetYear = now.getFullYear();

  // تجميع للشهر السابق
  let monthToSummarize = targetMonth === 0 ? 12 : targetMonth; // إذا كان يناير، نأخذ ديسمبر
  let yearToSummarize = targetMonth === 0 ? targetYear - 1 : targetYear;

  calculateAndSaveMonthlySummary(yearToSummarize, monthToSummarize);
  Logger.log('اكتملت مهمة تحديث الملخص المحاسبي الشهري.');
}

/**
 * دالة لإنشاء المشغلات المستندة إلى الوقت.
 * تُشغل يدوياً مرة واحدة لإعداد المشغلات.
 */
function createTimeDrivenTriggers() {
  // حذف أي مشغلات موجودة لتجنب التكرار
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'resetMonthlySessionsCronJob' ||
        trigger.getHandlerFunction() === 'updateMonthlyAccountingSummaryCronJob') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // إنشاء مشغل لإعادة تعيين الحصص في اليوم الأول من كل شهر، الساعة 1 صباحاً (مثلاً)
  ScriptApp.newTrigger('resetMonthlySessionsCronJob')
      .timeBased()
      .atMonthStart()
      .atHour(1)
      .create();

  // إنشاء مشغل لتحديث الملخص المالي في اليوم الأول من كل شهر، الساعة 1:30 صباحاً (بعد إعادة التعيين)
  ScriptApp.newTrigger('updateMonthlyAccountingSummaryCronJob')
      .timeBased()
      .atMonthStart()
      .atHour(1)
      .nearMinute(30)
      .create();

  Logger.log('تم إنشاء المشغلات المستندة إلى الوقت بنجاح.');
}

// ===========================================
//  وظائف مساعدة للواجهة الأمامية (timeHelpers.js)
// ===========================================

// يجب نسخ هذه الدوال إلى ملف JS منفصل يتم تضمينه في HTML (مثلاً: `JsHelpers.html`)
// أو تضمينها مباشرة في كل ملف HTML يستخدمها.
// هنا سأضعها في Apps Script لسهولة العرض، لكن يفضل أن تكون في ملف JS للواجهة الأمامية.

function formatTime12Hour(time24hrPart) {
  if (typeof time24hrPart !== 'string' || !time24hrPart.includes(':')) {
    return 'وقت غير صالح';
  }
  const [hours, minutes] = time24hrPart.split(':').map(Number);
  if (isNaN(hours) || isNaN(minutes)) {
    return 'وقت غير صالح';
  }
  const ampm = hours >= 12 ? 'م' : 'ص';
  const formattedHours = hours % 12 || 12;
  const formattedMinutes = minutes < 10 ? `0${minutes}` : minutes;
  return `${formattedHours}:${formattedMinutes} ${ampm}`;
}

function getTimeInMinutes(timeString) {
  if (typeof timeString !== 'string' || !timeString.includes(':')) {
    return -1;
  }
  const timePart = timeString.split(' - ')[0];
  const [hours, minutes] = timePart.split(':').map(Number);
  if (isNaN(hours) || isNaN(minutes)) {
    return -1;
  }
  return hours * 60 + minutes;
}




function verifyHashedPassword() {
  const passwordToHash = "1234"; // تأكد أن هذه هي الكلمة السرية التي تحاول تسجيل الدخول بها
  const expectedHashedPassword = hashPasswordSimple(passwordToHash);
  Logger.log("Expected Hash for 'adminpassword': " + expectedHashedPassword);
}







