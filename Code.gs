// Use the active spreadsheet to ensure we are reading/writing to the correct file
const ss = SpreadsheetApp.getActiveSpreadsheet();

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('HAPPY TUITIONS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDebugInfo() {
  try {
    const id = ss.getId();
    const name = ss.getName();
    const sheets = ss.getSheets().map(s => `${s.getName()} (Rows: ${s.getLastRow()})`).join(', ');
    return `<b>Connected File:</b> ${name} <br> <b>Spreadsheet ID:</b> ${id} <br> <b>Sheets Found:</b> ${sheets}`;
  } catch (e) {
    return "Error getting info: " + e.message;
  }
}

function authorizeScript() {
  // Run this function manually in the Apps Script editor to authorize permissions
  DriveApp.getRootFolder();
  SpreadsheetApp.getActiveSpreadsheet();
  console.log("Authorization complete. You can now use the app.");
}

function login(username, password) {
  try {
    const adminSheet = ss.getSheetByName('Admins');
    if (!adminSheet) return { success: false, message: "System Error: Admins DB not found." };
    
    const data = adminSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(username).trim() && 
          String(data[i][1]).trim() === String(password).trim() && 
          data[i][3] === 'active') {
        return { success: true, role: data[i][2] };
      }
    }
    return { success: false, message: "Invalid credentials" };
  } catch (e) {
    return { success: false, message: "Login Error: " + e.message };
  }
}

function changePassword(username, oldPassword, newPassword) {
  const adminSheet = ss.getSheetByName('Admins');
  const data = adminSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username && data[i][1] === oldPassword) {
      adminSheet.getRange(i + 1, 2).setValue(newPassword); // Update password column (index 2 in sheet, 1-based)
      return { success: true };
    }
  }
  return { success: false, message: "Invalid username or old password" };
}

/*
function addAdmin(username, password, role) {
  const adminSheet = ss.getSheetByName('Admins');
  adminSheet.appendRow([username, password, role, 'active']);
}

function getAdmins() {
  const adminSheet = ss.getSheetByName('Admins');
  return adminSheet.getDataRange().getValues().slice(1);
}

function updateAdmin(row, status) {
  const adminSheet = ss.getSheetByName('Admins');
  adminSheet.getRange(row + 1, 4).setValue(status); // Row is 0-indexed from data
}

function deleteAdmin(row) {
  const adminSheet = ss.getSheetByName('Admins');
  adminSheet.deleteRow(row + 2); // Headers + 1-based
}
*/


function logPaymentHistory(student) {
  let historySheet = ss.getSheetByName('PaymentHistory');
  if (!historySheet) {
    historySheet = ss.insertSheet('PaymentHistory');
    historySheet.appendRow(['AdmissionID', 'StudentName', 'Contact', 'Amount', 'PaymentDate', 'MonthYear', 'Timestamp']);
  }
  
  // Format Month-Year (e.g., Feb-2026)
  let monthYear = '';
  if (student.DateOfPayment) {
    const d = new Date(student.DateOfPayment);
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    monthYear = `${months[d.getMonth()]}-${d.getFullYear()}`;
  }

  historySheet.appendRow([
    student.AdmissionID,
    student.Name,
    student.ParentContact,
    student.PaidFee,
    student.DateOfPayment,
    monthYear,
    new Date()
  ]);
}

function addStudent(student) {
  const studentSheet = ss.getSheetByName('Students');
  const id = studentSheet.getLastRow(); // Auto ID starting from 1
  student.AdmissionID = id;
  studentSheet.appendRow([
    student.AdmissionID, student.DateOfJoining, student.Name, student.Gender,
    student.SchoolName, student.Class, student.ParentContact, '', // ParentName removed
    student.ActualFee, student.PaidFee, student.PaymentMode, student.DateOfPayment,
    '' // StudentImageURL removed
  ]);
  
  // Log Initial Payment
  logPaymentHistory(student);
}
function getPaymentHistoryDebug(admissionId, studentName) {
  try {
    const historySheet = ss.getSheetByName('PaymentHistory');
    
    if (!historySheet) {
      return [{ error: true, message: "System Error: PaymentHistory sheet not found." }];
    }
    
    const data = historySheet.getDataRange().getValues();
    if (data.length <= 1) {
       return [{ error: true, message: "No payment records found in system." }];
    }

    const searchId = String(admissionId).trim();
    const searchName = String(studentName).trim().toLowerCase();
    const rowsToCheck = data.slice(1);
    
    // Helper to safely serialize data
    const safeDate = (d) => (d instanceof Date) ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : d;
    const safeStr = (s) => (s === null || s === undefined) ? '' : String(s);

    // 1. Match by ID
    let history = rowsToCheck
      .filter(row => String(row[0]).trim() == searchId)
      .map(row => ({
        date: safeDate(row[4]),
        amount: safeStr(row[3]),
        monthYear: row[5], // Pass raw, formatted in final step
        id: safeStr(row[0]),
        studentName: safeStr(row[1])
      })).reverse();
    
    // 2. Match by Name
    if (history.length === 0 && searchName) {
      history = rowsToCheck
        .filter(row => String(row[1]).trim().toLowerCase().includes(searchName))
        .map(row => ({
          date: safeDate(row[4]),
          amount: safeStr(row[3]),
          monthYear: row[5], // Pass raw
          id: safeStr(row[0]),
          studentName: safeStr(row[1])
        })).reverse();
    }
    
    // 3. Fallback logic REMOVED as per user request.
    // Return empty array if no match found.
    
    // Formatting helper for Month-Year
    // If it's a date, format as MMM-yyyy. If string, keep as is.
    const formatMonthYear = (val) => {
      if (val instanceof Date) {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), 'MMM-yyyy');
      }
      return safeStr(val);
    };

    // Apply formatting to filtered history
    history = history.map(h => ({
      ...h,
      monthYear: formatMonthYear(h.monthYear) // Ensure Month-Year is clean
    }));

    return history; // Returns empty array [] if no match, which frontend handles.
  } catch (e) {
    return [{ error: true, message: "Server Error: " + e.toString() }];
  }
}


function getStudents(search = '', classFilter = '') {
  try {
    const studentSheet = ss.getSheetByName('Students');
    if (!studentSheet) throw new Error("Sheet 'Students' not found. Please create it or rename your tab.");
    if (studentSheet.getLastRow() <= 1) return []; // Only header exists
    let data = studentSheet.getDataRange().getValues().slice(1);
    
    // 1. Strict Filter by Class if provided
    if (classFilter && classFilter !== 'All') {
      data = data.filter(row => row[5] === classFilter);
    }

    // 2. Fuzzy Search if provided
    if (search) {
      search = search.toLowerCase();
      data = data.filter(row => 
        (row[2] != null && row[2].toString().toLowerCase().includes(search)) || // Name
        (row[7] != null && row[7].toString().toLowerCase().includes(search)) || // ParentName
        (row[6] != null && row[6].toString().includes(search)) // ParentContact
      );
    }
    // Convert Dates to Strings using Script TimeZone to avoid off-by-one errors
    return data.map(row => row.map(cell => (cell instanceof Date) ? Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd') : cell));
  } catch (e) {
    Logger.log("Error in getStudents: " + e.toString());
    throw new Error("Failed to fetch students. " + e.message);
  }
}

function updateStudent(id, student) {
  const studentSheet = ss.getSheetByName('Students');
  const data = studentSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      // Check if Payment Date changed
      const oldPaymentDate = data[i][11]; // Col 12 is Payment Date (index 11)
      const newPaymentDate = student.DateOfPayment;
      
      // Compare dates (as strings YYYY-MM-DD)
      const formattedOldDate = (oldPaymentDate instanceof Date) ? 
        Utilities.formatDate(oldPaymentDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : oldPaymentDate;
      
      if (formattedOldDate !== newPaymentDate) {
         // Log New Payment
         student.AdmissionID = id; // Ensure ID is set for history
         logPaymentHistory(student);
      }

      studentSheet.getRange(i + 1, 2, 1, 12).setValues([[
        student.DateOfJoining, student.Name, student.Gender,
        student.SchoolName, student.Class, student.ParentContact, '', // ParentName removed
        student.ActualFee, student.PaidFee, student.PaymentMode, student.DateOfPayment,
        '' // StudentImageURL removed
      ]]);
      break;
    }
  }
}

function deleteStudent(id) {
  const studentSheet = ss.getSheetByName('Students');
  const data = studentSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      studentSheet.deleteRow(i + 1);
      break;
    }
  }
}

function getStudentStats() {
  const students = getStudents();
  const total = students.length;
  const boys = students.filter(s => s[3] === 'Boy/Male').length;
  const girls = students.filter(s => s[3] === 'Girl/Female').length;
  const classCounts = {};
  
  let primaryCount = 0;
  let seniorCount = 0;
  
  const isSenior = (cls) => {
    const c = (cls || '').toString().toLowerCase().trim();
    if (!c) return false;
    const seniorKeywords = ['6', '7', '8', '9', '10', 'inter', 'senior', 'high school'];
    // If it contains any senior keywords
    if (seniorKeywords.some(k => c.includes(k))) return true;
    // Numeric check: if class number >= 6
    const classNum = parseInt(c.replace(/\D/g, ''));
    if (!isNaN(classNum) && classNum >= 6) return true;
    return false;
  };

  students.forEach(s => {
    const cls = (s[5] || '').toString().trim();
    if (cls) classCounts[cls] = (classCounts[cls] || 0) + 1;
    
    if (isSenior(cls)) {
      seniorCount++;
    } else {
      primaryCount++;
    }
  });
  return { total, boys, girls, classCounts, primaryCount, seniorCount };
}

function addTest(test) {
  const testSheet = ss.getSheetByName('WeeklyTests');
  const id = testSheet.getLastRow();
  test.ID = id;
  test.FinalMonthlyMarks = (parseFloat(test.Week1Marks) || 0) + (parseFloat(test.Week2Marks) || 0) +
    (parseFloat(test.Week3Marks) || 0) + (parseFloat(test.Week4Marks) || 0);
  testSheet.appendRow([
    test.ID, test.StudentName, test.Class, test.Week1Marks, test.Week2Marks,
    test.Week3Marks, test.Week4Marks, test.FinalMonthlyMarks, test.DateOfExam,
    test.ParentContact, test.Remarks
  ]);
}

function getTests(search = '') {
  try {
    const testSheet = ss.getSheetByName('WeeklyTests');
    if (!testSheet) throw new Error("Sheet 'WeeklyTests' not found. Please create it.");
    if (testSheet.getLastRow() <= 1) return [];
    let data = testSheet.getDataRange().getValues().slice(1);
    if (search) {
      search = search.toLowerCase();
      data = data.filter(row => 
        (row[1] != null && row[1].toString().toLowerCase().includes(search)) || // StudentName
        (row[9] != null && row[9].toString().includes(search)) // ParentContact
      );
    }
    // Convert Dates to Strings using Script TimeZone
    return data.map(row => row.map(cell => (cell instanceof Date) ? Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd') : cell));
  } catch (e) {
    throw new Error("Failed to fetch tests: " + e.message);
  }
}

function updateTest(id, test) {
  const testSheet = ss.getSheetByName('WeeklyTests');
  const data = testSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      test.FinalMonthlyMarks = (parseFloat(test.Week1Marks) || 0) + (parseFloat(test.Week2Marks) || 0) +
        (parseFloat(test.Week3Marks) || 0) + (parseFloat(test.Week4Marks) || 0);
      testSheet.getRange(i + 1, 2, 1, 10).setValues([[
        test.StudentName, test.Class, test.Week1Marks, test.Week2Marks,
        test.Week3Marks, test.Week4Marks, test.FinalMonthlyMarks, test.DateOfExam,
        test.ParentContact, test.Remarks
      ]]);
      break;
    }
  }
}

function deleteTest(id) {
  const testSheet = ss.getSheetByName('WeeklyTests');
  const data = testSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      testSheet.deleteRow(i + 1);
      break;
    }
  }
}

function getTestStats() {
  const tests = getTests();
  const totalStudents = new Set(tests.map(t => t[1])).size; // Unique students
  const week1 = tests.filter(t => t[3] !== '').length;
  const week2 = tests.filter(t => t[4] !== '').length;
  const week3 = tests.filter(t => t[5] !== '').length;
  const week4 = tests.filter(t => t[6] !== '').length;
  return { totalStudents, week1, week2, week3, week4 };
}

/*
function addFaculty(faculty) {
  const facultySheet = ss.getSheetByName('Faculty');
  const id = facultySheet.getLastRow();
  faculty.ID = id;
  facultySheet.appendRow([
    faculty.ID, faculty.Name, faculty.DateOfJoining, faculty.TeachingSubjects.join(','),
    faculty.MonthlySalary, faculty.Remarks
  ]);
}

function getFaculty(search = '') {
  try {
    const facultySheet = ss.getSheetByName('Faculty');
    if (!facultySheet) throw new Error("Sheet 'Faculty' not found. Please create it.");
    if (facultySheet.getLastRow() <= 1) return [];
    let data = facultySheet.getDataRange().getValues().slice(1);
    if (search) {
      search = search.toLowerCase();
      data = data.filter(row => row[1] != null && row[1].toString().toLowerCase().includes(search)); // Name
    }
    return data.map(row => {
      row[3] = (row[3] && typeof row[3] === 'string') ? row[3].split(',') : []; // Convert back to array safely
      // Convert remaining Dates using Script TimeZone
      return row.map(cell => (cell instanceof Date) ? Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd') : cell);
    });
  } catch (e) {
    throw new Error("Failed to fetch faculty: " + e.message);
  }
}

function updateFaculty(id, faculty) {
  const facultySheet = ss.getSheetByName('Faculty');
  const data = facultySheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      facultySheet.getRange(i + 1, 2, 1, 5).setValues([[
        faculty.Name, faculty.DateOfJoining, faculty.TeachingSubjects.join(','),
        faculty.MonthlySalary, faculty.Remarks
      ]]);
      break;
    }
  }
}

function deleteFaculty(id) {
  const facultySheet = ss.getSheetByName('Faculty');
  const data = facultySheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      facultySheet.deleteRow(i + 1);
      break;
    }
  }
}

function getFacultyStats() {
  const faculty = getFaculty();
  const total = faculty.length;
  const subjectCounts = {};
  faculty.forEach(f => {
    f[3].forEach(sub => {
      subjectCounts[sub] = (subjectCounts[sub] || 0) + 1;
    });
  });
  return { total, subjectCounts };
}
*/

function uploadImage(file) {
  const folder = DriveApp.getFolderById('11UtPwu8e7YcATDAh7xmttXoHV7khtzfZ'); // Create a folder in Drive and get ID
  const blob = Utilities.newBlob(Utilities.base64Decode(file.data), file.mimeType, file.name);
  const uploadedFile = folder.createFile(blob);
  uploadedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return "https://drive.google.com/uc?export=view&id=" + uploadedFile.getId();
}

    function exportToExcel(sheetName) {
      try {
        var sheet = ss.getSheetByName(sheetName);
        if (!sheet) throw new Error("Sheet not found: " + sheetName);
        var url = 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/export?format=xlsx&gid=' + sheet.getSheetId();
        return url;
      } catch (e) {
        throw new Error("Export failed: " + e.message);
      }
    }

    function exportToPDF(sheetName) {
      try {
        var sheet = ss.getSheetByName(sheetName);
        if (!sheet) throw new Error("Sheet not found: " + sheetName);
        var url = 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID + '/export?exportFormat=pdf&format=pdf&gid=' + sheet.getSheetId();
        return url;
      } catch (e) {
        throw new Error("Export failed: " + e.message);
      }
    }
// ------------------- ATTENDANCE FUNCTIONS -------------------

function markAttendance(date, records) {
  const sheet = ss.getSheetByName('Attendance');
  const today = date || new Date().toISOString().split('T')[0];
  
  // records = array of {admissionId, name, status, remarks}
  const rows = records.map(r => [
    today,
    r.admissionId,
    r.name,
    r.status,
    r.remarks || ''
  ]);
  
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
  }
}

function getAttendance(date) {
  const sheet = ss.getSheetByName('Attendance');
  let data = sheet.getDataRange().getValues().slice(1); // Skip header
  
  if (date) {
    data = data.filter(row => {
      const rowDate = row[0] instanceof Date ? row[0].toISOString().split('T')[0] : row[0];
      return rowDate === date;
    });
  }
  
  // Return formatted data including row index (using original index from full dataset is tricky after filter, 
  // so we'll just return data and handle updates by matching ID + Date)
  // Actually, for safer updates, let's return the 0-based index relative to the sheet's data range.
  const fullData = sheet.getDataRange().getValues();
  const result = [];
  
  for(let i = 1; i < fullData.length; i++) {
    const row = fullData[i];
    const rowDate = row[0] instanceof Date ? row[0].toISOString().split('T')[0] : row[0];
    if (!date || rowDate === date) {
      result.push({
        rowIndex: i + 1, // 1-based sheet row index
        date: rowDate,
        admissionId: row[1],
        name: row[2],
        status: row[3],
        remarks: row[4]
      });
    }
  }
  return result;
}

function updateAttendanceRecord(rowIndex, status, remarks) {
  const sheet = ss.getSheetByName('Attendance');
  // rowIndex is 1-based from getAttendance
  sheet.getRange(rowIndex, 4, 1, 2).setValues([[status, remarks]]);
}

function deleteAttendanceRecord(rowIndex) {
  const sheet = ss.getSheetByName('Attendance');
  sheet.deleteRow(rowIndex);
}

function getAttendanceSummary() {
  try {
    const sheet = ss.getSheetByName('Attendance');
    if (!sheet) throw new Error("Sheet 'Attendance' not found. Please create it.");
    if (sheet.getLastRow() <= 1) return { byDate: {}, byStudent: {} };
    const data = sheet.getDataRange().getValues().slice(1);
    
    const byDate = {};
    data.forEach(row => {
      const date = row[0] ? new Date(row[0]).toISOString().split('T')[0] : 'Unknown';
      if (!byDate[date]) byDate[date] = { present: 0, absent: 0, total: 0 };
      byDate[date].total++;
      if (row[3] === 'Present') byDate[date].present++;
      else byDate[date].absent++;
    });
    
    const studentAttendance = {};
    data.forEach(row => {
      const id = row[1];
      if (!id) return;
      if (!studentAttendance[id]) studentAttendance[id] = { name: row[2], present: 0, total: 0 };
      studentAttendance[id].total++;
      if (row[3] === 'Present') studentAttendance[id].present++;
    });
    
    return {
      byDate: byDate,
      byStudent: studentAttendance
    };
  } catch (e) {
    throw new Error("Failed to fetch attendance summary: " + e.message);
  }
}

function getStudentsForAttendance() {
  // Returns list of active students to mark attendance
  const students = ss.getSheetByName('Students').getDataRange().getValues().slice(1);
  return students.map(s => ({
    admissionId: s[0],
    name: s[2],
    class: s[5]
  }));
}

function getNextId(type) {
  try {
    let sheetName = '';
    if (type === 'student') sheetName = 'Students';
    else if (type === 'faculty') sheetName = 'Faculty';
    else if (type === 'test') sheetName = 'WeeklyTests';
    else if (type === 'event') sheetName = 'Events';
    
    if (!sheetName) return 'Auto Generated';
    
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return 1;
    // Assuming IDs correspond to row index or sequential. Use simple LastRow logic as placeholder.
    // If headers are there, ID starts from 1.
    // Logic depends on "add" function. AddStudent uses getLastRow() as ID. 
    // getLastRow returns total rows used. If header only (1 row), next row is 2. So ID used previously was header?
    // Wait, addStudent says: const id = studentSheet.getLastRow(); // Auto ID starting from 1
    // If sheet has 1 row (header), getLastRow is 1. Appends to row 2. ID becomes 1. Correct.
    return sheet.getLastRow(); 
  } catch (e) {
    return 'Error';
  }
}

// ------------------- NEW DASHBOARD FUNCTIONS -------------------

function addEvent(event) {
  let eventSheet = ss.getSheetByName('Events');
  if (!eventSheet) {
    eventSheet = ss.insertSheet('Events');
    eventSheet.appendRow(['ID', 'Date', 'EventName', 'AmountSpent']);
  }
  const id = eventSheet.getLastRow(); // Auto ID
  eventSheet.appendRow([
    id,
    event.date,
    event.eventName,
    event.amountSpent
  ]);
}

function getEvents(search = '') {
  try {
    const eventSheet = ss.getSheetByName('Events');
    if (!eventSheet) return []; // No events yet
    if (eventSheet.getLastRow() <= 1) return [];
    
    let data = eventSheet.getDataRange().getValues().slice(1);
    
    if (search) {
      search = search.toLowerCase();
      data = data.filter(row => row[2] && row[2].toString().toLowerCase().includes(search));
    }
    
    // Format dates using Script TimeZone
    return data.map(row => row.map(cell => (cell instanceof Date) ? Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd') : cell));
  } catch (e) {
    throw new Error("Failed to fetch events: " + e.message);
  }
}

function deleteEvent(id) {
  const eventSheet = ss.getSheetByName('Events');
  if (!eventSheet) return;
  const data = eventSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      eventSheet.deleteRow(i + 1);
      break;
    }
  }
}

function getEventStats() {
  const events = getEvents();
  const totalEvents = events.length;
  const totalExpenses = events.reduce((sum, e) => sum + (parseFloat(e[3]) || 0), 0);
  return { totalEvents, totalExpenses };
}

function getFeeDetails(search = '') {
  // Reuse existing getStudents but filter/map for specific table requirements
  const allStudents = getStudents(search);
  // Return format: Name, Class, Actual Fee, Paid Fee, Date of Payment
  return allStudents.map(s => ({
    name: s[2],
    class: s[5],
    actualFee: s[8],
    paidFee: s[9],
    dateOfPayment: s[11]
  }));
}

function getFeeStats() {
  const students = getStudents();
  const totalStudents = students.length;
  const totalFeeCollected = students.reduce((sum, s) => sum + (parseFloat(s[9]) || 0), 0);
  const totalActualFee = students.reduce((sum, s) => sum + (parseFloat(s[8]) || 0), 0);
  const totalDiscount = totalActualFee - totalFeeCollected; // As per requirement: diff between actual and paid
  
  // Calculate Payment Pending (Empty or 0 Paid Fee)
  const paymentPending = students.filter(s => !s[9] || s[9] == 0).length;

  // Calculate Total Expenses from Events
  const events = getEvents();
  const totalExpenses = events.reduce((sum, e) => sum + (parseFloat(e[3]) || 0), 0);
  
  return { totalStudents, totalFeeCollected, totalDiscount, totalExpenses, paymentPending };
}

function getRevenueData() {
  // Aggregate Monthly Data
  const fees = getStudents(); // Index 9 is PaidFee, Index 11 is DateOfPayment
  const expenses = getEvents(); // Index 1 is Date, Index 3 is AmountSpent
  
  const monthlyData = {};
  
  // Process Fees
  fees.forEach(s => {
    const dateStr = s[11]; // YYYY-MM-DD
    if (dateStr) {
      const monthKey = dateStr.substring(0, 7); // YYYY-MM
      if (!monthlyData[monthKey]) monthlyData[monthKey] = { fee: 0, expense: 0 };
      monthlyData[monthKey].fee += (parseFloat(s[9]) || 0);
    }
  });
  
  // Process Expenses
  expenses.forEach(e => {
    const dateStr = e[1]; // YYYY-MM-DD
    if (dateStr) {
      const monthKey = dateStr.substring(0, 7); // YYYY-MM
      if (!monthlyData[monthKey]) monthlyData[monthKey] = { fee: 0, expense: 0 };
      monthlyData[monthKey].expense += (parseFloat(e[3]) || 0);
    }
  });
  
  // Calculate Totals using logic similar to individual stats but globally
  const totalFeeCollected = Object.values(monthlyData).reduce((sum, m) => sum + m.fee, 0);
  const totalExpenses = Object.values(monthlyData).reduce((sum, m) => sum + m.expense, 0);
  
  return {
    totalFeeCollected,
    totalExpenses,
    monthlyData // Keyed by YYYY-MM
  };
}

function getMonthYearOptions() {
  try {
    const historySheet = ss.getSheetByName('PaymentHistory');
    if (!historySheet) return [];
    const data = historySheet.getDataRange().getValues().slice(1);
    
    // Helper to extract string representation of Month-Year
    const getMonthYearStr = (val) => {
      if (val instanceof Date) {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), 'MMM-yyyy');
      }
      return val ? String(val).trim() : '';
    };

    const months = [...new Set(data.map(row => getMonthYearStr(row[5])))].filter(m => m).sort((a, b) => {
      const partsA = a.split('-');
      const partsB = b.split('-');
      if (partsA.length < 2 || partsB.length < 2) return 0;
      
      const monthMap = { Jan: 0, Feb: 1, Mar: 2, Apr: 3, May: 4, Jun: 5, Jul: 6, Aug: 7, Sep: 8, Oct: 9, Nov: 10, Dec: 11 };
      
      const yearA = parseInt(partsA[1]);
      const yearB = parseInt(partsB[1]);
      
      if (yearA !== yearB) return yearB - yearA; // Year Desc
      return monthMap[partsB[0]] - monthMap[partsA[0]]; // Month Desc
    });
    return months;
  } catch (e) {
    console.error("Error in getMonthYearOptions:", e);
    return [];
  }
}

function getMonthWiseData(monthYear) {
  try {
    const historySheet = ss.getSheetByName('PaymentHistory');
    const studentSheet = ss.getSheetByName('Students');
    
    if (!historySheet || !studentSheet) return { error: true, message: "Sheets not found" };

    const historyData = historySheet.getDataRange().getValues().slice(1);
    const studentData = studentSheet.getDataRange().getValues().slice(1);
    
    const totalStudentsCount = studentData.length;
    
    const getMonthYearStr = (val) => {
      if (val instanceof Date) {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), 'MMM-yyyy');
      }
      return val ? String(val).trim() : '';
    };

    const filteredHistory = historyData.filter(row => getMonthYearStr(row[5]) === monthYear);
    
    const payments = filteredHistory.map(row => ({
      admissionId: row[0],
      studentName: row[1],
      contact: row[2],
      amount: row[3],
      paymentDate: (row[4] instanceof Date) ? Utilities.formatDate(row[4], Session.getScriptTimeZone(), 'yyyy-MM-dd') : row[4]
    }));

    const totalCollected = payments.reduce((sum, p) => sum + (parseFloat(p.amount) || 0), 0);
    const paidStudentsIds = [...new Set(payments.map(p => p.admissionId))];
    const totalStudentsPaid = paidStudentsIds.length;
    const pendingStudents = totalStudentsCount - totalStudentsPaid;

    return {
      payments: payments,
      stats: {
        totalCollected: totalCollected,
        totalStudentsPaid: totalStudentsPaid,
        pendingStudents: pendingStudents
      }
    };
  } catch (e) {
    return { error: true, message: e.toString() };
  }
}

function getPendingPaymentsData(monthYear) {
  try {
    const historySheet = ss.getSheetByName('PaymentHistory');
    const studentSheet = ss.getSheetByName('Students');
    
    if (!historySheet || !studentSheet) return { error: true, message: "Sheets not found" };

    const historyData = historySheet.getDataRange().getValues().slice(1);
    const studentData = studentSheet.getDataRange().getValues().slice(1);
    
    const getMonthYearStr = (val) => {
      if (val instanceof Date) {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), 'MMM-yyyy');
      }
      return val ? String(val).trim() : '';
    };

    // Find all admissions that PAID in this month
    const paidAdmissions = new Set(
      historyData
        .filter(row => getMonthYearStr(row[5]) === monthYear)
        .map(row => String(row[0]).trim())
    );

    // Filter students who have NOT paid
    const pendingPayments = studentData
      .filter(row => !paidAdmissions.has(String(row[0]).trim()))
      .map(row => ({
        admissionId: row[0],
        studentName: row[2],
        class: row[5],
        contact: row[6],
        actualFee: row[8]
      }));

    return {
      pending: pendingPayments,
      stats: {
        totalStudents: studentData.length,
        pendingCount: pendingPayments.length,
        paidCount: studentData.length - pendingPayments.length
      }
    };
  } catch (e) {
    return { error: true, message: e.toString() };
  }
}
