/**
 * ระบบเช็คชื่อนักเรียน - Google Apps Script Backend
 * เชื่อมต่อกับ Google Sheets เพื่อจัดเก็บข้อมูลการเข้าเรียน
 */

// ใช้ Spreadsheet ที่ผูกกับ Script นี้
const STUDENT_SHEET_NAME = 'Students';
const ATTENDANCE_SHEET_NAME = 'Attendance';

/**
 * ฟังก์ชันเริ่มต้นสำหรับสร้าง Sheets หากยังไม่มี
 */
function initializeSheets() {
  try {
    // ใช้ Spreadsheet ที่ผูกกับ Script นี้
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // สร้างหรือตรวจสอบ Sheet สำหรับข้อมูลนักเรียน
    let studentSheet = spreadsheet.getSheetByName(STUDENT_SHEET_NAME);
    if (!studentSheet) {
      studentSheet = spreadsheet.insertSheet(STUDENT_SHEET_NAME);
    }
    
    // ตรวจสอบว่า Sheet ว่างเปล่าหรือไม่มี header
    const studentLastRow = studentSheet.getLastRow();
    if (studentLastRow === 0 || studentSheet.getRange(1, 1).getValue() === '') {
      // สร้างหัวตาราง
      studentSheet.getRange(1, 1, 1, 4).setValues([
        ['รหัสนักเรียน', 'ชื่อ', 'นามสกุล', 'ห้องเรียน']
      ]);
      
      // จัดรูปแบบ header
      const headerRange = studentSheet.getRange(1, 1, 1, 4);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#E1BAFF'); // สีพาสเทลม่วง
      headerRange.setHorizontalAlignment('center');
      
      // ใส่ข้อมูลตัวอย่างถ้ายังไม่มีข้อมูล
      if (studentLastRow === 0) {
        studentSheet.getRange(2, 1, 5, 4).setValues([
          ['20001', 'สมชาย', 'ใจดี', 'ม.1/1'],
          ['20002', 'สมหญิง', 'ใจงาม', 'ม.1/1'],
          ['20003', 'ประเสริฐ', 'เก่งเก้า', 'ม.1/2'],
          ['20004', 'วรรณา', 'สวยงาม', 'ม.1/2'],
          ['20005', 'ชัยวัฒน์', 'รุ่งเรือง', 'ม.1/3']
        ]);
      }
    }
    
    // สร้างหรือตรวจสอบ Sheet สำหรับข้อมูลการเข้าเรียน
    let attendanceSheet = spreadsheet.getSheetByName(ATTENDANCE_SHEET_NAME);
    if (!attendanceSheet) {
      attendanceSheet = spreadsheet.insertSheet(ATTENDANCE_SHEET_NAME);
    }
    
    // ตรวจสอบว่า Sheet ว่างเปล่าหรือไม่มี header
    const attendanceLastRow = attendanceSheet.getLastRow();
    if (attendanceLastRow === 0 || attendanceSheet.getRange(1, 1).getValue() === '') {
      // สร้างหัวตาราง
      attendanceSheet.getRange(1, 1, 1, 6).setValues([
        ['วันที่', 'เวลา', 'รหัสนักเรียน', 'ชื่อ-นามสกุล', 'ห้องเรียน', 'สถานะ']
      ]);
      
      // จัดรูปแบบ header
      const headerRange = attendanceSheet.getRange(1, 1, 1, 6);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#BAFFC9'); // สีพาสเทลเขียว
      headerRange.setHorizontalAlignment('center');
    }
    
    return {
      success: true,
      spreadsheetId: spreadsheet.getId(),
      message: 'เตรียม Sheets เรียบร้อยแล้ว'
    };
    
  } catch (error) {
    console.error('Error initializing sheets:', error);
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการเตรียม Sheets: ' + error.message
    };
  }
}

/**
 * ฟังก์ชันดึงรายชื่อนักเรียนทั้งหมด
 */
function getAllStudents() {
  try {
    console.log('Getting all students...');
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const studentSheet = spreadsheet.getSheetByName(STUDENT_SHEET_NAME);
    
    if (!studentSheet) {
      console.error('Student sheet not found');
      return { success: false, message: 'ไม่พบ Sheet ข้อมูลนักเรียน' };
    }
    
    const data = studentSheet.getDataRange().getValues();
    const students = [];
    
    console.log('Total rows in student sheet:', data.length);
    
    // ข้าม header row (index 0)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) { // ตรวจสอบว่ามีรหัสนักเรียน
        students.push({
          id: data[i][0].toString(),
          firstName: data[i][1],
          lastName: data[i][2],
          classroom: data[i][3],
          fullName: data[i][1] + ' ' + data[i][2]
        });
      }
    }
    
    console.log('Found students:', students.length);
    
    return {
      success: true,
      students: students,
      total: students.length
    };
    
  } catch (error) {
    console.error('Error getting students:', error);
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการดึงข้อมูลนักเรียน: ' + error.message
    };
  }
}

/**
 * ฟังก์ชันค้นหานักเรียนตามรหัส
 */
function findStudentById(studentId) {
  try {
    const result = getAllStudents();
    if (!result.success) {
      return result;
    }
    
    const student = result.students.find(s => s.id === studentId.toString());
    
    if (student) {
      return {
        success: true,
        student: student
      };
    } else {
      return {
        success: false,
        message: 'ไม่พบนักเรียนรหัส ' + studentId
      };
    }
    
  } catch (error) {
    console.error('Error finding student:', error);
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการค้นหานักเรียน: ' + error.message
    };
  }
}

/**
 * ฟังก์ชันบันทึกการเข้าเรียน
 */
function recordAttendance(studentId, studentFullName, studentClassroom, status = 'เข้าเรียน') {
  try {
    console.log('Recording attendance for student:', studentId, studentFullName, studentClassroom, status);

    const now = new Date();
    const dateStr = Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/yyyy');
    const timeStr = Utilities.formatDate(now, 'Asia/Bangkok', 'HH:mm:ss');

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const attendanceSheet = spreadsheet.getSheetByName(ATTENDANCE_SHEET_NAME);

    if (!attendanceSheet) {
      return { success: false, message: 'ไม่พบ Sheet การเข้าเรียน' };
    }

    attendanceSheet.appendRow([
      dateStr,
      timeStr,
      studentId,
      studentFullName, // Use passed name
      studentClassroom, // Use passed classroom
      status
    ]);

    console.log('Attendance recorded successfully for:', studentFullName);

    return {
      success: true,
      message: 'บันทึกการเข้าเรียนเรียบร้อยแล้ว',
      student: { id: studentId, fullName: studentFullName, classroom: studentClassroom },
      timestamp: dateStr + ' ' + timeStr,
      status: status
    };

  } catch (error) {
    console.error('Error recording attendance:', error, error.stack);
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการบันทึกการเข้าเรียน: ' + error.message
    };
  }
}

/**
 * ฟังก์ชันดึงข้อมูลการเข้าเรียนวันนี้
 */
function getTodayAttendance() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const attendanceSheet = spreadsheet.getSheetByName(ATTENDANCE_SHEET_NAME);
    const today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yyyy');

    if (!attendanceSheet) {
      console.error('Attendance sheet not found:', ATTENDANCE_SHEET_NAME);
      return { success: false, message: 'ไม่พบ Sheet การเข้าเรียน (' + ATTENDANCE_SHEET_NAME + ')' };
    }

    const lastRow = attendanceSheet.getLastRow();
    const todayAttendance = [];

    if (lastRow > 1) { // Only process if there are data rows beyond the header
      const dataRange = attendanceSheet.getRange(2, 1, lastRow - 1, 6); // Start from row 2, get 6 columns
      const data = dataRange.getValues();
      console.log('Looking for attendance on:', today, 'in', data.length, 'data rows.');

      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (row[0] && row[0].toString() === today) {
          todayAttendance.push({
            date: row[0].toString(),
            time: row[1] ? row[1].toString() : '',
            studentId: row[2] ? row[2].toString() : '',
            studentName: row[3] ? row[3].toString() : '',
            classroom: row[4] ? row[4].toString() : '',
            status: row[5] ? row[5].toString() : ''
          });
        }
      }

      if (todayAttendance.length > 0) {
        todayAttendance.sort((a, b) => {
          try {
            const timeA = typeof a.time === 'string' && a.time.match(/\\d{2}:\\d{2}:\\d{2}/) ? a.time : '00:00:00';
            const timeB = typeof b.time === 'string' && b.time.match(/\\d{2}:\\d{2}:\\d{2}/) ? b.time : '00:00:00';
            return new Date('1970/01/01 ' + timeB) - new Date('1970/01/01 ' + timeA);
          } catch (sortError) {
            console.error('Error sorting attendance times:', sortError, 'a.time:', a.time, 'b.time:', b.time);
            return 0;
          }
        });
      }
    } else {
      console.log('Attendance sheet has no data rows (or only headers). Last row:', lastRow);
    }
    
    console.log('Found attendance records for today:', todayAttendance.length);
    return {
      success: true,
      attendance: todayAttendance,
      total: todayAttendance.length,
      date: today
    };

  } catch (error) {
    console.error('Error in getTodayAttendance:', error, error.stack);
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการดึงข้อมูลการเข้าเรียนวันนี้: ' + error.message
    };
  }
}

/**
 * ฟังก์ชันสำหรับหน้าเว็บ
 */
function doGet(e) {
  try {
    // Initialize sheets regardless of the page being served
    const initResult = initializeSheets();
    console.log('Initialize sheets result:', initResult);
  } catch (error) {
    console.error('Error initializing sheets in doGet:', error);
  }

  let page = e.parameter.page;
  if (page === 'dashboard') {
    return HtmlService.createTemplateFromFile('dashboard')
      .evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle('Dashboard - ระบบเช็คชื่อนักเรียน');
  } else {
    // Default to index page (attendance)
    return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setTitle('ระบบเช็คชื่อนักเรียน');
  }
}

/**
 * ฟังก์ชันรวม CSS/JS files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * ฟังก์ชันดึงรายชื่อห้องเรียนทั้งหมด
 */
function getAllClassrooms() {
  try {
    console.log('Getting all classrooms...');
    const result = getAllStudents();
    if (!result.success) {
      return result;
    }
    
    // สร้างรายการห้องเรียนที่ไม่ซ้ำ
    const classrooms = [...new Set(result.students.map(student => student.classroom))].sort();
    
    console.log('Found classrooms:', classrooms);
    
    return {
      success: true,
      classrooms: classrooms,
      total: classrooms.length
    };
    
  } catch (error) {
    console.error('Error getting classrooms:', error);
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการดึงข้อมูลห้องเรียน: ' + error.message
    };
  }
}

/**
 * ฟังก์ชันดึงรายชื่อนักเรียนตามห้องเรียน
 */
function getStudentsByClassroom(classroom) {
  try {
    console.log('Getting students for classroom:', classroom);
    const result = getAllStudents();
    if (!result.success) {
      return result;
    }
    
    // กรองนักเรียนตามห้องเรียน
    const studentsInClass = result.students.filter(student => student.classroom === classroom);
    
    console.log('Found students in classroom:', studentsInClass.length);
    
    return {
      success: true,
      students: studentsInClass,
      classroom: classroom,
      total: studentsInClass.length
    };
    
  } catch (error) {
    console.error('Error getting students by classroom:', error);
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการดึงข้อมูลนักเรียนในห้อง: ' + error.message
    };
  }
}

/**
 * ฟังก์ชันดึงสถิติการเข้าเรียนแบบละเอียด
 */
function getAttendanceStats(dateFrom = null, dateTo = null) {
  try {
    console.log('Getting attendance stats...');
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const attendanceSheet = spreadsheet.getSheetByName(ATTENDANCE_SHEET_NAME);
    
    if (!attendanceSheet) {
      return { success: false, message: 'ไม่พบ Sheet การเข้าเรียน' };
    }
    
    const data = attendanceSheet.getDataRange().getValues();
    const studentsResult = getAllStudents();
    
    if (!studentsResult.success) {
      return studentsResult;
    }
    
    const allStudents = studentsResult.students;
    const classrooms = [...new Set(allStudents.map(s => s.classroom))].sort();
    
    // กำหนดช่วงวันที่
    const today = new Date();
    const fromDate = dateFrom ? new Date(dateFrom) : new Date(today.getFullYear(), today.getMonth(), 1);
    const toDate = dateTo ? new Date(dateTo) : today;
    
    const fromDateStr = Utilities.formatDate(fromDate, 'Asia/Bangkok', 'dd/MM/yyyy');
    const toDateStr = Utilities.formatDate(toDate, 'Asia/Bangkok', 'dd/MM/yyyy');
    
    // กรองข้อมูลการเข้าเรียนตามช่วงวันที่
    const attendanceInRange = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        const recordDate = new Date(data[i][0].split('/').reverse().join('-'));
        if (recordDate >= fromDate && recordDate <= toDate) {
          attendanceInRange.push({
            date: data[i][0],
            time: data[i][1],
            studentId: data[i][2],
            studentName: data[i][3],
            classroom: data[i][4],
            status: data[i][5]
          });
        }
      }
    }
    
    // สถิติรวม
    const totalStudents = allStudents.length;
    const totalRecords = attendanceInRange.length;
    const uniqueStudentsAttended = [...new Set(attendanceInRange.map(r => r.studentId))].length;
    
    // สถิติตามห้องเรียน
    const classroomStats = classrooms.map(classroom => {
      const studentsInClass = allStudents.filter(s => s.classroom === classroom);
      const attendanceInClass = attendanceInRange.filter(r => r.classroom === classroom);
      const uniqueAttended = [...new Set(attendanceInClass.map(r => r.studentId))].length;
      
      return {
        classroom: classroom,
        totalStudents: studentsInClass.length,
        attendedStudents: uniqueAttended,
        attendancePercentage: studentsInClass.length > 0 ? Math.round((uniqueAttended / studentsInClass.length) * 100) : 0,
        totalRecords: attendanceInClass.length
      };
    });
    
    // สถิติรายวัน (7 วันล่าสุด)
    const dailyStats = [];
    for (let i = 6; i >= 0; i--) {
      const date = new Date(today);
      date.setDate(date.getDate() - i);
      const dateStr = Utilities.formatDate(date, 'Asia/Bangkok', 'dd/MM/yyyy');
      const dayName = Utilities.formatDate(date, 'Asia/Bangkok', 'EEEE');
      
      const dailyAttendance = attendanceInRange.filter(r => r.date === dateStr);
      const uniqueDaily = [...new Set(dailyAttendance.map(r => r.studentId))].length;
      
      dailyStats.push({
        date: dateStr,
        dayName: dayName,
        attendedStudents: uniqueDaily,
        totalRecords: dailyAttendance.length,
        attendancePercentage: totalStudents > 0 ? Math.round((uniqueDaily / totalStudents) * 100) : 0
      });
    }
    
    return {
      success: true,
      dateRange: {
        from: fromDateStr,
        to: toDateStr
      },
      overall: {
        totalStudents: totalStudents,
        totalRecords: totalRecords,
        uniqueStudentsAttended: uniqueStudentsAttended,
        attendancePercentage: totalStudents > 0 ? Math.round((uniqueStudentsAttended / totalStudents) * 100) : 0
      },
      classroomStats: classroomStats,
      dailyStats: dailyStats,
      topAttendees: getTopAttendees(attendanceInRange, allStudents, 10)
    };
    
  } catch (error) {
    console.error('Error getting attendance stats:', error);
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการดึงสถิติ: ' + error.message
    };
  }
}

/**
 * ฟังก์ชันดึงนักเรียนที่เข้าเรียนมากที่สุด
 */
function getTopAttendees(attendanceData, allStudents, limit = 10) {
  const attendanceCount = {};
  
  attendanceData.forEach(record => {
    if (attendanceCount[record.studentId]) {
      attendanceCount[record.studentId]++;
    } else {
      attendanceCount[record.studentId] = 1;
    }
  });
  
  const topAttendees = Object.entries(attendanceCount)
    .map(([studentId, count]) => {
      const student = allStudents.find(s => s.id === studentId);
      return {
        studentId: studentId,
        studentName: student ? student.fullName : 'ไม่พบข้อมูล',
        classroom: student ? student.classroom : '-',
        attendanceCount: count
      };
    })
    .sort((a, b) => b.attendanceCount - a.attendanceCount)
    .slice(0, limit);
  
  return topAttendees;
}

/**
 * ฟังก์ชันสำหรับ client-side เรียกใช้
 */
function checkAttendance(studentId, studentFullName, studentClassroom, status) {
  return recordAttendance(studentId, studentFullName, studentClassroom, status);
}

function getStudentsList() {
  return getAllStudents();
}

function getTodayAttendanceList() {
  return getTodayAttendance();
}

function searchStudent(studentId) {
  return findStudentById(studentId);
}

function getClassroomsList() {
  return getAllClassrooms();
}

function getClassroomStudents(classroom) {
  return getStudentsByClassroom(classroom);
}

function getDashboardStats(dateFrom, dateTo) {
  return getAttendanceStats(dateFrom, dateTo);
}
