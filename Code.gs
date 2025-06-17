/**
 * ระบบเช็คชื่อนักเรียน - Google Apps Script Backend
 * เชื่อมต่อกับ Google Sheets เพื่อจัดเก็บข้อมูลการเข้าเรียน
 */

// =====================
// CONFIGURATION & CONSTANTS
// =====================
const STUDENTS_SHEET_NAME = 'Students';
const CLASSROOMS_SHEET_NAME = 'Classrooms';
const ATTENDANCE_SHEET_NAME = 'Attendance';
const USERS_SHEET_NAME = 'Users';
const JWT_SECRET_KEY_PROPERTY = 'asdasdlglkbmkbtokb;ltmblmdfdfb';
const JWT_EXPIRATION_SECONDS = 30 * 24 * 60 * 60; // 30 days

// =====================
// UTILITY FUNCTIONS
// =====================
function generateSalt() {
  return Math.random().toString(36).substring(2, 15) + Math.random().toString(36).substring(2, 15);
}
function customBytesToHex(bytes) {
  if (!bytes) return null;
  let hex = '';
  for (let i = 0; i < bytes.length; i++) {
    let byte = bytes[i] & 0xFF;
    let hexByte = byte.toString(16);
    if (hexByte.length < 2) hexByte = '0' + hexByte;
    hex += hexByte;
  }
  return hex;
}
function hashPassword(password, salt) {
  Logger.log(`HASH_PASSWORD: Entered. Password provided: ${password ? 'Yes' : 'No'}, Salt provided: ${salt ? 'Yes' : 'No'}`);
  if (!password || !salt) {
    Logger.log('HASH_PASSWORD: Password or salt is missing.');
    return null;
  }
  const saltedPassword = password + salt;
  Logger.log(`HASH_PASSWORD: Salted password created.`);
  try {
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, saltedPassword, Utilities.Charset.UTF_8);
    Logger.log(`HASH_PASSWORD: Digest computed. Length: ${digest ? digest.length : 'null'}`);
    const hex = customBytesToHex(digest);
    if (hex) {
      Logger.log('HASH_PASSWORD: customBytesToHex conversion successful.');
      return hex;
    } else {
      Logger.log('HASH_PASSWORD: CRITICAL - customBytesToHex failed. Returning null.');
      throw new Error('customBytesToHex failed to convert digest.');
    }
  } catch (error) {
    Logger.log(`HASH_PASSWORD: ERROR during hashing process - ${error.toString()} Stack: ${error.stack}`);
    throw error;
  }
}
function base64UrlEncode(input) {
  let base64;
  if (typeof input === 'string') {
    base64 = Utilities.base64Encode(input, Utilities.Charset.UTF_8);
  } else if (Array.isArray(input)) {
    base64 = Utilities.base64EncodeWebSafe(input);
    base64 = base64.replace(/=+$/, '');
    return base64;
  } else {
    throw new Error('base64UrlEncode: Unsupported input type');
  }
  return base64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
}
function base64UrlDecode(str) {
  str = str.replace(/-/g, '+').replace(/_/g, '/');
  while (str.length % 4) str += '=';
  try {
    const decoded = Utilities.base64Decode(str);
    return Utilities.newBlob(decoded).getDataAsString();
  } catch (e) {
    Logger.log('BASE64_URL_DECODE_ERROR: ' + e.message);
    throw new Error('Failed to decode base64url: ' + e.message);
  }
}
function getJwtSecret() {
  let secret = PropertiesService.getScriptProperties().getProperty(JWT_SECRET_KEY_PROPERTY);
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid();
    PropertiesService.getScriptProperties().setProperty(JWT_SECRET_KEY_PROPERTY, secret);
    Logger.log('JWT_SECRET_KEY generated and stored.');
  }
  return secret;
}
function generateJwt(userInfo) {
  const secret = getJwtSecret();
  const header = { alg: 'HS256', typ: 'JWT' };
  const now = Math.floor(Date.now() / 1000);
  const payload = {
    user: { username: userInfo.username, role: userInfo.role, fullName: userInfo.fullName },
    iat: now,
    exp: now + JWT_EXPIRATION_SECONDS,
    iss: ScriptApp.getService().getUrl()
  };
  const encodedHeader = base64UrlEncode(JSON.stringify(header));
  const encodedPayload = base64UrlEncode(JSON.stringify(payload));
  const signatureInput = encodedHeader + '.' + encodedPayload;
  const signatureBytes = Utilities.computeHmacSha256Signature(signatureInput, secret, Utilities.Charset.UTF_8);
  const encodedSignature = base64UrlEncode(signatureBytes);
  return encodedHeader + '.' + encodedPayload + '.' + encodedSignature;
}
function verifyJwt(token) {
  if (!token) {
    Logger.log('VERIFY_JWT: Token is null or undefined.');
    return { valid: false, error: 'Token not provided' };
  }
  const parts = token.split('.');
  if (parts.length !== 3) {
    Logger.log('VERIFY_JWT: Token does not have 3 parts.');
    return { valid: false, error: 'Invalid token structure' };
  }
  const encodedHeader = parts[0];
  const encodedPayload = parts[1];
  const encodedSignature = parts[2];
  const secret = getJwtSecret();
  const signatureInput = encodedHeader + '.' + encodedPayload;
  const expectedSignatureBytes = Utilities.computeHmacSha256Signature(signatureInput, secret, Utilities.Charset.UTF_8);
  const expectedEncodedSignature = base64UrlEncode(expectedSignatureBytes);
  if (expectedEncodedSignature !== encodedSignature) {
    Logger.log('VERIFY_JWT: Signature mismatch.');
    return { valid: false, error: 'Invalid signature' };
  }
  let payload;
  try {
    payload = JSON.parse(base64UrlDecode(encodedPayload));
  } catch (e) {
    Logger.log('VERIFY_JWT: Error decoding payload: ' + e.message);
    return { valid: false, error: 'Invalid payload encoding' };
  }
  const now = Math.floor(Date.now() / 1000);
  if (payload.exp && now > payload.exp) {
    Logger.log('VERIFY_JWT: Token expired at ' + new Date(payload.exp * 1000) + '. Current time: ' + new Date(now * 1000));
    return { valid: false, error: 'Token expired', expired: true };
  }
  Logger.log('VERIFY_JWT: Token verified successfully for user: ' + payload.user.username);
  return { valid: true, payload: payload };
}

// =====================
// INITIALIZATION FUNCTIONS
// =====================
function initializeSheets() {
  Logger.log('ENTERING initializeSheets'); 
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // สร้างหรือตรวจสอบ Sheet สำหรับข้อมูลนักเรียน
    let studentSheet = spreadsheet.getSheetByName(STUDENTS_SHEET_NAME);
    if (!studentSheet) {
      studentSheet = spreadsheet.insertSheet(STUDENTS_SHEET_NAME);
    }
    const studentLastRow = studentSheet.getLastRow();
    if (studentLastRow === 0 || studentSheet.getRange(1, 1).getValue() === '') {
      studentSheet.getRange(1, 1, 1, 4).setValues([
        ['รหัสนักเรียน', 'ชื่อ', 'นามสกุล', 'ห้องเรียนID']
      ]);
      const headerRange = studentSheet.getRange(1, 1, 1, 4);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#E1BAFF');
      headerRange.setHorizontalAlignment('center');
      if (studentLastRow === 0) {
        studentSheet.getRange(2, 1, 5, 4).setValues([
          ['S20001', 'สมชาย', 'ใจดี', 'C101'],
          ['S20002', 'สมหญิง', 'ใจงาม', 'C101'],
          ['S20003', 'ประเสริฐ', 'เก่งเก้า', 'C102'],
          ['S20004', 'วรรณา', 'สวยงาม', 'C102'],
          ['S20005', 'ชัยวัฒน์', 'รุ่งเรือง', 'C103']
        ]);
      }
    }

    // สร้างหรือตรวจสอบ Sheet สำหรับข้อมูลห้องเรียน
    let classroomSheet = spreadsheet.getSheetByName(CLASSROOMS_SHEET_NAME);
    if (!classroomSheet) {
      classroomSheet = spreadsheet.insertSheet(CLASSROOMS_SHEET_NAME);
    }
    const classroomLastRow = classroomSheet.getLastRow();
    if (classroomLastRow === 0 || classroomSheet.getRange(1, 1).getValue() === '') {
      classroomSheet.getRange(1, 1, 1, 2).setValues([
        ['ห้องเรียนID', 'ชื่อห้องเรียน']
      ]);
      const headerRange = classroomSheet.getRange(1, 1, 1, 2);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#BAE1FF');
      headerRange.setHorizontalAlignment('center');
      if (classroomLastRow === 0) {
          classroomSheet.getRange(2, 1, 3, 2).setValues([
              ['C101', 'ม.1/1'],
              ['C102', 'ม.1/2'],
              ['C103', 'ม.1/3']
          ]);
      }
    }
    
    // สร้างหรือตรวจสอบ Sheet สำหรับข้อมูลการเข้าเรียน
    let attendanceSheet = spreadsheet.getSheetByName(ATTENDANCE_SHEET_NAME);
    if (!attendanceSheet) {
      attendanceSheet = spreadsheet.insertSheet(ATTENDANCE_SHEET_NAME);
    }    const attendanceLastRow = attendanceSheet.getLastRow();
    if (attendanceLastRow === 0 || attendanceSheet.getRange(1, 1).getValue() === '') {
      attendanceSheet.getRange(1, 1, 1, 8).setValues([
        ['Date', 'Time', 'StudentID', 'FirstName', 'LastName', 'Classroom', 'Status', 'RecordedBy']
      ]);
      const headerRange = attendanceSheet.getRange(1, 1, 1, 8);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#BAFFC9');
      headerRange.setHorizontalAlignment('center');
    } else {
      // Update existing headers if they're in Thai format
      const updateResult = updateAttendanceSheetHeaders();
      if (updateResult.success && updateResult.headersUpdated) {
        Logger.log('initializeSheets: Attendance sheet headers updated from Thai to English');
      }
    }
    
    Logger.log('initializeSheets: About to call initializeUsersSheet');
    initializeUsersSheet(); 
    Logger.log('initializeSheets: Returned from initializeUsersSheet SUCCESSFULLY');
    
    // ตรวจสอบและแก้ไขข้อมูลใน sheet Students
    Logger.log('initializeSheets: About to validate and fix student sheet data');
    const validationResult = validateAndFixStudentSheetData();
    if (validationResult.success) {
      Logger.log('initializeSheets: Student sheet validation complete: ' + validationResult.message);
      if (validationResult.studentsFixed) {
        Logger.log('initializeSheets: Some student data was fixed to match classroom data.');
      }
    } else {
      Logger.log('initializeSheets: Warning - Student sheet validation failed: ' + validationResult.message);
      // ไม่ throw error เพื่อให้โปรแกรมยังทำงานต่อไปได้
    }
    
    Logger.log('initializeSheets: About to get spreadsheet ID.');
    const id = spreadsheet.getId();
    Logger.log('initializeSheets: spreadsheet.getId() succeeded: ' + id);

    const returnValue = {
      success: true,
      spreadsheetId: id,
      message: 'เตรียม Sheets เรียบร้อยแล้ว' + 
        (validationResult.studentsFixed ? ' (ได้ทำการแก้ไขข้อมูลนักเรียนบางส่วนให้ตรงกับห้องเรียน)' : '')
    };
    Logger.log('initializeSheets: Success object constructed. About to return.');
    return returnValue;
    
  } catch (error) {
    Logger.log('ERROR INSIDE initializeSheets: ' + error.toString() + ' Stack: ' + error.stack);
    // console.error('Error initializing sheets:', error); // Original line
    // return { // Original block
    //   success: false,
    //   message: 'เกิดข้อผิดพลาดในการเตรียม Sheets: ' + error.message
    // };
    throw error; // Re-throw to be caught by doGet, to match current behavior and provide stack to client if possible
  }
}

function initializeUsersSheet() {
  Logger.log('ENTERING initializeUsersSheet'); // Diagnostic log
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(USERS_SHEET_NAME);
    let createdNewSheet = false;

    if (!sheet) {
      sheet = ss.insertSheet(USERS_SHEET_NAME);
      Logger.log(`Sheet '${USERS_SHEET_NAME}' created.`);
      createdNewSheet = true;
    }

    const expectedHeaders = ['Username', 'PasswordHash', 'Salt', 'Role', 'FullName', 'LastLogin'];
    let headersAreCorrect = false;
    if (sheet.getLastRow() > 0) {
      try {
        const actualHeaderValues = sheet.getRange(1, 1, 1, expectedHeaders.length).getValues()[0];
        headersAreCorrect = expectedHeaders.every((expectedHeader, i) => 
            actualHeaderValues[i] && actualHeaderValues[i].toString().trim().toLowerCase() === expectedHeader.toLowerCase()
        );
      } catch (e) {
        Logger.log(`Error reading headers from Users sheet, assuming headers are incorrect: ${e.message}`);
        headersAreCorrect = false;
      }
    }

    if (createdNewSheet || !headersAreCorrect) {
      if (!createdNewSheet && !headersAreCorrect && sheet.getLastRow() > 0) {
          Logger.log(`Sheet '${USERS_SHEET_NAME}' exists but headers are incorrect or missing. Re-initializing headers.`);
          // Optional: Clear existing content if headers are wrong and sheet is not empty.
          // sheet.clearContents(); // Be cautious with this in production.
      } else if (createdNewSheet) {
          Logger.log(`Sheet '${USERS_SHEET_NAME}' is new. Initializing headers.`);
      } else { // Sheet existed, was empty (getLastRow() === 0), and thus headersAreCorrect was false
          Logger.log(`Sheet '${USERS_SHEET_NAME}' is empty. Initializing headers.`);
      }

      // Ensure sheet has enough columns for headers
      if(sheet.getMaxColumns() < expectedHeaders.length){
          sheet.insertColumns(sheet.getMaxColumns() + 1, expectedHeaders.length - sheet.getMaxColumns());
      }
      // Set headers in the first row
      sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, expectedHeaders.length).setFontWeight('bold');
      Logger.log(`Headers set for sheet '${USERS_SHEET_NAME}'.`);

      // Add admin user with the specified password if headers were just (re)created
      addInitialAdminUser(sheet, 'admin1234'); 
    } else {
      Logger.log(`Sheet '${USERS_SHEET_NAME}' already exists and has correct headers.`);
      // Even if headers are correct, ensure the admin user exists.
      addInitialAdminUser(sheet, 'admin1234');
    }
    Logger.log('EXITING initializeUsersSheet SUCCESSFULLY'); // Diagnostic log
  } catch (e) {
    Logger.log('ERROR INSIDE initializeUsersSheet: ' + e.toString() + ' Stack: ' + e.stack); // Detailed error log
    throw e; // Re-throw the error
  }
}

function addInitialAdminUser(sheet, defaultPassword) { 
  Logger.log(`ADD_INITIAL_ADMIN: Entered for sheet '${sheet.getName()}'. Target password: '${defaultPassword ? '(provided)' : '(not provided)'}'.`);
  let adminExists = false;
  const lastRow = sheet.getLastRow();
  Logger.log(`ADD_INITIAL_ADMIN: Sheet last row is ${lastRow}.`);

  if (lastRow >= 2) { // Only check if there are potential data rows (beyond header)
    Logger.log(`ADD_INITIAL_ADMIN: Sheet has ${lastRow} rows. Checking for existing 'admin' user in data rows.`);
    try {
      // Read only the first column (Usernames) from the second row to the last data row
      const usernames = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      adminExists = usernames.some(row => row[0] && row[0].toString().trim().toLowerCase() === 'admin');
      Logger.log(`ADD_INITIAL_ADMIN: Existing 'admin' user check result: ${adminExists}.`);
    } catch (e) {
      Logger.log(`ADD_INITIAL_ADMIN: Error reading usernames: ${e.message}. Assuming 'admin' does not exist to be safe.`);
      adminExists = false; // Default to false on error to attempt creation
    }
  } else {
    Logger.log(`ADD_INITIAL_ADMIN: Sheet has ${lastRow} row(s) (header or empty). 'admin' user assumed not to exist yet.`);
    adminExists = false;
  }

  if (!adminExists) {
    Logger.log(`ADD_INITIAL_ADMIN: 'admin' user not found. Proceeding to add.`);
    const username = 'admin';
    const role = 'admin';
    const fullName = 'Administrator';
    
    const salt = generateSalt();
    Logger.log(`ADD_INITIAL_ADMIN: Salt generated: ${salt}`);
    const passwordToHash = defaultPassword;
    const passwordHash = hashPassword(passwordToHash, salt);
    
    if (passwordHash) {
      Logger.log(`ADD_INITIAL_ADMIN: Password hashed successfully. Appending admin user row.`);
      try {
        sheet.appendRow([username, passwordHash, salt, role, fullName, new Date()]);
        SpreadsheetApp.flush(); // Ensure changes are written immediately
        Logger.log(`ADD_INITIAL_ADMIN: Default admin user '${username}' successfully appended. Sheet should now have ${sheet.getLastRow()} rows.`);
      } catch (appendError) {
        Logger.log(`ADD_INITIAL_ADMIN: CRITICAL ERROR appending admin user row: ${appendError.message} Stack: ${appendError.stack}`);
      }
    } else {
      Logger.log(`ADD_INITIAL_ADMIN: Failed to hash password for default admin user '${username}'. User NOT added.`);
    }
  } else {
    Logger.log(`ADD_INITIAL_ADMIN: Default admin user 'admin' ALREADY EXISTS. No changes made to user. Sheet has ${sheet.getLastRow()} rows.`);
  }
}

// =====================
// MAIN LOGIC AND ENDPOINTS
// =====================
function doGet(e) {
  Logger.log('--- doGet STARTED --- Parameters: ' + JSON.stringify(e));
  
  try {
    Logger.log('doGet: About to call initializeSheets()');
    const initResult = initializeSheets(); // Store the result
    Logger.log('doGet: initializeSheets() returned. Result: ' + JSON.stringify(initResult));

    // Check if initializeSheets itself reported a failure in its return object
    if (initResult && initResult.success === false) {
        Logger.log('doGet: initializeSheets reported failure in its return object: ' + (initResult.message || 'Unknown error from init'));
        return HtmlService.createHtmlOutput("<b>System Initialization Error:</b> " + (initResult.message || 'Failed to initialize system components.'))
          .setTitle("System Error")
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    Logger.log('doGet: initializeSheets() appears to have succeeded based on its return object or did not return a success flag.');

  } catch (initError) {
    Logger.log('doGet: CAUGHT an error from the initializeSheets() call block.');
    let errorMessage = 'Unknown initialization error.';
    let errorStack = 'No stack available.';
    if (initError && typeof initError === 'object') {
        if (initError.message) errorMessage = initError.message;
        if (initError.stack) errorStack = initError.stack;
    } else if (initError) {
        errorMessage = initError.toString();
    }
    Logger.log(`CRITICAL ERROR during initializeSheets() in doGet: Message: ${errorMessage} Stack: ${errorStack}`);
    return HtmlService.createHtmlOutput(`<b>System Initialization Error:</b> Unable to initialize critical components. (Details: ${errorMessage}). Please try again later or contact support.`)
      .setTitle("System Error")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  Logger.log('doGet: Proceeding after initializeSheets() try...catch block.');

  // The 'page' parameter can still determine which base HTML to serve.
  // The client will then handle auth state.
  let pageToServe = e.parameter.page || 'index'; // Default to index, could be login
  let template;
  let title = 'Student Attendance System';
  // Always serve login.html if requested, or if it's the intended entry point for unauthenticated users.
  // The client-side script in index.html will check for a token and redirect to login if needed.
  if (pageToServe === 'login') {
      template = HtmlService.createTemplateFromFile('login');
      template.appUrl = ScriptApp.getService().getUrl(); // For form submissions or API calls
      // Pass other necessary parameters like error messages if redirected from server
      template.errorMessage = e.parameter.error || null;
      template.infoMessage = e.parameter.message || null;
      template.dest = e.parameter.dest || null; // Add the dest parameter for redirection after login
      title = 'Login - Student Attendance';
  } else if (pageToServe === 'index') {
      template = HtmlService.createTemplateFromFile('index');
      template.appUrl = ScriptApp.getService().getUrl();
      // User data will be fetched by client-side JS using the JWT
      title = 'Student Attendance';  } else if (pageToServe === 'dashboard') {
      template = HtmlService.createTemplateFromFile('dashboard');
      template.appUrl = ScriptApp.getService().getUrl();
      // User data and access control will be handled by client-side JS
      title = 'Dashboard - Student Attendance';
  } else if (pageToServe === 'search') {
      template = HtmlService.createTemplateFromFile('search');
      template.appUrl = ScriptApp.getService().getUrl();
      // User data and access control will be handled by client-side JS
      title = 'ค้นหาข้อมูลการเข้าแถว - Student Attendance';
  } else {
      // Fallback for unknown pages - serve index, client can decide what to do
      Logger.log(`Unknown page '${pageToServe}' requested. Serving index page.`);
      template = HtmlService.createTemplateFromFile('index');
      template.appUrl = ScriptApp.getService().getUrl();
      title = 'Student Attendance';
  }
  
  if (!template) { // Should not happen with the logic above, but as a safeguard
      Logger.log(`Error: No template resolved for page '${pageToServe}'. Serving a basic error.`);
      return HtmlService.createHtmlOutput("<b>Page Not Found:</b> The requested page could not be loaded.")
          .setTitle("Error")
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  let htmlOutput = template.evaluate().setTitle(title);
  return htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Client calls this with a token to get user data and verify session
function getUserDataFromToken(token) {
  Logger.log('Attempting to get user data from token.');
  const verificationResult = verifyJwt(token);
  if (verificationResult.valid) {
    Logger.log('Token valid. Returning user data for: ' + verificationResult.payload.user.username);
    return { 
      success: true, 
      user: verificationResult.payload.user,
      issuedAt: verificationResult.payload.iat,
      expiresAt: verificationResult.payload.exp
    };
  } else {
    Logger.log('Token invalid or expired: ' + verificationResult.error);
    return { success: false, error: verificationResult.error, expired: verificationResult.expired || false };
  }
}

function logoutUser(token) { // Token might be passed for logging or if a blocklist was implemented
  // For stateless JWT, logout is primarily client-side (deleting the token).
  // Server-side, we could log the logout attempt.
  // If a token blocklist were implemented, this is where you'd add the token to it.
  Logger.log('Logout requested. Client should discard token: ' + (token ? token.substring(0,20) + "..." : "No token provided"));
  return { success: true, message: 'Logged out. Please discard your token.' };
}

// Example of how a protected function would change:
function getStudents(authToken) { // Expects JWT to be passed by client
  const verificationResult = verifyJwt(authToken);
  if (!verificationResult.valid) {
    return { success: false, error: 'Authentication failed: ' + verificationResult.error, students: [] };
  }
  // Proceed if token is valid
  // Logger.log(`User ${verificationResult.payload.user.username} requesting students.`);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STUDENTS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  data.shift(); // Remove header row
  const students = data.map(row => ({ id: row[0], firstName: row[1], lastName: row[2], classroom: row[3] }));
  return { success: true, students: students };
}

function recordAttendance(studentId, status, authToken) { // Expects JWT
  const verificationResult = verifyJwt(authToken);
  if (!verificationResult.valid) {
    return { success: false, message: 'Authentication failed: ' + verificationResult.error, expired: verificationResult.expired || false };
  }
  const userMakingRequest = verificationResult.payload.user;
  Logger.log(`Attendance record attempt by ${userMakingRequest.username} (Role: ${userMakingRequest.role}) for student ${studentId} with status ${status}`);

  try {
    const attendanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ATTENDANCE_SHEET_NAME);
    const studentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STUDENTS_SHEET_NAME);

    if (!attendanceSheet || !studentSheet) {
      Logger.log(`RECORD_ATTENDANCE: Error - Missing sheet. Attendance: ${!attendanceSheet}, Student: ${!studentSheet}`);
      return { success: false, message: 'Data sheet missing.' };
    }    // Get student info
    const studentsData = studentSheet.getDataRange().getValues();
    const studentHeader = studentsData.shift(); // Remove header
    
    const studentIdColIdx = studentHeader.findIndex(h => h && (h.toString().trim() === 'รหัสนักเรียน' || h.toString().trim().toLowerCase() === 'studentid'));
    const studentFirstNameColIdx = studentHeader.findIndex(h => h && (h.toString().trim() === 'ชื่อ' || h.toString().trim().toLowerCase() === 'firstname'));
    const studentLastNameColIdx = studentHeader.findIndex(h => h && (h.toString().trim() === 'นามสกุล' || h.toString().trim().toLowerCase() === 'lastname'));
    const studentClassroomColIdx = studentHeader.findIndex(h => h && (
        h.toString().trim() === 'ห้องเรียนID' || 
        h.toString().trim() === 'ห้องเรียน' || // Added this variation
        h.toString().trim().toLowerCase() === 'classroomid' ||
        h.toString().trim().toLowerCase() === 'classroom' // Added this variation
    ));

    // Debug logging for header detection
    Logger.log(`RECORD_ATTENDANCE DEBUG: Student headers found: ${JSON.stringify(studentHeader)}`);
    Logger.log(`RECORD_ATTENDANCE DEBUG: Column indices - ID: ${studentIdColIdx}, FirstName: ${studentFirstNameColIdx}, LastName: ${studentLastNameColIdx}, Classroom: ${studentClassroomColIdx}`);
    studentHeader.forEach((header, index) => {
        Logger.log(`RECORD_ATTENDANCE DEBUG: Header ${index}: "${header}" (trimmed: "${header.toString().trim()}")`);
    });

    if ([studentIdColIdx, studentFirstNameColIdx, studentLastNameColIdx, studentClassroomColIdx].some(idx => idx === -1)) {
        let missingHeadersDetails = [];
        if (studentIdColIdx === -1) missingHeadersDetails.push("คอลัมน์สำหรับ 'รหัสนักเรียน' (Student ID)");
        if (studentFirstNameColIdx === -1) missingHeadersDetails.push("คอลัมน์สำหรับ 'ชื่อ' (First Name)");
        if (studentLastNameColIdx === -1) missingHeadersDetails.push("คอลัมน์สำหรับ 'นามสกุล' (Last Name)");        if (studentClassroomColIdx === -1) missingHeadersDetails.push("คอลัมน์สำหรับ 'ห้องเรียนID' หรือ 'ห้องเรียน' (Classroom ID)");
        
        const actualHeadersString = studentHeader ? studentHeader.map((h, idx) => `คอลัมน์ ${idx + 1}: "${h}"`).join(', ') : "ไม่พบข้อมูลส่วนหัว";
        const errorMessage = `พบปัญหาโครงสร้างชีตนักเรียน: ไม่พบส่วนหัวที่จำเป็น: ${missingHeadersDetails.join('; ')}. กรุณาตรวจสอบส่วนหัวในชีต 'Students'. ส่วนหัวที่คาดหวังเช่น ['รหัสนักเรียน', 'ชื่อ', 'นามสกุล', 'ห้องเรียนID' หรือ 'ห้องเรียน']. ส่วนหัวที่พบจริง: [${actualHeadersString}].`;
        Logger.log(`RECORD_ATTENDANCE: ${errorMessage}`);
        return { success: false, message: errorMessage, errorDetails: { expected: ['รหัสนักเรียน', 'ชื่อ', 'นามสกุล', 'ห้องเรียนID'], actual: studentHeader || [] } };
    }
    
    let studentInfo = null;
    for (let i = 0; i < studentsData.length; i++) { // Iterate over data rows (header already removed)
      if (studentsData[i][studentIdColIdx] && studentsData[i][studentIdColIdx].toString().trim() === studentId.toString().trim()) {
        studentInfo = {
          id: studentsData[i][studentIdColIdx],
          name: `${studentsData[i][studentFirstNameColIdx]} ${studentsData[i][studentLastNameColIdx]}`,
          classroom: studentsData[i][studentClassroomColIdx]
        };
        break;
      }
    }

    if (!studentInfo) {
      Logger.log(`RECORD_ATTENDANCE: Student with ID ${studentId} not found.`);
      return { success: false, message: 'Student not found.' };
    }

    const now = new Date();
    const currentDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const currentTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');

    const attendanceData = attendanceSheet.getDataRange().getValues();
    const attendanceHeader = attendanceData.shift(); // Remove header

    // Dynamically find column indices in Attendance sheet
    const dateColIdx = attendanceHeader.findIndex(h => h && (h.toString().trim() === 'วันที่' || h.toString().trim().toLowerCase() === 'date'));
    const studentIdAttColIdx = attendanceHeader.findIndex(h => h && (h.toString().trim() === 'รหัสนักเรียน' || h.toString().trim().toLowerCase() === 'studentid'));
    const timeColIdx = attendanceHeader.findIndex(h => h && (h.toString().trim() === 'เวลา' || h.toString().trim().toLowerCase() === 'time'));
    const statusColIdx = attendanceHeader.findIndex(h => h && (h.toString().trim() === 'สถานะ' || h.toString().trim().toLowerCase() === 'status'));
    // Optional: classroom and name columns if they need to be updated, though typically they wouldn't change for an existing record.
    // const nameAttColIdx = attendanceHeader.findIndex(h => h && (h.toString().trim() === 'ชื่อ-นามสกุล' || h.toString().trim().toLowerCase() === 'name'));
    // const classroomAttColIdx = attendanceHeader.findIndex(h => h && (h.toString().trim() === 'ห้องเรียน' || h.toString().trim().toLowerCase() === 'classroom'));


    if ([dateColIdx, studentIdAttColIdx, timeColIdx, statusColIdx].some(idx => idx === -1)) {
        Logger.log(`RECORD_ATTENDANCE: Error - Could not find all required columns in Attendance sheet. Indices - Date: ${dateColIdx}, StudentID: ${studentIdAttColIdx}, Time: ${timeColIdx}, Status: ${statusColIdx}`);
        return { success: false, message: 'Attendance sheet structure error.' };
    }

    let recordUpdated = false;
    let existingRowNumber = -1;

    for (let i = 0; i < attendanceData.length; i++) {
      const row = attendanceData[i];
      // Ensure row has enough columns and values are not null/undefined before accessing
      if (row.length > Math.max(dateColIdx, studentIdAttColIdx) && row[dateColIdx] && row[studentIdAttColIdx]) {
        let recordDateStr = row[dateColIdx];
        // Handle cases where date might be a Date object or a string
        if (recordDateStr instanceof Date) {
            recordDateStr = Utilities.formatDate(recordDateStr, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
            // Attempt to parse if it's a string that might not be in yyyy-MM-dd
            try {
                recordDateStr = Utilities.formatDate(new Date(recordDateStr), Session.getScriptTimeZone(), 'yyyy-MM-dd');
            } catch (e) {
                // If parsing fails, log and skip or use as is if it might match
                Logger.log(`RECORD_ATTENDANCE: Could not parse date string "${row[dateColIdx]}" for row ${i+2}. Skipping or using as is.`);
            }
        }

        if (recordDateStr === currentDate && row[studentIdAttColIdx].toString().trim() === studentInfo.id.toString().trim()) {
          // Found existing record for today and this student
          existingRowNumber = i + 2; // +1 for 1-based index, +1 because header was shifted
          attendanceSheet.getRange(existingRowNumber, timeColIdx + 1).setValue(currentTime);
          attendanceSheet.getRange(existingRowNumber, statusColIdx + 1).setValue(status);
          // Optionally update name and classroom if they might change or were initially wrong, though less common.
          // attendanceSheet.getRange(existingRowNumber, nameAttColIdx + 1).setValue(studentInfo.name);
          // attendanceSheet.getRange(existingRowNumber, classroomAttColIdx + 1).setValue(studentInfo.classroom);
          recordUpdated = true;
          Logger.log(`RECORD_ATTENDANCE: Updated record for student ${studentInfo.id} at row ${existingRowNumber}. Status: ${status}, Time: ${currentTime}`);
          break;
        }
      }
    }

    if (!recordUpdated) {
      // No existing record found, append a new one
      // Ensure all columns for appendRow are present as per sheet structure
      const newRecord = [];
      newRecord[dateColIdx] = currentDate;
      newRecord[timeColIdx] = currentTime;
      newRecord[studentIdAttColIdx] = studentInfo.id;
      
      // Find name and classroom columns for new record
      const nameAttColIdx = attendanceHeader.findIndex(h => h && (h.toString().trim() === 'ชื่อ-นามสกุล' || h.toString().trim().toLowerCase() === 'name'));
      const classroomAttColIdx = attendanceHeader.findIndex(h => h && (h.toString().trim() === 'ห้องเรียน' || h.toString().trim().toLowerCase() === 'classroom'));

      if (nameAttColIdx !== -1) newRecord[nameAttColIdx] = studentInfo.name;
      else Logger.log("RECORD_ATTENDANCE: 'ชื่อ-นามสกุล' column not found for new record, will be blank.");
      
      if (classroomAttColIdx !== -1) newRecord[classroomAttColIdx] = studentInfo.classroom;
      else Logger.log("RECORD_ATTENDANCE: 'ห้องเรียน' column not found for new record, will be blank.");
      
      newRecord[statusColIdx] = status;

      // Fill any gaps in the array with empty strings up to the max columns defined by header
      const finalRecord = [];
      for(let i=0; i < attendanceHeader.length; i++) {
          finalRecord[i] = newRecord[i] !== undefined ? newRecord[i] : "";
      }

      attendanceSheet.appendRow(finalRecord);
      Logger.log(`RECORD_ATTENDANCE: Appended new record for student ${studentInfo.id}. Status: ${status}`);
    }
    
    // Recalculate total present today for the response (optional, client might do this)
    // This is a simplified count, a more robust one would re-query like getDailyAttendanceStats
    let totalPresentToday = 0;
    const updatedAttendanceData = attendanceSheet.getDataRange().getValues();
    updatedAttendanceData.shift(); // remove header
    updatedAttendanceData.forEach(row => {
        if (row.length > Math.max(dateColIdx, statusColIdx) && row[dateColIdx]) {
            let recordDateStr = row[dateColIdx];
            if (recordDateStr instanceof Date) {
                recordDateStr = Utilities.formatDate(recordDateStr, Session.getScriptTimeZone(), 'yyyy-MM-dd');
            } else {
                 try { recordDateStr = Utilities.formatDate(new Date(recordDateStr), Session.getScriptTimeZone(), 'yyyy-MM-dd'); } catch(e){}
            }
            if (recordDateStr === currentDate && (row[statusColIdx] === 'present' || row[statusColIdx] === 'late')) {
                totalPresentToday++;
            }
        }
    });    return { 
      success: true, 
      message: 'Attendance ' + (recordUpdated ? 'updated' : 'recorded') + ' for ' + studentInfo.name + ' (' + status + ') by ' + userMakingRequest.username,
      action: recordUpdated ? 'updated' : 'created',
      studentId: studentInfo.id,
      status: status,
      newTotalPresentToday: totalPresentToday // Send updated count
    };

  } catch (error) {
    Logger.log(`RECORD_ATTENDANCE: CRITICAL ERROR - ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: 'An error occurred while recording attendance: ' + error.message };
  }
}

/**
 * Record attendance for multiple students at once to prevent race conditions
 */
function recordBulkAttendance(studentStatusMap, authToken) {
  const verificationResult = verifyJwt(authToken);
  if (!verificationResult.valid) {
    return { success: false, message: 'Authentication failed: ' + verificationResult.error, expired: verificationResult.expired || false };
  }
  const userMakingRequest = verificationResult.payload.user;
  Logger.log(`Bulk attendance record attempt by ${userMakingRequest.username} (Role: ${userMakingRequest.role}) for ${Object.keys(studentStatusMap).length} students`);

  try {
    const attendanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ATTENDANCE_SHEET_NAME);
    const studentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STUDENTS_SHEET_NAME);

    if (!attendanceSheet || !studentSheet) {
      Logger.log(`BULK_RECORD_ATTENDANCE: Error - Missing sheet. Attendance: ${!attendanceSheet}, Student: ${!studentSheet}`);
      return { success: false, message: 'Data sheet missing.' };
    }

    const now = new Date();
    const currentDate = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const currentTime = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');

    // Get student info
    const studentsData = studentSheet.getDataRange().getValues();
    const studentHeader = studentsData.shift(); // Remove header
    
    const studentIdColIdx = studentHeader.findIndex(h => h && (h.toString().trim() === 'รหัสนักเรียน' || h.toString().trim().toLowerCase() === 'studentid'));
    const studentFirstNameColIdx = studentHeader.findIndex(h => h && (h.toString().trim() === 'ชื่อ' || h.toString().trim().toLowerCase() === 'firstname'));
    const studentLastNameColIdx = studentHeader.findIndex(h => h && (h.toString().trim() === 'นามสกุล' || h.toString().trim().toLowerCase() === 'lastname'));
    const studentClassroomColIdx = studentHeader.findIndex(h => h && (
        h.toString().trim() === 'ห้องเรียนID' || 
        h.toString().trim() === 'ห้องเรียน' ||
        h.toString().trim().toLowerCase() === 'classroomid' ||
        h.toString().trim().toLowerCase() === 'classroom'
    ));

    // Debug logging for header detection
    Logger.log(`BULK_RECORD_ATTENDANCE DEBUG: Student headers found: ${JSON.stringify(studentHeader)}`);
    Logger.log(`BULK_RECORD_ATTENDANCE DEBUG: Column indices - ID: ${studentIdColIdx}, FirstName: ${studentFirstNameColIdx}, LastName: ${studentLastNameColIdx}, Classroom: ${studentClassroomColIdx}`);
    studentHeader.forEach((header, index) => {
        Logger.log(`BULK_RECORD_ATTENDANCE DEBUG: Header ${index}: "${header}" (trimmed: "${header.toString().trim()}")`);
    });

    if ([studentIdColIdx, studentFirstNameColIdx, studentLastNameColIdx, studentClassroomColIdx].some(idx => idx === -1)) {
        let missingHeadersDetails = [];
        if (studentIdColIdx === -1) missingHeadersDetails.push("คอลัมน์สำหรับ 'รหัสนักเรียน' (Student ID)");
        if (studentFirstNameColIdx === -1) missingHeadersDetails.push("คอลัมน์สำหรับ 'ชื่อ' (First Name)");
        if (studentLastNameColIdx === -1) missingHeadersDetails.push("คอลัมน์สำหรับ 'นามสกุล' (Last Name)");        if (studentClassroomColIdx === -1) missingHeadersDetails.push("คอลัมน์สำหรับ 'ห้องเรียนID' หรือ 'ห้องเรียน' (Classroom ID)");
        
        const actualHeadersString = studentHeader ? studentHeader.map((h, idx) => `คอลัมน์ ${idx + 1}: "${h}"`).join(', ') : "ไม่พบข้อมูลส่วนหัว";
        const errorMessage = `พบปัญหาโครงสร้างชีตนักเรียน: ไม่พบส่วนหัวที่จำเป็น: ${missingHeadersDetails.join('; ')}. กรุณาตรวจสอบส่วนหัวในชีต 'Students'. ส่วนหัวที่คาดหวังเช่น ['รหัสนักเรียน', 'ชื่อ', 'นามสกุล', 'ห้องเรียนID' หรือ 'ห้องเรียน']. ส่วนหัวที่พบจริง: [${actualHeadersString}].`;
        Logger.log(`BULK_RECORD_ATTENDANCE: ${errorMessage}`);
        return { success: false, message: errorMessage, errorDetails: { expected: ['รหัสนักเรียน', 'ชื่อ', 'นามสกุล', 'ห้องเรียนID'], actual: studentHeader || [] } };
    }
    
    let studentInfo = null;
    for (let i = 0; i < studentsData.length; i++) { // Iterate over data rows (header already removed)
      if (studentsData[i][studentIdColIdx] && studentsData[i][studentIdColIdx].toString().trim() === studentId.toString().trim()) {
        studentInfo = {
          id: studentsData[i][studentIdColIdx],
          name: `${studentsData[i][studentFirstNameColIdx]} ${studentsData[i][studentLastNameColIdx]}`,
          classroom: studentsData[i][studentClassroomColIdx]
        };
        break;
      }
    }

    if (!studentInfo) {
      Logger.log(`BULK_RECORD_ATTENDANCE: Student with ID ${studentId} not found.`);
      return { success: false, message: 'Student not found.' };
    }

    // Remove existing records for this date and these students to avoid duplicates
    const existingData = attendanceSheet.getDataRange().getValues();
    existingData.shift(); // Remove header
    const rowsToDelete = [];
    
    existingData.forEach((row, index) => {
      if (row[studentIdColIdx] && row[studentIdColIdx].toString().trim() === studentId.toString().trim()) {
        rowsToDelete.push(index + 2); // +2 because sheet rows are 1-indexed and we removed header
      }
    });    // Delete rows in reverse order to maintain correct indices
    rowsToDelete.reverse().forEach(rowIndex => {
      attendanceSheet.deleteRow(rowIndex);
    });

    const currentTimestamp = new Date();
    const timeString = Utilities.formatDate(currentTimestamp, Session.getScriptTimeZone(), 'HH:mm:ss');

    // Add new attendance records
    const recordsToAdd = [];
    Object.keys(studentStatusMap).forEach(studentId => {
      const status = studentStatusMap[studentId];
      const studentInfo = studentMap[studentId];
      
      if (studentInfo) {
        const record = new Array(studentHeader.length).fill(''); // Create empty array matching header length
        
        record[dateColIdx] = currentDate;
        record[studentIdColIdx] = studentId;
        record[firstNameColIdx] = studentInfo.firstName;
        record[lastNameColIdx] = studentInfo.lastName;
        record[statusColIdx] = status;
        if (timeColIdx !== -1) record[timeColIdx] = timeString;
        if (recordedByColIdx !== -1) record[recordedByColIdx] = userMakingRequest.username;
        if (classroomColIdx !== -1 && studentInfo.classroom) record[classroomColIdx] = studentInfo.classroom;
        
        recordsToAdd.push(record);
      }
    });
    
    if (recordsToAdd.length > 0) {
      attendanceSheet.getRange(attendanceSheet.getLastRow() + 1, 1, recordsToAdd.length, recordsToAdd[0].length).setValues(recordsToAdd);
    }
    
    Logger.log(`BULK_RECORD_ATTENDANCE: Successfully recorded ${recordsToAdd.length} attendance records`);
    return { 
      success: true, 
      message: `บันทึกการเข้าเรียนสำเร็จ ${recordsToAdd.length} รายการ`
    };

  } catch (error) {
    Logger.log(`BULK_RECORD_ATTENDANCE: CRITICAL ERROR - ${error.toString()} Stack: ${error.stack}`);
    return { success: false, message: 'An error occurred while recording bulk attendance: ' + error.message };
  }
}

/**
 * Record bulk attendance for multiple students with custom date
 * @param {Object} studentStatusMap - Object with studentId as key and status as value
 * @param {string} attendanceDate - Date in YYYY-MM-DD format
 * @param {string} authToken - JWT authentication token
 * @returns {Object} Result object with success status and message
 */
function recordBulkAttendanceWithDate(studentStatusMap, attendanceDate, authToken) {
  const verificationResult = verifyJwt(authToken);
  if (!verificationResult.valid) {
    return { 
      success: false, 
      message: 'Authentication failed: ' + verificationResult.error, 
      expired: verificationResult.expired || false 
    };
  }
  
  try {
    // Parse date from string if needed
    let attendanceDateObj;
    try {
      attendanceDateObj = new Date(attendanceDate);
      if (isNaN(attendanceDateObj.getTime())) {
        throw new Error('Invalid date format');
      }
    } catch (e) {
      return {
        success: false,
        message: 'Invalid date format: ' + e.message
      };
    }
    
    const userMakingRequest = verificationResult.payload.user;
    Logger.log(`User ${userMakingRequest.username} recording bulk attendance for date: ${attendanceDate}`);
    
    if (!studentStatusMap || Object.keys(studentStatusMap).length === 0) {
      return {
        success: false,
        message: 'No attendance data provided'
      };
    }
    
    // Get sheets
    const attendanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ATTENDANCE_SHEET_NAME);
    const studentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STUDENTS_SHEET_NAME);
    
    if (!attendanceSheet || !studentSheet) {
      return {
        success: false,
        message: 'Required sheets not found'
      };
    }
    
    // Get attendance headers
    const attendanceHeaders = attendanceSheet.getRange(1, 1, 1, attendanceSheet.getLastColumn()).getValues()[0];
    const dateColIdx = attendanceHeaders.indexOf('Date');
    const studentIdColIdx = attendanceHeaders.indexOf('StudentID');
    const firstNameColIdx = attendanceHeaders.indexOf('FirstName');
    const lastNameColIdx = attendanceHeaders.indexOf('LastName');
    const statusColIdx = attendanceHeaders.indexOf('Status');
    const timeColIdx = attendanceHeaders.indexOf('Time');
    const recordedByColIdx = attendanceHeaders.indexOf('RecordedBy');
    const classroomColIdx = attendanceHeaders.indexOf('Classroom');
    
    // Validate headers
    const requiredCols = [dateColIdx, studentIdColIdx, statusColIdx];
    if (requiredCols.some(idx => idx === -1)) {
      return {
        success: false,
        message: 'Attendance sheet is missing required columns'
      };
    }
    
    // Get student data to populate names
    const studentsData = studentSheet.getDataRange().getValues();
    const studentHeader = studentsData.shift(); // Remove header
    
    const studentIdColIdx_students = studentHeader.findIndex(h => h && (h.toString().trim() === 'รหัสนักเรียน' || h.toString().trim().toLowerCase() === 'studentid'));
    const studentFirstNameColIdx = studentHeader.findIndex(h => h && (h.toString().trim() === 'ชื่อ' || h.toString().trim().toLowerCase() === 'firstname'));
    const studentLastNameColIdx = studentHeader.findIndex(h => h && (h.toString().trim() === 'นามสกุล' || h.toString().trim().toLowerCase() === 'lastname'));
    const studentClassroomColIdx = studentHeader.findIndex(h => h && (
        h.toString().trim() === 'ห้องเรียนID' || 
        h.toString().trim() === 'ห้องเรียน' ||
        h.toString().trim().toLowerCase() === 'classroomid' ||
        h.toString().trim().toLowerCase() === 'classroom'
    ));
    
    if ([studentIdColIdx_students, studentFirstNameColIdx, studentLastNameColIdx].some(idx => idx === -1)) {
      return {
        success: false,
        message: 'Student sheet is missing required columns'
      };
    }
    
    // Build student map for easy lookup
    const studentMap = {};
    let classroomOfCurrentStudents = null;
    
    // First, identify which classroom these students belong to
    for (let studentId in studentStatusMap) {
      // Find this student in the student sheet to get their classroom
      for (let i = 0; i < studentsData.length; i++) {
        if (studentsData[i][studentIdColIdx_students] && 
            studentsData[i][studentIdColIdx_students].toString() === studentId.toString()) {
          classroomOfCurrentStudents = studentClassroomColIdx !== -1 ? 
            studentsData[i][studentClassroomColIdx] : '';
          break;
        }
      }
      if (classroomOfCurrentStudents) break; // Stop once we've found the classroom
    }
    
    Logger.log(`Identified classroom for this batch: ${classroomOfCurrentStudents}`);
    
    // Continue building the student map
    studentsData.forEach(row => {
      if (row[studentIdColIdx_students]) {
        const studentId = row[studentIdColIdx_students].toString();
        const classroom = studentClassroomColIdx !== -1 ? row[studentClassroomColIdx] : '';
        
        studentMap[studentId] = {
          firstName: row[studentFirstNameColIdx],
          lastName: row[studentLastNameColIdx],
          classroom: classroom
        };
      }
    });
    
    // Delete ALL existing attendance records for this classroom and date
    const currentDate = Utilities.formatDate(attendanceDateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const existingData = attendanceSheet.getDataRange().getValues();
    existingData.shift(); // Remove header
    const rowsToDelete = [];
    
    if (!classroomOfCurrentStudents) {
      Logger.log("WARNING: Could not determine classroom for this batch of students. Cannot safely delete existing records.");
      return {
        success: false,
        message: 'Unable to determine the classroom for these students. Cannot update attendance.'
      };
    }
    
    // Find ALL records for this classroom on this date, not just for the students in studentStatusMap
    existingData.forEach((row, index) => {
      // Check if this record is for the same date
      if (row[dateColIdx] && 
          Utilities.formatDate(new Date(row[dateColIdx]), Session.getScriptTimeZone(), 'yyyy-MM-dd') === currentDate) {
        
        // Check if this record is for a student in the same classroom
        const studentId = row[studentIdColIdx];
        if (studentId && studentId.toString()) {
          // If we have classroom column in attendance, use that directly
          if (classroomColIdx !== -1 && row[classroomColIdx] === classroomOfCurrentStudents) {
            rowsToDelete.push(index + 2); // +2 because sheet rows are 1-indexed and we removed header
          } 
          // Otherwise, check if this student is in the target classroom according to student map
          else {
            const studentInfo = studentMap[studentId.toString()];
            if (studentInfo && studentInfo.classroom === classroomOfCurrentStudents) {
              rowsToDelete.push(index + 2);
            }
          }
        }
      }
    });

    Logger.log(`Found ${rowsToDelete.length} existing attendance records to delete for classroom ${classroomOfCurrentStudents} on ${currentDate}`);

    // Delete rows in reverse order to maintain correct indices
    if (rowsToDelete.length > 0) {
      rowsToDelete.sort((a, b) => b - a); // Sort in descending order
      rowsToDelete.forEach(rowIndex => {
        attendanceSheet.deleteRow(rowIndex);
      });
      Logger.log(`Deleted ${rowsToDelete.length} existing attendance records for classroom ${classroomOfCurrentStudents}`);
    }

    const currentTimestamp = new Date();
    const timeString = Utilities.formatDate(currentTimestamp, Session.getScriptTimeZone(), 'HH:mm:ss');

    // Add new attendance records
    const recordsToAdd = [];
    Object.keys(studentStatusMap).forEach(studentId => {
      const status = studentStatusMap[studentId];
      const studentInfo = studentMap[studentId];
      
      if (studentInfo) {
        const record = new Array(attendanceSheet.getLastColumn()).fill(''); // Create empty array matching header length
        
        record[dateColIdx] = currentDate;
        record[studentIdColIdx] = studentId;
        record[firstNameColIdx] = studentInfo.firstName;
        record[lastNameColIdx] = studentInfo.lastName;
        record[statusColIdx] = status;
        if (timeColIdx !== -1) record[timeColIdx] = timeString;
        if (recordedByColIdx !== -1) record[recordedByColIdx] = userMakingRequest.username;
        if (classroomColIdx !== -1 && studentInfo.classroom) record[classroomColIdx] = studentInfo.classroom;
        
        recordsToAdd.push(record);
      }
    });
    
    if (recordsToAdd.length > 0) {
      attendanceSheet.getRange(attendanceSheet.getLastRow() + 1, 1, recordsToAdd.length, recordsToAdd[0].length).setValues(recordsToAdd);
    }
    
    Logger.log(`RECORD_BULK_ATTENDANCE_WITH_DATE: Successfully recorded ${recordsToAdd.length} attendance records for date ${currentDate}`);
    return { 
      success: true, 
      message: `บันทึกการเข้าเรียนสำเร็จ ${recordsToAdd.length} รายการ สำหรับวันที่ ${currentDate}`,
      recordsCount: recordsToAdd.length,
      date: currentDate
    };

  } catch (error) {
    Logger.log(`Error in recordBulkAttendanceWithDate: ${error.message}. Stack: ${error.stack}`);
    return { success: false, message: 'Error recording attendance: ' + error.message };
  }
}

/**
 * Validates and fixes student sheet data to ensure it's in sync with classroom data
 * - Ensures all students have valid classroom IDs
 * - Fixes student data when classroom IDs have changed
 * - Reports any issues found during validation
 * @returns {Object} Result object with success status and message
 */
function validateAndFixStudentSheetData() {
  try {
    Logger.log('Starting validation of student sheet data...');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const studentsSheet = ss.getSheetByName(STUDENTS_SHEET_NAME);
    const classroomsSheet = ss.getSheetByName(CLASSROOMS_SHEET_NAME);
    
    if (!studentsSheet || !classroomsSheet) {
      return {
        success: false,
        message: 'Missing Student or Classroom sheets',
        studentsFixed: false
      };
    }
    
    // Get data from sheets
    const studentsData = studentsSheet.getDataRange().getValues();
    const classroomsData = classroomsSheet.getDataRange().getValues();
    
    if (studentsData.length <= 1 || classroomsData.length <= 1) {
      // Only headers or empty
      return {
        success: true, 
        message: 'No data to validate',
        studentsFixed: false
      };
    }
      // Extract header rows
    const studentsHeader = studentsData[0];
    const classroomsHeader = classroomsData[0];
    
    // Find index of classroom ID column in students sheet
    const classroomIdColIndex = studentsHeader.findIndex(header => 
      header && (header.toString().trim() === 'ห้องเรียนID' || 
                header.toString().trim() === 'ห้องเรียน' ||
                header.toString().trim().toLowerCase() === 'classroomid' ||
                header.toString().trim().toLowerCase() === 'classroom' ||
                header.toString().trim().toLowerCase() === 'classroom id')
    );
    
    if (classroomIdColIndex === -1) {
      Logger.log('Warning: Could not find classroom ID column in students sheet');
      return {
        success: false,
        message: 'Could not find classroom ID column in students sheet',
        studentsFixed: false
      };
    }
    
    // Create a set of valid classroom IDs from classrooms sheet (first column should be ID)
    const validClassroomIds = new Set();
    for (let i = 1; i < classroomsData.length; i++) {
      if (classroomsData[i][0]) {
        validClassroomIds.add(classroomsData[i][0].toString().trim());
      }
    }
    
    Logger.log(`Found ${validClassroomIds.size} valid classroom IDs: ${Array.from(validClassroomIds).join(', ')}`);
    
    // Track students with invalid classroom IDs
    let studentsWithInvalidIds = 0;
    let studentsWithEmptyIds = 0;
    let studentsFixed = false;
    
    // Check each student and fix if necessary
    for (let i = 1; i < studentsData.length; i++) {
      const classroomId = studentsData[i][classroomIdColIndex];
      
      // Skip empty rows
      if (!studentsData[i][0]) continue;
      
      if (!classroomId) {
        studentsWithEmptyIds++;
        Logger.log(`Student at row ${i+1} has empty classroom ID`);
        
        // We could set a default classroom ID here if needed
        if (validClassroomIds.size > 0) {
          const defaultClassroomId = Array.from(validClassroomIds)[0];
          studentsSheet.getRange(i + 1, classroomIdColIndex + 1).setValue(defaultClassroomId);
          Logger.log(`Fixed student at row ${i+1} by setting default classroom ID: ${defaultClassroomId}`);
          studentsFixed = true;
        }
      } 
      else if (!validClassroomIds.has(classroomId.toString().trim())) {
        studentsWithInvalidIds++;
        Logger.log(`Student at row ${i+1} has invalid classroom ID: ${classroomId}`);
        
        // Fix invalid classroom ID by setting a valid one if available
        if (validClassroomIds.size > 0) {
          const defaultClassroomId = Array.from(validClassroomIds)[0];
          studentsSheet.getRange(i + 1, classroomIdColIndex + 1).setValue(defaultClassroomId);
          Logger.log(`Fixed student at row ${i+1} by replacing invalid ID ${classroomId} with valid ID: ${defaultClassroomId}`);
          studentsFixed = true;
        }
      }
    }
    
    const totalStudents = studentsData.length - 1;
    const validStudents = totalStudents - studentsWithInvalidIds - studentsWithEmptyIds;
    
    let resultMessage = `Validated ${totalStudents} students: ${validStudents} valid, ${studentsWithInvalidIds} with invalid classroom IDs, ${studentsWithEmptyIds} with empty classroom IDs.`;
    if (studentsFixed) {
      resultMessage += ' Some student records were automatically fixed.';
    }
    
    return {
      success: true,
      message: resultMessage,
      studentsFixed: studentsFixed,
      totalStudents: totalStudents,
      validStudents: validStudents,
      studentsWithInvalidIds: studentsWithInvalidIds,
      studentsWithEmptyIds: studentsWithEmptyIds
    };
    
  } catch (error) {
    Logger.log('Error validating student sheet data: ' + error.message);
    return {
      success: false,
      message: 'Error during validation: ' + error.message,
      studentsFixed: false
    };
  }
}

// =====================
// DASHBOARD AND STATISTICS FUNCTIONS
// =====================

/**
 * Gets daily attendance statistics for the specified number of days
 * @param {number} days - Number of days to include in stats (defaults to 7)
 * @param {string} token - JWT authentication token
 * @return {object} Response containing daily attendance statistics
 */
function getDailyAttendanceStats(days = 7, token) {
  try {
    // Verify JWT token
    const verification = verifyJwt(token);
    if (!verification.valid) {
      return { 
        success: false, 
        error: verification.error || 'Authentication failed', 
        expired: verification.expired || false 
      };
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const attendanceSheet = spreadsheet.getSheetByName(ATTENDANCE_SHEET_NAME);
    
    if (!attendanceSheet) {
      return { success: false, error: 'Attendance sheet not found' };
    }
    
    // Get all attendance data
    const data = attendanceSheet.getDataRange().getValues();
    const headers = data[0];
      // Find column indices
    const dateColIdx = headers.indexOf('Date');
    const statusColIdx = headers.indexOf('Status');
    
    if (dateColIdx === -1 || statusColIdx === -1) {
      return { success: false, error: 'Required columns not found in attendance sheet' };
    }
    
    // Calculate date range (today - days)
    const today = new Date();
    const startDate = new Date(today);
    startDate.setDate(startDate.getDate() - days);
    
    // Initialize daily stats object
    const dailyStats = [];
    for (let i = 0; i <= days; i++) {
      const date = new Date(today);
      date.setDate(date.getDate() - i);
      const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      dailyStats.push({
        date: dateStr,
        total: 0,
        present: 0,
        absent: 0,
        percentage: 0
      });
    }
    
    // Calculate statistics for each day
    data.slice(1).forEach(row => {
      if (row.length <= Math.max(dateColIdx, statusColIdx)) return;
      
      let recordDate = row[dateColIdx];
      if (typeof recordDate === 'string') {
        try {
          recordDate = new Date(recordDate);
        } catch(e) {
          return; // Skip invalid dates
        }
      }
      
      if (!(recordDate instanceof Date) || isNaN(recordDate.getTime())) return;
      
      if (recordDate >= startDate && recordDate <= today) {
        const dateStr = Utilities.formatDate(recordDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        const status = row[statusColIdx];
        
        const dayStat = dailyStats.find(stat => stat.date === dateStr);
        if (dayStat) {
          dayStat.total++;
          if (status === 'present' || status === 'มา') {
            dayStat.present++;
          } else if (status === 'absent' || status === 'ขาด') {
            dayStat.absent++;
          }
        }
      }
    });
    
    // Calculate percentages
    dailyStats.forEach(stat => {
      if (stat.total > 0) {
        stat.percentage = Math.round((stat.present / stat.total) * 100);
      }
    });
    
    // Get total students for today's percentage calculation
    const studentsSheet = spreadsheet.getSheetByName(STUDENTS_SHEET_NAME);
    const totalStudents = studentsSheet ? studentsSheet.getLastRow() - 1 : 0;
    
    // Find today's stats
    const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const todayStat = dailyStats.find(stat => stat.date === todayStr);
    const totalPresentToday = todayStat ? todayStat.present : 0;
    
    return {
      success: true,
      totalStudents: totalStudents,
      totalPresentToday: totalPresentToday,
      dailyStats: dailyStats
    };
  } catch (error) {
    Logger.log('ERROR in getDailyAttendanceStats: ' + error.toString());
    return {
      success: false,
      error: 'Failed to get daily attendance stats: ' + error.toString()
    };
  }
}

/**
 * Gets attendance records for a specific date
 * @param {string} date - Date in YYYY-MM-DD format
 * @param {string} token - JWT authentication token
 * @return {object} Response containing attendance records
 */
function getAttendanceByDate(date, token) {
  try {
    // Verify JWT token
    const verification = verifyJwt(token);
    if (!verification.valid) {
      return { 
        success: false, 
        error: verification.error || 'Authentication failed', 
        expired: verification.expired || false 
      };
    }
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const attendanceSheet = spreadsheet.getSheetByName(ATTENDANCE_SHEET_NAME);
    
    if (!attendanceSheet) {
      return { success: false, error: 'Attendance sheet not found' };
    }
    
    // Get all attendance data
    const data = attendanceSheet.getDataRange().getValues();
    const headers = data[0];
      // Find column indices
    const dateColIdx = headers.indexOf('Date');
    const studentIdColIdx = headers.indexOf('StudentID');
    const statusColIdx = headers.indexOf('Status');
    
    if (dateColIdx === -1 || studentIdColIdx === -1 || statusColIdx === -1) {
      return { success: false, error: 'Required columns not found in attendance sheet' };
    }
    
    // Filter attendance records for the specified date
    const attendance = [];
    data.slice(1).forEach(row => {
      if (row.length <= Math.max(dateColIdx, studentIdColIdx, statusColIdx)) return;
      
      let recordDate = row[dateColIdx];
      if (typeof recordDate === 'string') {
        try {
          recordDate = new Date(recordDate);
        } catch(e) {
          return; // Skip invalid dates
        }
      }
      
      if (!(recordDate instanceof Date) || isNaN(recordDate.getTime())) return;
      
      const recordDateStr = Utilities.formatDate(recordDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      
      if (recordDateStr === date) {
        attendance.push({
          studentId: row[studentIdColIdx],
          status: row[statusColIdx],
          date: recordDateStr
        });
      }
    });
    
    return {
      success: true,
      attendance: attendance,
      date: date
    };
  } catch (error) {
    Logger.log('ERROR in getAttendanceByDate: ' + error.toString());
    return {
      success: false,
      error: 'Failed to get attendance data: ' + error.toString()
    };
  }
}

/**
 * Gets all classroom data (ID and Name)
 * @param {string} authToken - JWT authentication token
 * @return {object} Response containing classroom data
 */
function getClassrooms(authToken) {
  const verificationResult = verifyJwt(authToken);
  if (!verificationResult.valid) {
    return { 
      success: false, 
      error: 'Authentication failed: ' + verificationResult.error, 
      expired: verificationResult.expired || false,
      classrooms: [] 
    };
  }
  
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const classroomSheet = spreadsheet.getSheetByName(CLASSROOMS_SHEET_NAME);
    
    if (!classroomSheet) {
      return { success: false, error: 'Classrooms sheet not found', classrooms: [] };
    }
    
    const data = classroomSheet.getDataRange().getValues();
    data.shift(); // Remove header row
    
    const classrooms = data.map(row => ({
      id: row[0],      // ห้องเรียนID
      name: row[1]     // ชื่อห้องเรียน
    }));
    
    return { success: true, classrooms: classrooms };
  } catch (error) {
    Logger.log('ERROR in getClassrooms: ' + error.toString());
    return {
      success: false,
      error: 'Failed to get classrooms: ' + error.toString(),
      classrooms: []
    };
  }
}

/**
 * Updates attendance sheet headers from Thai to English if needed
 */
function updateAttendanceSheetHeaders() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const attendanceSheet = spreadsheet.getSheetByName(ATTENDANCE_SHEET_NAME);
    
    if (!attendanceSheet) {
      Logger.log('updateAttendanceSheetHeaders: Attendance sheet not found');
      return { success: false, error: 'Attendance sheet not found' };
    }
    
    // Get current headers
    const currentHeaders = attendanceSheet.getRange(1, 1, 1, attendanceSheet.getLastColumn()).getValues()[0];
    
    // Check if headers need updating (if they're in Thai)
    const needsUpdate = currentHeaders.includes('วันที่') || 
                       currentHeaders.includes('รหัสนักเรียน') || 
                       currentHeaders.includes('สถานะ');
    
    if (needsUpdate) {
      Logger.log('updateAttendanceSheetHeaders: Updating headers from Thai to English');
      
      // Create mapping for header translation
      const headerMapping = {
        'วันที่': 'Date',
        'เวลา': 'Time', 
        'รหัสนักเรียน': 'StudentID',
        'ชื่อ-นามสกุล': 'FirstName',
        'ชื่อ': 'FirstName',
        'นามสกุล': 'LastName',
        'ห้องเรียน': 'Classroom',
        'ห้องเรียนID': 'Classroom',
        'สถานะ': 'Status'
      };
      
      // Update headers
      const newHeaders = currentHeaders.map(header => {
        const headerStr = header ? header.toString().trim() : '';
        return headerMapping[headerStr] || headerStr;
      });
      
      // Add missing required headers if they don't exist
      const requiredHeaders = ['Date', 'Time', 'StudentID', 'FirstName', 'LastName', 'Classroom', 'Status', 'RecordedBy'];
      
      requiredHeaders.forEach(requiredHeader => {
        if (!newHeaders.includes(requiredHeader)) {
          newHeaders.push(requiredHeader);
        }
      });
      
      // Update the sheet headers
      attendanceSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
      
      // Format headers
      const headerRange = attendanceSheet.getRange(1, 1, 1, newHeaders.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#BAFFC9');
      headerRange.setHorizontalAlignment('center');
      
      Logger.log('updateAttendanceSheetHeaders: Headers updated successfully');
      return { success: true, message: 'Headers updated successfully', headersUpdated: true };
    } else {
      Logger.log('updateAttendanceSheetHeaders: Headers are already in English format');
      return { success: true, message: 'Headers are already correct', headersUpdated: false };
    }
    
  } catch (error) {
    Logger.log('ERROR in updateAttendanceSheetHeaders: ' + error.toString());
    return { success: false, error: 'Failed to update headers: ' + error.toString() };
  }
}
