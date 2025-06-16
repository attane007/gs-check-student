/**
 * ระบบเช็คชื่อนักเรียน - Google Apps Script Backend
 * เชื่อมต่อกับ Google Sheets เพื่อจัดเก็บข้อมูลการเข้าเรียน
 */

// ใช้ Spreadsheet ที่ผูกกับ Script นี้
const STUDENTS_SHEET_NAME = 'Students';
const CLASSROOMS_SHEET_NAME = 'Classrooms';
const ATTENDANCE_SHEET_NAME = 'Attendance';
const USERS_SHEET_NAME = 'Users'; // New sheet for users

/**
 * ฟังก์ชันเริ่มต้นสำหรับสร้าง Sheets หากยังไม่มี
 */
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
    }
    const attendanceLastRow = attendanceSheet.getLastRow();
    if (attendanceLastRow === 0 || attendanceSheet.getRange(1, 1).getValue() === '') {
      attendanceSheet.getRange(1, 1, 1, 6).setValues([
        ['วันที่', 'เวลา', 'รหัสนักเรียน', 'ชื่อ-นามสกุล', 'ห้องเรียน', 'สถานะ']
      ]);
      const headerRange = attendanceSheet.getRange(1, 1, 1, 6);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#BAFFC9');
      headerRange.setHorizontalAlignment('center');
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

// --- User and Auth Utilities ---
function generateSalt() {
  return Math.random().toString(36).substring(2, 15) + Math.random().toString(36).substring(2, 15);
}

// Custom function to convert byte array to hex string
function customBytesToHex(bytes) {
  if (!bytes) return null;
  let hex = '';
  for (let i = 0; i < bytes.length; i++) {
    let byte = bytes[i] & 0xFF; // Ensure byte is positive
    let hexByte = byte.toString(16);
    if (hexByte.length < 2) {
      hexByte = '0' + hexByte;
    }
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
    
    // Use the custom bytesToHex function as a workaround
    const hex = customBytesToHex(digest);
    if (hex) {
        Logger.log('HASH_PASSWORD: customBytesToHex conversion successful.');
        return hex;
    } else {
        Logger.log('HASH_PASSWORD: CRITICAL - customBytesToHex failed. Returning null.');
        throw new Error('customBytesToHex failed to convert digest.');
    }  } catch (error) {
    Logger.log(`HASH_PASSWORD: ERROR during hashing process - ${error.toString()} Stack: ${error.stack}`);
    throw error; 
  }
}

// Simple test function to isolate Utilities object behavior
function testUtilitiesObject() {
  Logger.log('TEST_UTILITIES: Starting test.');
  try {
    Logger.log(`TEST_UTILITIES: typeof Utilities = ${typeof Utilities}`);
    if (Utilities) {
      Logger.log(`TEST_UTILITIES: typeof Utilities.bytesToHex (direct access) = ${typeof Utilities.bytesToHex}`);
      const testBytes = [1, 2, 3, 4, 5, 10, 15, 16, 255];
      if (typeof Utilities.bytesToHex === 'function') {
        Logger.log(`TEST_UTILITIES: Utilities.bytesToHex IS a function. Test conversion: ${Utilities.bytesToHex(testBytes)}`);
      } else {
        Logger.log('TEST_UTILITIES: Utilities.bytesToHex is NOT a function.');
      }
    } else {
      Logger.log('TEST_UTILITIES: Global Utilities object is not defined.');
    }
  } catch (e) {
    Logger.log(`TEST_UTILITIES: Error during test: ${e.message} Stack: ${e.stack}`);
  }
  Logger.log('TEST_UTILITIES: Test finished.');
}

// --- JWT Configuration ---
const JWT_SECRET_KEY_PROPERTY = 'asdasdlglkbmkbtokb;ltmblmdfdfb';
const JWT_EXPIRATION_SECONDS = 30 * 24 * 60 * 60; // 30 days (30 days × 24 hours × 60 minutes × 60 seconds)

function getJwtSecret() {
  let secret = PropertiesService.getScriptProperties().getProperty(JWT_SECRET_KEY_PROPERTY);
  if (!secret) {
    secret = Utilities.getUuid() + Utilities.getUuid(); // Generate a strong random secret
    PropertiesService.getScriptProperties().setProperty(JWT_SECRET_KEY_PROPERTY, secret);
    Logger.log('JWT_SECRET_KEY generated and stored.');
  }
  return secret;
}

// --- Base64 URL Encoding/Decoding Helpers ---
function base64UrlEncode(input) {
  let base64;
  if (typeof input === 'string') {
    // Input is a string, encode it
    base64 = Utilities.base64Encode(input, Utilities.Charset.UTF_8);
  } else if (Array.isArray(input)) {
    // Input is a byte array (from computeHmacSha256Signature, etc.)
    base64 = Utilities.base64EncodeWebSafe(input);
    // Remove padding '=' since we'll add it back in decode if needed
    base64 = base64.replace(/=+$/, '');
    return base64;  // Already web-safe
  } else {
    // Handle other cases (like Blob) if needed
    throw new Error('base64UrlEncode: Unsupported input type');
  }
  // Replace standard base64 chars with URL-safe versions
  return base64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
}

function base64UrlDecode(str) {
  // Restore base64 standard characters
  str = str.replace(/-/g, '+').replace(/_/g, '/');
  // Add padding if needed
  while (str.length % 4) {
    str += '=';
  }
  try {
    // Decode the string
    const decoded = Utilities.base64Decode(str);
    return Utilities.newBlob(decoded).getDataAsString();
  } catch (e) {
    Logger.log('BASE64_URL_DECODE_ERROR: ' + e.message);
    throw new Error('Failed to decode base64url: ' + e.message);
  }
}

// --- JWT Generation ---
function generateJwt(userInfo) {
  const secret = getJwtSecret();
  const header = {
    alg: 'HS256',
    typ: 'JWT'
  };
  const now = Math.floor(Date.now() / 1000);
  const payload = {
    user: { // Only include necessary, non-sensitive user info
      username: userInfo.username,
      role: userInfo.role,
      fullName: userInfo.fullName
    },
    iat: now, // Issued at
    exp: now + JWT_EXPIRATION_SECONDS, // Expiration time
    iss: ScriptApp.getService().getUrl() // Issuer (this script)
  };

  const encodedHeader = base64UrlEncode(JSON.stringify(header));
  const encodedPayload = base64UrlEncode(JSON.stringify(payload));
  const signatureInput = encodedHeader + '.' + encodedPayload;
  const signatureBytes = Utilities.computeHmacSha256Signature(signatureInput, secret, Utilities.Charset.UTF_8);
  const encodedSignature = base64UrlEncode(signatureBytes);

  return encodedHeader + '.' + encodedPayload + '.' + encodedSignature;
}

// --- JWT Verification ---
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
  
  // Optional: Check issuer (iss) if needed
  // if (payload.iss !== ScriptApp.getService().getUrl()) {
  //   Logger.log('VERIFY_JWT: Issuer mismatch.');
  //   return { valid: false, error: 'Invalid issuer' };
  // }

  Logger.log('VERIFY_JWT: Token verified successfully for user: ' + payload.user.username);
  return { valid: true, payload: payload };
}

// Function to verify user credentials against the Users sheet
function verifyUser(username, password) {
  try {
    if (!username || !password) {
      Logger.log("VERIFY_USER: Username or password missing");
      return null;
    }
    
    Logger.log(`VERIFY_USER: Attempting to verify user: ${username}`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const usersSheet = ss.getSheetByName(USERS_SHEET_NAME);
    
    if (!usersSheet) {
      Logger.log("VERIFY_USER: Users sheet not found");
      return null;
    }
    
    const data = usersSheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log("VERIFY_USER: No user data found (only header or empty sheet)");
      return null;
    }
    
    // Get header row to find column indexes
    const headers = data[0].map(header => String(header).toLowerCase().trim());
    const usernameColIndex = headers.indexOf("username");
    const passwordHashColIndex = headers.indexOf("passwordhash");
    const saltColIndex = headers.indexOf("salt");
    const roleColIndex = headers.indexOf("role");
    const fullNameColIndex = headers.indexOf("fullname");
    
    if (usernameColIndex === -1 || passwordHashColIndex === -1 || saltColIndex === -1) {
      Logger.log("VERIFY_USER: Required column headers not found in Users sheet");
      return null;
    }
    
    // Look for matching username
    let userRow = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][usernameColIndex] && data[i][usernameColIndex].toString().toLowerCase() === username.toLowerCase()) {
        userRow = data[i];
        break;
      }
    }
    
    if (!userRow) {
      Logger.log(`VERIFY_USER: User ${username} not found`);
      return null;
    }
    
    const storedHash = userRow[passwordHashColIndex];
    const salt = userRow[saltColIndex];
    
    if (!storedHash || !salt) {
      Logger.log(`VERIFY_USER: Stored hash or salt missing for user ${username}`);
      return null;
    }
    
    // Hash the provided password with the stored salt
    const providedPasswordHash = hashPassword(password, salt);
    
    if (providedPasswordHash === storedHash) {
      Logger.log(`VERIFY_USER: Authentication successful for ${username}`);
      
      // Update last login timestamp
      const lastLoginColIndex = headers.indexOf("lastlogin");
      if (lastLoginColIndex !== -1) {
        usersSheet.getRange(data.indexOf(userRow) + 1, lastLoginColIndex + 1).setValue(new Date());
      }
      
      // Return user info for JWT
      return {
        username: userRow[usernameColIndex],
        role: roleColIndex !== -1 ? userRow[roleColIndex] : "user", // Default to "user" if role not found
        fullName: fullNameColIndex !== -1 ? userRow[fullNameColIndex] : username // Default to username if fullName not found
      };
    } else {
      Logger.log(`VERIFY_USER: Password verification failed for ${username}`);
      return null;
    }
  } catch (error) {
    Logger.log(`VERIFY_USER ERROR: ${error.message} Stack: ${error.stack}`);
    return null;
  }
}

// New function to be called from login.html
function loginUser(username, password) {
  try {
    Logger.log(`Attempting login for user: ${username}`);
    const userInfo = verifyUser(username, password); // verifyUser still uses the Users sheet
    if (userInfo) {
      const token = generateJwt(userInfo);
      Logger.log(`Login successful for ${username}. JWT generated.`);
      // The client will receive this token and store it.
      // Redirection will be handled by the client.
      return { 
        success: true, 
        message: 'Login successful!', 
        token: token, // Send JWT to client
        user: { // Send basic user info for immediate use by client
            username: userInfo.username,
            role: userInfo.role,
            fullName: userInfo.fullName
        }
      };
    } else {
      Logger.log(`Login failed for user: ${username}`);
      return { success: false, message: 'Invalid username or password.' };
    }
  } catch (error) {
    Logger.log(`Error during loginUser: ${error.message} Stack: ${error.stack}`);
    console.error('Login error:', error);
    return { success: false, message: 'An error occurred during login: ' + error.message };
  }
}

// --- Web App Endpoints ---
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
      title = 'Student Attendance';
  } else if (pageToServe === 'dashboard') {
      template = HtmlService.createTemplateFromFile('dashboard');
      template.appUrl = ScriptApp.getService().getUrl();
      // User data and access control will be handled by client-side JS
      title = 'Dashboard - Student Attendance';
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
    });

    return { 
      success: true, 
      message: `Attendance ${recordUpdated ? 'updated' : 'recorded'} for ${studentInfo.name} (${status}) by ${userMakingRequest.username}`,
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

// New function to get classrooms
function getClassrooms(authToken) {
  const verificationResult = verifyJwt(authToken);
  if (!verificationResult.valid) {
    return { success: false, error: 'Authentication failed: ' + verificationResult.error, classrooms: [], expired: verificationResult.expired || false };
  }
  Logger.log(`User ${verificationResult.payload.user.username} requesting classrooms.`);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CLASSROOMS_SHEET_NAME);
    if (!sheet) {
        Logger.log('Classrooms sheet not found.');
        return { success: false, error: 'Classrooms data source not found.', classrooms: [] };
    }
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) { // Only header or empty
        return { success: true, classrooms: [] }; // No classrooms data
    }
    data.shift(); // Remove header row
    const classrooms = data.map(row => ({ id: row[0], name: row[1] }));
    return { success: true, classrooms: classrooms };
  } catch (error) {
    Logger.log('Error fetching classrooms: ' + error.message);
    return { success: false, error: 'Error fetching classrooms: ' + error.message, classrooms: [] };
  }
}

// New function to get students by classroom
function getStudentsByClassroom(classroomId, authToken) {
  const verificationResult = verifyJwt(authToken);
  if (!verificationResult.valid) {
    return { success: false, error: 'Authentication failed: ' + verificationResult.error, students: [], expired: verificationResult.expired || false };
  }
  Logger.log(`User ${verificationResult.payload.user.username} requesting students for classroom ID: ${classroomId}`);
  
  if (!classroomId) {
    return { success: false, error: 'Classroom ID not provided.', students: [] };
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STUDENTS_SHEET_NAME);
    if (!sheet) {
        Logger.log('Students sheet not found.');
        return { success: false, error: 'Students data source not found.', students: [] };
    }
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
        return { success: true, students: [] }; // No students data
    }    const headers = data.shift(); // Remove header row and get headers
    
    // Assuming student ID is col 0, firstName col 1, lastName col 2, classroomId col 3
    // Find index of classroom column, which might be labeled 'ห้องเรียนID' or similar
    const classroomIdColIndex = headers.findIndex(header => 
        header && (
            header.toString().trim() === 'ห้องเรียนID' || 
            header.toString().trim() === 'ห้องเรียน' ||
            header.toString().trim().toLowerCase() === 'classroomid' ||
            header.toString().trim().toLowerCase() === 'classroom id' ||
            header.toString().trim().toLowerCase() === 'classroom'
        )
    );
    
    // If not found at known labels, assume it's the 4th column (index 3) as per initialization pattern
    if (classroomIdColIndex === -1) {
        Logger.log('Warning: Could not find exact classroom ID header. Using column index 3 (4th column) based on initialization pattern.');
        
        // Check if we have at least 4 columns
        if (headers.length < 4) {
            Logger.log('Error: Students sheet does not have enough columns. Expected at least 4 columns.');
            return { success: false, error: 'Students sheet does not have enough columns. Expected at least 4 columns.', students: [] };
        }
        
        // Use column 3 (4th column) as the classroom ID column
        const students = data
            .filter(row => row.length > 3 && row[3] && row[3].toString().trim() === classroomId.toString().trim())
            .map(row => ({ 
                id: row[0],          // รหัสนักเรียน
                firstName: row[1],   // ชื่อ
                lastName: row[2]     // นามสกุล
            }));        return { success: true, students: students };
    }
    
    // Use the found classroom ID column index
    const students = data
      .filter(row => row[classroomIdColIndex] && row[classroomIdColIndex].toString().trim() === classroomId.toString().trim())
      .map(row => ({ 
          id: row[0],          // รหัสนักเรียน
          firstName: row[1],   // ชื่อ
          lastName: row[2]     // นามสกุล
          // classroom: row[classroomIdColIndex] // classroom ID, not needed by current frontend renderStudent
      }));
    return { success: true, students: students };
  } catch (error) {
    Logger.log('Error fetching students by classroom: ' + error.message);
    return { success: false, error: 'Error fetching students: ' + error.message, students: [] };
  }
}

// Function to get dashboard data
function getDashboardData(dateFromString, dateToString, authToken) {
  const verificationResult = verifyJwt(authToken);
  if (!verificationResult.valid) {
    return { success: false, error: 'Authentication failed: ' + verificationResult.error, expired: verificationResult.expired || false };
  }
  Logger.log(`User ${verificationResult.payload.user.username} requesting dashboard data. From: ${dateFromString}, To: ${dateToString}`);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const attendanceSheet = ss.getSheetByName(ATTENDANCE_SHEET_NAME);
    const classroomSheet = ss.getSheetByName(CLASSROOMS_SHEET_NAME);
    const studentSheet = ss.getSheetByName(STUDENTS_SHEET_NAME);

    if (!attendanceSheet || !classroomSheet || !studentSheet) {
      return { success: false, error: 'One or more data sheets are missing.' };
    }    // ข้อมูลจากตาราง
    const attendanceData = attendanceSheet.getDataRange().getValues();
    const classroomData = classroomSheet.getDataRange().getValues();
    const studentData = studentSheet.getDataRange().getValues();

    // ข้อมูล header ของแต่ละตาราง
    const attendanceHeader = attendanceData.shift(); // Remove and get header
    const classroomHeader = classroomData.shift(); // Remove and get header
    const studentHeader = studentData.shift(); // Remove and get header

    Logger.log('Dashboard - Processing sheet data with headers: Students: ' + JSON.stringify(studentHeader) + 
              ', Attendance: ' + JSON.stringify(attendanceHeader));

    // ค้นหา index ของคอลัมน์ที่สำคัญใน sheet Students - ไม่ควรแฮร์ดโค้ด
    // Find student ID column (ควรจะอยู่คอลัมน์แรกเสมอ แต่ตรวจสอบเพื่อความแน่นอน)
    const studentIDColIndex = studentHeader.findIndex(header => 
        header && (header.toString().trim() === 'รหัสนักเรียน' || 
                  header.toString().trim().toLowerCase() === 'student id' ||
                  header.toString().trim().toLowerCase() === 'studentid')
    );
    const actualStudentIDColIndex = studentIDColIndex !== -1 ? studentIDColIndex : 0;
    
    // Find first name column
    const studentFirstNameColIndex = studentHeader.findIndex(header => 
        header && (header.toString().trim() === 'ชื่อ' || 
                  header.toString().trim().toLowerCase() === 'first name' ||
                  header.toString().trim().toLowerCase() === 'firstname' ||
                  header.toString().trim().toLowerCase() === 'name')
    );
    const actualStudentFirstNameColIndex = studentFirstNameColIndex !== -1 ? studentFirstNameColIndex : 1;
    
    // Find last name column
    const studentLastNameColIndex = studentHeader.findIndex(header => 
        header && (header.toString().trim() === 'นามสกุล' || 
                  header.toString().trim().toLowerCase() === 'last name' ||
                  header.toString().trim().toLowerCase() === 'lastname' ||
                  header.toString().trim().toLowerCase() === 'surname')
    );
    const actualStudentLastNameColIndex = studentLastNameColIndex !== -1 ? studentLastNameColIndex : 2;
    
    // หา index ของคอลัมน์ห้องเรียนใน sheet Students (อาจเป็น index ที่ 3 หรือชื่ออื่น)
    const studentClassroomColIndex = studentHeader.findIndex(header => 
        header && (
            header.toString().trim() === 'ห้องเรียนID' || 
            header.toString().trim().toLowerCase() === 'classroomid' ||
            header.toString().trim() === 'ห้องเรียน' ||
            header.toString().trim().toLowerCase() === 'classroom' ||
            header.toString().trim().toLowerCase() === 'classroom id'
        )
    );
    
    const actualStudentClassroomColIndex = studentClassroomColIndex !== -1 ? studentClassroomColIndex : 3;
    Logger.log(`Dashboard: Student Classroom column found at index ${actualStudentClassroomColIndex}`);    // ค้นหา index ของคอลัมน์ที่สำคัญใน sheet Attendance
    // ส่วนใหญ่จะมีโครงสร้างคล้ายกัน แต่เพื่อความแน่นอนเราจะค้นหา index ที่ถูกต้อง
    
    // Find date column
    const attendanceDateColIndex = attendanceHeader.findIndex(header => 
        header && (header.toString().trim() === 'วันที่' || 
                  header.toString().trim().toLowerCase() === 'date')
    );
    const actualAttendanceDateColIndex = attendanceDateColIndex !== -1 ? attendanceDateColIndex : 0;
    
    // Find student ID column
    const attendanceStudentIDColIndex = attendanceHeader.findIndex(header => 
        header && (header.toString().trim() === 'รหัสนักเรียน' || 
                  header.toString().trim().toLowerCase() === 'student id' ||
                  header.toString().trim().toLowerCase() === 'studentid')
    );
    const actualAttendanceStudentIDColIndex = attendanceStudentIDColIndex !== -1 ? attendanceStudentIDColIndex : 2;
    
    // Find classroom column
    const attendanceClassroomColIndex = attendanceHeader.findIndex(header => 
        header && (header.toString().trim() === 'ห้องเรียน' || 
                  header.toString().trim().toLowerCase() === 'classroom' ||
                  header.toString().trim() === 'ห้องเรียนชื่อ' ||
                  header.toString().trim().toLowerCase() === 'classroom name')
    );
    const actualAttendanceClassroomColIndex = attendanceClassroomColIndex !== -1 ? attendanceClassroomColIndex : 4;
    
    // Find status column
    const attendanceStatusColIndex = attendanceHeader.findIndex(header => 
        header && (header.toString().trim() === 'สถานะ' || 
                  header.toString().trim().toLowerCase() === 'status' ||
                  header.toString().trim() === 'การเข้าเรียน' ||
                  header.toString().trim().toLowerCase() === 'attendance')
    );
    const actualAttendanceStatusColIndex = attendanceStatusColIndex !== -1 ? attendanceStatusColIndex : 5;
    
    // Log detected column indexes
    Logger.log(`Dashboard - Detected column indexes for Attendance: 
      Date: ${actualAttendanceDateColIndex}, 
      StudentID: ${actualAttendanceStudentIDColIndex}, 
      Classroom: ${actualAttendanceClassroomColIndex}, 
      Status: ${actualAttendanceStatusColIndex}`);
      
    Logger.log(`Dashboard - Detected column indexes for Students: 
      ID: ${actualStudentIDColIndex}, 
      FirstName: ${actualStudentFirstNameColIndex}, 
      LastName: ${actualStudentLastNameColIndex}, 
      ClassroomID: ${actualStudentClassroomColIndex}`);    // แปลงข้อมูลจากตาราง - ตรวจสอบข้อมูลที่อ่านมา
    const classroomIdIndex = 0;  // expected to be first column
    const classroomNameIndex = classroomHeader.findIndex(header => 
        header && (header.toString().trim() === 'ชื่อห้องเรียน' || 
                  header.toString().trim().toLowerCase() === 'classroom name' ||
                  header.toString().trim() === 'ชื่อ' || 
                  header.toString().trim().toLowerCase() === 'name')
    );
    const actualClassroomNameIndex = classroomNameIndex !== -1 ? classroomNameIndex : 1; // default to second column
    
    Logger.log(`Dashboard - Classroom sheet: Using name column at index ${actualClassroomNameIndex}`);

    // Map classroom data with detected column indexes
    const classrooms = classroomData.map(row => ({ 
        id: row[classroomIdIndex], 
        name: row[actualClassroomNameIndex] 
    }));
    
    // Create a lookup map for classroom names
    const classroomNameMap = new Map();
    classrooms.forEach(classroom => {
      classroomNameMap.set(classroom.id.toString().trim(), classroom.name);
    });
    
    // Map student data with detected column indexes
    const students = studentData.map(row => { 
      const classroomId = row[actualStudentClassroomColIndex];
      return {
        id: row[actualStudentIDColIndex], 
        name: `${row[actualStudentFirstNameColIndex]} ${row[actualStudentLastNameColIndex]}`, 
        classroomId: classroomId,
        classroomName: classroomNameMap.get(classroomId ? classroomId.toString().trim() : '') || 'ไม่ระบุห้องเรียน'
      };
    });

    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    let filterStartDate = dateFromString ? new Date(dateFromString) : null;
    let filterEndDate = dateToString ? new Date(dateToString) : null;

    if (filterStartDate && isNaN(filterStartDate.getTime())) filterStartDate = null;
    if (filterEndDate && isNaN(filterEndDate.getTime())) filterEndDate = null;

    // Adjust end date to include the whole day
    if (filterEndDate) {
      filterEndDate.setHours(23, 59, 59, 999);
    }
    
    // Log actual data counts
    Logger.log(`Dashboard - Data counts: Classrooms: ${classrooms.length}, Students: ${students.length}, Attendance rows: ${attendanceData.length}`);
    if (classrooms.length === 0 || students.length === 0) {
      Logger.log('Warning: One or more critical data sets are empty! This may cause incorrect statistics.');
    }    const stats = {
      totalPresentToday: 0,
      totalAbsentToday: 0,
      totalLateToday: 0,
      totalExcusedToday: 0, // เพิ่มสถานะ "ลา"
      overallAttendanceRateToday: 0,
      overallClassroomBreakdown: [],
      topAttendees: []
    };
    const todayClassroomStatsMap = new Map();
    const overallClassroomStatsMap = new Map();
    const studentAttendanceCount = new Map();    attendanceData.forEach(row => {
      // Use the detected column indexes
      const recordDateStr = row[actualAttendanceDateColIndex]; 
      let recordDate;
      
      // Ensure date is properly formatted
      if (recordDateStr instanceof Date) {
        recordDate = recordDateStr;
      } else {
        try {
          recordDate = new Date(recordDateStr);
        } catch (e) {
          Logger.log(`Error parsing date: ${recordDateStr}`);
          return; // Skip this row if date can't be parsed
        }
      }
      
      const recordDateFormatted = Utilities.formatDate(recordDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      
      // Access other fields with appropriate column indexes
      const studentId = row[actualAttendanceStudentIDColIndex];
      const status = row[actualAttendanceStatusColIndex] ? row[actualAttendanceStatusColIndex].toString().toLowerCase() : '';
      let classroomId = row[actualAttendanceClassroomColIndex]; // This might be either ID or name depending on format
      
      // Find the proper classroom name
      let classroomName;
      
      // Try to determine if the value is an ID or already a name by checking our mapping
      if (classroomId && classroomNameMap.has(classroomId.toString().trim())) {
        classroomName = classroomNameMap.get(classroomId.toString().trim());
      } else {
        // If not found in mapping, the value might already be a name or an unknown ID
        classroomName = classroomId || 'ไม่ระบุห้องเรียน';
      }      // Today's overall stats - use formatted date for comparison
      const isToday = recordDateFormatted === today;
      if (isToday) {
        // Check for various present status formats
        if (status === 'present' || status === 'มา') {
          stats.totalPresentToday++;
        } 
        // Check for late status
        else if (status === 'late' || status === 'สาย') {
          stats.totalLateToday++;
          stats.totalPresentToday++; // Count late as present for overall stats
        }
        // Check for excused status (ลา)
        else if (status === 'excused' || status === 'ลา') {
          stats.totalExcusedToday++;
        }
        // Check for various absent status formats
        else if (status === 'absent' || status === 'ขาด') {
          stats.totalAbsentToday++;
        }
      }// Today's classroom stats
      if (isToday && classroomName) {
        let classroomTodayStat = todayClassroomStatsMap.get(classroomName);        if (!classroomTodayStat) {
          classroomTodayStat = { 
            classroomName: classroomName, 
            present: 0, 
            late: 0,
            excused: 0, // เพิ่มสถานะ "ลา"
            absent: 0, 
            total: 0 
          };
          todayClassroomStatsMap.set(classroomName, classroomTodayStat);
        }
        
        // Check for various status formats with separated counts
        if (status === 'present' || status === 'มา') {
          classroomTodayStat.present++;
        } 
        else if (status === 'late' || status === 'สาย') {
          classroomTodayStat.late++;
        }
        else if (status === 'excused' || status === 'ลา') {
          classroomTodayStat.excused++;
        }
        else if (status === 'absent' || status === 'ขาด') {
          classroomTodayStat.absent++;
        }
        
        // Count in total if any recognized status
        if (status === 'present' || status === 'late' || status === 'excused' || status === 'absent' || 
            status === 'มา' || status === 'สาย' || status === 'ลา' || status === 'ขาด') {
          classroomTodayStat.total++;
        }      }
      
      // Date range filtering for overall classroom breakdown and top attendees
      let isInDateRange = true;
      if (filterStartDate && recordDate < filterStartDate) isInDateRange = false;
      if (filterEndDate && recordDate > filterEndDate) isInDateRange = false;

      if (isInDateRange && classroomName) {        // Overall classroom breakdown for selected range
        let classroomOverallStat = overallClassroomStatsMap.get(classroomName);
        if (!classroomOverallStat) {
          classroomOverallStat = { 
            name: classroomName, 
            present: 0, 
            late: 0,
            excused: 0, // เพิ่มสถานะ "ลา"
            absent: 0, 
            total: 0 
          };
          overallClassroomStatsMap.set(classroomName, classroomOverallStat);
        }
        
        // Check for various status formats with separated counts
        if (status === 'present' || status === 'มา') {
          classroomOverallStat.present++;
          
          // Track top attendees for present statuses
          if (studentId) {
            studentAttendanceCount.set(studentId, (studentAttendanceCount.get(studentId) || 0) + 1);
          }
        } 
        else if (status === 'late' || status === 'สาย') {
          classroomOverallStat.late++;
          
          // Track top attendees for late statuses too
          if (studentId) {
            studentAttendanceCount.set(studentId, (studentAttendanceCount.get(studentId) || 0) + 1);
          }
        }
        else if (status === 'excused' || status === 'ลา') {
          classroomOverallStat.excused++;
        }
        else if (status === 'absent' || status === 'ขาด') {
          classroomOverallStat.absent++;
        }
        
        // Count in total if any recognized status
        if (status === 'present' || status === 'late' || status === 'excused' || status === 'absent' || 
            status === 'มา' || status === 'สาย' || status === 'ลา' || status === 'ขาด') {
          classroomOverallStat.total++;
        }
      }
    });

    // Calculate today's overall attendance rate
    const totalMarkedToday = stats.totalPresentToday + stats.totalAbsentToday;
    if (totalMarkedToday > 0) {
      stats.overallAttendanceRateToday = (stats.totalPresentToday / totalMarkedToday) * 100;
    }

    // Format today's classroom stats
    const todayClassroomStats = [];
    todayClassroomStatsMap.forEach(cs => {
      cs.rate = cs.total > 0 ? ((cs.present + cs.late) / cs.total) * 100 : 0;
      todayClassroomStats.push(cs);
    });

    // Format overall classroom breakdown for selected range
    overallClassroomStatsMap.forEach(cs => {
      cs.rate = cs.total > 0 ? ((cs.present + cs.late) / cs.total) * 100 : 0;
      stats.overallClassroomBreakdown.push(cs);
    });

    // Format top attendees with proper student names
    const sortedTopAttendees = Array.from(studentAttendanceCount.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10) // Top 10
      .map(([studentId, daysPresent]) => {
        const studentDetail = students.find(s => s.id.toString() === studentId.toString());
        const studentName = studentDetail ? studentDetail.name : 'Unknown Student';
        // Include classroom name if available
        const classroomName = studentDetail ? studentDetail.classroomName : '';
        
        return { 
          studentName: studentName, 
          daysPresent: daysPresent,
          classroomName: classroomName
        };
      });
    stats.topAttendees = sortedTopAttendees;

    // Include detailed logging for troubleshooting
    Logger.log('getDashboardData returning - Stats structure: ' + 
              JSON.stringify({
                totalPresent: stats.totalPresentToday,
                totalAbsent: stats.totalAbsentToday,
                totalLate: stats.totalLateToday,
                rate: stats.overallAttendanceRateToday,
                hasClassrooms: stats.overallClassroomBreakdown.length > 0,
                hasTopAttendees: stats.topAttendees.length > 0,
                todayStatsCount: todayClassroomStats.length
              }));
              
    return { 
        success: true, 
        stats: stats, 
        todayClassroomStats: todayClassroomStats 
    };

  } catch (error) {
    Logger.log('Error in getDashboardData: ' + error.message + ' Stack: ' + error.stack);
    return { success: false, error: 'Error processing dashboard data: ' + error.message };
  }
}

// Function to get daily attendance stats for the chart
function getDailyAttendanceStats(days, authToken) {
  const verificationResult = verifyJwt(authToken);
  if (!verificationResult.valid) {
    return { success: false, error: 'Authentication failed: ' + verificationResult.error, dailyStats: [], expired: verificationResult.expired || false };
  }
  Logger.log(`User ${verificationResult.payload.user.username} requesting daily attendance stats for last ${days} days.`);

  if (!days || isNaN(parseInt(days)) || parseInt(days) <= 0) {
    days = 7; // Default to 7 days
  }
  days = parseInt(days); // Ensure 'days' is an integer

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const attendanceSheet = ss.getSheetByName(ATTENDANCE_SHEET_NAME);
    if (!attendanceSheet) {
      Logger.log('Error in getDailyAttendanceStats: Attendance data sheet not found.');
      return { success: false, error: 'Attendance data sheet not found.', dailyStats: [] };
    }    const attendanceDataRange = attendanceSheet.getDataRange();
    const attendanceData = attendanceDataRange.getValues();
    
    if (attendanceData.length <= 1) { // Only header or empty
        Logger.log('No attendance data found to process for daily stats.');
        return { success: true, dailyStats: [] }; // No data to process
    }
    
    // Get header before removing it
    const attendanceHeader = attendanceData[0];
    attendanceData.shift(); // Remove header

    const dailyStatsMap = new Map();
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Normalize today to the start of the day    // Initialize map for the last 'days'
    for (let i = 0; i < days; i++) {
      const targetDate = new Date(today);
      targetDate.setDate(today.getDate() - i);
      const dateString = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      dailyStatsMap.set(dateString, { date: dateString, present: 0, late: 0, absent: 0, excused: 0 });    }

    // Get column indexes for attendance data using the header we saved
    // Find date column
    const dateDolIndex = attendanceHeader.findIndex(header => 
        header && (header.toString().trim() === 'วันที่' || 
                  header.toString().trim().toLowerCase() === 'date')
    );
    const actualDateColIndex = dateDolIndex !== -1 ? dateDolIndex : 0;
    
    // Find status column
    const statusColIndex = attendanceHeader.findIndex(header => 
        header && (header.toString().trim() === 'สถานะ' || 
                  header.toString().trim().toLowerCase() === 'status' ||
                  header.toString().trim() === 'การเข้าเรียน' ||
                  header.toString().trim().toLowerCase() === 'attendance')
    );
    const actualStatusColIndex = statusColIndex !== -1 ? statusColIndex : 5;
    
    Logger.log(`Daily stats - Using column indexes - Date: ${actualDateColIndex}, Status: ${actualStatusColIndex}`);
    Logger.log(`Daily stats - Header structure: ${JSON.stringify(attendanceHeader)}`);
    // Populate stats from attendance data
    attendanceData.forEach(row => {
      if (!row || row.length <= Math.max(actualDateColIndex, actualStatusColIndex)) { 
          Logger.log('Skipping invalid row in attendance data: ' + JSON.stringify(row));
          return; 
      }
      
      // Handle date properly
      let recordDate;
      const dateValue = row[actualDateColIndex];
      
      if (dateValue instanceof Date) {
        recordDate = dateValue;
      } else {
        try {
          recordDate = new Date(dateValue);
        } catch (e) {
          Logger.log(`Error parsing date in daily stats: ${dateValue}`);
          return; // Skip this row if date can't be parsed
        }
      }
        const recordDateStr = Utilities.formatDate(recordDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const status = row[actualStatusColIndex] ? row[actualStatusColIndex].toString().toLowerCase() : '';

      // Debug log for each status found
      if (status) {
        Logger.log(`Processing record: Date=${recordDateStr}, Status='${status}', Has date in map: ${dailyStatsMap.has(recordDateStr)}`);
      }if (recordDateStr && dailyStatsMap.has(recordDateStr)) {
        let stat = dailyStatsMap.get(recordDateStr);
          // Check for various present status formats
        if (status === 'present' || status === 'มา') {
          stat.present++;
          Logger.log(`✓ Incremented PRESENT for ${recordDateStr}: now ${stat.present}`);
        } 
        // Check for late status
        else if (status === 'late' || status === 'สาย') {
          stat.late++;
          Logger.log(`✓ Incremented LATE for ${recordDateStr}: now ${stat.late}`);
        }
        // Check for excused status (ลา) - separate from absent
        else if (status === 'excused' || status === 'ลา') {
          stat.excused++;
          Logger.log(`✓ Incremented EXCUSED for ${recordDateStr}: now ${stat.excused}`);
        }
        // Check for absent status (not including excused)
        else if (status === 'absent' || status === 'ขาด') {
          stat.absent++;
          Logger.log(`✓ Incremented ABSENT for ${recordDateStr}: now ${stat.absent}`);
        }
        else if (status) {
          Logger.log(`⚠️ Unknown status '${status}' for ${recordDateStr} - not counted`);
        }
        // No need to dailyStatsMap.set(recordDateStr, stat) as stat is a reference
      }
    });    const dailyStats = Array.from(dailyStatsMap.values()).sort((a,b) => new Date(a.date) - new Date(b.date)); // Sort by date ascending

    Logger.log(`getDailyAttendanceStats returning ${dailyStats.length} days of data. Sample: ${JSON.stringify(dailyStats.slice(0, 2))}`);
    Logger.log(`Full daily stats data structure: ${JSON.stringify(dailyStats)}`);
    
    // Log summary of all data
    dailyStats.forEach(day => {
      Logger.log(`📊 ${day.date}: Present=${day.present}, Late=${day.late}, Absent=${day.absent}, Excused=${day.excused}`);
    });
    
    return { 
      success: true, 
      dailyStats: dailyStats 
    };

  } catch (error) {
    Logger.log('Error in getDailyAttendanceStats: ' + error.message + ' Stack: ' + error.stack);
    return { success: false, error: 'Error processing daily attendance stats: ' + error.message, dailyStats: [] };
  }
}

// getCurrentUserRoleClient is replaced by getUserDataFromToken or similar client-side logic
// Remove: function getUserRoleClient(sessionId) { ... }
// Remove: function getCurrentUser() { ... }

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

// Function to get today's attendance records for a specific classroom
function getTodayAttendanceByClassroom(classroomId, authToken) {
  const verificationResult = verifyJwt(authToken);
  if (!verificationResult.valid) {
    return { success: false, error: 'Authentication failed: ' + verificationResult.error, attendanceRecords: {}, expired: verificationResult.expired || false };
  }
  Logger.log(`User ${verificationResult.payload.user.username} requesting today's attendance for classroom ID: ${classroomId}`);
  
  if (!classroomId) {
    return { success: false, error: 'Classroom ID not provided.', attendanceRecords: {} };
  }

  try {
    const attendanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ATTENDANCE_SHEET_NAME);
    const studentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STUDENTS_SHEET_NAME);
    
    if (!attendanceSheet || !studentSheet) {
      Logger.log(`getTodayAttendanceByClassroom: Missing sheet. Attendance: ${!attendanceSheet}, Student: ${!studentSheet}`);
      return { success: false, error: 'Required data sheet missing.', attendanceRecords: {} };
    }
    
    // Get today's date in the correct format
    const today = new Date();
    const todayString = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    Logger.log(`getTodayAttendanceByClassroom: Finding records for date ${todayString} and classroom ${classroomId}`);
    
    // Get column indices for Attendance sheet
    const attendanceData = attendanceSheet.getDataRange().getValues();
    if (attendanceData.length <= 1) {
      return { success: true, attendanceRecords: {} }; // No attendance data
    }
    
    const attendanceHeader = attendanceData.shift(); // Remove and get header row
    
    // Find required column indices in Attendance sheet
    const dateColIdx = attendanceHeader.findIndex(h => h && (h.toString().trim() === 'วันที่' || h.toString().trim().toLowerCase() === 'date'));
    const studentIdColIdx = attendanceHeader.findIndex(h => h && (h.toString().trim() === 'รหัสนักเรียน' || h.toString().trim().toLowerCase() === 'studentid'));
    const classroomColIdx = attendanceHeader.findIndex(h => h && (
        h.toString().trim() === 'ห้องเรียน' || 
        h.toString().trim().toLowerCase() === 'classroom' ||
        h.toString().trim() === 'ห้องเรียนID' ||
        h.toString().trim().toLowerCase() === 'classroomid'
    ));
    const statusColIdx = attendanceHeader.findIndex(h => h && (h.toString().trim() === 'สถานะ' || h.toString().trim().toLowerCase() === 'status'));
    
    // Ensure we have all required columns
    if ([dateColIdx, studentIdColIdx, classroomColIdx, statusColIdx].some(idx => idx === -1)) {
      let missingCols = [];
      if (dateColIdx === -1) missingCols.push("Date");
      if (studentIdColIdx === -1) missingCols.push("StudentID");
      if (classroomColIdx === -1) missingCols.push("Classroom");
      if (statusColIdx === -1) missingCols.push("Status");
      
      Logger.log(`getTodayAttendanceByClassroom: Missing columns in Attendance sheet: ${missingCols.join(', ')}`);
      return { success: false, error: `Attendance sheet structure error: missing ${missingCols.join(', ')}`, attendanceRecords: {} };
    }
    
    // Find students in this classroom
    const studentData = studentSheet.getDataRange().getValues();
    const studentHeader = studentData.shift();
    
    // Find classroom column index in Students sheet
    const studentClassroomColIdx = studentHeader.findIndex(h => h && (
        h.toString().trim() === 'ห้องเรียนID' || 
        h.toString().trim() === 'ห้องเรียน' || 
        h.toString().trim().toLowerCase() === 'classroomid' ||
        h.toString().trim().toLowerCase() === 'classroom'
    ));
    
    const studentIdStudentColIdx = studentHeader.findIndex(h => h && (
        h.toString().trim() === 'รหัสนักเรียน' || 
        h.toString().trim().toLowerCase() === 'studentid'
    ));
    
    if (studentClassroomColIdx === -1 || studentIdStudentColIdx === -1) {
      Logger.log(`getTodayAttendanceByClassroom: Students sheet missing required columns. ClassroomIdx: ${studentClassroomColIdx}, StudentIdIdx: ${studentIdStudentColIdx}`);
      return { success: false, error: 'Students sheet structure error', attendanceRecords: {} };
    }
    
    // Find all students in this classroom
    const classroomStudents = studentData
      .filter(row => row[studentClassroomColIdx] && row[studentClassroomColIdx].toString().trim() === classroomId.toString().trim())
      .map(row => row[studentIdStudentColIdx].toString().trim());
    
    Logger.log(`getTodayAttendanceByClassroom: Found ${classroomStudents.length} students in classroom ${classroomId}`);
    
    // Build attendance records for today
    const attendanceRecords = {};
    
    // Go through attendance records for today
    attendanceData.forEach(row => {
      // Handle case when date is a Date object vs string
      let recordDateStr;
      if (row[dateColIdx] instanceof Date) {
        recordDateStr = Utilities.formatDate(row[dateColIdx], Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        try {
          const recordDate = new Date(row[dateColIdx]);
          recordDateStr = Utilities.formatDate(recordDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } catch (e) {
          // Skip records with invalid dates
          return;
        }
      }
      
      // Check if record is for today
      if (recordDateStr === todayString) {
        const studentId = row[studentIdColIdx] ? row[studentIdColIdx].toString().trim() : '';
        let recordClassroom = row[classroomColIdx] ? row[classroomColIdx].toString().trim() : '';
        const status = row[statusColIdx] ? row[statusColIdx].toString().trim() : '';

        // If this student is in our target classroom or record's classroom matches
        if ((classroomStudents.includes(studentId)) || recordClassroom === classroomId.toString().trim()) {
          // Save this attendance record with studentId as key
          attendanceRecords[studentId] = status;
        }
      }
    });
    
    Logger.log(`getTodayAttendanceByClassroom: Returning ${Object.keys(attendanceRecords).length} attendance records for classroom ${classroomId}`);
    
    return {
      success: true,
      attendanceRecords: attendanceRecords
    };
    
  } catch (error) {
    Logger.log(`Error in getTodayAttendanceByClassroom: ${error.message}. Stack: ${error.stack}`);
    return { success: false, error: 'Error processing attendance data: ' + error.message, attendanceRecords: {} };
  }
}
