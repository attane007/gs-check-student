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
    
    Logger.log('initializeSheets: About to get spreadsheet ID.');
    const id = spreadsheet.getId();
    Logger.log('initializeSheets: spreadsheet.getId() succeeded: ' + id);

    const returnValue = {
      success: true,
      spreadsheetId: id,
      message: 'เตรียม Sheets เรียบร้อยแล้ว'
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
const JWT_EXPIRATION_SECONDS = 3600; // 1 hour, adjust as needed

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
    return { success: false, message: 'Authentication failed: ' + verificationResult.error };
  }
  const userMakingRequest = verificationResult.payload.user;
  Logger.log(`Attendance record attempt by ${userMakingRequest.username} (Role: ${userMakingRequest.role})`);

  // Optional: Add role-based check here if needed
  // if (userMakingRequest.role !== 'teacher' && userMakingRequest.role !== 'admin') {
  //   return { success: false, message: 'Permission denied. Insufficient role.' };
  // }

  const attendanceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ATTENDANCE_SHEET_NAME);
  // ... (rest of the recordAttendance logic remains similar)
  const studentSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STUDENTS_SHEET_NAME);
  
  const studentsData = studentSheet.getDataRange().getValues();
  let studentInfo = null;
  for (let i = 1; i < studentsData.length; i++) {
    if (studentsData[i][0].toString() === studentId.toString()) {
      studentInfo = {
        id: studentsData[i][0],
        name: `${studentsData[i][1]} ${studentsData[i][2]}`,
        classroom: studentsData[i][3]
      };
      break;
    }
  }
  
  if (!studentInfo) {
    return { success: false, message: 'Student not found.' };
  }
  
  const now = new Date();
  const date = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const time = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
  
  attendanceSheet.appendRow([
    date,
    time,
    studentInfo.id,
    studentInfo.name,
    studentInfo.classroom,
    status
  ]);
  
  return { success: true, message: `Attendance recorded for ${studentInfo.name} (${status}) by ${userMakingRequest.username}` };
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
    }
    const headers = data.shift(); // Remove header row and get headers
    // Assuming student ID is col 0, firstName col 1, lastName col 2, classroomId col 3
    // Find index of 'ห้องเรียนID'
    const classroomIdColIndex = headers.findIndex(header => header.toString().trim() === 'ห้องเรียนID');
    if (classroomIdColIndex === -1) {
        Logger.log('Header "ห้องเรียนID" not found in Students sheet.');
        return { success: false, error: 'Students sheet is not configured correctly (missing ห้องเรียนID header).', students: [] };
    }

    const students = data
      .filter(row => row[classroomIdColIndex] && row[classroomIdColIndex].toString().trim() === classroomId.toString().trim())
      .map(row => ({ 
          id: row[0],          // รหัสนักเรียน
          firstName: row[1],   // ชื่อ
          lastName: row[2]     // นามสกุล
          // classroom: row[3] // classroom ID, not needed by current frontend renderStudent
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
    }

    const attendanceData = attendanceSheet.getDataRange().getValues();
    const classroomData = classroomSheet.getDataRange().getValues();
    const studentData = studentSheet.getDataRange().getValues();

    attendanceData.shift(); // Remove header
    classroomData.shift(); // Remove header
    studentData.shift(); // Remove header

    const classrooms = classroomData.map(row => ({ id: row[0], name: row[1] }));
    const students = studentData.map(row => ({ id: row[0], name: `${row[1]} ${row[2]}`, classroomId: row[3] }));

    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    let filterStartDate = dateFromString ? new Date(dateFromString) : null;
    let filterEndDate = dateToString ? new Date(dateToString) : null;

    if (filterStartDate && isNaN(filterStartDate.getTime())) filterStartDate = null;
    if (filterEndDate && isNaN(filterEndDate.getTime())) filterEndDate = null;

    // Adjust end date to include the whole day
    if (filterEndDate) {
      filterEndDate.setHours(23, 59, 59, 999);
    }

    const stats = {
      totalPresentToday: 0,
      totalAbsentToday: 0,
      overallAttendanceRateToday: 0,
      overallClassroomBreakdown: [],
      topAttendees: []
    };
    const todayClassroomStatsMap = new Map();
    const overallClassroomStatsMap = new Map();
    const studentAttendanceCount = new Map();

    attendanceData.forEach(row => {
      const recordDateStr = row[0]; // Assuming date is in first column as string 'yyyy-MM-dd'
      const recordDate = new Date(recordDateStr);
      const studentId = row[2];
      const status = row[5] ? row[5].toString().toLowerCase() : '';
      const classroomNameFromAttendance = row[4]; // Classroom name as recorded in attendance

      // Today's overall stats
      if (recordDateStr === today) {
        if (status === 'present' || status === 'late') {
          stats.totalPresentToday++;
        } else if (status === 'absent') {
          stats.totalAbsentToday++;
        }
      }

      // Today's classroom stats
      if (recordDateStr === today) {
        let classroomTodayStat = todayClassroomStatsMap.get(classroomNameFromAttendance);
        if (!classroomTodayStat) {
          classroomTodayStat = { classroomName: classroomNameFromAttendance, present: 0, absent: 0, total: 0 };
          todayClassroomStatsMap.set(classroomNameFromAttendance, classroomTodayStat);
        }
        if (status === 'present' || status === 'late') {
          classroomTodayStat.present++;
        } else if (status === 'absent') {
          classroomTodayStat.absent++;
        }
        if (status === 'present' || status === 'late' || status === 'absent') {
            classroomTodayStat.total++;
        }
      }

      // Date range filtering for overall classroom breakdown and top attendees
      let isInDateRange = true;
      if (filterStartDate && recordDate < filterStartDate) isInDateRange = false;
      if (filterEndDate && recordDate > filterEndDate) isInDateRange = false;

      if (isInDateRange) {
        // Overall classroom breakdown for selected range
        let classroomOverallStat = overallClassroomStatsMap.get(classroomNameFromAttendance);
        if (!classroomOverallStat) {
          classroomOverallStat = { name: classroomNameFromAttendance, present: 0, absent: 0, total: 0 };
          overallClassroomStatsMap.set(classroomNameFromAttendance, classroomOverallStat);
        }
        if (status === 'present' || status === 'late') {
          classroomOverallStat.present++;
        } else if (status === 'absent') {
          classroomOverallStat.absent++;
        }
         if (status === 'present' || status === 'late' || status === 'absent') {
            classroomOverallStat.total++;
        }

        // Top attendees for selected range
        if (status === 'present' || status === 'late') {
          studentAttendanceCount.set(studentId, (studentAttendanceCount.get(studentId) || 0) + 1);
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
      cs.rate = cs.total > 0 ? (cs.present / cs.total) * 100 : 0;
      todayClassroomStats.push(cs);
    });

    // Format overall classroom breakdown for selected range
    overallClassroomStatsMap.forEach(cs => {
      cs.rate = cs.total > 0 ? (cs.present / cs.total) * 100 : 0;
      stats.overallClassroomBreakdown.push(cs);
    });

    // Format top attendees
    const sortedTopAttendees = Array.from(studentAttendanceCount.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10) // Top 10
      .map(([studentId, daysPresent]) => {
        const studentDetail = students.find(s => s.id.toString() === studentId.toString());
        return { studentName: studentDetail ? studentDetail.name : 'Unknown Student', daysPresent };
      });
    stats.topAttendees = sortedTopAttendees;

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
    }
    const attendanceDataRange = attendanceSheet.getDataRange();
    const attendanceData = attendanceDataRange.getValues();
    
    if (attendanceData.length <= 1) { // Only header or empty
        Logger.log('No attendance data found to process for daily stats.');
        return { success: true, dailyStats: [] }; // No data to process
    }
    attendanceData.shift(); // Remove header

    const dailyStatsMap = new Map();
    const today = new Date();
    today.setHours(0, 0, 0, 0); // Normalize today to the start of the day

    // Initialize map for the last 'days'
    for (let i = 0; i < days; i++) {
      const targetDate = new Date(today);
      targetDate.setDate(today.getDate() - i);
      const dateString = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      dailyStatsMap.set(dateString, { date: dateString, present: 0, absent: 0 });
    }

    // Populate stats from attendance data
    attendanceData.forEach(row => {
      if (!row || row.length < 6) { // Basic check for valid row structure
          Logger.log('Skipping invalid row in attendance data: ' + JSON.stringify(row));
          return; 
      }
      const recordDateStr = row[0] ? Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : null;
      const status = row[5] ? row[5].toString().toLowerCase() : '';

      if (recordDateStr && dailyStatsMap.has(recordDateStr)) {
        let stat = dailyStatsMap.get(recordDateStr);
        if (status === 'present' || status === 'late') {
          stat.present++;
        } else if (status === 'absent') {
          stat.absent++;
        }
        // No need to dailyStatsMap.set(recordDateStr, stat) as stat is a reference
      }
    });

    const dailyStats = Array.from(dailyStatsMap.values()).sort((a,b) => new Date(a.date) - new Date(b.date)); // Sort by date ascending

    return { success: true, dailyStats: dailyStats };

  } catch (error) {
    Logger.log('Error in getDailyAttendanceStats: ' + error.message + ' Stack: ' + error.stack);
    return { success: false, error: 'Error processing daily attendance stats: ' + error.message, dailyStats: [] };
  }
}

// getCurrentUserRoleClient is replaced by getUserDataFromToken or similar client-side logic
// Remove: function getUserRoleClient(sessionId) { ... }
// Remove: function getCurrentUser() { ... }

// Ensure initializeUsersSheet and addInitialAdminUser are still present and correct
// ... (initializeUsersSheet, addInitialAdminUser, hashPassword, customBytesToHex, generateSalt, verifyUser, addUser functions should remain largely unchanged as they deal with user persistence, not session/auth token type)
