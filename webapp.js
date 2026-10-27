const SESSION_KEY = 'session_token';
const SESSION_TTL_MS = 7 * 24 * 60 * 60 * 1000; // 7 days, sliding (extends on every use)
let CURRENT_TOKEN_ = '';
let CURRENT_SESSION_ = null;

// Function to handle GET requests and route pages
function doGet(e) {
  e = e || { parameter: {} };

  const urlToken = (e.parameter.t || '').toString().trim();

  // Handle logout requests
  if (e.parameter.page === 'logout') {
    clearUserSession();
    if (urlToken) clearTokenSession_(urlToken);
    return servePage_('Login', 'Login');
  }

  // Restore session from URL token (survives refresh)
  if (urlToken) {
    const tokenSession = getTokenSession_(urlToken);
    if (tokenSession) {
      CURRENT_TOKEN_ = urlToken;
      CURRENT_SESSION_ = tokenSession;
    }
  }

  // Identity comes ONLY from the user's own URL token (no shared cache)
  const userData = CURRENT_SESSION_;

  // If there's no session, show the login page
  if (!userData) {
    return servePage_('Login', 'Login');
  }

  const userRole = userData.role;

  // Routing for users with the 'Doer' role
  if (userRole === 'Doer') {
    if (e.parameter.page === 'index' || !e.parameter.page) {
      return servePage_('DoerDashboard', 'Doer Dashboard');
    } else if (e.parameter.page === 'doer') {
      return servePage_('DoerDashboard', 'Doer Dashboard');
    } else if (e.parameter.page === 'CompanyChart') {
      return servePage_('CompanyChart', 'Company Chart');
    } else if (e.parameter.page === 'scorecard') {
      return servePage_('ScoreCard', 'Score Card');
    } else if (e.parameter.page === 'Expense') {
      return servePage_('Expense', 'Expense');
    } else if (e.parameter.page === 'VendorMangement') {
      return servePage_('VendorMangement', 'VendorMangement');
    } else if (e.parameter.page === 'MOM') {
      return serveMomPage_();
    } else if (e.parameter.page === 'checklist') {
      return servePage_('Checklist', 'Checklist');
    } else if (e.parameter.page === 'Help') {
      return servePage_('Help', 'Help');
    } else if (e.parameter.page === 'MISScoring') {
      return servePage_('MISScoring', 'MIS Scoring');
    }
  } else {
    // Routing for other user roles (Admin, etc.)
    if (e.parameter.page === 'doer') {
      return servePage_('DoerDashboard', 'Doer Dashboard');
    } else if (e.parameter.page === 'scorecard') {
      return servePage_('ScoreCard', 'Score Card');
    } else if (e.parameter.page === 'CompanyChart') {
      return servePage_('CompanyChart', 'Company Chart');
    } else if (e.parameter.page === 'checklist') {
      return servePage_('Checklist', 'Checklist');
    } else if (e.parameter.page === 'Dashboard') {
      return servePage_('Dashboard', 'Dashboard');
    } else if (e.parameter.page === 'Setting') {
      return servePage_('Setting', 'Setting');
    } else if (e.parameter.page === 'Expense') {
      return servePage_('Expense', 'Expense');
    } else if (e.parameter.page === 'VendorMangement') {
      return servePage_('VendorMangement', 'VendorMangement');
    } else if (e.parameter.page === 'MOM') {
      return serveMomPage_();
    } else if (e.parameter.page === 'Help') {
      return servePage_('Help', 'Help');
    } else if (e.parameter.page === 'MISScoring') {
      return servePage_('MISScoring', 'MIS Scoring');
    } else if (e.parameter.page === 'index' || !e.parameter.page) {
      return servePage_('Index', 'Delegation');
    }
  }

  // Default fallback routing based on role
  if (userRole === 'Doer') {
    return servePage_('DoerDashboard', 'Doer Dashboard');
  } else {
    return servePage_('Index', 'Delegation');
  }
}

function servePage_(fileName, title) {
  return HtmlService.createTemplateFromFile(fileName)
    .evaluate()
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://www.google.com/images/favicon.ico');
}

function serveMomPage_() {
  const template = HtmlService.createTemplateFromFile('MOM');
  template.projects = getProjects();
  template.raTeamMembers = getRATeamMembers();
  return template.evaluate()
    .setTitle('MOM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setFaviconUrl('https://www.google.com/images/favicon.ico');
}


function getNeedHelpCount(userName, userRole) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Master');
  const data = sheet.getDataRange().getValues(); // get all data
  const header = data[0];
  
  const statusCol = 9; // column J (0-indexed is 9 if counting from 0, but getValues is 2D array starting 0)
  const nameCol = 3;   // column D

  let count = 0;

  for (let i = 1; i < data.length; i++) { // skip header row
    const row = data[i];
    const status = row[statusCol];
    const name = row[nameCol];
    
    if (status === "Need Help") {
      if (userRole.toLowerCase() === "admin") {
        count++;
      } else if (userRole.toLowerCase() === "doer" && name === userName) {
        count++;
      }
    }
  }

  return count;
}


// Permanent session handling (UserCache + UserProperties)
function setUserSession(userData) {
  try {
    userData.timestamp = new Date().toISOString();
    const payload = JSON.stringify(userData);
    CacheService.getUserCache().put(SESSION_KEY, payload, 21600);
    PropertiesService.getUserProperties().setProperty(SESSION_KEY, payload);
    return true;
  } catch (error) {
    console.error('Error storing session:', error);
    return false;
  }
}

function getUserSession() {
  // Shared cache disabled – identity only from the user's own token
  return CURRENT_SESSION_ || null;
}

function getUserSessionLegacy_() {
  try {
    let payload = CacheService.getUserCache().get(SESSION_KEY);

    if (!payload) {
      payload = PropertiesService.getUserProperties().getProperty(SESSION_KEY);
      if (!payload) return null;

      const restored = JSON.parse(payload);
      if (restored.timestamp && (new Date() - new Date(restored.timestamp)) > SESSION_TTL_MS) {
        clearUserSession();
        return null;
      }
      CacheService.getUserCache().put(SESSION_KEY, payload, 21600); // re-warm cache
      return restored;
    }

    return JSON.parse(payload);
  } catch (error) {
    console.error('Error retrieving session:', error);
    return null;
  }
}

function clearUserSession() {
  try {
    CacheService.getUserCache().remove(SESSION_KEY);
    PropertiesService.getUserProperties().deleteProperty(SESSION_KEY);
    return true;
  } catch (error) {
    console.error('Error clearing session:', error);
    return false;
  }
}

// ── Token sessions (stored script-wide, keyed by random token) ──
function createTokenSession_(userData) {
  const token = Utilities.getUuid().replace(/-/g, '');
  const payload = JSON.stringify({
    email: userData.email, name: userData.name, role: userData.role, deviceId: userData.deviceId || '',
    token: token, timestamp: new Date().toISOString()
  });
  CacheService.getScriptCache().put('S_' + token, payload, 21600);
  PropertiesService.getScriptProperties().setProperty('S_' + token, payload);
  return token;
}

function getTokenSession_(token) {
  try {
    const key = 'S_' + token;
    let payload = CacheService.getScriptCache().get(key);
    if (!payload) {
      payload = PropertiesService.getScriptProperties().getProperty(key);
      if (!payload) return null;
      CacheService.getScriptCache().put(key, payload, 21600);
    }
    const s = JSON.parse(payload);
    const lastSeen = new Date(s.lastSeen || s.timestamp || 0);
    if ((new Date() - lastSeen) > SESSION_TTL_MS) {
      clearTokenSession_(token);
      return null;
    }

    // Sliding expiry: refresh "lastSeen" at most every 30 min
    if ((new Date() - lastSeen) > 30 * 60 * 1000) {
      s.lastSeen = new Date().toISOString();
      const updated = JSON.stringify(s);
      CacheService.getScriptCache().put(key, updated, 21600);
      PropertiesService.getScriptProperties().setProperty(key, updated);
    }
    return s;
  } catch (err) {
    return null;
  }
}

function clearTokenSession_(token) {
  try {
    CacheService.getScriptCache().remove('S_' + token);
    PropertiesService.getScriptProperties().deleteProperty('S_' + token);
  } catch (e) {}
}

// tokenPair = "token|deviceId" – session valid only on the device that logged in
function getSessionForCall_(tokenPair) {
  if (!tokenPair) return null;
  const parts = String(tokenPair).split('|');
  const s = getTokenSession_(parts[0]);
  if (!s || !s.deviceId || s.deviceId !== (parts[1] || '')) return null;
  return s;
}

// ── SSO to O2D FMS ──
const O2D_EXEC_URL = 'https://script.google.com/a/macros/akashenterprise.in/s/AKfycbynEHa78UJ6rZrJ2VrCIDMXvkLQbBiuwYVfU5Ev16lvUQSVLpCMF7Jhd5_TPwdbCzBhDw/exec';

// Run ONCE from the editor → copy the value from Logs into O2D's Script Properties
function generateSsoSecret() {
  const secret = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty('SSO_SECRET', secret);
  Logger.log('SSO_SECRET = ' + secret);
}

function b64url_(bytes) {
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}

// Called on click – returns a one-time O2D login URL for the logged-in doer
function getO2DSsoUrl(tokenPair) {
  const s = getSessionForCall_(tokenPair);
  if (!s) throw new Error('Session expired. Please login again.');

  const secret = PropertiesService.getScriptProperties().getProperty('SSO_SECRET') ||
    '26d62c3d5eb54346b1f0e80b617a0b10ea6d642ec380462091152c5d5a585b85';

  const payload = JSON.stringify({
    u: s.name,                          // doer name
    exp: Date.now() + 2 * 60 * 1000,    // valid 2 minutes
    n: Utilities.getUuid()              // one-time nonce
  });
  const p = b64url_(Utilities.newBlob(payload).getBytes());
  const sig = b64url_(Utilities.computeHmacSha256Signature(p, secret));
  return O2D_EXEC_URL + '?sso=' + encodeURIComponent(p + '.' + sig);
}

function checkSession(tokenPair) {
  return !!getSessionForCall_(tokenPair);
}

// Used by Login page to auto-resume a saved session on the SAME device
function resumeSession(tokenPair) {
  const s = getSessionForCall_(tokenPair);
  if (!s) return { success: false };
  return {
    success: true,
    name: s.name,
    role: s.role,
    token: String(tokenPair).split('|')[0]
  };
}

function getSessionTokenForPage() {
  if (CURRENT_TOKEN_) return CURRENT_TOKEN_;
  const s = getUserSession();
  return (s && s.token) ? s.token : '';
}

// Set a daily time-driven trigger on this to remove expired tokens
function cleanupExpiredSessions() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  Object.keys(all).forEach(k => {
    if (k.indexOf('S_') !== 0) return;
    try {
      const s = JSON.parse(all[k]);
      const seen = s.lastSeen || s.timestamp;
      if (!seen || (new Date() - new Date(seen)) > SESSION_TTL_MS) props.deleteProperty(k);
    } catch (e) { props.deleteProperty(k); }
  });
}


function authenticateUser(name, password, deviceId) {
  try {
    // Get the sheet named 'Doer List'
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Doer List');

    if (!sheet) {
      return {
        success: false,
        message: 'User database not found. Please contact administrator.'
      };
    }

    // Get all data from the sheet
    const data = sheet.getDataRange().getValues();
    
    // Clean up the input from the user (remove accidental spaces)
    const cleanInputName = name.toString().toLowerCase().trim();
    const cleanInputPassword = password.toString().trim();

    // Loop through rows, starting from the second row (index 1) to skip the header
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Column Mapping:
      const rowName = row[0];
      const rowPassword = row[1];
      const rowRole = row[2];
      const rowEmail = row[3];
      
      // Skip empty rows
      if (!rowName) continue;

      // Clean up the data from the Google Sheet (remove accidental spaces)
      const cleanSheetName = rowName.toString().toLowerCase().trim();
      const cleanSheetPassword = String(rowPassword).trim();

      // Check if the provided name and password match the current row perfectly
      if (cleanSheetName === cleanInputName && cleanSheetPassword === cleanInputPassword) {

        // User is found and authenticated. Create the session data object.
        const userData = {
          email: rowEmail,
          name: rowName.toString().trim(), // Save the clean version
          role: rowRole.toString().trim()
        };

        // Create URL token + store in session cache
        userData.deviceId = String(deviceId || '');
        if (!userData.deviceId) {
          return { success: false, message: 'Browser storage is blocked. Please allow site data/cookies and try again.' };
        }
        userData.token = createTokenSession_(userData);
        setUserSession(userData);

        return {
          success: true,
          userData: userData,
          role: rowRole,
          token: userData.token
        };
      }
    }

    // If the loop finishes without finding a match, the user is not found
    return {
      success: false,
      message: 'Invalid name or password. Please try again.'
    };

  } catch (error) {
    console.error('Error in authenticateUser:', error);
    return {
      success: false,
      message: 'Authentication error: ' + error.toString()
    };
  }
}




function getUserEmails() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Doer List');

    if (!sheet) {
      throw new Error('Doer List sheet not found');
    }

    // Get all data from the sheet
    const data = sheet.getDataRange().getValues();
    const emails = [];

    // Loop through rows starting from row 2 (skip header)
    for (let i = 1; i < data.length; i++) {
      const email = data[i][0]; // Column D (index 3) contains emails
      if (email && email.toString().trim()) {
        emails.push(email.toString().trim());
      }
    }

    // Remove duplicates and return
    return [...new Set(emails)];

  } catch (error) {
    console.error('Error in getUserEmails:', error);
    throw error;
  }
}



// Include this function to get the user data in the HTML files
function includeUserData() {
  const userData = getUserSession();
  return userData ? JSON.stringify(userData) : 'null';
}


function getData(token) {
  try {
    Logger.log("=== START getData ===");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName('Master');
    const userData = getSessionForCall_(token);

    if (!userData) {
      Logger.log("ERROR: No user session.");
      return JSON.stringify([]);
    }

    if (!masterSheet) {
      Logger.log("ERROR: Master sheet not found.");
      return JSON.stringify([]);
    }

    const allValues = masterSheet.getDataRange().getValues();
    
    if (!allValues || allValues.length <= 1) {
      return JSON.stringify([]);
    }

    allValues.shift();

    let filteredValues = [];

    if (userData.role && userData.role.toLowerCase() === 'doer') {
      filteredValues = allValues.filter(row => {
        return row && row.length > 3 && 
               row[3] && 
               row[3].toString().toLowerCase().trim() === userData.name.toLowerCase().trim();
      });
    } else {
      filteredValues = allValues;
    }

    const formattedValues = filteredValues.map(row => {
      if (!row) return null;
      return row.map((cell, index) => {
        // Format date columns
        if ([0, 6, 7, 18, 20].includes(index) && cell instanceof Date && !isNaN(cell)) {
          const day = cell.getDate().toString().padStart(2, '0');
          const month = (cell.getMonth() + 1).toString().padStart(2, '0');
          const year = cell.getFullYear();
          const hours = cell.getHours().toString().padStart(2, '0');
          const minutes = cell.getMinutes().toString().padStart(2, '0');
          const seconds = cell.getSeconds().toString().padStart(2, '0');
          return `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
        }
        
        // FIX: Column W (index 22) - Approved Count - force to number
        if (index === 22) {
          if (cell instanceof Date) return 0;
          const num = parseInt(cell);
          return isNaN(num) ? 0 : num;
        }
        
        // FIX: Column I (index 8) - Revision Count - force to number
        if (index === 8) {
          if (cell instanceof Date) return 0;
          const num = parseInt(cell);
          return isNaN(num) ? 0 : num;
        }

        return cell !== null && cell !== undefined ? cell : '';
      });
    }).filter(row => row !== null);

    const finalValues = formattedValues.filter(row => {
      if (!row || row.length === 0) return false;
      const isNotEmpty = row[0] !== '' && row[0] !== null && row[0] !== undefined;
      const status = row.length > 9 ? row[9].toString().trim().toLowerCase() : '';
      const isNotShifted = status !== 'week shifted';
      return isNotEmpty && isNotShifted;
    });

    Logger.log("Final rows: " + finalValues.length);
    Logger.log("=== END getData ===");

    return JSON.stringify(finalValues);
    
  } catch (error) {
    Logger.log("EXCEPTION in getData: " + error.toString());
    return JSON.stringify([]);
  }
}
// Updated function for adding MD Remarks with status updates (removed count functionality)
function addMDRemarkWithStatus(taskId, newRemark, statusValue) {
  try {
    // Log input parameters for debugging
    Logger.log('Adding MD Remark for Task ID: ' + taskId);
    Logger.log('New Remark: ' + newRemark);
    Logger.log('Status Update: ' + (statusValue || 'No change'));

    // Get the spreadsheet and Master sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Master');

    // Find the row with the matching Task ID (Search column B for Task Id)
    const taskIdColumn = sheet.getRange('B2:B').getValues();

    let rowIndex = -1;
    for (let i = 0; i < taskIdColumn.length; i++) {
      // Convert both to strings and trim for comparison
      const cellValue = String(taskIdColumn[i][0]).trim();
      const searchValue = String(taskId).trim();

      Logger.log(`Comparing: [${cellValue}] with [${searchValue}]`);

      if (cellValue === searchValue) {
        rowIndex = i + 2; // +2 because we start at row 2 and arrays are zero-indexed
        Logger.log('Match found at row: ' + rowIndex);
        break;
      }
    }

    if (rowIndex === -1) {
      Logger.log('Task ID not found in column B');
      throw new Error('Task ID not found: ' + taskId);
    }

    // Update status if a new status is provided (Column K is the 11th column)
    if (statusValue && statusValue.trim() !== '') {
      const statusCell = sheet.getRange(rowIndex, 10); // Corrected from 10 to 11 for Column K
      Logger.log('Updating status from: ' + statusCell.getValue() + ' to: ' + statusValue);
      statusCell.setValue(statusValue);
    }

    // Get the current MD Remark cell (Column R is the 18th column)
    const remarkCell = sheet.getRange(rowIndex, 18); // Corrected from 16 to 18 for Column R
    const currentRemark = remarkCell.getValue();

    // Append the new remark or set it if there's nothing yet
    let updatedRemark;
    if (currentRemark && String(currentRemark).trim()) {
      updatedRemark = currentRemark + '\n' + newRemark;
    } else {
      updatedRemark = newRemark;
    }

    Logger.log('Updated MD remark to be set: ' + updatedRemark);

    // Update the remark cell
    remarkCell.setValue(updatedRemark);

    // Force the spreadsheet to save changes
    SpreadsheetApp.flush();

    Logger.log('MD Remark successfully added.');

    // Return success
    return {
      success: true,
      message: 'MD Remark added successfully' + (statusValue ? ' with status update' : '')
    };

  } catch (error) {
    // Log the error and return failure
    Logger.log('Error adding MD Remark with status: ' + error.toString());
    return {
      success: false,
      message: error.toString()
    };
  }
}


// Lock prevents two simultaneous assignments writing into the same row
function assignTaskWithAttachments(taskData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    return assignTaskCore_(taskData);
  } catch (err) {
    return { success: false, error: err.toString() };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function assignTaskCore_(taskData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Master');

    const assigneeName = taskData.assigneeName;
    const doerName = taskData.doerName;
    const projectName = taskData.projectName;
    const taskDescription = taskData.taskDescription;
    const dueDate = taskData.dueDate || '';
    const evidenceRequired = taskData.evidenceRequired;
    const approvalRequired = taskData.approvalRequired; 
    const recurring = taskData.recurring;               
    const taskType = taskData.taskType || 'Regular Task';

    const columnB = sheet.getRange('B:B').getValues();
    let firstEmptyRow = 0;
    for (let i = 1; i < columnB.length; i++) {
      if (!columnB[i][0]) {
        firstEmptyRow = i + 1;
        break;
      }
    }
    if (firstEmptyRow === 0) firstEmptyRow = columnB.length + 1;

    const taskId = 'Task/' + new Date().getTime().toString().slice(-6);
    const now = new Date();
    const currentTimestamp = Utilities.formatDate(now, "GMT+5:30", "dd/MM/yyyy HH:mm:ss");
    const status = dueDate ? 'Pending' : 'Without Plan';

    sheet.getRange(firstEmptyRow, 1).setValue(currentTimestamp); 
    sheet.getRange(firstEmptyRow, 2).setValue(taskId);           
    sheet.getRange(firstEmptyRow, 3).setValue(assigneeName);      
    sheet.getRange(firstEmptyRow, 4).setValue(doerName);          
    sheet.getRange(firstEmptyRow, 5).setValue(projectName);       
    sheet.getRange(firstEmptyRow, 6).setValue(taskDescription);   
    sheet.getRange(firstEmptyRow, 7).setValue(dueDate);           
    sheet.getRange(firstEmptyRow, 10).setValue(status);           
    sheet.getRange(firstEmptyRow, 15).setValue(evidenceRequired); 
    sheet.getRange(firstEmptyRow, 16).setValue(approvalRequired); 
    sheet.getRange(firstEmptyRow, 17).setValue(recurring);        
    sheet.getRange(firstEmptyRow, 22).setValue(taskType);         

    // FIX: Set revision count and approved count as plain numbers with number format
    sheet.getRange(firstEmptyRow, 9).setValue(0);
    sheet.getRange(firstEmptyRow, 9).setNumberFormat('0');
    
    sheet.getRange(firstEmptyRow, 23).setValue(0);
    sheet.getRange(firstEmptyRow, 23).setNumberFormat('0');

    const folderId = '1c3tpbea50bUKvv6AS30wAjclk-ZFOQcU';
    const folder = DriveApp.getFolderById(folderId);

    if (taskData.voiceNote) {
      const voiceBlob = Utilities.newBlob(
        Utilities.base64Decode(taskData.voiceNote.data), 
        taskData.voiceNote.type, 
        taskData.voiceNote.name
      );
      const file = folder.createFile(voiceBlob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      sheet.getRange(firstEmptyRow, 12).setValue(file.getUrl()); 
    }

    if (taskData.attachments && taskData.attachments.length > 0) {
      let links = taskData.attachments.map(att => {
        const blob = Utilities.newBlob(
          Utilities.base64Decode(att.data), 
          att.type, 
          att.name
        );
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return file.getUrl();
      });
      sheet.getRange(firstEmptyRow, 13).setValue(links.join(', ')); 
    }

    SpreadsheetApp.flush();
    return { success: true, taskId: taskId };
    
  } catch (error) {
    Logger.log(error);
    return { success: false, error: error.toString() };
  }
}




function updateTask(updateData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Master');
    const taskIdColumn = sheet.getRange('B:B').getValues();
    let rowIndex = -1;

    for (let i = 1; i < taskIdColumn.length; i++) {
      if (taskIdColumn[i][0] === updateData.taskId) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) throw new Error('Task not found');

    sheet.getRange(rowIndex, 3).setValue(updateData.assigneeName);    
    sheet.getRange(rowIndex, 4).setValue(updateData.doerName);        
    sheet.getRange(rowIndex, 5).setValue(updateData.projectName);     
    sheet.getRange(rowIndex, 6).setValue(updateData.taskDescription); 
    sheet.getRange(rowIndex, 7).setValue(updateData.dueDate);         
    sheet.getRange(rowIndex, 15).setValue(updateData.evidenceRequired); 
    sheet.getRange(rowIndex, 16).setValue(updateData.approvalRequired); 
    sheet.getRange(rowIndex, 17).setValue(updateData.recurring);        

    if (updateData.taskType) {
      sheet.getRange(rowIndex, 22).setValue(updateData.taskType);
    }

    // FIX: Ensure revision count and approved count stay as numbers
    const currentRevision = sheet.getRange(rowIndex, 9).getValue();
    if (currentRevision instanceof Date) {
      sheet.getRange(rowIndex, 9).setValue(0);
    }
    sheet.getRange(rowIndex, 9).setNumberFormat('0');

    const currentApproved = sheet.getRange(rowIndex, 23).getValue();
    if (currentApproved instanceof Date) {
      sheet.getRange(rowIndex, 23).setValue(0);
    }
    sheet.getRange(rowIndex, 23).setNumberFormat('0');

    if (updateData.newVoiceNote) {
      const folderId = '1c3tpbea50bUKvv6AS30wAjclk-ZFOQcU';
      const folder = DriveApp.getFolderById(folderId);
      const voiceNoteBlob = Utilities.newBlob(
        Utilities.base64Decode(updateData.newVoiceNote.data),
        updateData.newVoiceNote.type,
        `voice_note_${updateData.taskId}_${new Date().getTime()}.wav`
      );
      const voiceNoteFile = folder.createFile(voiceNoteBlob);
      voiceNoteFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      sheet.getRange(rowIndex, 12).setValue(voiceNoteFile.getUrl()); 
    }

    SpreadsheetApp.flush();
    return { success: true, message: 'Task updated successfully' };
    
  } catch (error) {
    Logger.log(error);
    throw new Error('Failed to update task: ' + error.toString());
  }
}

function sendTaskNotificationEmail(doerName, taskId, taskDescription, dueDate, voiceNoteLink, attachmentLinks) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const doerListSheet = ss.getSheetByName('Doer List');

    // Get all data from Doer List sheet
    const doerData = doerListSheet.getDataRange().getValues();

    // Find the doer's email by matching their name
    let doerEmail = null;
    for (let i = 1; i < doerData.length; i++) { // Start from 1 to skip header row
      if (doerData[i][1] === doerName) { // Column B (index 1) has doer name
        doerEmail = doerData[i][0]; // Column A (index 0) has email
        break;
      }
    }

    if (!doerEmail) {
      Logger.log('Could not find email for doer: ' + doerName);
      return;
    }

    // Format due date for display
    let formattedDueDate = dueDate;
    if (dueDate instanceof Date) {
      formattedDueDate = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'MMM dd, yyyy');
    }

    // Prepare email content
    const subject = `Task Assignment: ${taskId}`;

    let body = `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>New Task Assignment</title>
        <style>
          body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333333;
            margin: 0;
            padding: 0;
          }
          .container {
            max-width: 600px;
            margin: 0 auto;
            padding: 20px;
          }
          .header {
            background-color: #4285f4;
            color: white;
            padding: 20px;
            text-align: center;
            border-radius: 5px 5px 0 0;
          }
          .content {
            background-color: #ffffff;
            padding: 20px;
            border-left: 1px solid #dddddd;
            border-right: 1px solid #dddddd;
          }
          .task-card {
            background-color: #f9f9f9;
            border-left: 4px solid #4285f4;
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 4px;
          }
          .task-detail {
            margin-bottom: 10px;
          }
          .label {
            font-weight: bold;
            color: #555555;
            display: inline-block;
            width: 100px;
          }
          .attachments {
            background-color: #f5f5f5;
            padding: 15px;
            border-radius: 4px;
            margin-top: 20px;
          }
          .attachments h3 {
            margin-top: 0;
            color: #555555;
            font-size: 16px;
          }
          .attachments ul {
            padding-left: 20px;
          }
          .attachments a {
            color: #4285f4;
            text-decoration: none;
          }
          .attachments a:hover {
            text-decoration: underline;
          }
          .footer {
            background-color: #f5f5f5;
            padding: 15px;
            text-align: center;
            font-size: 12px;
            color: #777777;
            border-radius: 0 0 5px 5px;
            border: 1px solid #dddddd;
          }
          .button {
            display: inline-block;
            background-color: #4285f4;
            color: white;
            padding: 10px 20px;
            text-decoration: none;
            border-radius: 4px;
            margin-top: 15px;
            font-weight: bold;
          }
          .due-date {
            color: ${isDateSoon(dueDate) ? '#ff5252' : '#333333'};
            font-weight: ${isDateSoon(dueDate) ? 'bold' : 'normal'};
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h1>New Task Assignment</h1>
          </div>
          
          <div class="content">
            <p>Hello ${doerName},</p>
            <p>You have been assigned a new task. Details are below:</p>
            
            <div class="task-card">
              <div class="task-detail">
                <span class="label">Task ID:</span> ${taskId}
              </div>
              <div class="task-detail">
                <span class="label">Due Date:</span> <span class="due-date">${formattedDueDate}</span>
              </div>
              <div class="task-detail">
                <span class="label">Description:</span>
              </div>
              <div style="margin-top: 8px; line-height: 1.5;">
                ${taskDescription}
              </div>
            </div>
    `;

    // Add links to attachments if they exist
    const hasVoiceNote = voiceNoteLink && voiceNoteLink.trim() !== '';
    const hasAttachments = Array.isArray(attachmentLinks) && attachmentLinks.length > 0;

    if (hasVoiceNote || hasAttachments) {
      body += `
            <div class="attachments">
              <h3>Attachments</h3>
              <ul>
      `;

      // Add voice note if available
      if (hasVoiceNote) {
        body += `<li><a href="${voiceNoteLink}">Voice Note</a></li>`;
      }

      // Add all file attachments if available
      if (hasAttachments) {
        // If attachmentLinks is an array of URLs
        attachmentLinks.forEach((link, index) => {
          // Try to get a filename from the URL or use a default name
          const fileName = getFilenameFromURL(link) || `Attachment ${index + 1}`;
          body += `<li><a href="${link}">${fileName}</a></li>`;
        });
      } else if (typeof attachmentLinks === 'string' && attachmentLinks.trim() !== '') {
        // For backward compatibility - if attachmentLinks is a single string
        // Check if it contains multiple links separated by commas
        const links = attachmentLinks.split(',').map(url => url.trim()).filter(url => url !== '');

        if (links.length > 1) {
          // Multiple links in a string separated by commas
          links.forEach((link, index) => {
            const fileName = getFilenameFromURL(link) || `Attachment ${index + 1}`;
            body += `<li><a href="${link}">${fileName}</a></li>`;
          });
        } else if (links.length === 1) {
          // Single link as a string
          const fileName = getFilenameFromURL(links[0]) || 'File Attachment';
          body += `<li><a href="${links[0]}">${fileName}</a></li>`;
        }
      }

      body += `
              </ul>
            </div>
      `;
    }

    // Add a button to acknowledge receipt
    body += `
            <div style="text-align: center; margin-top: 25px;">
              <a href="https://script.google.com/a/codearchitects.in/macros/s/AKfycbzawv8hZ0lQRGOhemghTb4sjFEZdjdoE-C2iG32tHdG/dev?page=Index" class="button">View Your All Tasks</a>
            </div>
            
            <p style="margin-top: 25px;">Please complete this task by the due date. If you have any questions, reply to this email or contact your supervisor.</p>
            <p>Thank you!</p>
          </div>
          
          <div class="footer">
            <p>This is an automated message. Please do not reply directly to this email.</p>
            <p>© ${new Date().getFullYear()} Your Company Name</p>
          </div>
        </div>
      </body>
    </html>
    `;

    // Send the email
    GmailApp.sendEmail(doerEmail, subject, "", {
      htmlBody: body
    });

    Logger.log('Task notification email sent to: ' + doerEmail);
    return true;

  } catch (error) {
    Logger.log('Failed to send task notification email: ' + error.toString());
    return false;
  }
}

// Helper function to extract filename from URL
function getFilenameFromURL(url) {
  try {
    if (!url) return '';

    // Try to extract filename from the URL path
    const urlObj = new URL(url);
    const pathParts = urlObj.pathname.split('/');
    let fileName = pathParts[pathParts.length - 1];

    // Remove query parameters if present
    fileName = fileName.split('?')[0];

    // Decode URI components
    fileName = decodeURIComponent(fileName);

    // If filename is empty, use a default name
    if (!fileName || fileName === '/') {
      return '';
    }

    return fileName;
  } catch (e) {
    // If URL parsing fails, return empty string
    return '';
  }
}

// Helper function to check if a date is coming soon (within 3 days)
function isDateSoon(dueDate) {
  try {
    if (!dueDate) return false;

    let dueDateObj;
    if (typeof dueDate === 'string') {
      dueDateObj = new Date(dueDate);
    } else if (dueDate instanceof Date) {
      dueDateObj = dueDate;
    } else {
      return false;
    }

    const today = new Date();
    const diffTime = dueDateObj.getTime() - today.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

    return diffDays >= 0 && diffDays <= 3;
  } catch (e) {
    return false;
  }
}


function getDoerNamesWithRoles() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const doerSheet = ss.getSheetByName('Doer List');

    if (!doerSheet) {
      throw new Error('Doer List sheet not found');
    }

    // Get all values from columns B and C starting from row 2
    const dataRange = doerSheet.getRange('A2:B');
    const dataValues = dataRange.getValues();

    // Filter out empty values and return array of objects with name and role
    return dataValues
      .filter(row => row[0] !== '') // Filter out empty names
      .map(row => ({
        name: row[0],
        role: row[1] || '' // Role from column C, default to empty string if null
      }));

  } catch (error) {
    Logger.log(error);
    throw new Error('Failed to fetch doer names with roles: ' + error.toString());
  }
}

// Keep the old function for backward compatibility
function getDoerNames() {
  try {
    const doerData = getDoerNamesWithRoles();
    return doerData.map(item => item.name);
  } catch (error) {
    Logger.log(error);
    throw new Error('Failed to fetch doer names: ' + error.toString());
  }
}

// Helper function to check if a date is valid
function isValidDate(d) {
  return d instanceof Date && !isNaN(d);
}

function getDoerData(token) {
  try {
    SpreadsheetApp.flush();
    const userSession = getSessionForCall_(token);
    if (!userSession) return [];

    const { role: userRole, name: userName } = userSession;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const masterSheet = ss.getSheetByName('Master');
    const masterValues = masterSheet.getDataRange().getValues();
    const masterMap = {};
    masterValues.forEach(r => {
      if(r[1]) {
        masterMap[r[1]] = { 
          taskType: r[21] || 'Regular Task',
          approvedCount: r[22] || 0
        };
      }
    });

    const sheet = ss.getSheetByName('DoerDashboard');
    if (!sheet) {
      Logger.log("DoerDashboard sheet not found");
      return [];
    }

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) {
      Logger.log("DoerDashboard has no data rows");
      return [];
    }

    const dataRows = values.slice(1); 
    
    // FIXED: Safe filter
    const filteredRows = dataRows.filter(row => {
      if (!row[0] || row[0].toString().trim() === '') return false; 
      if (userRole === 'Doer') {
        const rowName = row[2] ? row[2].toString().trim().toLowerCase() : '';
        return rowName === userName.trim().toLowerCase();
      }
      return true; 
    });

    Logger.log("Filtered rows count: " + filteredRows.length);

    // FIXED: Safe sort - handle non-Date values
    const sortedRows = filteredRows.sort((a, b) => {
      const statusA = (a[14] || '').toString().trim();
      const statusB = (b[14] || '').toString().trim();
      const dateA = a[4];
      const dateB = b[4];

      const isEscalatedA = statusA.toLowerCase() === 'escalated to b.o';
      const isEscalatedB = statusB.toLowerCase() === 'escalated to b.o';
      if (isEscalatedA && !isEscalatedB) return 1;
      if (!isEscalatedA && isEscalatedB) return -1;

      // FIXED: Safe date comparison
      const hasDateA = dateA instanceof Date && !isNaN(dateA);
      const hasDateB = dateB instanceof Date && !isNaN(dateB);
      
      if (!hasDateA && !hasDateB) return 0;
      if (!hasDateA) return -1;
      if (!hasDateB) return 1;
      return dateA - dateB;
    });

    const formattedRows = sortedRows.map(row => {
      const formattedRow = [...row];
      
      // Format planned date (col E = index 4)
      if (formattedRow[4] instanceof Date && !isNaN(formattedRow[4])) {
        formattedRow[4] = Utilities.formatDate(
          formattedRow[4], 
          Session.getScriptTimeZone(), 
          "dd/MM/yyyy HH:mm:ss"
        );
      }
      
      // Append task type & approved count from Master
      const taskId = formattedRow[0];
      formattedRow[15] = masterMap[taskId] ? masterMap[taskId].taskType : 'Regular Task';
      formattedRow[16] = masterMap[taskId] ? masterMap[taskId].approvedCount : 0;
      
      return formattedRow;
    });

    Logger.log("getDoerData returning: " + formattedRows.length + " rows");
    return formattedRows;
    
  } catch (error) {
    Logger.log('Error in getDoerData: ' + error.toString() + ' | Stack: ' + error.stack);
    return [];
  }
}



// Helper function to check if a date is valid
function isValidDate(d) {
  return d instanceof Date && !isNaN(d);
}


function getActiveUserEmail() {
  return Session.getActiveUser().getEmail();
}


// Add this function to handle image uploads to Drive


function updateTaskDateWithImage(taskId, nextDate, reason, action, isShifted, imageUrl) {
  const res = updateTaskDateWithImageCore_(taskId, nextDate, reason, action, isShifted, imageUrl);
  SpreadsheetApp.flush(); // commit before the page refreshes
  return res;
}

function updateTaskDateWithImageCore_(taskId, nextDate, reason, action, isShifted, imageUrl) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName("Master");

    if (!masterSheet) throw new Error("Master sheet not found");

    const data = masterSheet.getDataRange().getValues();

    let taskRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][1].toString() === taskId.toString()) {
        taskRow = i + 1;
        break;
      }
    }

    if (taskRow === -1) throw new Error("Task ID not found: " + taskId);

    if (imageUrl) {
      masterSheet.getRange(taskRow, 14).setValue(imageUrl);
    }

    const now = new Date();
    const day = now.getDate().toString().padStart(2, '0');
    const month = (now.getMonth() + 1).toString().padStart(2, '0');
    const year = now.getFullYear();
    const hours = now.getHours().toString().padStart(2, '0');
    const minutes = now.getMinutes().toString().padStart(2, '0');
    const seconds = now.getSeconds().toString().padStart(2, '0');
    const timestamp = `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;

    // ── COMPLETED ──────────────────────────────────────────────
    if (action === 'completed') {
      const columnPValue = masterSheet.getRange(taskRow, 16).getValue();
      let statusToSet = "";

      if (columnPValue && columnPValue.toString().toLowerCase() === "yes") {
        statusToSet = "Approval Waiting";
      } else {
        const columnUTimestamp = masterSheet.getRange(taskRow, 21).getValue();
        if (columnUTimestamp) {
          const deadlineDate = new Date(columnUTimestamp);
          statusToSet = now > deadlineDate ? "Completed With Delay" : "Completed";
        } else {
          statusToSet = "Completed";
        }
      }

      masterSheet.getRange(taskRow, 10).setValue(statusToSet);
      masterSheet.getRange(taskRow, 19).setValue(timestamp);

      return {
        status: "success",
        message: "Task marked as " + statusToSet + (imageUrl ? " with image" : "")
      };

    // ── REVISE AND COMPLETE ────────────────────────────────────
    } else if (action === 'revise_and_complete') {
      if (!nextDate) throw new Error("Next date is required for revise and complete action");

      const columnPValue = masterSheet.getRange(taskRow, 16).getValue();
      let statusToSet = "";

      if (columnPValue && columnPValue.toString().toLowerCase() === "yes") {
        statusToSet = "Approval Waiting";
      } else {
        const columnUTimestamp = masterSheet.getRange(taskRow, 21).getValue();
        if (columnUTimestamp) {
          const deadlineDate = new Date(columnUTimestamp);
          statusToSet = now > deadlineDate ? "Completed With Delay" : "Completed";
        } else {
          statusToSet = "Completed";
        }
      }

      masterSheet.getRange(taskRow, 10).setValue(statusToSet);
      masterSheet.getRange(taskRow, 19).setValue(timestamp);

      const columnAValues = masterSheet.getRange("A:A").getValues();
      let lastNonBlankRow = 0;
      for (let i = columnAValues.length - 1; i >= 0; i--) {
        if (columnAValues[i][0] !== "" && columnAValues[i][0] !== null) {
          lastNonBlankRow = i + 1;
          break;
        }
      }
      const newRow = lastNonBlankRow + 1;

      const rowData = masterSheet.getRange(taskRow, 1, 1, 20).getValues()[0];
      const randomNum = Math.floor(100000 + Math.random() * 900000);
      const newTaskId = "Task/" + randomNum;

      const dateObj = new Date(nextDate);
      const formattedDate = Utilities.formatDate(
        dateObj,
        Session.getScriptTimeZone(),
        "dd/MM/yyyy HH:mm:ss"
      );

      rowData[0] = timestamp;
      rowData[1] = newTaskId;
      rowData[6] = formattedDate;
      rowData[7] = "";
      rowData[8] = 0;
      rowData[9] = "Pending";
      rowData[13] = "";
      rowData[18] = "";

      let remarkText = timestamp + " - Task revised and copied from Task ID: " + taskId;
      if (reason && reason.trim() !== '') remarkText += ": " + reason;
      if (imageUrl) remarkText = "Image attached. " + remarkText;
      rowData[10] = remarkText;

      masterSheet.getRange(newRow, 1, 1, 20).setValues([rowData]);

      return {
        status: "success",
        message: "Task completed and new revision created with ID: " + newTaskId + (imageUrl ? " with image" : "")
      };

    // ── NEED HELP ──────────────────────────────────────────────
    } else if (action === 'need_help') {
      masterSheet.getRange(taskRow, 10).setValue("Need Help");

      let remarkText = timestamp + " - Need Help";
      if (reason && reason.trim() !== '') remarkText += ": " + reason;
      if (imageUrl) remarkText = "Image attached. " + remarkText;

      const currentRemark = masterSheet.getRange(taskRow, 11).getValue();
      const newRemark = currentRemark ? (currentRemark + "\n" + remarkText) : remarkText;
      masterSheet.getRange(taskRow, 11).setValue(newRemark);

      return {
        status: "success",
        message: "Help request submitted" + (imageUrl ? " with image" : "")
      };

    // ── SHIFTED (Follow-up & Regular date updates) ──────────────
    } else if (action === 'shifted') {
      const dateObj = new Date(nextDate);
      const formattedDate = Utilities.formatDate(
        dateObj,
        Session.getScriptTimeZone(),
        "dd/MM/yyyy HH:mm:ss"
      );

      // Check task type from Column V (index 21, col 22)
      const taskTypeValue = masterSheet.getRange(taskRow, 22).getValue();
      const isFollowUpTask = taskTypeValue && taskTypeValue.toString().trim() === 'Follow-up Task';

      // Update revision count in Column I (index 9)
      const revisionRange = masterSheet.getRange(taskRow, 9);
      let revisionCount = parseInt(revisionRange.getValue()) || 0;
      revisionCount++;
      revisionRange.setValue(revisionCount);
      revisionRange.setNumberFormat('0');

      masterSheet.getRange(taskRow, 10).setValue("Pending");

      const currentFirstDate = masterSheet.getRange(taskRow, 7).getValue();
      let isInitialDateSet = false;

      // Ensure Column G (index 7, 1-based index) always keeps the original first date.
      // Any shifted dates go into Column H (index 8, 1-based index) for both task types.
      if (!currentFirstDate || currentFirstDate === "") {
        masterSheet.getRange(taskRow, 7).setValue(formattedDate);
        isInitialDateSet = true;
      } else {
        masterSheet.getRange(taskRow, 8).setValue(formattedDate);
      }

      let remarkText = "";
      let successMessage = "";

      if (isFollowUpTask) {
        remarkText = timestamp + " - Follow-up date updated";
        successMessage = "Follow-up date updated";
      } else {
        if (isInitialDateSet) {
          remarkText = timestamp + " - Initial date set";
          successMessage = "Initial date set";
        } else {
          remarkText = timestamp + " - Date updated";
          successMessage = "Date updated";
        }
      }

      if (reason && reason.trim() !== '') remarkText += ": " + reason;
      if (imageUrl) remarkText = "Image attached. " + remarkText;

      const currentRemark = masterSheet.getRange(taskRow, 11).getValue();
      const newRemark = currentRemark ? (currentRemark + "\n" + remarkText) : remarkText;
      masterSheet.getRange(taskRow, 11).setValue(newRemark);

      return {
        status: "success",
        message: successMessage + (imageUrl ? " with image" : "")
      };
    }

    return { status: "error", message: "Unknown action: " + action };

  } catch (error) {
    console.error("Error in updateTaskDateWithImage:", error);
    return { status: "error", message: error.toString() };
  }
}

// For backward compatibility, keep the original function and make it call the new one
function updateTaskDate(taskId, nextDate, reason, action, isShifted) {
  return updateTaskDateWithImage(taskId, nextDate, reason, action, isShifted, null);
}

function getScoreCardData() {
  try {
    // Get user session data using browser-based cache
    const userSession = getUserSession();

    // If no user session, return empty array
    if (!userSession) {
      return [];
    }

    // Get user info
    const userRole = userSession.role;
    const userName = userSession.name;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ScoreCard');

    if (!sheet) {
      console.error('ScoreCard sheet not found');
      return [];
    }

    const range = sheet.getDataRange();
    const values = range.getValues();

    // Skip header row
    const dataRows = values.slice(1);

    // Filter out empty rows, apply role-based access control, and format date
    return dataRows
      .filter(row => {
        // Skip empty rows
        if (row[0] === '') return false;

        // If user is Admin, show all rows
        if (userRole === 'Admin') return true;

        // For Doer users, only show rows where their name matches column A (index 0)
        return row[0].toString().trim().toLowerCase() === userName.trim().toLowerCase();
      })
      .map(row => {
        // Create a copy of the row to avoid modifying the original data
        const formattedRow = [...row];

        // Format the date in column C (index 2)
        if (formattedRow[2] instanceof Date && !isNaN(formattedRow[2])) {
          const date = formattedRow[2];
          const day = date.getDate().toString().padStart(2, '0');
          const month = (date.getMonth() + 1).toString().padStart(2, '0');
          const year = date.getFullYear();
          formattedRow[2] = `${day}/${month}/${year}`;
        }
        return formattedRow;
      });
  } catch (error) {
    console.error('Error in getScoreCardData:', error);
    return [];
  }
}


function sendPendingTasksSummaryEmails() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const doerDashboardSheet = ss.getSheetByName('DoerDashboard');

    if (!doerDashboardSheet) {
      Logger.log('Error: DoerDashboard sheet not found');
      return;
    }

    // Get all data from DoerDashboard sheet
    const dashboardData = doerDashboardSheet.getDataRange().getValues();

    // Skip header row and check if there's data
    if (dashboardData.length <= 1) {
      Logger.log('No pending tasks found to send');
      return;
    }

    // Group tasks by email
    const tasksByEmail = {};

    // Start from row 1 (skip header at row 0)
    for (let i = 1; i < dashboardData.length; i++) {
      const row = dashboardData[i];
      const taskId = row[0];          // Column A - Task ID
      const email = row[1];           // Column B - Email
      const doerName = row[2];        // Column C - Doer Name
      const taskDescription = row[3]; // Column D - Task Description
      const dueDate = row[4];         // Column E - Due Date
      const voiceNoteLink = row[6];   // Column G - Voice Note Link
      const attachmentLink = row[7];  // Column H - Attachment Link

      // Skip rows with no email
      if (!email) continue;

      // Format due date
      let formattedDueDate = dueDate;
      if (dueDate instanceof Date) {
        formattedDueDate = Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'MMM dd, yyyy');
      }

      // Create task object
      const task = {
        taskId: taskId,
        doerName: doerName,
        taskDescription: taskDescription,
        dueDate: dueDate, // Keep Date object for sorting
        formattedDueDate: formattedDueDate,
        voiceNoteLink: voiceNoteLink,
        attachmentLink: attachmentLink
      };

      // Add task to the email group
      if (!tasksByEmail[email]) {
        tasksByEmail[email] = {
          doerName: doerName,
          tasks: []
        };
      }

      tasksByEmail[email].tasks.push(task);
    }

    // Send emails to each doer with their pending tasks
    let emailsSent = 0;

    for (const email in tasksByEmail) {
      const doerInfo = tasksByEmail[email];
      const doerName = doerInfo.doerName;
      const tasks = doerInfo.tasks;

      // Sort tasks by due date (ascending)
      tasks.sort((a, b) => {
        if (a.dueDate instanceof Date && b.dueDate instanceof Date) {
          return a.dueDate.getTime() - b.dueDate.getTime();
        }
        return 0;
      });

      // Check how many tasks are urgent (due within 3 days)
      const today = new Date();
      const urgentTasks = tasks.filter(task => {
        if (task.dueDate instanceof Date) {
          const diffTime = task.dueDate.getTime() - today.getTime();
          const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
          return diffDays >= 0 && diffDays <= 3;
        }
        return false;
      });

      // Send the email with all pending tasks
      sendPendingTasksEmail(email, doerName, tasks, urgentTasks.length);
      emailsSent++;
    }

    Logger.log(`Successfully sent pending tasks summary emails to ${emailsSent} doers`);
    return emailsSent;

  } catch (error) {
    Logger.log('Failed to send pending tasks summary emails: ' + error.toString());
    return 0;
  }
}


function sendPendingTasksEmail(email, doerName, tasks, urgentCount) {
  // Prepare email content
  const subject = urgentCount > 0
    ? `❗ ${urgentCount} Urgent Task${urgentCount > 1 ? 's' : ''} - Pending Tasks Summary`
    : 'Your Pending Tasks Summary';

  let body = `
  <!DOCTYPE html>
  <html>
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Pending Tasks Summary</title>
      <style>
        body {
          font-family: Arial, sans-serif;
          line-height: 1.6;
          color: #333333;
          margin: 0;
          padding: 0;
        }
        .container {
          max-width: 800px;
          margin: 0 auto;
          padding: 20px;
        }
        .header {
          background-color: #3c78d8;
          color: white;
          padding: 20px;
          text-align: center;
          border-radius: 5px 5px 0 0;
        }
        .content {
          background-color: #ffffff;
          padding: 20px;
          border-left: 1px solid #dddddd;
          border-right: 1px solid #dddddd;
        }
        .summary {
          background-color: #f8f9fa;
          padding: 15px;
          margin-bottom: 20px;
          border-radius: 4px;
          border-left: 4px solid #3c78d8;
        }
        table {
          width: 100%;
          border-collapse: collapse;
          margin-bottom: 20px;
        }
        th {
          background-color: #f2f2f2;
          text-align: left;
          padding: 12px;
          border: 1px solid #dddddd;
        }
        td {
          padding: 12px;
          border: 1px solid #dddddd;
          vertical-align: top;
        }
        tr:nth-child(even) {
          background-color: #f9f9f9;
        }
        .urgent {
          background-color: #ffe6e6 !important;
        }
        .past-due {
          background-color: #ffcccc !important;
        }
        .footer {
          background-color: #f5f5f5;
          padding: 15px;
          text-align: center;
          font-size: 12px;
          color: #777777;
          border-radius: 0 0 5px 5px;
          border: 1px solid #dddddd;
        }
        .button {
          display: inline-block;
          background-color: #3c78d8;
          color: white;
          padding: 10px 20px;
          text-decoration: none;
          border-radius: 4px;
          margin-top: 15px;
        }
        .attachment-link {
          color: #3c78d8;
          text-decoration: none;
        }
        .attachment-link:hover {
          text-decoration: underline;
        }
      </style>
    </head>
    <body>
      <div class="container">
        <div class="header">
          <h1>Your Pending Tasks</h1>
        </div>
        
        <div class="content">
          <p>Hello ${doerName},</p>
          
          <div class="summary">
            <p><strong>Tasks Summary:</strong> You have ${tasks.length} pending task${tasks.length !== 1 ? 's' : ''}${urgentCount > 0 ? `, with ${urgentCount} marked as urgent (due within 3 days)` : ''}.</p>
          </div>
          
          <p>Here is a list of all your pending tasks:</p>
          
          <table>
            <tr>
              <th>Task ID</th>
              <th>Description</th>
              <th>Due Date</th>
              <th>Attachments</th>
            </tr>
  `;

  // Add row for each task
  const today = new Date();

  tasks.forEach(task => {
    const dueDate = task.dueDate instanceof Date ? task.dueDate : new Date(task.dueDate);
    const diffTime = dueDate.getTime() - today.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

    let rowClass = '';
    if (diffDays < 0) {
      rowClass = 'past-due'; // Past due
    } else if (diffDays <= 3) {
      rowClass = 'urgent'; // Due within 3 days
    }

    body += `
            <tr class="${rowClass}">
              <td>${task.taskId}</td>
              <td>${task.taskDescription}</td>
              <td>${task.formattedDueDate}${diffDays < 0 ? ' <strong>(PAST DUE)</strong>' : diffDays <= 3 ? ' <strong>(URGENT)</strong>' : ''}</td>
              <td>
    `;

    // Add attachment links if they exist
    if (task.voiceNoteLink || task.attachmentLink) {
      body += '<ul style="margin: 0; padding-left: 20px;">';

      if (task.voiceNoteLink) {
        body += `<li><a href="${task.voiceNoteLink}" class="attachment-link">Voice Note</a></li>`;
      }

      if (task.attachmentLink) {
        body += `<li><a href="${task.attachmentLink}" class="attachment-link">File Attachment</a></li>`;
      }

      body += '</ul>';
    } else {
      body += 'None';
    }

    body += `
              </td>
            </tr>
    `;
  });

  body += `
          </table>
          
          <p>Please prioritize tasks marked as <strong>URGENT</strong> and complete all tasks by their respective due dates.</p>
          
          <div style="text-align: center;">
            <a href="#" class="button">View All Tasks in Dashboard</a>
          </div>
          
          <p>If you have any questions or need clarification on any task, please contact your supervisor.</p>
          <p>Thank you for your attention to these matters.</p>
        </div>
        
        <div class="footer">
          <p>This is an automated message sent on ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy')}.</p>
          <p>© ${new Date().getFullYear()} Your Company Name</p>
        </div>
      </div>
    </body>
  </html>
  `;

  // Send the email
  GmailApp.sendEmail(email, subject, "", {
    htmlBody: body
  });

  Logger.log('Pending tasks summary email sent to: ' + email);
}



// function getChecklistData() {
//   try {
//     // Get user session data using browser-based cache
//     const userSession = getUserSession();

//     // If no user session, return empty array
//     if (!userSession) {
//       return [];
//     }

//     // Get user info
//     const userRole = userSession.role;
//     const userName = userSession.name;

//     // Get the spreadsheet and data
//     const ss = SpreadsheetApp.getActiveSpreadsheet();
//     const sheet = ss.getSheetByName('Checklist');

//     if (!sheet) {
//       console.error('Checklist sheet not found');
//       return [];
//     }

//     const range = sheet.getDataRange();
//     const values = range.getValues();

//     // Skip header row
//     const dataRows = values.slice(1);

//     // Filter out empty rows, apply role-based access control, and format dates
//     return dataRows
//       .filter(row => {
//         // Skip empty rows
//         if (row[0] === '') return false;

//         // If user is Admin, show all rows
//         if (userRole === 'Admin') return true;

//         // For Doer users, only show rows where their name matches column A (index 0)
//         return row[1].toString().trim().toLowerCase() === userName.trim().toLowerCase();
//       })
//       .map(row => {
//         // Create a copy of the row to avoid modifying the original data
//         const formattedRow = [...row];

//         // Format column F (index 5) - Planned Date with timestamp
//         if (formattedRow[5] instanceof Date && !isNaN(formattedRow[5])) {
//           const date = formattedRow[5];
//           const day = date.getDate().toString().padStart(2, '0');
//           const month = (date.getMonth() + 1).toString().padStart(2, '0');
//           const year = date.getFullYear();
//           const hours = date.getHours().toString().padStart(2, '0');
//           const minutes = date.getMinutes().toString().padStart(2, '0');
//           const seconds = date.getSeconds().toString().padStart(2, '0');
//           formattedRow[5] = `${day}/${month}/${year}, ${hours}:${minutes}:${seconds}`;
//         }

//         // Format column G (index 6) - Actual Date with timestamp
//         if (formattedRow[6] instanceof Date && !isNaN(formattedRow[6])) {
//           const date = formattedRow[6];
//           const day = date.getDate().toString().padStart(2, '0');
//           const month = (date.getMonth() + 1).toString().padStart(2, '0');
//           const year = date.getFullYear();
//           const hours = date.getHours().toString().padStart(2, '0');
//           const minutes = date.getMinutes().toString().padStart(2, '0');
//           const seconds = date.getSeconds().toString().padStart(2, '0');
//           formattedRow[6] = `${day}/${month}/${year}, ${hours}:${minutes}:${seconds}`;
//         }

//         return formattedRow;
//       });
//   } catch (error) {
//     console.error('Error in getChecklistData:', error);
//     return [];
//   }
// }


// Changes By Arjun
// function getChecklistData() { 
//   try {
//     const userSession = getUserSession();
//     if (!userSession) return [];

//     const userRole = userSession.role;
//     const userName = userSession.name;
//     const ss = SpreadsheetApp.getActiveSpreadsheet();
//     const sheet = ss.getSheetByName('Checklist');

//     if (!sheet) return [];

//     const values = sheet.getDataRange().getValues();
//     const dataRows = values.slice(1);

//     // ✅ Max Date set (31 Dec 2027)
//     const maxDate = new Date(2027, 11, 31, 23, 59, 59, 999);

//     return dataRows
//       .filter(row => {
//         if (row[0] === '') return false;

//         // --- DATE FILTER LOGIC START ---
//         const plannedDate = row[5];
        
//         if (plannedDate instanceof Date && !isNaN(plannedDate)) {
//           // ❗ Only filter if beyond 2027
//           if (plannedDate > maxDate) {
//             return false;
//           }
//         } else if (typeof plannedDate === 'string' && plannedDate.trim() !== '') {
//           const parts = plannedDate.split(/[\s,/:-]+/); 
//           if (parts.length >= 3) {
//             const day = parseInt(parts[0], 10);
//             const month = parseInt(parts[1], 10) - 1; 
//             const year = parseInt(parts[2], 10);
//             const parsedDate = new Date(year, month, day);

//             if (parsedDate > maxDate) return false;
//           }
//         }
//         // --- DATE FILTER LOGIC END ---

//         // ✅ Role-based filtering
//         if (userRole === 'Admin') return true;

//         const isDoer = row[1] && row[1].toString().trim().toLowerCase() === userName.trim().toLowerCase();
//         const isChecker = row[12] && row[12].toString().trim().toLowerCase() === userName.trim().toLowerCase();
        
//         return isDoer || isChecker;
//       })
//       .map(row => {
//         const formattedRow = [...row];

//         // Format Planned Date
//         if (formattedRow[5] instanceof Date && !isNaN(formattedRow[5])) {
//           const date = formattedRow[5];
//           formattedRow[5] = `${date.getDate().toString().padStart(2, '0')}/${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getFullYear()}, ${date.getHours().toString().padStart(2, '0')}:${date.getMinutes().toString().padStart(2, '0')}:${date.getSeconds().toString().padStart(2, '0')}`;
//         }

//         // Format Actual Date
//         if (formattedRow[6] instanceof Date && !isNaN(formattedRow[6])) {
//           const date = formattedRow[6];
//           formattedRow[6] = `${date.getDate().toString().padStart(2, '0')}/${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getFullYear()}, ${date.getHours().toString().padStart(2, '0')}:${date.getMinutes().toString().padStart(2, '0')}:${date.getSeconds().toString().padStart(2, '0')}`;
//         }

    

//         return formattedRow;
//       });

//   } catch (error) {
//     console.error('Error in getChecklistData:', error);
//     return [];
//   }
// }
// Updated getChecklistData with Column P formatting
// Updated getChecklistData to clear Checker info when not required
// REVISED: getChecklistData - This version sends ALL tasks to the dashboard
// FETCH ONLY PENDING TASKS
function getChecklistData(token) { 
  try {
    const userSession = getSessionForCall_(token);
    if (!userSession) return [];

    const userRole = userSession.role;
    const userName = userSession.name;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Checklist');

    if (!sheet) return [];

    const values = sheet.getDataRange().getValues();
    const dataRows = values.slice(1);

    const maxDate = new Date(2027, 11, 31, 23, 59, 59, 999);

    return dataRows
      .filter(row => {
        if (row[0] === '') return false;

        // --- STRICT FILTER: Completely block "Completed" or "Done" tasks ---
        const status = row[13] ? row[13].toString().trim().toLowerCase() : '';
        if (status === 'completed' || status === 'done') {
          return false; // Skip this row completely
        }
        // -------------------------------------------------------------------

        const plannedDate = row[5];
        if (plannedDate instanceof Date && !isNaN(plannedDate)) {
          if (plannedDate > maxDate) return false;
        } else if (typeof plannedDate === 'string' && plannedDate.trim() !== '') {
          const parts = plannedDate.split(/[\s,/:-]+/); 
          if (parts.length >= 3) {
            const day = parseInt(parts[0], 10);
            const month = parseInt(parts[1], 10) - 1; 
            const year = parseInt(parts[2], 10);
            const parsedDate = new Date(year, month, day);
            if (parsedDate > maxDate) return false;
          }
        }

        if (userRole === 'Admin') return true;
        const isDoer = row[1] && row[1].toString().trim().toLowerCase() === userName.trim().toLowerCase();
        const isChecker = row[12] && row[12].toString().trim().toLowerCase() === userName.trim().toLowerCase();
        
        return isDoer || isChecker;
      })
      .map(row => {
        const formattedRow = [...row];
        const proofRequired = formattedRow[7] ? formattedRow[7].toString().trim().toLowerCase() : '';

        // Clear Checker info if proof is not required
        if (proofRequired === 'no') {
          formattedRow[12] = ''; 
          formattedRow[15] = ''; 
        }

        // Format Dates
        if (formattedRow[5] instanceof Date && !isNaN(formattedRow[5])) {
          const date = formattedRow[5];
          formattedRow[5] = `${date.getDate().toString().padStart(2, '0')}/${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getFullYear()}, ${date.getHours().toString().padStart(2, '0')}:${date.getMinutes().toString().padStart(2, '0')}:${date.getSeconds().toString().padStart(2, '0')}`;
        }

        if (formattedRow[6] instanceof Date && !isNaN(formattedRow[6])) {
          const date = formattedRow[6];
          formattedRow[6] = `${date.getDate().toString().padStart(2, '0')}/${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getFullYear()}, ${date.getHours().toString().padStart(2, '0')}:${date.getMinutes().toString().padStart(2, '0')}:${date.getSeconds().toString().padStart(2, '0')}`;
        }

        if (formattedRow[15] instanceof Date && !isNaN(formattedRow[15])) {
          const date = formattedRow[15];
          formattedRow[15] = `${date.getDate().toString().padStart(2, '0')}/${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getFullYear()}, ${date.getHours().toString().padStart(2, '0')}:${date.getMinutes().toString().padStart(2, '0')}:${date.getSeconds().toString().padStart(2, '0')}`;
        }

        return formattedRow;
      });

  } catch (error) {
    console.error('Error in getChecklistData:', error);
    return [];
  }
}
//New Funtion add by arjun gehlot
// function updateCheckerStatus(taskId, status) {
//   try {
//     const ss = SpreadsheetApp.getActiveSpreadsheet();
//     const sheet = ss.getSheetByName('Checklist');
//     const data = sheet.getDataRange().getValues();
    
//     let rowIndex = -1;
//     for (let i = 0; i < data.length; i++) {
//       if (data[i][2] === taskId) { // Column C contains Task ID
//         rowIndex = i + 1; 
//         break;
//       }
//     }

//     if (rowIndex === -1) throw new Error('Task not found: ' + taskId);

//     // Update Status column (Column N, index 14)
//     sheet.getRange(rowIndex, 14).setValue(status);
//     return { success: true };
//   } catch (error) {
//     Logger.log('Error updating checker status: ' + error);
//     return { success: false, error: error.toString() };
//   }
// }

// Updated Backend Function to save Status and Remarks
function updateCheckerStatus(taskId, status, remark) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Checklist');
    const data = sheet.getDataRange().getValues();
    
    let rowIndex = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i][2] === taskId) { // Column C contains Task ID
        rowIndex = i + 1; 
        break;
      }
    }

    if (rowIndex === -1) throw new Error('Task not found: ' + taskId);

    // Update Status column (Column N, index 14)
    sheet.getRange(rowIndex, 14).setValue(status);
    
    // Update Reviewer Remarks column (Column O, index 15)
    // Ensure you have a header "Reviewer Remarks" in Column O of your Checklist sheet
    sheet.getRange(rowIndex, 15).setValue(remark); 

    if (status === 'Need Correction') {
      sheet.getRange(rowIndex, 7).setValue('');  // Clear Column G (Actual)
      sheet.getRange(rowIndex, 16).setValue(''); // Clear Column P (Checker Plan Date)
    }
    
    return { success: true };
  } catch (error) {
    Logger.log('Error updating checker status: ' + error);
    return { success: false, error: error.toString() };
  }
}


function uploadImageToDrive(fileName, base64Data, mimeType) {
  try {
    // Get the target folder where images will be stored
    const folderId = '1c3tpbea50bUKvv6AS30wAjclk-ZFOQcU';
    
    // Create a unique file name to avoid duplicates
    const uniqueFileName = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss") + "_" + fileName;

    // Convert base64 data to blob
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, uniqueFileName);

    // --- REPLACED DriveApp with Advanced Drive Service V3 ---
    const fileMetadata = {
      name: uniqueFileName,
      parents: [folderId]
    };
    
    // Create the file
    const file = Drive.Files.create(fileMetadata, blob);

    // Try to make the file accessible via link
    try {
      Drive.Permissions.create({
        role: 'reader',
        type: 'anyone'
      }, file.id);
    } catch (sharingError) {
      console.warn("Could not set public sharing (likely blocked by organization workspace policy): " + sharingError.toString());
    }

    // Return the viewable URL of the file
    return `https://drive.google.com/uc?id=${file.id}&export=view`;
  } catch (error) {
    console.error("Error uploading image to Drive:", error);
    throw new Error("Failed to upload image: " + error.toString());
  }
}

function uploadTaskProof(fileData, taskId, folderId) {
  try {
    if (!fileData || !fileData.contents) {
      throw new Error("File data is missing.");
    }

    const targetFolderId = folderId || '1c3tpbea50bUKvv6AS30wAjclk-ZFOQcU';
    const timestamp = new Date().getTime();
    const safeTaskId = taskId || 'UnknownTask';
    const safeFileName = fileData.name || 'document_upload';
    const fileName = `${safeTaskId}_proof_${timestamp}_${safeFileName}`;

    const contentType = fileData.type || 'application/octet-stream';
    const bytes = Utilities.base64Decode(fileData.contents);
    const blob = Utilities.newBlob(bytes, contentType, fileName);

    // --- REPLACED DriveApp with Advanced Drive Service V3 ---
    const fileMetadata = {
      name: fileName,
      parents: [targetFolderId]
    };
    
    // Create the file
    const file = Drive.Files.create(fileMetadata, blob);
    
    // Try to set sharing permissions safely
    try {
      Drive.Permissions.create({
        role: 'reader',
        type: 'anyone'
      }, file.id);
    } catch (sharingError) {
      Logger.log("Could not set public sharing (likely blocked by organization workspace policy). Proceeding anyway... " + sharingError.toString());
    }

    // Return the viewable URL
    return `https://drive.google.com/uc?id=${file.id}&export=view`;

  } catch (error) {
    Logger.log('Error uploading proof document: ' + error.toString());
    throw new Error(error.message || "Failed to upload to Drive."); 
  }
}


// function updateTaskCompletion(taskId, fileUrl, remark) {
//   try {
//     const ss = SpreadsheetApp.getActiveSpreadsheet();
//     const sheet = ss.getSheetByName('Checklist');
//     const data = sheet.getDataRange().getValues();
    
//     let rowIndex = -1;
//     let proofRequired = 'yes'; // Default to yes
    
//     for (let i = 0; i < data.length; i++) {
//       if (data[i][2] && data[i][2].toString() === taskId.toString()) { 
//         rowIndex = i + 1; // +1 because sheet rows are 1-indexed
        
//         // 🔥 FIXED: Robust case-insensitive read of Column H
//         proofRequired = data[i][7] ? data[i][7].toString().trim().toLowerCase() : 'yes'; 
//         break;
//       }
//     }

//     if (rowIndex === -1) {
//       throw new Error('Task not found: ' + taskId);
//     }

//     const today = new Date();
//     const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

//     // 1. Update the Actual Date column (Column G, index 7)
//     sheet.getRange(rowIndex, 7).setValue(formattedDate);

//     // >>> NEW: Calculate and Update Checker Planned Date (Column P, index 16) <<<
//     const checkerPlannedDate = new Date(today);
//     checkerPlannedDate.setDate(today.getDate() + 3); // Add 3 days
//     const formattedCheckerDate = Utilities.formatDate(checkerPlannedDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
//     sheet.getRange(rowIndex, 16).setValue(formattedCheckerDate);

//     // 2. Update the Proof Document column (Column I, index 9)
//     if (fileUrl && fileUrl.toString().trim() !== '') {
//       sheet.getRange(rowIndex, 9).setValue(fileUrl.toString());
//     }

//     // 3. Update the Remark (Column J, index 10)
//     if (remark && remark.toString().trim() !== '') {
//       sheet.getRange(rowIndex, 10).setValue(remark.toString());
//     }

//     // 🔥 4. FIXED: Auto-update Status to "Completed" in the Google Sheet if Proof is strictly "No"
//     if (proofRequired === 'no') {
//       sheet.getRange(rowIndex, 14).setValue('Completed'); // Column N (Status)
//     }

//     // Force flush to ensure it saves to the sheet immediately
//     SpreadsheetApp.flush();

//     return true;
//   } catch (error) {
//     Logger.log('Error updating task completion: ' + error.toString());
//     throw new Error(error.toString());
//   }
// }

function updateTaskCompletion(taskId, fileUrl, remark) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Checklist');
    const data = sheet.getDataRange().getValues();
    
    let rowIndex = -1;
    let proofRequired = 'yes'; // Default to yes
    let taskDescription = '';  // Variable to store the task description
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][2] && data[i][2].toString() === taskId.toString()) { 
        rowIndex = i + 1; // +1 because sheet rows are 1-indexed
        
        // Extract Task Description from Column E (index 4)
        taskDescription = data[i][4] ? data[i][4].toString().trim() : '';
        
        // Robust case-insensitive read of Column H (index 7) for Proof Required
        proofRequired = data[i][7] ? data[i][7].toString().trim().toLowerCase() : 'yes'; 
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error('Task not found: ' + taskId);
    }

    const today = new Date();
    const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    // 1. Update the Actual Date column (Column G, index 7)
    sheet.getRange(rowIndex, 7).setValue(formattedDate);

    // >>> LOGIC: Calculate and Update Checker Planned Date (Column P, index 16) <<<
    // Check if taskDescription is entirely in CAPITAL LETTERS. 
    // We use /[A-Z]/ to ensure the string actually contains letters.
    const isCapitalEntry = taskDescription === taskDescription.toUpperCase() && /[A-Z]/.test(taskDescription);
    
    // If Capitals (Bank Entry) -> 3 days, else -> 2 days.
    const daysToAdd = isCapitalEntry ? 3 : 2;
    
    const checkerPlannedDate = new Date(today);
    checkerPlannedDate.setDate(today.getDate() + daysToAdd); 
    
    const formattedCheckerDate = Utilities.formatDate(checkerPlannedDate, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    sheet.getRange(rowIndex, 16).setValue(formattedCheckerDate);

    // 2. Update the Proof Document column (Column I, index 9)
    if (fileUrl && fileUrl.toString().trim() !== '') {
      sheet.getRange(rowIndex, 9).setValue(fileUrl.toString());
    }

    // 3. Update the Remark (Column J, index 10)
    if (remark && remark.toString().trim() !== '') {
      sheet.getRange(rowIndex, 10).setValue(remark.toString());
    }

    // 4. Auto-update Status to "Completed" in the Google Sheet if Proof is strictly "No"
    if (proofRequired === 'no') {
      sheet.getRange(rowIndex, 14).setValue('Completed'); // Column N (Status)
    }

    // Force flush to ensure it saves to the sheet immediately
    SpreadsheetApp.flush();

    return true;
  } catch (error) {
    Logger.log('Error updating task completion: ' + error.toString());
    throw new Error(error.toString());
  }
}

function getDoerList() {
  try {
    // Get the active spreadsheet and Doer List sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Doer List');

    // Get the range of emails (A2:A) and names (B2:B)
    const emailRange = sheet.getRange('D2:D' + sheet.getLastRow());
    const nameRange = sheet.getRange('A2:A' + sheet.getLastRow());

    // Get the values
    const emails = emailRange.getValues();
    const names = nameRange.getValues();

    // Create an array of objects with name and email
    const doerList = [];
    for (let i = 0; i < emails.length; i++) {
      // Skip empty rows
      if (emails[i][0] === '' || names[i][0] === '') continue;

      doerList.push({
        email: emails[i][0],
        name: names[i][0]
      });
    }

    return doerList;
  } catch (error) {
    // Log the error and return an empty array
    Logger.log('Error getting doer list: ' + error);
    return [];
  }
}

function addNewTask(name, email, task, freq, date, proofRequired, checker,companyName) {
  try {
    // Get the active spreadsheet and Checklist sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Checklist');

    // Get Holiday List sheet and holiday dates
    const holidaySheet = ss.getSheetByName('Holiday List');
    const holidayDates = [];

    if (holidaySheet) {
      // Get all holiday dates from the sheet
      const holidayRange = holidaySheet.getRange("A2:A" + holidaySheet.getLastRow());
      const holidayValues = holidayRange.getValues();

      // Store all holiday dates in an array for quick lookup
      for (let i = 0; i < holidayValues.length; i++) {
        if (holidayValues[i][0] instanceof Date) {
          holidayDates.push(holidayValues[i][0].toDateString());
        }
      }
    }

    // Get the last task ID from the sheet to determine the next task number
    const lastTaskId = getLastTaskId(sheet);
    let nextTaskNumber = 1; // Default starting number

    if (lastTaskId) {
      // Extract the number from the last task ID (format: Task-X)
      const match = lastTaskId.match(/Task-(\d+)/);
      if (match && match[1]) {
        nextTaskNumber = parseInt(match[1]) + 1;
      }
    }

    // Convert the date string to a Date object for the starting date 
    // (This will be the 1 April 2026 date you select in the form)
    const startDate = new Date(date);

    // --- NEW DYNAMIC END DATE LOGIC ---
    let endDate;
    if (freq === "Yearly") {
      // For Yearly: Generate up to 2028-29 (Ends March 31, 2029)
      // Note: JavaScript months are 0-indexed (0 = Jan, 1 = Feb, 2 = March)
      endDate = new Date(2029, 2, 31); 
    } else {
      // For Daily, Weekly, Monthly, Quarterly: Generate up to May 31, 2027
      // Note: JavaScript months are 0-indexed (4 = May)
      endDate = new Date(2027, 4, 31); 
    }

    // Array to store all rows that will be added
    const rowsToAdd = [];

    // Generate dates based on frequency
    let currentDate = new Date(startDate);

    while (currentDate <= endDate) {
      // Skip Sundays for Daily frequency
      if (freq === "Daily" && currentDate.getDay() === 0) {
        currentDate.setDate(currentDate.getDate() + 1);
        continue;
      }

      // Skip holidays
      if (holidayDates.includes(currentDate.toDateString())) {
        currentDate.setDate(currentDate.getDate() + 1);
        continue;
      }

      // Generate a sequential task ID
      const taskId = `Task-${nextTaskNumber}`;
      nextTaskNumber++;

      // Create the row data with the proof required field (column N)
     const rowData = [
        email,          // A (0)
        name,           // B (1)
        taskId,         // C (2)
        freq,           // D (3)
        task,           // E (4)
        new Date(currentDate), // F (5)
        '',             // G (6)
        proofRequired,  // H (7)
        '',             // I (8)
        '',             // J (9)
        '',             // K (10)
        0,              // L (11)
        checker,        // M (12)
        'Pending',      // N (13)
        '',             // O (14)
        '',             // P (15)
        companyName     // Q (16) << ADD THIS
      ];

      rowsToAdd.push(rowData);

      // Increment date based on frequency
      switch (freq) {
        case "Daily":
          currentDate.setDate(currentDate.getDate() + 1);
          break;
        case "Weekly":
          currentDate.setDate(currentDate.getDate() + 7);
          break;
        case "Monthly":
          currentDate.setMonth(currentDate.getMonth() + 1);
          break;
        case "Quarterly":
          currentDate.setMonth(currentDate.getMonth() + 3);
          break;
        case "Yearly":
          currentDate.setFullYear(currentDate.getFullYear() + 1);
          break;
        default:
          // For any other frequency, just add one instance
          currentDate = new Date(endDate.getTime() + 1);
          break;
      }
    }

    // If we have rows to add, add them all at once
    if (rowsToAdd.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length)
        .setValues(rowsToAdd);
    }

    // Return success
    return true;
  } catch (error) {
    // Log the error and return false
    Logger.log('Error adding new tasks: ' + error);
    return false;
  }
}

function getLastTaskId(sheet) {
  // Get all task IDs from column C (index 2)
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return null; // Sheet is empty or only has headers
  }

  // Get all task IDs from the Task ID column (column C, index 2)
  const taskIdRange = sheet.getRange(2, 3, lastRow - 1, 1);
  const taskIds = taskIdRange.getValues();

  // Filter out empty values and find the last task ID that matches the pattern
  let lastTaskId = null;
  for (let i = taskIds.length - 1; i >= 0; i--) {
    const id = taskIds[i][0];
    if (id && typeof id === 'string' && id.match(/Task-\d+/)) {
      lastTaskId = id;
      break;
    }
  }

  return lastTaskId;
}




function addMDRemarkChecklist(taskId, remark) {
  try {
    // Get the active spreadsheet and Checklist sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Checklist');

    // Find the row with the matching task ID
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 0; i < data.length; i++) {
      if (data[i][2] === taskId) { // Column C contains Task ID
        rowIndex = i + 1; // +1 because sheet rows are 1-indexed
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error('Task not found: ' + taskId);
    }

    // Get current MD remarks, if any (column J, index 9)
    const currentRemarks = data[rowIndex - 1][9] || '';

    // Add new remark to existing remarks
    let updatedRemarks = '';
    if (currentRemarks.trim()) {
      updatedRemarks = currentRemarks + '\n' + remark;
    } else {
      updatedRemarks = remark;
    }

    // Update the MD Remarks column (column J, index 9)
    sheet.getRange(rowIndex, 10).setValue(updatedRemarks);

    return true;
  } catch (error) {
    Logger.log('Error adding MD remark: ' + error.message);
    throw new Error('Failed to add MD remark: ' + error.message);
  }
}


function addEARemarkChecklist(taskId, remark) {
  try {
    // Get the active spreadsheet and Checklist sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Checklist');

    // Find the row with the matching task ID
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 0; i < data.length; i++) {
      if (data[i][2] === taskId) { // Column C contains Task ID
        rowIndex = i + 1; // +1 because sheet rows are 1-indexed
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error('Task not found: ' + taskId);
    }

    // Get current EA remarks, if any (column K, index 10)
    const currentRemarks = data[rowIndex - 1][10] || '';

    // Add new remark to existing remarks
    let updatedRemarks = '';
    if (currentRemarks.trim()) {
      updatedRemarks = currentRemarks + '\n' + remark;
    } else {
      updatedRemarks = remark;
    }

    // Update the EA Remarks column (column K, index 10)
    sheet.getRange(rowIndex, 11).setValue(updatedRemarks);

    // Get current follow up count from column L (index 11)
    let followUpCount = data[rowIndex - 1][11] || 0;

    // If follow up count is not a number, set it to 0
    if (isNaN(followUpCount)) {
      followUpCount = 0;
    }

    // Increment follow up count
    followUpCount = parseInt(followUpCount) + 1;

    // Update follow up count in column L (index 11)
    sheet.getRange(rowIndex, 12).setValue(followUpCount);

    return true;
  } catch (error) {
    Logger.log('Error adding EA remark: ' + error.message);
    throw new Error('Failed to add EA remark: ' + error.message);
  }
}



function addBulkMDRemarks(taskIds, remark) {
  try {
    // Get the active spreadsheet and Checklist sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Checklist');

    // Get all data from the sheet
    const data = sheet.getDataRange().getValues();

    // Track success and failures
    const results = {
      success: true,
      processed: 0,
      failed: 0,
      errors: []
    };

    // Process each task ID
    taskIds.forEach(taskId => {
      try {
        // Find the row with the matching task ID
        let rowIndex = -1;

        for (let i = 0; i < data.length; i++) {
          if (data[i][2] === taskId) { // Column C contains Task ID
            rowIndex = i + 1; // +1 because sheet rows are 1-indexed
            break;
          }
        }

        if (rowIndex === -1) {
          // Task not found, add to errors
          results.failed++;
          results.errors.push(`Task ID not found: ${taskId}`);
          return;
        }

        // Get current MD remarks, if any (column J, index 9)
        const currentRemarks = data[rowIndex - 1][9] || '';

        // Add new remark to existing remarks
        let updatedRemarks = '';
        if (currentRemarks.trim()) {
          updatedRemarks = currentRemarks + '\n' + remark;
        } else {
          updatedRemarks = remark;
        }

        // Update the MD Remarks column (column J, index 9)
        sheet.getRange(rowIndex, 10).setValue(updatedRemarks);

        // Increment success counter
        results.processed++;
      } catch (error) {
        // Log the error and continue with next task
        results.failed++;
        results.errors.push(`Error updating task ${taskId}: ${error.message}`);
        Logger.log(`Error updating task ${taskId}: ${error.message}`);
      }
    });

    // Set overall success based on failures
    if (results.failed > 0) {
      results.success = false;
    }

    return results;
  } catch (error) {
    Logger.log('Error adding bulk MD remark: ' + error.message);
    return {
      success: false,
      message: 'Failed to add bulk MD remark: ' + error.message
    };
  }
}



function addBulkEARemarks(taskIds, remark) {
  try {
    // Get the active spreadsheet and Checklist sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Checklist');

    // Get all data from the sheet
    const data = sheet.getDataRange().getValues();

    // Track success and failures
    const results = {
      success: true,
      processed: 0,
      failed: 0,
      errors: []
    };

    // Process each task ID
    taskIds.forEach(taskId => {
      try {
        // Find the row with the matching task ID
        let rowIndex = -1;

        for (let i = 0; i < data.length; i++) {
          if (data[i][2] === taskId) { // Column C contains Task ID
            rowIndex = i + 1; // +1 because sheet rows are 1-indexed
            break;
          }
        }

        if (rowIndex === -1) {
          // Task not found, add to errors
          results.failed++;
          results.errors.push(`Task ID not found: ${taskId}`);
          return;
        }

        // Get current EA remarks, if any (column K, index 10)
        const currentRemarks = data[rowIndex - 1][10] || '';

        // Add new remark to existing remarks
        let updatedRemarks = '';
        if (currentRemarks.trim()) {
          updatedRemarks = currentRemarks + '\n' + remark;
        } else {
          updatedRemarks = remark;
        }

        // Update the EA Remarks column (column K, index 10)
        sheet.getRange(rowIndex, 11).setValue(updatedRemarks);

        // Get current follow up count from column L (index 11)
        let followUpCount = data[rowIndex - 1][11] || 0;

        // If follow up count is not a number, set it to 0
        if (isNaN(followUpCount)) {
          followUpCount = 0;
        }

        // Increment follow up count
        followUpCount = parseInt(followUpCount) + 1;

        // Update follow up count in column L (index 11)
        sheet.getRange(rowIndex, 12).setValue(followUpCount);

        // Increment success counter
        results.processed++;
      } catch (error) {
        // Log the error and continue with next task
        results.failed++;
        results.errors.push(`Error updating task ${taskId}: ${error.message}`);
        Logger.log(`Error updating task ${taskId}: ${error.message}`);
      }
    });

    // Set overall success based on failures
    if (results.failed > 0) {
      results.success = false;
    }

    return results;
  } catch (error) {
    Logger.log('Error adding bulk EA remark: ' + error.message);
    return {
      success: false,
      message: 'Failed to add bulk EA remark: ' + error.message
    };
  }
}


function markTaskAsNonFunctional(taskId) {
  try {
    // Get the active spreadsheet and Checklist sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Checklist');

    // Find the row with the matching task ID
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 0; i < data.length; i++) {
      if (data[i][2] === taskId) { // Column C contains Task ID
        rowIndex = i + 1; // +1 because sheet rows are 1-indexed
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error('Task not found: ' + taskId);
    }

    // Get today's date in DD/MM/YYYY format
    const today = new Date();
    const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    // Update the Actual Date column (column G, index 6)
    sheet.getRange(rowIndex, 7).setValue(formattedDate);

    // Update the Proof Document column (column I, index 8) with "Non-Functional" text
    sheet.getRange(rowIndex, 9).setValue("Non-Functional");

    return true;
  } catch (error) {
    Logger.log('Error marking task as non-functional: ' + error);
    return false;
  }
}



function deleteTaskByDoerAndTaskName(doerName, taskName, token) {
  try {
    // Verify user authorization
    const userSession = getSessionForCall_(token);
    if (!userSession) {
      throw new Error("User not authenticated");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Checklist");

    if (!sheet) {
      throw new Error("Checklist sheet not found");
    }

    const data = sheet.getDataRange().getValues();

    // Find ALL rows with matching doer name and task name
    const matchingRows = [];

    for (let i = 1; i < data.length; i++) {
      // Check if both doer name (Column B, index 1) and task name (Column E, index 4) match
      if (data[i][1] === doerName && data[i][4] === taskName) {
        // Check authorization - only Admin or the assigned doer can delete
        if (userSession.role !== 'Admin' &&
          data[i][1].toLowerCase() !== userSession.name.toLowerCase()) {
          throw new Error("You are not authorized to delete this task");
        }

        // Add to matching rows list
        matchingRows.push(i);
      }
    }

    if (matchingRows.length === 0) {
      throw new Error(`No tasks found for doer "${doerName}" with task "${taskName}"`);
    }

    // Delete rows in reverse order to avoid shifting issues
    matchingRows.sort((a, b) => b - a);

    // Count deletions
    let deletedCount = 0;

    // Delete all matching rows
    for (const rowIndex of matchingRows) {
      sheet.deleteRow(rowIndex + 1); // +1 because sheet rows are 1-indexed
      deletedCount++;
    }

    // Log the action
    Logger.log(`${deletedCount} task(s) for doer "${doerName}" with task "${taskName}" deleted by ${userSession.name}`);

    return true;
  } catch (error) {
    Logger.log("Error deleting task: " + error);
    throw error;
  }
}


function updateTaskByDoerAndTaskName(doerName, taskName, newDoerName, newTaskDesc, token) {
  try {
    // Verify user authorization
    const userSession = getSessionForCall_(token);
    if (!userSession) {
      throw new Error("User not authenticated");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Checklist");

    if (!sheet) {
      throw new Error("Checklist sheet not found");
    }

    // Find user's email from the DoerList sheet
    const doerListSheet = ss.getSheetByName("Doer List");

    if (!doerListSheet) {
      throw new Error("Doer List sheet not found");
    }

    const doerListData = doerListSheet.getDataRange().getValues();
    let doerEmail = "";
    let doerFound = false;

    // Find the email for the new doer name
    for (let i = 1; i < doerListData.length; i++) {
      if (doerListData[i][1] === newDoerName) { // Name is in column B
        doerEmail = doerListData[i][0];         // Email is in column A
        doerFound = true;
        break;
      }
    }

    if (!doerFound) {
      throw new Error("New doer not found in the system");
    }

    // Find ALL rows with matching doer name and task name
    const data = sheet.getDataRange().getValues();
    const matchingRows = [];

    for (let i = 1; i < data.length; i++) {
      // Check if both doer name (Column B, index 1) and task name (Column E, index 4) match
      if (data[i][1] === doerName && data[i][4] === taskName) {
        // Check authorization - only Admin or the assigned doer can update
        if (userSession.role !== 'Admin' &&
          data[i][1].toLowerCase() !== userSession.name.toLowerCase()) {
          throw new Error("You are not authorized to update this task");
        }

        // Store row index and original data
        matchingRows.push({
          rowIndex: i,
          originalData: data[i]
        });
      }
    }

    if (matchingRows.length === 0) {
      throw new Error(`No tasks found for doer "${doerName}" with task "${taskName}"`);
    }

    // Get current date/time for remarks
    const now = new Date();
    const day = now.getDate().toString().padStart(2, '0');
    const month = (now.getMonth() + 1).toString().padStart(2, '0');
    const year = now.getFullYear();
    const hours = now.getHours().toString().padStart(2, '0');
    const minutes = now.getMinutes().toString().padStart(2, '0');

    // Count updates
    let updatedCount = 0;

    // Update ALL matching rows
    for (const match of matchingRows) {
      const { rowIndex, originalData } = match;

      // Update task information
      sheet.getRange(rowIndex + 1, 2).setValue(newDoerName); // Name (Column B)
      sheet.getRange(rowIndex + 1, 1).setValue(doerEmail);   // Email (Column A)
      sheet.getRange(rowIndex + 1, 5).setValue(newTaskDesc); // Task (Column E)

      updatedCount++;
    }

    // Log the action
    Logger.log(`${updatedCount} task(s) for doer "${doerName}" with task "${taskName}" updated to "${newDoerName}" with task "${newTaskDesc}" by ${userSession.name}`);

    return true;
  } catch (error) {
    Logger.log("Error updating task: " + error);
    throw error;
  }
}



function getDashboardData() {
  try {
    // Get user session data using browser-based cache
    const userSession = getUserSession();

    // If no user session, return empty array
    if (!userSession) {
      return [];
    }

    // Get user info
    const userRole = userSession.role;
    const userName = userSession.name;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Dashboard');

    if (!sheet) {
      console.error('Dashboard sheet not found');
      return [];
    }

    const range = sheet.getDataRange();
    const values = range.getValues();

    // Skip header row
    const dataRows = values.slice(1);

    // Filter out empty rows, apply role-based access control, and format date
    return dataRows
      .filter(row => {
        // Skip empty rows
        if (row[0] === '') return false;

        // If user is Admin, show all rows
        if (userRole === 'Admin') return true;

        // For Doer users, only show rows where their name matches column A (index 0)
        return row[0].toString().trim().toLowerCase() === userName.trim().toLowerCase();
      })
      .map(row => {
        // Create a copy of the row to avoid modifying the original data
        const formattedRow = [...row];

        // Format the date in column C (index 2)
        if (formattedRow[2] instanceof Date && !isNaN(formattedRow[2])) {
          const date = formattedRow[2];
          const day = date.getDate().toString().padStart(2, '0');
          const month = (date.getMonth() + 1).toString().padStart(2, '0');
          const year = date.getFullYear();
          formattedRow[2] = `${day}/${month}/${year}`;
        }
        return formattedRow;
      });
  } catch (error) {
    console.error('Error in getScoreCardData:', error);
    return [];
  }
}


function savePromiseScores(scoresData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Promise Score');

    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet('Promise Score');
      // Add headers
      sheet.getRange(1, 1, 1, 5).setValues([['Name', 'Week', 'Green', 'Yellow', 'Red']]);
    }

    // Get current week number
    const currentWeek = getCurrentWeekNumber();

    // Get existing data to avoid duplicates
    const existingData = sheet.getDataRange().getValues();
    const existingKeys = existingData.slice(1).map(row => `${row[0]}_${row[1]}`);

    // Prepare data for insertion
    const dataToInsert = [];

    scoresData.forEach(score => {
      const key = `${score.name}_${currentWeek}`;
      const existingRowIndex = existingKeys.indexOf(key);

      if (existingRowIndex !== -1) {
        // Update existing row
        const rowNumber = existingRowIndex + 2; // +2 because of header and 0-based index
        sheet.getRange(rowNumber, 3, 1, 3).setValues([[score.green, score.yellow, score.red]]);
      } else {
        // Add new row with current week
        dataToInsert.push([score.name, currentWeek, score.green, score.yellow, score.red]);
      }
    });

    // Insert new rows if any
    if (dataToInsert.length > 0) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, dataToInsert.length, 5).setValues(dataToInsert);
    }

    return { success: true, message: 'Scores saved successfully' };

  } catch (error) {
    console.error('Error saving scores:', error);
    throw new Error('Failed to save scores: ' + error.toString());
  }
}

// Helper function to get current week number in format "Week XX/YYYY"
function getCurrentWeekNumber() {
  const today = new Date();
  const year = today.getFullYear();

  // Get first day of the year
  const firstDay = new Date(year, 0, 1);

  // Calculate the number of milliseconds since the first day of the year
  const pastDaysOfYear = (today - firstDay) / 86400000;

  // Calculate the week number (add 1 because week starts from 1, not 0)
  const weekNumber = Math.ceil((pastDaysOfYear + firstDay.getDay() + 1) / 7);

  return `Week ${weekNumber}/${year}`;
}



function sendApprovalWaitingEmails() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const approvalWaitingSheet = ss.getSheetByName('Approval Waiting');

    if (!approvalWaitingSheet) {
      Logger.log('Error: Approval Waiting sheet not found');
      return 0;
    }

    // Get all data from Approval Waiting sheet
    const approvalData = approvalWaitingSheet.getDataRange().getValues();

    // Skip header row and check if there's data
    if (!approvalData || approvalData.length <= 1) {
      Logger.log('No approval waiting tasks found to send');
      return 0;
    }

    // Group tasks by assignee email
    const tasksByAssigneeEmail = {};

    // Start from row 1 (skip header at row 0)
    for (let i = 1; i < approvalData.length; i++) {
      const row = approvalData[i];

      // Check if row has enough columns
      if (!row || row.length < 6) {
        Logger.log(`Skipping row ${i}: insufficient data`);
        continue;
      }

      const taskId = row[0];          // Column A - Task Id
      const assigneeEmail = row[1];   // Column B - Assignee Email
      const assignee = row[2];        // Column C - Assignee
      const doerName = row[3];        // Column D - Doer Name
      const taskDescription = row[4]; // Column E - Task
      const plannedDate = row[5];     // Column F - Planned Date

      // Skip rows with no assignee email or task ID
      if (!assigneeEmail || !taskId) {
        Logger.log(`Skipping row ${i}: missing assignee email or task ID`);
        continue;
      }

      // Format planned date
      let formattedPlannedDate = 'No planned date';
      if (plannedDate) {
        if (plannedDate instanceof Date) {
          formattedPlannedDate = Utilities.formatDate(plannedDate, Session.getScriptTimeZone(), 'MMM dd, yyyy');
        } else if (typeof plannedDate === 'string' && plannedDate.trim() !== '') {
          // Try to parse string date
          try {
            const parsedDate = new Date(plannedDate);
            if (!isNaN(parsedDate.getTime())) {
              formattedPlannedDate = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), 'MMM dd, yyyy');
            } else {
              formattedPlannedDate = plannedDate; // Use original string if can't parse
            }
          } catch (e) {
            formattedPlannedDate = plannedDate; // Use original string if error
          }
        }
      }

      // Create task object
      const task = {
        taskId: taskId || 'N/A',
        assignee: assignee || 'Unknown',
        doerName: doerName || 'Unknown',
        taskDescription: taskDescription || 'No description',
        plannedDate: plannedDate, // Keep original for sorting
        formattedPlannedDate: formattedPlannedDate
      };

      // Add task to the assignee email group
      if (!tasksByAssigneeEmail[assigneeEmail]) {
        tasksByAssigneeEmail[assigneeEmail] = {
          assigneeName: assignee || 'Unknown',
          tasks: []
        };
      }

      tasksByAssigneeEmail[assigneeEmail].tasks.push(task);
    }

    // Check if we have any tasks to send
    const emailCount = Object.keys(tasksByAssigneeEmail).length;
    if (emailCount === 0) {
      Logger.log('No valid approval waiting tasks found to send emails');
      return 0;
    }

    // Send emails to each assignee with their approval waiting tasks
    let emailsSent = 0;

    for (const assigneeEmail in tasksByAssigneeEmail) {
      try {
        const assigneeInfo = tasksByAssigneeEmail[assigneeEmail];
        const assigneeName = assigneeInfo.assigneeName;
        const tasks = assigneeInfo.tasks;

        // Validate tasks array
        if (!tasks || !Array.isArray(tasks) || tasks.length === 0) {
          Logger.log(`Skipping email to ${assigneeEmail}: no valid tasks`);
          continue;
        }

        // Sort tasks by planned date (ascending)
        tasks.sort((a, b) => {
          const dateA = a.plannedDate instanceof Date ? a.plannedDate : new Date(a.plannedDate);
          const dateB = b.plannedDate instanceof Date ? b.plannedDate : new Date(b.plannedDate);

          // Handle invalid dates
          const timeA = isNaN(dateA.getTime()) ? new Date('9999-12-31').getTime() : dateA.getTime();
          const timeB = isNaN(dateB.getTime()) ? new Date('9999-12-31').getTime() : dateB.getTime();

          return timeA - timeB;
        });

        // Check how many tasks are overdue (planned date has passed)
        const today = new Date();
        const overdueTasks = tasks.filter(task => {
          if (task.plannedDate instanceof Date) {
            const diffTime = task.plannedDate.getTime() - today.getTime();
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
            return diffDays < 0;
          } else if (typeof task.plannedDate === 'string' && task.plannedDate.trim() !== '') {
            try {
              const parsedDate = new Date(task.plannedDate);
              if (!isNaN(parsedDate.getTime())) {
                const diffTime = parsedDate.getTime() - today.getTime();
                const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                return diffDays < 0;
              }
            } catch (e) {
              // Ignore parsing errors
            }
          }
          return false;
        });

        // Send the email with all approval waiting tasks
        sendApprovalWaitingEmail(assigneeEmail, assigneeName, tasks, overdueTasks.length);
        emailsSent++;

      } catch (emailError) {
        Logger.log(`Failed to send approval email to ${assigneeEmail}: ${emailError.toString()}`);
      }
    }

    Logger.log(`Successfully sent approval waiting emails to ${emailsSent} assignees`);
    return emailsSent;

  } catch (error) {
    Logger.log('Failed to send approval waiting emails: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    return 0;
  }
}

function sendApprovalWaitingEmail(assigneeEmail, assigneeName, tasks, overdueCount) {
  try {
    // Validate input parameters
    if (!assigneeEmail || typeof assigneeEmail !== 'string') {
      throw new Error('Invalid assignee email address');
    }

    if (!tasks || !Array.isArray(tasks)) {
      throw new Error('Tasks must be an array');
    }

    if (tasks.length === 0) {
      Logger.log(`No tasks to send for ${assigneeEmail}`);
      return;
    }

    const safeAssigneeName = assigneeName || 'User';
    const safeOverdueCount = typeof overdueCount === 'number' ? overdueCount : 0;

    // Prepare email content
    const subject = safeOverdueCount > 0
      ? `🚨 ${safeOverdueCount} Overdue Task${safeOverdueCount > 1 ? 's' : ''} - Tasks Awaiting Your Approval`
      : 'Tasks Awaiting Your Approval';

    let body = `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Tasks Awaiting Approval</title>
        <style>
          body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333333;
            margin: 0;
            padding: 0;
          }
          .container {
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
          }
          .header {
            background-color: #ff6b35;
            color: white;
            padding: 20px;
            text-align: center;
            border-radius: 5px 5px 0 0;
          }
          .content {
            background-color: #ffffff;
            padding: 20px;
            border-left: 1px solid #dddddd;
            border-right: 1px solid #dddddd;
          }
          .summary {
            background-color: #fff8e1;
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 4px;
            border-left: 4px solid #ff6b35;
          }
          .status-notice {
            background-color: #e8f5e8;
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 4px;
            border-left: 4px solid #4caf50;
            font-weight: bold;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
          }
          th {
            background-color: #f2f2f2;
            text-align: left;
            padding: 12px;
            border: 1px solid #dddddd;
          }
          td {
            padding: 12px;
            border: 1px solid #dddddd;
            vertical-align: top;
          }
          tr:nth-child(even) {
            background-color: #f9f9f9;
          }
          .overdue {
            background-color: #ffebee !important;
          }
          .footer {
            background-color: #f5f5f5;
            padding: 15px;
            text-align: center;
            font-size: 12px;
            color: #777777;
            border-radius: 0 0 5px 5px;
            border: 1px solid #dddddd;
          }
          .button {
            display: inline-block;
            background-color: #ff6b35;
            color: white;
            padding: 10px 20px;
            text-decoration: none;
            border-radius: 4px;
            margin-top: 15px;
          }
          .approve-button {
            background-color: #4caf50;
            color: white;
            padding: 8px 16px;
            text-decoration: none;
            border-radius: 4px;
            font-size: 12px;
            display: inline-block;
            margin-top: 5px;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h1>⏳ Tasks Awaiting Your Approval</h1>
          </div>
          
          <div class="content">
            <p>Hello ${safeAssigneeName},</p>
            
            <div class="status-notice">
              📋 <strong>Status Update:</strong> Doers have completed their tasks and are waiting for your approval to proceed.
            </div>
            
            <div class="summary">
              <p><strong>Approval Summary:</strong> You have ${tasks.length} task${tasks.length !== 1 ? 's' : ''} awaiting your approval${safeOverdueCount > 0 ? `, with ${safeOverdueCount} overdue` : ''}.</p>
            </div>
            
            <p>The following tasks have been completed by doers and require your review and approval:</p>
            
            <table>
              <tr>
                <th>Task ID</th>
                <th>Doer Name</th>
                <th>Task Description</th>
                <th>Planned Date</th>
                <th>Action Required</th>
              </tr>
    `;

    // Add row for each task
    const today = new Date();

    tasks.forEach((task, index) => {
      try {
        // Safely access task properties
        const taskId = task.taskId || 'N/A';
        const doerName = task.doerName || 'Unknown';
        const taskDescription = task.taskDescription || 'No description';
        const formattedPlannedDate = task.formattedPlannedDate || 'No planned date';

        // Calculate days difference for urgency
        let rowClass = '';
        let urgencyText = '';

        if (task.plannedDate) {
          const plannedDate = task.plannedDate instanceof Date ? task.plannedDate : new Date(task.plannedDate);

          if (!isNaN(plannedDate.getTime())) {
            const diffTime = plannedDate.getTime() - today.getTime();
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

            if (diffDays < 0) {
              rowClass = 'overdue';
              urgencyText = ' <strong>(OVERDUE)</strong>';
            }
          }
        }

        body += `
                <tr class="${rowClass}">
                  <td>${taskId}</td>
                  <td>${doerName}</td>
                  <td>${taskDescription}</td>
                  <td>${formattedPlannedDate}${urgencyText}</td>
                  <td>
                    <div style="text-align: center;">
                      <a href="https://script.google.com/a/codearchitects.in/macros/s/AKfycbzawv8hZ0lQRGOhemghTb4sjFEZdjdoE-C2iG32tHdG/dev?page=Index" class="approve-button">Review & Approve</a>
                    </div>
                  </td>
                </tr>
        `;

      } catch (taskError) {
        Logger.log(`Error processing approval task ${index} for ${assigneeEmail}: ${taskError.toString()}`);
      }
    });

    body += `
            </table>
            
            <div class="status-notice">
              ✅ <strong>Important:</strong> These tasks have been completed by the respective doers. Please review the work and take appropriate action to approve or provide feedback.
            </div>
            
            <p><strong>Next Steps:</strong></p>
            <ul>
              <li>Review each completed task thoroughly</li>
              <li>Approve tasks that meet the required standards</li>
              <li>Provide feedback for tasks that need revision</li>
              <li>Prioritize overdue approvals to maintain workflow efficiency</li>
            </ul>
            
            <div style="text-align: center;">
              <a href="https://script.google.com/a/codearchitects.in/macros/s/AKfycbzawv8hZ0lQRGOhemghTb4sjFEZdjdoE-C2iG32tHdG/dev?page=Index" class="button">Open Task Dashboard</a>
            </div>
            
            <p>If you have any questions about any of these tasks, please contact the respective doers or your supervisor.</p>
            <p>Thank you for your prompt attention to these approval requests.</p>
          </div>
          
          <div class="footer">
            <p>This is an automated approval notification sent on ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy')}.</p>
            <p>© ${new Date().getFullYear()} Your Company Name</p>
          </div>
        </div>
      </body>
    </html>
    `;

    // Send the email
    GmailApp.sendEmail(assigneeEmail, subject, "", {
      htmlBody: body
    });

    Logger.log('Approval waiting email sent to: ' + assigneeEmail);

  } catch (error) {
    Logger.log(`Failed to send approval email to ${assigneeEmail}: ${error.toString()}`);
    Logger.log('Stack trace: ' + error.stack);
    throw error; // Re-throw to be caught by calling function
  }
}



function sendEmailChecklistReminders() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const emailChecklistSheet = ss.getSheetByName('Email Checklist');

    if (!emailChecklistSheet) {
      Logger.log('Error: Email Checklist sheet not found');
      return 0;
    }

    // Get all data from Email Checklist sheet
    const checklistData = emailChecklistSheet.getDataRange().getValues();

    // Skip header row and check if there's data
    if (!checklistData || checklistData.length <= 1) {
      Logger.log('No checklist tasks found to send');
      return 0;
    }

    // Group tasks by email
    const tasksByEmail = {};

    // Start from row 1 (skip header at row 0)
    for (let i = 1; i < checklistData.length; i++) {
      const row = checklistData[i];

      // Check if row has enough columns
      if (!row || row.length < 7) {
        Logger.log(`Skipping row ${i}: insufficient data`);
        continue;
      }

      const name = row[0];            // Column A - Name
      const email = row[1];           // Column B - Email
      const taskId = row[2];          // Column C - Task ID
      const frequency = row[3];       // Column D - Freq
      const taskType = row[4];        // Column E - Task Type
      const taskDescription = row[5]; // Column F - Task
      const plannedDate = row[6];     // Column G - Planned Date

      // Skip rows with no email or task ID
      if (!email || !taskId) {
        Logger.log(`Skipping row ${i}: missing email or task ID`);
        continue;
      }

      // Format planned date
      let formattedPlannedDate = 'No planned date';
      if (plannedDate) {
        if (plannedDate instanceof Date) {
          formattedPlannedDate = Utilities.formatDate(plannedDate, Session.getScriptTimeZone(), 'MMM dd, yyyy');
        } else if (typeof plannedDate === 'string' && plannedDate.trim() !== '') {
          // Try to parse string date
          try {
            const parsedDate = new Date(plannedDate);
            if (!isNaN(parsedDate.getTime())) {
              formattedPlannedDate = Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), 'MMM dd, yyyy');
            } else {
              formattedPlannedDate = plannedDate; // Use original string if can't parse
            }
          } catch (e) {
            formattedPlannedDate = plannedDate; // Use original string if error
          }
        }
      }

      // Create task object
      const task = {
        taskId: taskId || 'N/A',
        name: name || 'Unknown',
        frequency: frequency || 'N/A',
        taskType: taskType || 'N/A',
        taskDescription: taskDescription || 'No description',
        plannedDate: plannedDate, // Keep original for sorting
        formattedPlannedDate: formattedPlannedDate
      };

      // Add task to the email group
      if (!tasksByEmail[email]) {
        tasksByEmail[email] = {
          userName: name || 'Unknown',
          tasks: []
        };
      }

      tasksByEmail[email].tasks.push(task);
    }

    // Check if we have any tasks to send
    const emailCount = Object.keys(tasksByEmail).length;
    if (emailCount === 0) {
      Logger.log('No valid checklist tasks found to send emails');
      return 0;
    }

    // Send emails to each user with their checklist tasks
    let emailsSent = 0;

    for (const email in tasksByEmail) {
      try {
        const userInfo = tasksByEmail[email];
        const userName = userInfo.userName;
        const tasks = userInfo.tasks;

        // Validate tasks array
        if (!tasks || !Array.isArray(tasks) || tasks.length === 0) {
          Logger.log(`Skipping email to ${email}: no valid tasks`);
          continue;
        }

        // Sort tasks by planned date (ascending), then by task type
        tasks.sort((a, b) => {
          const dateA = a.plannedDate instanceof Date ? a.plannedDate : new Date(a.plannedDate);
          const dateB = b.plannedDate instanceof Date ? b.plannedDate : new Date(b.plannedDate);

          // Handle invalid dates
          const timeA = isNaN(dateA.getTime()) ? new Date('9999-12-31').getTime() : dateA.getTime();
          const timeB = isNaN(dateB.getTime()) ? new Date('9999-12-31').getTime() : dateB.getTime();

          // First sort by date
          if (timeA !== timeB) {
            return timeA - timeB;
          }

          // Then sort by task type
          return (a.taskType || '').localeCompare(b.taskType || '');
        });

        // Check how many tasks are overdue (planned date has passed)
        const today = new Date();
        const overdueTasks = tasks.filter(task => {
          if (task.plannedDate instanceof Date) {
            const diffTime = task.plannedDate.getTime() - today.getTime();
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
            return diffDays < 0;
          } else if (typeof task.plannedDate === 'string' && task.plannedDate.trim() !== '') {
            try {
              const parsedDate = new Date(task.plannedDate);
              if (!isNaN(parsedDate.getTime())) {
                const diffTime = parsedDate.getTime() - today.getTime();
                const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
                return diffDays < 0;
              }
            } catch (e) {
              // Ignore parsing errors
            }
          }
          return false;
        });

        // Count critical tasks
        const criticalTasks = tasks.filter(task =>
          (task.taskType || '').toLowerCase().includes('critical')
        );

        // Send the email with all checklist tasks
        sendEmailChecklistReminder(email, userName, tasks, overdueTasks.length, criticalTasks.length);
        emailsSent++;

      } catch (emailError) {
        Logger.log(`Failed to send checklist email to ${email}: ${emailError.toString()}`);
      }
    }

    Logger.log(`Successfully sent checklist reminder emails to ${emailsSent} users`);
    return emailsSent;

  } catch (error) {
    Logger.log('Failed to send checklist reminder emails: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    return 0;
  }
}

function sendEmailChecklistReminder(email, userName, tasks, overdueCount, criticalCount) {
  try {
    // Validate input parameters
    if (!email || typeof email !== 'string') {
      throw new Error('Invalid email address');
    }

    if (!tasks || !Array.isArray(tasks)) {
      throw new Error('Tasks must be an array');
    }

    if (tasks.length === 0) {
      Logger.log(`No tasks to send for ${email}`);
      return;
    }

    const safeUserName = userName || 'User';
    const safeOverdueCount = typeof overdueCount === 'number' ? overdueCount : 0;
    const safeCriticalCount = typeof criticalCount === 'number' ? criticalCount : 0;

    // Prepare email content
    let subject = 'Your Checklist - Pending Tasks Reminder';
    if (safeOverdueCount > 0 && safeCriticalCount > 0) {
      subject = `🚨 ${safeOverdueCount} Overdue & ${safeCriticalCount} Critical - Checklist Tasks Reminder`;
    } else if (safeOverdueCount > 0) {
      subject = `⚠️ ${safeOverdueCount} Overdue Task${safeOverdueCount > 1 ? 's' : ''} - Checklist Reminder`;
    } else if (safeCriticalCount > 0) {
      subject = `🔴 ${safeCriticalCount} Critical Task${safeCriticalCount > 1 ? 's' : ''} - Checklist Reminder`;
    }

    let body = `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Checklist Tasks Reminder</title>
        <style>
          body {
            font-family: Arial, sans-serif;
            line-height: 1.6;
            color: #333333;
            margin: 0;
            padding: 0;
          }
          .container {
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
          }
          .header {
            background-color: #6c5ce7;
            color: white;
            padding: 20px;
            text-align: center;
            border-radius: 5px 5px 0 0;
          }
          .content {
            background-color: #ffffff;
            padding: 20px;
            border-left: 1px solid #dddddd;
            border-right: 1px solid #dddddd;
          }
          .summary {
            background-color: #f8f7ff;
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 4px;
            border-left: 4px solid #6c5ce7;
          }
          .status-notice {
            background-color: #e8f4fd;
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 4px;
            border-left: 4px solid #0984e3;
            font-weight: bold;
          }
          table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
          }
          th {
            background-color: #f2f2f2;
            text-align: left;
            padding: 12px;
            border: 1px solid #dddddd;
          }
          td {
            padding: 12px;
            border: 1px solid #dddddd;
            vertical-align: top;
          }
          tr:nth-child(even) {
            background-color: #f9f9f9;
          }
          .overdue {
            background-color: #ffebee !important;
          }
          .critical {
            background-color: #fff3e0 !important;
          }
          .overdue-critical {
            background-color: #ffcdd2 !important;
          }
          .task-type-badge {
            display: inline-block;
            padding: 4px 8px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: bold;
            text-transform: uppercase;
          }
          .critical-badge {
            background-color: #ff5722;
            color: white;
          }
          .normal-badge {
            background-color: #2196f3;
            color: white;
          }
          .frequency-badge {
            display: inline-block;
            padding: 2px 6px;
            border-radius: 8px;
            font-size: 10px;
            background-color: #e0e0e0;
            color: #333;
          }
          .footer {
            background-color: #f5f5f5;
            padding: 15px;
            text-align: center;
            font-size: 12px;
            color: #777777;
            border-radius: 0 0 5px 5px;
            border: 1px solid #dddddd;
          }
          .button {
            display: inline-block;
            background-color: #6c5ce7;
            color: white;
            padding: 10px 20px;
            text-decoration: none;
            border-radius: 4px;
            margin-top: 15px;
          }
        </style>
      </head>
      <body>
        <div class="container">
          <div class="header">
            <h1>📋 Your Checklist Tasks</h1>
          </div>
          
          <div class="content">
            <p>Hello ${safeUserName},</p>
            
            <div class="status-notice">
              📝 <strong>Status Update:</strong> These are your checklist pending tasks that require your attention.
            </div>
            
            <div class="summary">
              <p><strong>Checklist Summary:</strong> You have ${tasks.length} pending checklist task${tasks.length !== 1 ? 's' : ''}${safeOverdueCount > 0 ? `, with ${safeOverdueCount} overdue` : ''}${safeCriticalCount > 0 ? `, including ${safeCriticalCount} critical task${safeCriticalCount > 1 ? 's' : ''}` : ''}.</p>
            </div>
            
            <p>Please review and complete the following checklist tasks:</p>
            
            <table>
              <tr>
                <th>Task ID</th>
                <th>Task Description</th>
                <th>Type & Frequency</th>
                <th>Planned Date</th>
                <th>Priority</th>
              </tr>
    `;

    // Add row for each task
    const today = new Date();

    tasks.forEach((task, index) => {
      try {
        // Safely access task properties
        const taskId = task.taskId || 'N/A';
        const taskDescription = task.taskDescription || 'No description';
        const formattedPlannedDate = task.formattedPlannedDate || 'No planned date';
        const taskType = task.taskType || 'Normal';
        const frequency = task.frequency || 'One-time';

        // Calculate days difference for urgency
        let rowClass = '';
        let urgencyText = '';
        let priorityText = 'Normal';

        const isCritical = taskType.toLowerCase().includes('critical');
        let isOverdue = false;

        if (task.plannedDate) {
          const plannedDate = task.plannedDate instanceof Date ? task.plannedDate : new Date(task.plannedDate);

          if (!isNaN(plannedDate.getTime())) {
            const diffTime = plannedDate.getTime() - today.getTime();
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

            if (diffDays < 0) {
              isOverdue = true;
              urgencyText = ' <strong>(OVERDUE)</strong>';
              priorityText = 'Overdue';
            }
          }
        }

        // Set row class and priority based on conditions
        if (isOverdue && isCritical) {
          rowClass = 'overdue-critical';
          priorityText = 'URGENT - Overdue Critical';
        } else if (isOverdue) {
          rowClass = 'overdue';
          priorityText = 'Overdue';
        } else if (isCritical) {
          rowClass = 'critical';
          priorityText = 'Critical';
        }

        // Create task type badge
        const taskTypeBadge = isCritical ?
          `<span class="task-type-badge critical-badge">${taskType}</span>` :
          `<span class="task-type-badge normal-badge">${taskType}</span>`;

        const frequencyBadge = `<span class="frequency-badge">${frequency}</span>`;

        body += `
                <tr class="${rowClass}">
                  <td>${taskId}</td>
                  <td>${taskDescription}</td>
                  <td>
                    ${taskTypeBadge}<br>
                    ${frequencyBadge}
                  </td>
                  <td>${formattedPlannedDate}${urgencyText}</td>
                  <td><strong>${priorityText}</strong></td>
                </tr>
        `;

      } catch (taskError) {
        Logger.log(`Error processing checklist task ${index} for ${email}: ${taskError.toString()}`);
      }
    });

    body += `
            </table>
            
            <div class="status-notice">
              ✅ <strong>Action Required:</strong> Please complete these checklist tasks according to their priority and planned dates.
            </div>
            
            <p><strong>Priority Guidelines:</strong></p>
            <ul>
              <li><strong>URGENT - Overdue Critical:</strong> Immediate attention required</li>
              <li><strong>Overdue:</strong> Should be completed as soon as possible</li>
              <li><strong>Critical:</strong> High priority, complete by planned date</li>
              <li><strong>Normal:</strong> Complete according to schedule</li>
            </ul>
            
            <p><strong>Task Type Legend:</strong></p>
            <ul>
              <li><span class="task-type-badge critical-badge">Critical</span> - High priority tasks requiring immediate attention</li>
              <li><span class="task-type-badge normal-badge">Normal</span> - Regular checklist tasks</li>
            </ul>
            
            <div style="text-align: center;">
              <a href="https://script.google.com/a/codearchitects.in/macros/s/AKfycbzawv8hZ0lQRGOhemghTb4sjFEZdjdoE-C2iG32tHdG/dev?page=Index" class="button">Open Checklist Dashboard</a>
            </div>
            
            <p>Please ensure all checklist tasks are completed on time to maintain workflow efficiency.</p>
            <p>If you have any questions about any of these tasks, please contact your supervisor.</p>
            <p>Thank you for your attention to these checklist items.</p>
          </div>
          
          <div class="footer">
            <p>This is an automated checklist reminder sent on ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy')}.</p>
            <p>© ${new Date().getFullYear()} Your Company Name</p>
          </div>
        </div>
      </body>
    </html>
    `;

    // Send the email
    GmailApp.sendEmail(email, subject, "", {
      htmlBody: body
    });

    Logger.log('Checklist reminder email sent to: ' + email);

  } catch (error) {
    Logger.log(`Failed to send checklist email to ${email}: ${error.toString()}`);
    Logger.log('Stack trace: ' + error.stack);
    throw error; // Re-throw to be caught by calling function
  }
}



function getDoerListData() {
  try {
    // Get user session data using browser-based cache
    const userSession = getUserSession();

    // If no user session, return empty array
    if (!userSession) {
      return [];
    }

    // Get the spreadsheet and data
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Doer List');

    if (!sheet) {
      console.error('Doer List sheet not found');
      return [];
    }

    const range = sheet.getDataRange();
    const values = range.getValues();

    // Skip header row
    const dataRows = values.slice(1);

    // Filter out empty rows
    return dataRows
      .filter(row => {
        // Skip empty rows
        return row[0] !== '';
      })
      .map(row => {
        // Return the row as-is since there are no dates to format
        return [...row];
      });
  } catch (error) {
    console.error('Error in getDoerListData:', error);
    return [];
  }
}


function addNewDoer(formData) {
  try {
    const userSession = getUserSession();
    if (!userSession) {
      return { success: false, error: 'No valid session' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Doer List');

    if (!sheet) {
      return { success: false, error: 'Doer List sheet not found' };
    }

    // Add new row
    sheet.appendRow([
      formData.email,
      formData.name,
      formData.role,
      formData.mobile
    ]);

    return { success: true };
  } catch (error) {
    console.error('Error in addNewDoer:', error);
    return { success: false, error: error.toString() };
  }
}

function updateDoer(rowIndex, formData) {
  try {
    const userSession = getUserSession();
    if (!userSession) {
      return { success: false, error: 'No valid session' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Doer List');

    if (!sheet) {
      return { success: false, error: 'Doer List sheet not found' };
    }

    // Update the row
    sheet.getRange(rowIndex, 1, 1, 4).setValues([[
      formData.email,
      formData.name,
      formData.role,
      formData.mobile
    ]]);

    return { success: true };
  } catch (error) {
    console.error('Error in updateDoer:', error);
    return { success: false, error: error.toString() };
  }
}

function deleteDoer(rowIndex) {
  try {
    const userSession = getUserSession();
    if (!userSession) {
      return { success: false, error: 'No valid session' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Doer List');

    if (!sheet) {
      return { success: false, error: 'Doer List sheet not found' };
    }

    // Delete the row
    sheet.deleteRow(rowIndex);

    return { success: true };
  } catch (error) {
    console.error('Error in deleteDoer:', error);
    return { success: false, error: error.toString() };
  }
}





function getDoerChart() {
  try {
    // Get user session data using browser-based cache
    const userSession = getUserSession();

    // If no user session, return empty array
    if (!userSession) {
      return [];
    }

    // Get the spreadsheet and data
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Chart');

    if (!sheet) {
      console.error('Doer List sheet not found');
      return [];
    }

    const range = sheet.getDataRange();
    const values = range.getValues();

    // Skip header row
    const dataRows = values.slice(1);

    // Filter out empty rows
    return dataRows
      .filter(row => {
        // Skip empty rows
        return row[0] !== '';
      })
      .map(row => {
        // Return the row as-is since there are no dates to format
        return [...row];
      });
  } catch (error) {
    console.error('Error in getDoerListData:', error);
    return [];
  }
}




// Main spreadsheet for storing MOM data
const SPREADSHEET_ID = '11zYFlya8SMnMkCn9KfQjiA74gfOXmURoNxVlF0464qA'; // <-- IMPORTANT: The ID of the sheet where MOMs are saved

// NEW: Spreadsheet for the list of projects
const PROJECTS_SPREADSHEET_ID = '11zYFlya8SMnMkCn9KfQjiA74gfOXmURoNxVlF0464qA'; // <-- IMPORTANT: The ID of the sheet containing your 'Projects' tab

const MOM_SHEET_NAME = 'MOMs';
const PROJECT_SHEET_NAME = 'Projects';



function getProjects() {
  try {
    const spreadsheet = SpreadsheetApp.openById(PROJECTS_SPREADSHEET_ID);
    const projectSheet = spreadsheet.getSheetByName(PROJECT_SHEET_NAME);
    if (!projectSheet) {
      console.warn(`Sheet '${PROJECT_SHEET_NAME}' not found in the projects spreadsheet.`);
      return ['Error: Projects Sheet Not Found'];
    }
    const lastRow = projectSheet.getLastRow();
    if (lastRow < 1) return [];
    const projects = projectSheet.getRange(1, 1, lastRow, 1).getValues();
    const projectNames = projects.flat().filter(String);
    console.log('Fetched project names:', projectNames);  // Added console log to show the project names
    return projectNames;
  } catch (e) {
    console.error('Error in getProjects: ' + e.toString());
    return ['Error: Could not fetch projects'];
  }
}

// Add this function to fetch RA team members
function getRATeamMembers() {
  try {
    const spreadsheet = SpreadsheetApp.openById(PROJECTS_SPREADSHEET_ID);
    const projectSheet = spreadsheet.getSheetByName(PROJECT_SHEET_NAME);
    if (!projectSheet) {
      console.warn(`Sheet '${PROJECT_SHEET_NAME}' not found in the projects spreadsheet.`);
      return ['Error: RA Team Sheet Not Found'];
    }
    const lastRow = projectSheet.getLastRow();
    if (lastRow < 1) return [];
    // Get column B (2nd column) for RA team members
    const raTeamMembers = projectSheet.getRange(1, 2, lastRow, 1).getValues();
    const memberNames = raTeamMembers.flat().filter(String);
    console.log('Fetched RA team member names:', memberNames);
    return memberNames;
  } catch (e) {
    console.error('Error in getRATeamMembers: ' + e.toString());
    return ['Error: Could not fetch RA team members'];
  }
}


/**
 * Finds a row number based on a unique MOM ID in the main data sheet.
 * @param {Sheet} sheet The sheet to search in.
 * @param {string} momId The MOM ID to find.
 * @returns {number} The row number (1-indexed), or -1 if not found.
 */
function findRowByMomId(sheet, momId) {
  if (!momId) return -1;
  const momIdColumn = sheet.getRange("A:A").getValues();
  for (let i = 0; i < momIdColumn.length; i++) {
    if (momIdColumn[i][0] == momId) {
      return i + 1; // Return 1-indexed row number
    }
  }
  return -1;
}

/**
 * Saves a new or updates an existing MOM record in the main data sheet.
 * @param {object} data The MOM data from the form.
 * @param {string|null} momId The ID of the MOM to update, or null to create a new one.
 * @returns {string} A success message.
 */
function saveOrUpdateMOM(data, momId) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID); // Uses the main SPREADSHEET_ID
    let sheet = spreadsheet.getSheetByName(MOM_SHEET_NAME);
    const headers = [
      "MOM ID", "Timestamp", "Project", "Date", "Time", "Location",
      "RA Team Attendees", "Client Team Attendees", "Vendor Team Attendees", "Other Attendees",
      "Minutes"
    ];

    if (!sheet) {
      sheet = spreadsheet.insertSheet(MOM_SHEET_NAME);
      sheet.appendRow(headers);
    } else if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
    }

    const rowData = [
      momId || `MOM-${new Date().getTime()}`,
      new Date(), data.project, new Date(data.date).toLocaleDateString("en-US"), data.time,
      data.location, data.attendeesRA, data.attendeesClient, data.attendeesVendor,
      data.attendeesOther, data.minutes,
    ];

    if (momId) {
      const rowToUpdate = findRowByMomId(sheet, momId);
      if (rowToUpdate !== -1) {
        sheet.getRange(rowToUpdate, 1, 1, rowData.length).setValues([rowData]);
        return "MOM updated successfully!";
      } else { throw new Error("MOM ID not found for update."); }
    } else {
      sheet.appendRow(rowData);
      return "MOM created successfully!";
    }
  } catch (e) {
    console.error('Error in saveOrUpdateMOM: ' + e.toString());
    throw new Error('Could not save data.');
  }
}

/**
 * Retrieves all MOM records from the main data sheet.
 * @returns {Array<Object>} An array of MOM records.
 */
function getMOMs() {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID); // Uses the main SPREADSHEET_ID
    const sheet = spreadsheet.getSheetByName(MOM_SHEET_NAME);

    if (!sheet) {
      console.log('Sheet not found, returning empty array');
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      console.log('No data rows found, returning empty array');
      return [];
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values.shift();

    const headerMap = {};
    headers.forEach((header, index) => {
      headerMap[header.trim()] = index;
    });

    const result = values.map(row => {
      let formattedDate = '';
      const dateValue = row[headerMap["Date"]];
      if (dateValue) {
        try {
          if (dateValue instanceof Date) {
            formattedDate = dateValue.toLocaleDateString("en-US");
          } else {
            const parsedDate = new Date(dateValue);
            if (!isNaN(parsedDate.getTime())) {
              formattedDate = parsedDate.toLocaleDateString("en-US");
            } else {
              formattedDate = String(dateValue);
            }
          }
        } catch (e) {
          formattedDate = String(dateValue) || '';
        }
      }

      return {
        momId: row[headerMap["MOM ID"]] ? String(row[headerMap["MOM ID"]]) : '',
        project: row[headerMap["Project"]] ? String(row[headerMap["Project"]]) : '',
        date: formattedDate,
        time: row[headerMap["Time"]] ? String(row[headerMap["Time"]]) : '',
        location: row[headerMap["Location"]] ? String(row[headerMap["Location"]]) : '',
        attendeesRA: row[headerMap["RA Team Attendees"]] ? String(row[headerMap["RA Team Attendees"]]) : '',
        attendeesClient: row[headerMap["Client Team Attendees"]] ? String(row[headerMap["Client Team Attendees"]]) : '',
        attendeesVendor: row[headerMap["Vendor Team Attendees"]] ? String(row[headerMap["Vendor Team Attendees"]]) : '',
        attendeesOther: row[headerMap["Other Attendees"]] ? String(row[headerMap["Other Attendees"]]) : '',
        minutes: row[headerMap["Minutes"]] ? String(row[headerMap["Minutes"]]) : '[]',
      };
    }).reverse();

    console.log('getMOMs returning:', result.length, 'records');
    return result;

  } catch (e) {
    console.error('Error in getMOMs: ' + e.toString());
    return [];
  }
}

function generatePdf(momId) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName(MOM_SHEET_NAME);

    if (!sheet) {
      throw new Error('MOM sheet not found');
    }

    const rowToFind = findRowByMomId(sheet, momId);
    if (rowToFind === -1) {
      throw new Error('MOM record not found');
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];
    const momData = values[rowToFind - 1];

    const headerMap = {};
    headers.forEach((header, index) => {
      headerMap[header.trim()] = index;
    });

    const mom = {
      momId: momData[headerMap['MOM ID']] || '',
      timestamp: momData[headerMap['Timestamp']] || '',
      project: momData[headerMap['Project']] || '',
      date: momData[headerMap['Date']] || '',
      time: momData[headerMap['Time']] || '',
      location: momData[headerMap['Location']] || '',
      attendeesRA: momData[headerMap['RA Team Attendees']] || '',
      attendeesClient: momData[headerMap['Client Team Attendees']] || '',
      attendeesVendor: momData[headerMap['Vendor Team Attendees']] || '',
      attendeesOther: momData[headerMap['Other Attendees']] || '',
      minutes: momData[headerMap['Minutes']] || '[]'
    };

    const tempDoc = DocumentApp.create(`MOM_${momId}_${new Date().getTime()}`);
    const body = tempDoc.getBody();

    body.clear();

    addContentToDoc(body, mom);

    tempDoc.saveAndClose();

    const docFile = DriveApp.getFileById(tempDoc.getId());
    const pdfBlob = docFile.getAs('application/pdf');

    const pdfFile = DriveApp.createFile(pdfBlob);
    pdfFile.setName(`MOM_${momId}.pdf`);

    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    DriveApp.getFileById(tempDoc.getId()).setTrashed(true);

    return pdfFile.getUrl();

  } catch (e) {
    console.error('Error generating PDF: ' + e.toString());
    throw new Error('Failed to generate PDF: ' + e.message);
  }
}

/**
 * Adds formatted content to the Google Doc with compact spacing.
 */
function addContentToDoc(body, mom) {
  // Reduced margins for more content
  body.setMarginTop(20);
  body.setMarginBottom(20);
  body.setMarginLeft(30);
  body.setMarginRight(30);

  createHeader(body, mom);

  // Minimal spacing between sections
  body.appendParagraph('').setSpacingAfter(2);

  createMembersTable(body, mom);

  body.appendParagraph('').setSpacingAfter(2);

  createMinutesTable(body, mom);

  body.appendParagraph('').setSpacingAfter(2);

  const footer = body.appendParagraph('**END**');
  footer.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  footer.setAttributes({
    [DocumentApp.Attribute.FONT_SIZE]: 8,
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial'
  });
  footer.setSpacingBefore(2);
  footer.setSpacingAfter(2);
}

/**
 * Creates compact header with logo image.
 */
function createHeader(body, mom) {
  const title = body.insertParagraph(0, 'Minutes of Meeting');
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  title.setAttributes({
    [DocumentApp.Attribute.FONT_SIZE]: 14,
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.BACKGROUND_COLOR]: '#2E5B8A',
    [DocumentApp.Attribute.FOREGROUND_COLOR]: '#FFFFFF',
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial'
  });
  title.setSpacingBefore(2);
  title.setSpacingAfter(2);

  const projectTitle = body.appendParagraph(mom.project || 'PROJECT NAME');
  projectTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  projectTitle.setAttributes({
    [DocumentApp.Attribute.FONT_SIZE]: 12,
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial'
  });
  projectTitle.setSpacingBefore(2);
  projectTitle.setSpacingAfter(2);

  const headerTable = body.appendTable();
  const infoRow = headerTable.appendTableRow();

  // Logo cell with actual image - Remove background color to show logo fully
  const logoCell = infoRow.appendTableCell();
  logoCell.setAttributes({
    [DocumentApp.Attribute.PADDING_TOP]: 8,
    [DocumentApp.Attribute.PADDING_BOTTOM]: 8,
    [DocumentApp.Attribute.PADDING_LEFT]: 8,
    [DocumentApp.Attribute.PADDING_RIGHT]: 8,
    [DocumentApp.Attribute.VERTICAL_ALIGNMENT]: DocumentApp.VerticalAlignment.CENTER
  });

  try {
    // Fetch and insert the logo image from Google Drive
    const logoFileId = '1EaaG_B9QeVnRqhiJX4ovh3UgRb-B52qt';
    const logoFile = DriveApp.getFileById(logoFileId);
    const blob = logoFile.getBlob();

    // Clear the cell first
    logoCell.clear();

    // Create a paragraph and append the image to it
    const logoPara = logoCell.appendParagraph('');
    logoPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    const image = logoPara.appendInlineImage(blob);
    image.setWidth(90);
    image.setHeight(90);

    Logger.log('Logo image inserted successfully');
  } catch (e) {
    Logger.log('Error fetching logo image: ' + e.toString());
    // Fallback to text if image fetch fails
    logoCell.clear();
    const logoPara = logoCell.appendParagraph('RAW ASSOCIATES\nLOGO');
    logoPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    logoPara.setAttributes({
      [DocumentApp.Attribute.FONT_SIZE]: 10,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.FONT_FAMILY]: 'Arial'
    });
  }

  const detailsCell = infoRow.appendTableCell();
  detailsCell.setPaddingTop(2);
  detailsCell.setPaddingBottom(2);

  const detailsTable = detailsCell.appendTable();

  const locationRow = detailsTable.appendTableRow();
  locationRow.appendTableCell('Meeting Location').setAttributes({
    [DocumentApp.Attribute.BACKGROUND_COLOR]: '#E8E8E8',
    [DocumentApp.Attribute.PADDING_TOP]: 2,
    [DocumentApp.Attribute.PADDING_BOTTOM]: 2,
    [DocumentApp.Attribute.PADDING_LEFT]: 6
  }).getChild(0).asParagraph().setAttributes({
    [DocumentApp.Attribute.FONT_SIZE]: 9,
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial'
  });
  locationRow.appendTableCell(mom.location || '').setAttributes({
    [DocumentApp.Attribute.PADDING_TOP]: 2,
    [DocumentApp.Attribute.PADDING_BOTTOM]: 2,
    [DocumentApp.Attribute.PADDING_LEFT]: 6
  }).getChild(0).asParagraph().setAttributes({
    [DocumentApp.Attribute.FONT_SIZE]: 9,
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial'
  });

  const dateRow = detailsTable.appendTableRow();
  dateRow.appendTableCell('Meeting Date').setAttributes({
    [DocumentApp.Attribute.BACKGROUND_COLOR]: '#E8E8E8',
    [DocumentApp.Attribute.PADDING_TOP]: 2,
    [DocumentApp.Attribute.PADDING_BOTTOM]: 2,
    [DocumentApp.Attribute.PADDING_LEFT]: 6
  }).getChild(0).asParagraph().setAttributes({
    [DocumentApp.Attribute.FONT_SIZE]: 9,
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial'
  });
  dateRow.appendTableCell(`${formatDate(mom.date)} ${mom.time || ''}`).setAttributes({
    [DocumentApp.Attribute.PADDING_TOP]: 2,
    [DocumentApp.Attribute.PADDING_BOTTOM]: 2,
    [DocumentApp.Attribute.PADDING_LEFT]: 6
  }).getChild(0).asParagraph().setAttributes({
    [DocumentApp.Attribute.FONT_SIZE]: 9,
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial'
  });

  headerTable.setColumnWidth(0, 100);
  headerTable.setColumnWidth(1, 400);
}

/**
 * Creates compact members table.
 */
function createMembersTable(body, mom) {
  const membersTitle = body.appendParagraph('Members Present List');
  membersTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  membersTitle.setAttributes({
    [DocumentApp.Attribute.FONT_SIZE]: 11,
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial'
  });
  membersTitle.setSpacingBefore(2);
  membersTitle.setSpacingAfter(2);

  const membersTable = body.appendTable();

  const headerRow = membersTable.appendTableRow();
  const headers = ['S.No.', 'RA Team', 'Client Team', 'Vendor Team', 'Other'];
  headers.forEach((header) => {
    const cell = headerRow.appendTableCell(header);
    cell.setAttributes({
      [DocumentApp.Attribute.BACKGROUND_COLOR]: '#D4D4D4',
      [DocumentApp.Attribute.PADDING_TOP]: 2,
      [DocumentApp.Attribute.PADDING_BOTTOM]: 2,
      [DocumentApp.Attribute.PADDING_LEFT]: 4,
      [DocumentApp.Attribute.PADDING_RIGHT]: 4
    });
    cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    cell.getChild(0).asParagraph().setAttributes({
      [DocumentApp.Attribute.FONT_SIZE]: 8,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.FONT_FAMILY]: 'Arial'
    });
  });

  const raAttendees = mom.attendeesRA ? mom.attendeesRA.split('\n').filter(Boolean) : [];
  const clientAttendees = mom.attendeesClient ? mom.attendeesClient.split('\n').filter(Boolean) : [];
  const vendorAttendees = mom.attendeesVendor ? mom.attendeesVendor.split('\n').filter(Boolean) : [];
  const otherAttendees = mom.attendeesOther ? mom.attendeesOther.split('\n').filter(Boolean) : [];

  const maxRows = Math.max(raAttendees.length, clientAttendees.length, vendorAttendees.length, otherAttendees.length, 1);

  for (let i = 0; i < maxRows; i++) {
    const dataRow = membersTable.appendTableRow();
    const cells = [
      (i + 1).toString(),
      raAttendees[i] || '',
      clientAttendees[i] || '',
      vendorAttendees[i] || '',
      otherAttendees[i] || ''
    ];
    cells.forEach((cellContent, index) => {
      const cell = dataRow.appendTableCell(cellContent);
      cell.setAttributes({
        [DocumentApp.Attribute.PADDING_TOP]: 2,
        [DocumentApp.Attribute.PADDING_BOTTOM]: 2,
        [DocumentApp.Attribute.PADDING_LEFT]: 4,
        [DocumentApp.Attribute.PADDING_RIGHT]: 4
      });
      cell.getChild(0).asParagraph().setAttributes({
        [DocumentApp.Attribute.FONT_SIZE]: 8,
        [DocumentApp.Attribute.FONT_FAMILY]: 'Arial'
      });
      if (index === 0) {
        cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      }
    });
  }

  membersTable.setColumnWidth(0, 35);
  membersTable.setColumnWidth(1, 115);
  membersTable.setColumnWidth(2, 115);
  membersTable.setColumnWidth(3, 115);
  membersTable.setColumnWidth(4, 115);
}

/**
 * Creates compact minutes table.
 */
function createMinutesTable(body, mom) {
  const minutesTitle = body.appendParagraph('Meeting Minutes');
  minutesTitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  minutesTitle.setAttributes({
    [DocumentApp.Attribute.FONT_SIZE]: 11,
    [DocumentApp.Attribute.BOLD]: true,
    [DocumentApp.Attribute.FONT_FAMILY]: 'Arial'
  });
  minutesTitle.setSpacingBefore(2);
  minutesTitle.setSpacingAfter(2);

  const minutesTable = body.appendTable();

  const headerRow = minutesTable.appendTableRow();
  const headers = ['S.No.', 'Minutes Discussed', 'Action By', 'Plan Start', 'Plan End', 'Actual End', 'Remarks'];
  headers.forEach((header) => {
    const cell = headerRow.appendTableCell(header);
    cell.setAttributes({
      [DocumentApp.Attribute.BACKGROUND_COLOR]: '#7CB3E8',
      [DocumentApp.Attribute.PADDING_TOP]: 2,
      [DocumentApp.Attribute.PADDING_BOTTOM]: 2,
      [DocumentApp.Attribute.PADDING_LEFT]: 3,
      [DocumentApp.Attribute.PADDING_RIGHT]: 3
    });
    cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    cell.getChild(0).asParagraph().setAttributes({
      [DocumentApp.Attribute.FONT_SIZE]: 7,
      [DocumentApp.Attribute.BOLD]: true,
      [DocumentApp.Attribute.FONT_FAMILY]: 'Arial'
    });
  });

  let minutes = [];
  try {
    minutes = JSON.parse(mom.minutes || '[]');
  } catch (e) {
    console.error('Error parsing minutes:', e);
  }

  // Only add rows for actual data, no blank rows
  const totalRows = Math.max(minutes.length, 1);

  for (let i = 0; i < totalRows; i++) {
    const minute = minutes[i] || {};
    const dataRow = minutesTable.appendTableRow();
    const cells = [
      (i + 1).toString(),
      minute.discussed || '',
      minute.actionBy || '',
      formatDateShort(minute.plannedStart),
      formatDateShort(minute.plannedCompletion),
      formatDateShort(minute.actualCompletion),
      minute.remarks || ''
    ];
    cells.forEach((cellContent, index) => {
      const cell = dataRow.appendTableCell(cellContent);
      cell.setAttributes({
        [DocumentApp.Attribute.PADDING_TOP]: 2,
        [DocumentApp.Attribute.PADDING_BOTTOM]: 2,
        [DocumentApp.Attribute.PADDING_LEFT]: 3,
        [DocumentApp.Attribute.PADDING_RIGHT]: 3
      });
      cell.getChild(0).asParagraph().setAttributes({
        [DocumentApp.Attribute.FONT_SIZE]: 7,
        [DocumentApp.Attribute.FONT_FAMILY]: 'Arial'
      });
      if (index === 0) {
        cell.getChild(0).asParagraph().setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      }
    });
  }

  // Optimized column widths for compact layout
  minutesTable.setColumnWidth(0, 25);   // S.No.
  minutesTable.setColumnWidth(1, 140);  // Minutes Discussed
  minutesTable.setColumnWidth(2, 60);   // Action By
  minutesTable.setColumnWidth(3, 55);   // Plan Start
  minutesTable.setColumnWidth(4, 55);   // Plan End
  minutesTable.setColumnWidth(5, 55);   // Actual End
  minutesTable.setColumnWidth(6, 105);  // Remarks
}

/**
 * Formats date for compact display (MM/DD/YY format).
 */
function formatDateShort(date) {
  if (!date) return '';

  try {
    let parsedDate;
    if (date instanceof Date) {
      parsedDate = date;
    } else {
      parsedDate = new Date(date);
    }

    if (!isNaN(parsedDate.getTime())) {
      const month = String(parsedDate.getMonth() + 1).padStart(2, '0');
      const day = String(parsedDate.getDate()).padStart(2, '0');
      const year = String(parsedDate.getFullYear()).slice(-2);
      return `${month}/${day}/${year}`;
    }
    return String(date);
  } catch (e) {
    return String(date);
  }
}

/**
 * Formats date for display.
 */
function formatDate(date) {
  if (!date) return '';

  try {
    if (date instanceof Date) {
      return date.toLocaleDateString('en-US');
    } else {
      const parsedDate = new Date(date);
      if (!isNaN(parsedDate.getTime())) {
        return parsedDate.toLocaleDateString('en-US');
      } else {
        return String(date);
      }
    }
  } catch (e) {
    return String(date);
  }
}



function testDocumentCreation() {
  try {
    const tempDoc = DocumentApp.create('Test_Document_' + new Date().getTime());
    tempDoc.getBody().appendParagraph('This is a test document.');
    tempDoc.saveAndClose();

    // Optionally, trash the test document after creation to clean up
    DriveApp.getFileById(tempDoc.getId()).setTrashed(true);

    Logger.log('Test document created successfully. Permissions granted.');
    return 'Permissions granted successfully!';
  } catch (e) {
    Logger.log('Error in testDocumentCreation: ' + e.toString());
    throw new Error('Failed to create test document: ' + e.message);
  }
}




function getChecklistDData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Checklist');
  const range = sheet.getRange('B2:G');
  const values = range.getValues();

  // Select all columns from B to G: B(0), C(1), D(2), E(3), F(4), G(5)
  // In new array: 0=B, 1=C, 2=D, 3=E, 4=F, 5=G
  const formattedValues = values.map(row => {
    const selected = [
      row[0], // B - Name
      row[1], // C - Task ID
      row[2], // D - Freq
      row[3], // E - Task
      row[4], // F - Planned
      row[5]  // G - Actual
    ];

    // Format date columns in the new array: index 4 (F) and 5 (G)
    const formattedSelected = selected.map((cell, index) => {
      if ((index === 4 || index === 5) && cell instanceof Date && !isNaN(cell)) {
        const day = cell.getDate().toString().padStart(2, '0');
        const month = (cell.getMonth() + 1).toString().padStart(2, '0');
        const year = cell.getFullYear();
        const hours = cell.getHours().toString().padStart(2, '0');
        const minutes = cell.getMinutes().toString().padStart(2, '0');
        const seconds = cell.getSeconds().toString().padStart(2, '0');
        return `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
      }
      return cell;
    });

    // Calculate status
    const plannedStr = formattedSelected[4];
    const actualStr = formattedSelected[5];

    let status = '';

    // Parse planned timestamp
    let plannedDate = null;
    if (plannedStr && typeof plannedStr === 'string') {
      try {
        const [datePart, timePart] = plannedStr.split(' ');
        const [day, month, year] = datePart.split('/');
        const [hours, minutes, seconds] = timePart.split(':');
        plannedDate = new Date(
          parseInt(year),
          parseInt(month) - 1,
          parseInt(day),
          parseInt(hours),
          parseInt(minutes),
          parseInt(seconds)
        );
      } catch (e) {
        plannedDate = null;
      }
    }

    // Parse actual timestamp
    let actualDate = null;
    if (actualStr && typeof actualStr === 'string') {
      try {
        const [datePart, timePart] = actualStr.split(' ');
        const [day, month, year] = datePart.split('/');
        const [hours, minutes, seconds] = timePart.split(':');
        actualDate = new Date(
          parseInt(year),
          parseInt(month) - 1,
          parseInt(day),
          parseInt(hours),
          parseInt(minutes),
          parseInt(seconds)
        );
      } catch (e) {
        actualDate = null;
      }
    }

    if (actualDate === null) {
      // Not completed - check against current time
      const now = new Date();
      if (plannedDate && plannedDate >= now) {
        status = 'Pending Within Limit';
      } else {
        status = 'Pending Overdue';
      }
    } else {
      // Completed
      if (plannedDate && actualDate <= plannedDate) {
        status = 'Completed On Time';
      } else {
        status = 'Completed With Delay';
      }
    }

    // Add status to the array
    formattedSelected.push(status);

    return formattedSelected;
  });

  // Get today's date at midnight for comparison (for filtering)
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Filter out empty rows (based on column B, which is now index 0 in the new array)
  // and only include rows where planned date (index 4) is today or earlier
  return formattedValues.filter(row => {
    if (row[0] === '') return false;

    // Parse planned date (index 4) for filtering
    const plannedDateStr = row[4];
    if (!plannedDateStr || typeof plannedDateStr !== 'string') return false;

    try {
      const [datePart] = plannedDateStr.split(' ');
      const [day, month, year] = datePart.split('/');
      const plannedDate = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
      plannedDate.setHours(0, 0, 0, 0);

      return plannedDate <= today;
    } catch (error) {
      return false;
    }
  });
}






function getDelegationData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Master');
  const range = sheet.getRange('A2:U');
  const values = range.getValues();

  // Helper function to format dates consistently
  function formatDate(cell) {
    if (cell instanceof Date && !isNaN(cell)) {
      const day = cell.getDate().toString().padStart(2, '0');
      const month = (cell.getMonth() + 1).toString().padStart(2, '0');
      const year = cell.getFullYear();
      const hours = cell.getHours().toString().padStart(2, '0');
      const minutes = cell.getMinutes().toString().padStart(2, '0');
      const seconds = cell.getSeconds().toString().padStart(2, '0');
      return `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
    }
    return cell;
  }

  // Helper function to parse formatted date string
  function parseFormattedDate(dateStr) {
    if (!dateStr || typeof dateStr !== 'string') return null;
    try {
      const [datePart, timePart] = dateStr.trim().split(' ');
      const [day, month, year] = datePart.split('/');
      const [hours = '0', minutes = '0', seconds = '0'] = (timePart || '').split(':');
      return new Date(
        parseInt(year),
        parseInt(month) - 1,
        parseInt(day),
        parseInt(hours),
        parseInt(minutes),
        parseInt(seconds)
      );
    } catch (e) {
      return null;
    }
  }

  const formattedValues = values.map(row => {
    // Filter out rows where Created At (index 0) is empty
    if (!row[0] || row[0] === '') return null;

    const selected = [
      row[0], // Created AT
      row[1], // task id
      row[2], // Assigned by
      String(row[3] || '').trim(), // Assigned to (trimmed)
      row[5], // Task
      row[20] || row[7] || row[6], // Planned Time → fallback Revision Date → First Date
      row[8], // Revision
      row[9], // Status
      row[18] // Actual Time
    ];

    // Format date columns: Created At (0), Planned Time (5), Actual Time (8)
    const formattedSelected = selected.map((cell, index) => {
      if ((index === 0 || index === 5 || index === 8) && cell instanceof Date && !isNaN(cell)) {
        return formatDate(cell);
      }
      return cell;
    });

    // Update status based on planned time and current status
    const status = formattedSelected[7];
    if (status === 'Pending' || status === 'Work Not Done' || status === 'Need Help') {
      const plannedStr = formattedSelected[5];
      const now = new Date();
      const plannedDate = parseFormattedDate(plannedStr);

      if (plannedDate && plannedDate >= now) {
        formattedSelected[7] = 'Pending Within Limit';
      } else if (plannedDate && plannedDate < now) {
        formattedSelected[7] = 'Pending Overdue';
      }
    } else if (status === 'Completed') {
      const plannedStr = formattedSelected[5];
      const actualStr = formattedSelected[8];
      const plannedDate = parseFormattedDate(plannedStr);
      const actualDate = parseFormattedDate(actualStr);

      if (plannedDate && actualDate) {
        if (actualDate <= plannedDate) {
          formattedSelected[7] = 'Completed On Time';
        } else {
          formattedSelected[7] = 'Completed With Delay';
        }
      }
    }

    return formattedSelected;
  }).filter(row => row !== null); // Remove null rows

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  return formattedValues.filter(row => {
    if (row[7] === "Week Shifted" || row[7] === "Without Plan") return false;

    const plannedDate = parseFormattedDate(String(row[5] || ''));
    if (!plannedDate) return false;

    // Pending / open tasks: show ALL (incl. future & newly assigned)
    const isCompleted = /^completed/i.test(String(row[7] || ''));
    if (!isCompleted) return true;

    // Completed tasks: only up to today
    plannedDate.setHours(0, 0, 0, 0);
    return plannedDate <= today;
  });
}

function escalateStatuses() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Master");

  if (!sheet) {
    Logger.log("Sheet 'Master' not found.");
    return;
  }

  // Get the data range (assuming data starts from row 2, adjust if needed)
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  // Current date/time for comparison
  var now = new Date();

  // Loop through each row starting from index 1 (row 2 in sheet)
  for (var i = 1; i < data.length; i++) {
    var valueI = Number(data[i][8]); // Column I (0-indexed: A=0, I=8) -> Revision Count
    var status = data[i][9]; // Column J (0-indexed: J=9) -> Status
    var taskType = data[i][21]; // Column V (0-indexed: V=21) -> Task Type
    var approvedCount = Number(data[i][22]) || 0; // Column W (0-indexed: W=22) -> Approved Count

    // Determine the effective deadline date for comparison
    var firstDate = data[i][6]; // Column G (Initial Planned Date)
    var revisionDate = data[i][7]; // Column H (Follow-up / Shifted Date)
    var uDate = data[i][20]; // Column U (Planned Time - Fallback)
    
    var effectiveDateStr = revisionDate ? revisionDate : firstDate;
    var effectiveDate = null;

    if (effectiveDateStr instanceof Date) {
      effectiveDate = effectiveDateStr;
    } else if (typeof effectiveDateStr === 'string' && effectiveDateStr.trim() !== '') {
      // Parse "dd/MM/yyyy HH:mm:ss" safely into a readable date object
      var parts = effectiveDateStr.split(' ');
      var dParts = parts[0].split('/');
      if (dParts.length === 3) {
        var tParts = parts[1] ? parts[1].split(':') : [0,0,0];
        effectiveDate = new Date(dParts[2], dParts[1] - 1, dParts[0], tParts[0]||0, tParts[1]||0, tParts[2]||0);
      }
    }

    // Fallback to Column U if parsing failed
    if (!effectiveDate || isNaN(effectiveDate.getTime())) {
      if (uDate instanceof Date) {
        effectiveDate = uDate;
      } else {
        continue; // Cannot determine a valid date, skip this row
      }
    }

    var isFollowUp = taskType && taskType.toString().trim() === 'Follow-up Task';
    
    // For Follow-up Tasks: maximum revisions is dynamically scaled as 3 + (approvedCount * 3)
    // For Regular Tasks: maximum revisions before escalation defaults to 2
    var limit = isFollowUp ? (3 + (approvedCount * 3)) : 2;

    // Check if revision count >= limit, status is "Pending" or "Need Help", and the newly parsed deadline has passed
    if (valueI >= limit && (status === "Pending" || status === "Need Help") && effectiveDate < now) {
      // Update status to "Work Not Done"
      sheet.getRange(i + 1, 10).setValue("Work Not Done"); // Row i+1, Column 10 (J)
    }
  }

  Logger.log("Escalation check completed.");
}



function saveAdminAction(taskId, remark) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Master");
  const data = sheet.getDataRange().getValues();
  
  // Task ID column B (index 1) mein search karein
  for (let i = 1; i < data.length; i++) {
    if (data[i][1].toString() === taskId.toString()) {
      const rowNum = i + 1;
      const timestamp = Utilities.formatDate(new Date(), "GMT+5:30", "yyyy-MM-dd HH:mm:ss");
      const finalRemark = remark + " (Time: " + timestamp + ")";
      
      // Column J (Status) - Index 10 (Google Sheets 1-based index)
      sheet.getRange(rowNum, 10).setValue("Need Help(Closed)");
      
      // Column R (Remark) - Index 18
      sheet.getRange(rowNum, 18).setValue(finalRemark);
      
      return "Success";
    }
  }
  throw new Error("Task ID nahi mila!");
}

function getNeedHelpCount(userName, userRole) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Master');
    if (!sheet) return 0;

    const data = sheet.getDataRange().getValues();

    const statusCol = 9;  // Column J
    const nameCol = 3;    // Column D

    let count = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[statusCol];
      const name = row[nameCol];

      if (status === "Need Help") {
        if (userRole && userRole.toLowerCase() === "admin") {
          count++;
        } else if (
          userRole &&
          userRole.toLowerCase() === "doer" &&
          name &&
          name.toString().trim().toLowerCase() === userName.toString().trim().toLowerCase()
        ) {
          count++;
        }
      }
    }

    return count;
  } catch (error) {
    Logger.log('Error in getNeedHelpCount: ' + error.toString());
    return 0;
  }
}

// Fetch MIS Scoring data for the logged-in user from the Combine sheet
function getMyMISData(userName, userRole) {
  try {
    Logger.log("=== START getMyMISData ===");
    Logger.log("1. Original userName passed from front-end: '" + userName + "' | Role: '" + userRole + "'");

    // The exact Spreadsheet ID from your "Combine MIS Report"
    var ssId = '19KV5ZQBt6sAzDjMqngVbIwqtKu6qLZbfwsrvC-tYvYk';
    var ss = SpreadsheetApp.openById(ssId);
    
    // Specifically target the sheet named "Combine"
    var sheet = ss.getSheetByName('Combine'); 
    
    // Fallback just in case the name changes
    if (!sheet) {
      Logger.log("2. WARNING: Sheet named 'Combine' NOT found. Falling back to the first tab.");
      sheet = ss.getSheets()[0]; 
    } else {
      Logger.log("2. SUCCESS: Sheet named 'Combine' found and opened.");
    }
    
    var data = sheet.getDataRange().getDisplayValues();
    Logger.log("3. Total rows fetched from sheet (including header): " + data.length);

    if (data.length <= 1) {
      Logger.log("ERROR: Sheet is empty or only has a header row. Returning empty array.");
      return[]; // Return empty if no data
    }
    
    var rows = data.slice(1); // Remove header row

    // --- ADMIN / B.O CHECK ---
    // If the user is Admin or B.O, skip the name filter and return ALL rows
    if (userRole && (userRole.toLowerCase() === 'admin' || userRole.toLowerCase() === 'b.o')) {
      Logger.log("4. Access Level: ADMIN/B.O. Bypassing name filter. Returning all records.");
      
      // Filter out completely empty rows, but keep everyone's data
      var allValidRows = rows.filter(function(row) { 
        return row[0] && row[0].toString().trim() !== ''; 
      });
      
      Logger.log("5. Total matched rows successfully found for Admin: " + allValidRows.length);
      Logger.log("=== END getMyMISData ===");
      return allValidRows;
    }
    
    // --- DOER CHECK (SMART NAME MATCHER) ---
    // 1. Take the user's login name (e.g. "Prakash Roy - 8017342858")
    // 2. Split by "-" to remove the number, and make it lowercase
    // Result: "prakash roy"
    var cleanUserName = userName.split('-')[0].toLowerCase().trim();
    Logger.log("4. Access Level: DOER. Cleaned Logged-in UserName for matching: '" + cleanUserName + "'");
    
    // Filter the rows for this specific Doer
    var matchCount = 0;
    var matchedRows = rows.filter(function(row, index) {
      if (!row[0]) return false; // Skip empty rows
      
      // 1. Get the name from Column A (e.g. "Mr. Prakash Roy")
      var rawColName = row[0].toString();
      
      // 2. Make it lowercase
      var colName = rawColName.toLowerCase().trim();
      
      // 3. Remove "mr. ", "ms. ", or "mrs. " from the beginning
      colName = colName.replace(/^mr\.\s*/, '').replace(/^ms\.\s*/, '').replace(/^mrs\.\s*/, '').trim();
      
      // Compare the names
      var isMatch = cleanUserName.includes(colName) || colName.includes(cleanUserName);
      
      // LOGGER: Show exactly what is happening for the first 3 rows so we can debug the logic
      if (index < 3) {
        Logger.log("   -> Row " + (index+1) + " Check: Raw Sheet Name: '" + rawColName + "' | Cleaned Sheet Name: '" + colName + "' | Is Match? " + isMatch);
      }

      if (isMatch) {
        matchCount++;
      }
      
      return isMatch;
    });
    
    Logger.log("5. Total matched rows successfully found for this user: " + matchCount);
    Logger.log("=== END getMyMISData ===");
    
    return matchedRows;
    
  } catch (error) {
    Logger.log("EXCEPTION / ERROR in getMyMISData: " + error.toString());
    return[];
  }}
// --- NEW FUNCTION: submitFollowUpProof ---
function submitFollowUpProof(taskId, imageUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Master');
  const data = sheet.getDataRange().getValues();
  
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][1] === taskId) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) throw new Error("Task not found");
  
  sheet.getRange(rowIndex, 10).setValue('Pending Follow-up Approval');
  
  if (imageUrl) {
    const currentDocs = sheet.getRange(rowIndex, 14).getValue();
    const newDocs = currentDocs ? currentDocs + ", " + imageUrl : imageUrl;
    sheet.getRange(rowIndex, 14).setValue(newDocs);
  }
  return true;
}

// --- NEW FUNCTION: processFollowUpAdminBackend ---
function processFollowUpAdminBackend(taskId, action, finalStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Master');
  const data = sheet.getDataRange().getValues();
  
  let rowIndex = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][1] === taskId) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) throw new Error("Task not found");
  
  if (action === 'approve') {
    const currentCount = parseInt(sheet.getRange(rowIndex, 23).getValue()) || 0;
    sheet.getRange(rowIndex, 23).setValue(currentCount + 1);
    sheet.getRange(rowIndex, 10).setValue('Pending'); // Unfreeze Task
  } else if (action === 'disapprove') {
    sheet.getRange(rowIndex, 10).setValue(finalStatus);
  }
  
  return true;
}
function testGetData() {
  const result = getData();
  Logger.log('Result type: ' + typeof result);
  Logger.log('Is array: ' + Array.isArray(result));
  Logger.log('Length: ' + (result ? result.length : 'null'));
  Logger.log('First row: ' + (result && result[0] ? JSON.stringify(result[0]) : 'none'));
}
// Run this ONCE to fix all existing rows in the sheet
function fixExistingNumberColumns() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Master');
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) return 'No data rows found';
    
    // Fix Column I (index 9, col 9) - Revision Count
    const revisionRange = sheet.getRange(2, 9, lastRow - 1, 1);
    const revisionValues = revisionRange.getValues();
    const fixedRevisions = revisionValues.map(row => {
      const val = row[0];
      if (val instanceof Date) return [0];
      const num = parseInt(val);
      return [isNaN(num) ? 0 : num];
    });
    revisionRange.setValues(fixedRevisions);
    revisionRange.setNumberFormat('0');
    
    // Fix Column W (index 23, col 23) - Approved Count
    const approvedRange = sheet.getRange(2, 23, lastRow - 1, 1);
    const approvedValues = approvedRange.getValues();
    const fixedApproved = approvedValues.map(row => {
      const val = row[0];
      if (val === '' || val === null || val === undefined) return [0];
      if (val instanceof Date) return [0];
      const num = parseInt(val);
      return [isNaN(num) ? 0 : num];
    });
    approvedRange.setValues(fixedApproved);
    approvedRange.setNumberFormat('0');
    
    SpreadsheetApp.flush();
    Logger.log('Fixed ' + (lastRow - 1) + ' rows successfully');
    return 'Fixed ' + (lastRow - 1) + ' rows successfully';
    
  } catch (error) {
    Logger.log('Error in fixExistingNumberColumns: ' + error.toString());
    return 'Error: ' + error.toString();
  }
}
function testDoerDataDirect() {
  // Temporarily force a session to test
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DoerDashboard');
  
  Logger.log('Sheet found: ' + (sheet ? 'YES' : 'NO'));
  
  if (sheet) {
    const values = sheet.getDataRange().getValues();
    Logger.log('Total rows: ' + values.length);
    Logger.log('First data row: ' + JSON.stringify(values[1]));
    
    // Check column A values
    const nonEmpty = values.slice(1).filter(r => r[0] !== '');
    Logger.log('Non-empty rows: ' + nonEmpty.length);
  }
  
  // Check session
  const session = getUserSession();
  Logger.log('Session: ' + JSON.stringify(session));
}