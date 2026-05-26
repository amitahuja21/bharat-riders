// ============================================================
//  ROAD WARRIOR EV CHALLENGE — Google Apps Script Backend
// ============================================================

const CONFIG = {
  SHEET_ID:       '1F-3Rl7bk0mkAgaqutqPtP-wdKRE8uPidfIrqqb4HE7I',
  MAIN_WEBSITE:   'https://amitahuja21.github.io/bharat-riders/',
  ADMIN_PASSWORD: 'AsdfQwer!234',
  SCRIPT_URL:     'https://script.google.com/macros/s/1YNG2H6ShH-2ZaMBuDZ3C4Olr7dqAHSiKVl1zwRk3xPCOd6kU3E1SitjR/exec'
};

const SHEETS = {
  REGISTRATIONS: 'Registrations',
  CLICKS:        'Clicks'
};

function doGet(e) {
  const p = e.parameter || {};

  // ── 0. FORM SUBMISSION → Save to Responses sheet ──────────
  if (p.action === 'submit') {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = getOrCreateSheet(ss, 'Responses', [
      'Timestamp','Ref Code','Referred By','Language',
      'Name','Phone','State','City','Platforms',
      'Vehicle Type','Current Vehicle','Fuel Type',
      'EV Speed','Charging Method','Charge Time',
      'Ownership','EMI/Rent','Monthly Expenses',
      'Has Insurance','Insurance Type','Challenges',
      'RTO Training','Customer Training','App Challenges'
    ]);

    // ── CHECK DUPLICATE PHONE ──────────────────────────────
    const phone = (p.phone || '').replace(/\D/g, '');
    if (phone) {
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][5]).replace(/\D/g, '') === phone) {
          return jsonResponse({ success: false, duplicate: true, message: 'You have already submitted a response.' });
        }
      }
    }

    sheet.appendRow([
      p.timestamp || new Date().toISOString(),
      p.refCode || '', p.referredBy || '', p.language || 'en',
      p.name || '', p.phone || '', p.state || '', p.city || '',
      p.platforms || '', p.vehicleType || '', p.currentVehicle || '',
      p.fuelType || '', p.evSpeed || '', p.chargingMethod || '',
      p.chargeTime || '', p.ownership || '', p.emiRent || '',
      p.monthlyExpenses || '', p.hasInsurance || '', p.insuranceType || '',
      p.challenges || '', p.rtoTraining || '', p.customerTraining || '',
      p.appChallenges || ''
    ]);
    return jsonResponse({ success: true });
  } catch(err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

  // ── 1. AFFILIATE LINK CLICK → Log + Redirect ──────────────
  if (p.ref) {
    try { logClick(p.ref); } catch(err) {}
    const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta http-equiv="refresh" content="0;url=${CONFIG.MAIN_WEBSITE}">
  <title>Road Warrior EV Challenge</title>
</head>
<body>
  <script>window.location.replace("${CONFIG.MAIN_WEBSITE}");<\/script>
  <p>Redirecting...</p>
</body>
</html>`;
    return HtmlService.createHtmlOutput(html)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // ── 2. ADMIN — Get All Data ────────────────────────────────
  if (p.action === 'getData' && p.pwd === CONFIG.ADMIN_PASSWORD) {
    try {
      return jsonResponse({ success: true, data: getAllRegistrations() });
    } catch(err) {
      return jsonResponse({ success: false, message: err.toString() });
    }
  }

  // ── 3. LOOKUP — Get My Referral Link by Phone ─────────────
  if (p.action === 'getMyLink' && p.phone) {
    try {
      const user = findByPhone(p.phone);
      if (user) {
        return jsonResponse({
          success: true,
          name: user.name,
          code: user.code,
          referralLink: `${CONFIG.SCRIPT_URL}?ref=${user.code}`,
          clicks: user.clicks,
          referrals: countReferrals(user.code)
        });
      } else {
        return jsonResponse({ success: false, message: 'Phone number not found.' });
      }
    } catch(err) {
      return jsonResponse({ success: false, message: err.toString() });
    }
  }

  return jsonResponse({ success: false, message: 'Invalid request' });
}

// ──────────────────────────────────────────────────────────────
//  POST HANDLER
// ──────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'register') {
      return jsonResponse(registerUser(data));
    }
    return jsonResponse({ success: false, message: 'Unknown action' });
  } catch(err) {
    return jsonResponse({ success: false, message: 'Server error: ' + err.toString() });
  }
}

// ──────────────────────────────────────────────────────────────
//  REGISTRATION LOGIC
// ──────────────────────────────────────────────────────────────
function registerUser(data) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEETS.REGISTRATIONS, [
    'ID', 'Name', 'Phone', 'City', 'Bike Model', 'Email',
    'Referral Code', 'Referred By Code', 'Registered At', 'Clicks'
  ]);
  if (!data.name || !data.phone) {
    return { success: false, message: 'Name and phone are required.' };
  }
  const phone = String(data.phone).replace(/\D/g, '');
  if (phone.length < 10) {
    return { success: false, message: 'Please enter a valid 10-digit phone number.' };
  }
  const existing = findByPhone(phone);
  if (existing) {
    return {
      success: false, alreadyRegistered: true,
      message: 'You are already registered!',
      name: existing.name, code: existing.code,
      referralLink: `${CONFIG.SCRIPT_URL}?ref=${existing.code}`,
      clicks: existing.clicks
    };
  }
  const code = generateUniqueCode(sheet, data.name);
  const id = 'IR' + Date.now().toString().slice(-7);
  sheet.appendRow([
    id, data.name.trim(), phone,
    (data.city || '').trim(), (data.bike || '').trim(), (data.email || '').trim(),
    code, (data.referredBy || '').trim().toUpperCase(),
    new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }), 0
  ]);
  return {
    success: true, message: 'Welcome! 🎉',
    id, name: data.name.trim(), code,
    referralLink: `${CONFIG.SCRIPT_URL}?ref=${code}`
  };
}

// ──────────────────────────────────────────────────────────────
//  CLICK TRACKING
// ──────────────────────────────────────────────────────────────
function logClick(refCode) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const clickSheet = getOrCreateSheet(ss, SHEETS.CLICKS, ['Timestamp', 'Ref Code', 'IST Time']);
  clickSheet.appendRow([
    new Date().toISOString(), refCode.toUpperCase(),
    new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' })
  ]);
  const regSheet = ss.getSheetByName(SHEETS.REGISTRATIONS);
  if (!regSheet) return;
  const rows = regSheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][6]).toUpperCase() === refCode.toUpperCase()) {
      regSheet.getRange(i + 1, 10).setValue((Number(rows[i][9]) || 0) + 1);
      break;
    }
  }
}

// ──────────────────────────────────────────────────────────────
//  DATA RETRIEVAL
// ──────────────────────────────────────────────────────────────
function getAllRegistrations() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.REGISTRATIONS);
  if (!sheet || sheet.getLastRow() <= 1) return [];
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    obj['Referrals'] = countReferrals(String(row[6]));
    return obj;
  }).reverse();
}

function findByPhone(phone) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.REGISTRATIONS);
  if (!sheet || sheet.getLastRow() <= 1) return null;
  const rows = sheet.getDataRange().getValues();
  const cleanPhone = String(phone).replace(/\D/g, '');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][2]).replace(/\D/g, '') === cleanPhone) {
      return { id: rows[i][0], name: rows[i][1], phone: rows[i][2], city: rows[i][3],
               bike: rows[i][4], code: rows[i][6], referredBy: rows[i][7],
               registeredAt: rows[i][8], clicks: Number(rows[i][9]) || 0 };
    }
  }
  return null;
}

function countReferrals(code) {
  if (!code) return 0;
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.REGISTRATIONS);
  if (!sheet || sheet.getLastRow() <= 1) return 0;
  const rows = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][7]).toUpperCase() === code.toUpperCase()) count++;
  }
  return count;
}

// ──────────────────────────────────────────────────────────────
//  HELPERS
// ──────────────────────────────────────────────────────────────
function generateUniqueCode(sheet, name) {
  const prefix = name.replace(/[^a-zA-Z]/g, '').toUpperCase().slice(0, 4).padEnd(4, 'X');
  let code, isUnique = false, attempts = 0;
  while (!isUnique && attempts < 20) {
    const suffix = Math.random().toString(36).slice(2, 6).toUpperCase();
    code = `${prefix}-${suffix}`;
    isUnique = !codeExists(sheet, code);
    attempts++;
  }
  return code;
}

function codeExists(sheet, code) {
  if (sheet.getLastRow() <= 1) return false;
  const data = sheet.getRange(2, 7, sheet.getLastRow() - 1, 1).getValues();
  return data.some(row => row[0] === code);
}

function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doOptions(e) {
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}
