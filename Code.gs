// ─────────────────────────────────────────────────────────────────────────────
// Google Apps Script — Employee Onboarding Backend
// Deploy as: Extensions → Apps Script → Deploy → New deployment
//            Type: Web app | Execute as: Me | Who has access: Anyone
// After deploying, copy the Web App URL into the HTML file (SCRIPT_URL constant)
// ─────────────────────────────────────────────────────────────────────────────

const SECRET     = "Bnpower@2026";   // ← same key in the HTML file
const SHEET_NAME = "Employees";
const FOLDER_NAME = "Employee Photos";

// ── Entry point for POST requests ────────────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.secret !== SECRET) return respond({ ok: false, error: "Unauthorized" });
    const action = payload.action || "save";
    if (action === "save")   return respond(saveEmployee(payload.data));
    if (action === "update") return respond(updateEmployee(payload.data));
    if (action === "delete") return respond(deleteEmployee(payload.id));
    if (action === "list")   return respond(listEmployees());
    return respond({ ok: false, error: "Unknown action" });
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

function doGet(e) {
  if (e.parameter.secret !== SECRET) return respond({ ok: false, error: "Unauthorized" });
  return respond(listEmployees());
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Write header row
    sheet.appendRow([
      "ID","SL","Joining Date","Site","Designation","Name","DOB","Age",
      "Gender","Marital Status","Blood Group","Relation Name","Mobile",
      "Emergency Contact","Aadhar","UAN","ESIC","Wages","Temp Address",
      "Perm Address","Pincode","Bank Name","Account No","IFSC","Branch",
      "Nominee","Nominee Relation","Nominee Contact","Education",
      "Technical Qual","Training","Experience","Place","Photo (Drive Link)","Added On"
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 35).setFontWeight("bold").setBackground("#1f6f54").setFontColor("#ffffff");
  }
  return sheet;
}

function getPhotoFolder() {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
}

function savePhoto(base64Data, employeeName, employeeId) {
  if (!base64Data) return "";
  try {
    // Strip data URL prefix  e.g. "data:image/jpeg;base64,..."
    const match = base64Data.match(/^data:([^;]+);base64,(.+)$/);
    if (!match) return "";
    const mimeType = match[1];
    const blob = Utilities.newBlob(Utilities.base64Decode(match[2]), mimeType, employeeName + "_" + employeeId + ".jpg");
    const file = getPhotoFolder().createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (err) {
    return "";
  }
}

function rowToObj(headers, row) {
  const obj = {};
  headers.forEach((h, i) => obj[h] = row[i]);
  return obj;
}

// ── CRUD operations ───────────────────────────────────────────────────────────
function saveEmployee(d) {
  const sheet = getSheet();
  const photoUrl = savePhoto(d.photo, d.name, d.id);
  const exps = JSON.stringify(d.exps || []);

  sheet.appendRow([
    d.id, d.sl, d.joining, d.site, d.desig, d.name, d.dob, d.age,
    d.gender, d.marital, d.blood, d.rel, d.mobile, d.emerg,
    d.aadhar, d.uan, d.esic, d.wage, d.taddr, d.paddr, d.pin,
    d.bank, d.acc, d.ifsc, d.branch, d.nom, d.nrel, d.ncon,
    d.edu, d.tech, d.train, exps, d.place, photoUrl, d.added
  ]);

  return { ok: true, photoUrl };
}

function updateEmployee(d) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const idCol = 0; // column A = ID

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(d.id)) {
      // Only update photo if a new one was provided
      let photoUrl = data[i][33]; // existing Drive link
      if (d.photo && d.photo.startsWith("data:")) {
        photoUrl = savePhoto(d.photo, d.name, d.id);
      }
      const exps = JSON.stringify(d.exps || []);
      const row = i + 1; // sheet rows are 1-indexed
      sheet.getRange(row, 1, 1, 35).setValues([[
        d.id, d.sl, d.joining, d.site, d.desig, d.name, d.dob, d.age,
        d.gender, d.marital, d.blood, d.rel, d.mobile, d.emerg,
        d.aadhar, d.uan, d.esic, d.wage, d.taddr, d.paddr, d.pin,
        d.bank, d.acc, d.ifsc, d.branch, d.nom, d.nrel, d.ncon,
        d.edu, d.tech, d.train, exps, d.place, photoUrl, d.added
      ]]);
      return { ok: true, photoUrl };
    }
  }
  return { ok: false, error: "Record not found" };
}

function deleteEmployee(id) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: "Record not found" };
}

function listEmployees() {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { ok: true, data: [] };
  const headers = rows[0];
  const employees = rows.slice(1).map(r => rowToObj(headers, r));
  return { ok: true, data: employees };
}
