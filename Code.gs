const SECRET      = "Bnpower@2026";
const SHEET_NAME  = "Employees";
const FOLDER_NAME = "Employee Photos";

function doGet(e) {
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  try {
    if (e.parameter.secret !== SECRET) { out.setContent(JSON.stringify({ok:false,error:"Unauthorized"})); return out; }
    const action = e.parameter.action || "list";
    if (action === "list")   { out.setContent(JSON.stringify(listEmployees())); return out; }
    const payload = JSON.parse(e.parameter.data || "{}");
    if (action === "save")   { out.setContent(JSON.stringify(saveEmployee(payload.data))); return out; }
    if (action === "update") { out.setContent(JSON.stringify(updateEmployee(payload.data))); return out; }
    if (action === "delete") { out.setContent(JSON.stringify(deleteEmployee(payload.id))); return out; }
    out.setContent(JSON.stringify({ok:false,error:"Unknown action"})); return out;
  } catch(err) {
    out.setContent(JSON.stringify({ok:false,error:err.message})); return out;
  }
}

function doPost(e) {
  const out = ContentService.createTextOutput();
  out.setMimeType(ContentService.MimeType.JSON);
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.secret !== SECRET) { out.setContent(JSON.stringify({ok:false,error:"Unauthorized"})); return out; }
    const action = payload.action || "save";
    if (action === "save")   { out.setContent(JSON.stringify(saveEmployee(payload.data))); return out; }
    if (action === "update") { out.setContent(JSON.stringify(updateEmployee(payload.data))); return out; }
    if (action === "delete") { out.setContent(JSON.stringify(deleteEmployee(payload.id))); return out; }
    if (action === "list")   { out.setContent(JSON.stringify(listEmployees())); return out; }
    out.setContent(JSON.stringify({ok:false,error:"Unknown action"})); return out;
  } catch(err) {
    out.setContent(JSON.stringify({ok:false,error:err.message})); return out;
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["ID","SL","Joining Date","Site","Designation","Name","DOB","Age","Gender","Marital Status","Blood Group","Relation Name","Mobile","Emergency Contact","Aadhar","UAN","ESIC","Wages","Temp Address","Perm Address","Pincode","Bank Name","Account No","IFSC","Branch","Nominee","Nominee Relation","Nominee Contact","Education","Technical Qual","Training","Experience","Place","Photo (Drive Link)","Added On"]);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,35).setFontWeight("bold").setBackground("#1f6f54").setFontColor("#ffffff");
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
    const match = base64Data.match(/^data:([^;]+);base64,(.+)$/);
    if (!match) return "";
    const blob = Utilities.newBlob(Utilities.base64Decode(match[2]), match[1], employeeName+"_"+employeeId+".jpg");
    const file = getPhotoFolder().createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch(err) { return ""; }
}

function rowToObj(headers, row) {
  const obj = {};
  headers.forEach((h,i) => obj[h] = row[i]);
  return obj;
}

function saveEmployee(d) {
  const sheet = getSheet();
  const photoUrl = savePhoto(d.photo, d.name, d.id);
  sheet.appendRow([d.id,d.sl,d.joining,d.site,d.desig,d.name,d.dob,d.age,d.gender,d.marital,d.blood,d.rel,d.mobile,d.emerg,d.aadhar,d.uan,d.esic,d.wage,d.taddr,d.paddr,d.pin,d.bank,d.acc,d.ifsc,d.branch,d.nom,d.nrel,d.ncon,d.edu,d.tech,d.train,JSON.stringify(d.exps||[]),d.place,photoUrl,d.added]);
  return {ok:true, photoUrl};
}

function updateEmployee(d) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(d.id)) {
      let photoUrl = data[i][33];
      if (d.photo && d.photo.startsWith("data:")) photoUrl = savePhoto(d.photo, d.name, d.id);
      sheet.getRange(i+1,1,1,35).setValues([[d.id,d.sl,d.joining,d.site,d.desig,d.name,d.dob,d.age,d.gender,d.marital,d.blood,d.rel,d.mobile,d.emerg,d.aadhar,d.uan,d.esic,d.wage,d.taddr,d.paddr,d.pin,d.bank,d.acc,d.ifsc,d.branch,d.nom,d.nrel,d.ncon,d.edu,d.tech,d.train,JSON.stringify(d.exps||[]),d.place,photoUrl,d.added]]);
      return {ok:true, photoUrl};
    }
  }
  return {ok:false, error:"Record not found"};
}

function deleteEmployee(id) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return {ok:true}; }
  }
  return {ok:false, error:"Record not found"};
}

function listEmployees() {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return {ok:true, data:[]};
  const headers = rows[0];
  return {ok:true, data: rows.slice(1).map(r => rowToObj(headers, r))};
}
