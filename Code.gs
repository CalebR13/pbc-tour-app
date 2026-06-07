// ============================================================
// PBC TOUR APP — Google Apps Script
// Includes Google Drive file upload for receipts & interest forms
// ============================================================

var FOLDER_RECEIPTS  = "1tCBpsUqS4FUyo00X8Nus6gg3HoAcjBvq"; // Tour Receipts
var FOLDER_MAVERICK  = "1Ip2NYKB9NBlE1rxPcHenk6ZCGGEA59t6"; // Tour Maverick Club

var SHEET_NAMES = {
  evaluations: "Church Evaluations",
  expenses:    "Expenses",
  income:      "Income",
  infocards:   "Interest Cards",
  socials:     "Social Photos",
  schedule:    "Tour Schedule",
  lodging:     "Tour Lodging",
  log:         "Submission Log"
};

var HEADERS = {
  "Church Evaluations": [
    "Timestamp", "Tour Group", "Visit Date", "Church Name / City / State", "Pastor Name",
    "Teenagers in Attendance (est.)", "PBC Connections", "PBC Connection Details",
    "Other Bible College Connections", "Provided by Church",
    "Soul-Winning Rating (1-5)", "Differences from Standards",
    "Notes for Future Visits", "Recommend Returning (0-5)", "Submitted By"
  ],
  "Expenses": [
    "Timestamp", "Tour Group", "Expense Date", "Amount ($)", "Category", "Category Details", "Receipt Link", "Submitted By", "Paid with Cash"
  ],
  "Income": [
    "Timestamp", "Tour Group", "Church Name / City / State", "Income Date",
    "Amount ($)", "Income Category", "Income Type", "Submitted By"
  ],
  "Interest Cards": [
    "Timestamp", "Tour Group", "First Name", "Last Name", "Gender", "Date of Birth",
    "HS Grad Year", "Email", "Phone", "City", "State", "Zip",
    "Church Name / City / State", "Pastor Name",
    "Interests", "Majors of Interest", "Interest Form Photo Link", "Submitted By"
  ],
  "Submission Log": [
    "Timestamp", "Type", "Sub-Type", "Summary"
  ],
  "Social Photos": [
    "Timestamp", "Tour Group", "Location", "Caption", "Submitted By"
  ],
  "Tour Schedule": [
    "Tour Group", "Date", "Day of Week", "Event", "Destination", "Destination City, State", "Pastor", "Event Time", "Leader Notes"
  ],
  "Tour Lodging": [
    "Tour Group", "Date", "Lodging Location", "Address", "Contact", "Notes"
  ]
};

// ============================================================
// SAFE SETUP — only adds MISSING tabs, never deletes existing data
// ============================================================
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var created = [];
  for (var sheetName in HEADERS) {
    var existing = ss.getSheetByName(sheetName);
    if (!existing) {
      var sheet = ss.insertSheet(sheetName);
      var hdrs = HEADERS[sheetName];
      sheet.appendRow(hdrs);
      sheet.getRange(1, 1, 1, hdrs.length)
        .setFontWeight("bold")
        .setBackground("#1C1A17")
        .setFontColor("#F5F2EC");
      sheet.setFrozenRows(1);
      sheet.autoResizeColumns(1, hdrs.length);
      created.push(sheetName);
    }
  }
  if (created.length) {
    Logger.log("Created new sheets: " + created.join(", "));
  } else {
    Logger.log("All sheets already exist — nothing changed.");
  }
}

// ============================================================
// DANGER — only run this if you want to wipe ALL data and start fresh
// Renamed to make it harder to run accidentally
// ============================================================
function DANGER_deleteAndRebuildAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert("WARNING", "This will DELETE all data in all tabs. Are you absolutely sure?", ui.ButtonSet.YES_NO);
  if (result !== ui.Button.YES) { Logger.log("Cancelled."); return; }
  for (var sheetName in HEADERS) {
    var existing = ss.getSheetByName(sheetName);
    if (existing) ss.deleteSheet(existing);
    var sheet = ss.insertSheet(sheetName);
    var hdrs = HEADERS[sheetName];
    sheet.appendRow(hdrs);
    sheet.getRange(1, 1, 1, hdrs.length)
      .setFontWeight("bold")
      .setBackground("#1C1A17")
      .setFontColor("#F5F2EC");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, hdrs.length);
  }
  Logger.log("All sheets rebuilt.");
}

// ============================================================
// POST — receives form submissions including file uploads
// ============================================================
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // File upload only — attach to last row of relevant sheet
    if (data.action === "uploadFile") {
      handleFileUpload(ss, data);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === "updateSchedule") {
      var updated = updateScheduleRow(ss, data);
      return ContentService
        .createTextOutput(JSON.stringify({ success: updated, error: updated ? "" : "Schedule row not found" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === "addScheduleEvent") {
      addScheduleEvent(ss, data);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === "addLodging") {
      addLodging(ss, data);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === "updateLodging") {
      var updatedLodging = updateLodging(ss, data);
      return ContentService
        .createTextOutput(JSON.stringify({ success: updatedLodging, error: updatedLodging ? "" : "No existing lodging row found for this date/group" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === "deleteLodging") {
      var deletedLodging = deleteLodging(ss, data);
      return ContentService
        .createTextOutput(JSON.stringify({ success: deletedLodging, error: deletedLodging ? "" : "Lodging row not found" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.action === "deleteScheduleEvent") {
      var deleted = deleteScheduleEvent(ss, data);
      return ContentService
        .createTextOutput(JSON.stringify({ success: deleted, error: deleted ? "" : "Schedule row not found" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (data.role === "evaluation") {
      writeEvaluation(ss, data);
      if (data.incomeAddon && data.incomeAddon.entries && data.incomeAddon.entries.length > 0) {
        writeIncome(ss, { tourGroup: data.tourGroup, church: data.incomeAddon.church, date: data.incomeAddon.date, entries: data.incomeAddon.entries });
      }
    } else if (data.role === "expenses" && data.type === "Expense") {
      writeExpense(ss, data);
    } else if (data.role === "expenses" && data.type === "Income") {
      writeIncome(ss, data);
    } else if (data.role === "infocard") {
      writeInfoCard(ss, data);
    } else if (data.role === "socials") {
      writeSocials(ss, data);
    }

    writeLog(ss, data);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// GET — returns all data for the dashboard
// ============================================================
function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var data = JSON.stringify({
      evaluations: sheetToObjects(ss.getSheetByName(SHEET_NAMES.evaluations)),
      expenses:    sheetToObjects(ss.getSheetByName(SHEET_NAMES.expenses)),
      income:      sheetToObjects(ss.getSheetByName(SHEET_NAMES.income)),
      infocards:   sheetToObjects(ss.getSheetByName(SHEET_NAMES.infocards)),
      schedule:    sheetToObjects(ss.getSheetByName(SHEET_NAMES.schedule)),
      lodging:     sheetToObjects(ss.getSheetByName(SHEET_NAMES.lodging)),
      log:         sheetToObjects(ss.getSheetByName(SHEET_NAMES.log))
    });

    var callback = e.parameter.callback;
    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + data + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return ContentService
      .createTextOutput(data)
      .setMimeType(ContentService.MimeType.TEXT);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function doOptions(e) {
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ============================================================
// FILE UPLOAD HELPER
// Saves base64 file to a Google Drive folder, returns its URL
// ============================================================
function saveFileToDrive(folderId, fileName, mimeType, base64Data) {
  try {
    var folder = DriveApp.getFolderById(folderId);
    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, mimeType, fileName);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/file/d/' + file.getId() + '/view?usp=sharing';
  } catch (err) {
    Logger.log("File upload error: " + err.toString());
    return "Upload failed: " + err.toString();
  }
}

// ============================================================
// WRITERS
// ============================================================
function writeEvaluation(ss, d) {
  var sheet = ss.getSheetByName(SHEET_NAMES.evaluations);
  if (!sheet) return;
  sheet.appendRow([
    new Date(),
    d.tourGroup    || "Unknown",
    d.date           || "",
    d.church         || "",
    d.pastor         || "",
    d.teenagers      || "",
    d.pbcConnections || "",
    d.pbcDetails     || "",
    d.otherColleges  || "",
    d.provided       || "",
    d.soulWinning    || 0,
    d.differences    || "",
    d.notes          || "",
    d.recommend      || "",
    d.submittedBy    || ""
  ]);
}

function writeExpense(ss, d) {
  var sheet = ss.getSheetByName(SHEET_NAMES.expenses);
  if (!sheet) return;

  var receiptUrl = "";
  if (d.receiptFile && d.receiptFile.data) {
    var fileName = "Receipt_" + (d.date || new Date().toISOString().split("T")[0]) + "_" + d.receiptFile.name;
    receiptUrl = saveFileToDrive(FOLDER_RECEIPTS, fileName, d.receiptFile.mimeType, d.receiptFile.data);
  }

  sheet.appendRow([
    new Date(),
    d.tourGroup    || "Unknown",
    d.date       || "",
    parseFloat(d.amount) || 0,
    d.category   || "",
    d.categoryDetail || "",
    receiptUrl   || "No receipt",
    d.submittedBy    || "",
    d.paidWithCash || ""
  ]);
}

function writeIncome(ss, d) {
  var sheet = ss.getSheetByName(SHEET_NAMES.income);
  if (!sheet) return;
  var entries = d.entries || [];
  entries.forEach(function(entry) {
    sheet.appendRow([
      new Date(),
      d.tourGroup      || "Unknown",
      d.church             || "",
      d.date               || "",
      parseFloat(entry.amount) || 0,
      entry.category       || "",
      entry.type           || "",
      d.submittedBy        || ""
    ]);
  });
}

function writeInfoCard(ss, d) {
  var sheet = ss.getSheetByName(SHEET_NAMES.infocards);
  if (!sheet) return;

  var photoUrl = "";
  if (d.formPhoto && d.formPhoto.data) {
    var fileName = "InterestForm_" + (d.lastName || "") + "_" + (d.firstName || "") + "_" + new Date().toISOString().split("T")[0] + "_" + d.formPhoto.name;
    photoUrl = saveFileToDrive(FOLDER_MAVERICK, fileName, d.formPhoto.mimeType, d.formPhoto.data);
  }

  sheet.appendRow([
    new Date(),
    d.tourGroup    || "Unknown",
    d.firstName    || "",
    d.lastName     || "",
    d.gender       || "",
    d.dob          || "",
    d.gradYear     || "",
    d.email        || "",
    d.phone        || "",
    d.city         || "",
    d.state        || "",
    d.zip          || "",
    d.church       || "",
    d.pastor       || "",
    d.interests    || "",
    d.majors       || "",
    photoUrl       || "No photo",
    d.submittedBy    || ""
  ]);
}

function handleFileUpload(ss, d) {
  if (!d.file || !d.file.data) return;
  var folderId = (d.role === "infocard") ? FOLDER_MAVERICK : FOLDER_RECEIPTS;
  var label = d.label || "file";
  var date = d.date || new Date().toISOString().split("T")[0];
  var fileName = label + "_" + date + "_" + (d.file.name || "upload.jpg");
  var fileUrl = saveFileToDrive(folderId, fileName, d.file.mimeType, d.file.data);

  // Find the last row in the relevant sheet and update the file link column
  var sheetName = d.role === "infocard" ? SHEET_NAMES.infocards : SHEET_NAMES.expenses;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var linkCol = headers.indexOf(d.role === "infocard" ? "Interest Form Photo Link" : "Receipt Link") + 1;
  if (linkCol > 0 && lastRow > 1) {
    sheet.getRange(lastRow, linkCol).setValue(fileUrl);
  }
}

function writeSocials(ss, d) {
  var sheet = ss.getSheetByName(SHEET_NAMES.socials);
  if (!sheet) return;
  sheet.appendRow([
    new Date(),
    d.tourGroup    || "",
    d.location     || "",
    d.caption      || "",
    d.submittedBy  || ""
  ]);
}

function deleteScheduleEvent(ss, d) {
  var sheet = ss.getSheetByName(SHEET_NAMES.schedule);
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idx = scheduleIndexes(headers);
  if (!d.original) throw new Error("Missing original row signature for delete");
  for (var i = data.length - 1; i >= 1; i--) {
    if (scheduleRowMatches(data[i], idx, d.original)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function addScheduleEvent(ss, d) {
  var sheet = ss.getSheetByName(SHEET_NAMES.schedule);
  if (!sheet) return;
  sheet.appendRow([
    d.tourGroup   || "",
    d.date        || "",
    d.dayOfWeek   || dayOfWeekFromDateStr(d.date),
    d.eventType   || "",
    d.churchName  || "",
    d.cityState   || "",
    d.pastor      || "",
    d.serviceTimes|| "",
    d.notes       || ""
  ]);
}

function dayOfWeekFromDateStr(val) {
  if (!val) return "";
  var d = new Date(val);
  if (isNaN(d)) return "";
  var days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  return days[d.getDay()];
}

function scheduleIndexes(headers) {
  return {
    tourGroup: headers.indexOf("Tour Group"),
    date: headers.indexOf("Date"),
    dayOfWeek: headers.indexOf("Day of Week"),
    event: headers.indexOf("Event"),
    destination: headers.indexOf("Destination"),
    cityState: headers.indexOf("Destination City, State"),
    pastor: headers.indexOf("Pastor"),
    eventTime: indexOfAny(headers, ["Event Time", "Event time"]),
    notes: headers.indexOf("Leader Notes")
  };
}

function indexOfAny(headers, names) {
  for (var i = 0; i < names.length; i++) {
    var idx = headers.indexOf(names[i]);
    if (idx > -1) return idx;
  }
  return -1;
}

function scheduleRowMatches(row, idx, original) {
  return String(row[idx.tourGroup] || "") === String(original.tourGroup || "") &&
    normalizeDate(String(row[idx.date] || "")) === normalizeDate(String(original.date || "")) &&
    String(row[idx.event] || "") === String(original.eventType || "") &&
    String(row[idx.destination] || "") === String(original.destination || "") &&
    String(row[idx.cityState] || "") === String(original.cityState || "") &&
    String(row[idx.pastor] || "") === String(original.pastor || "") &&
    String(row[idx.eventTime] || "") === String(original.eventTime || "") &&
    String(row[idx.notes] || "") === String(original.notes || "");
}

function normalizeDate(val) {
  if (!val) return '';
  var s = String(val);
  // If it looks like an ISO timestamp, convert to readable
  if (s.indexOf('T') > -1) {
    var d = new Date(s);
    return (d.getMonth()+1) + '/' + d.getDate() + '/' + d.getFullYear();
  }
  // If it's a Date object from sheets (number), convert it
  if (!isNaN(s) && s.length > 4) {
    var d = new Date(parseFloat(s));
    return (d.getMonth()+1) + '/' + d.getDate() + '/' + d.getFullYear();
  }
  return s;
}

function updateScheduleRow(ss, d) {
  var sheet = ss.getSheetByName(SHEET_NAMES.schedule);
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idx = scheduleIndexes(headers);
  if (!d.original) throw new Error("Missing original row signature for update");
  
  for (var i = 1; i < data.length; i++) {
    if (scheduleRowMatches(data[i], idx, d.original)) {
      sheet.getRange(i + 1, idx.tourGroup + 1).setValue(d.tourGroup || "");
      sheet.getRange(i + 1, idx.date + 1).setValue(d.date || "");
      sheet.getRange(i + 1, idx.dayOfWeek + 1).setValue(d.dayOfWeek || dayOfWeekFromDateStr(d.date));
      sheet.getRange(i + 1, idx.event + 1).setValue(d.eventType || "");
      sheet.getRange(i + 1, idx.destination + 1).setValue(d.churchName || "");
      sheet.getRange(i + 1, idx.cityState + 1).setValue(d.cityState || "");
      sheet.getRange(i + 1, idx.pastor + 1).setValue(d.pastor || "");
      sheet.getRange(i + 1, idx.eventTime + 1).setValue(d.serviceTimes || "");
      sheet.getRange(i + 1, idx.notes + 1).setValue(d.notes || "");
      return true;
    }
  }
  return false;
}

function writeLog(ss, d) {
  var sheet = ss.getSheetByName(SHEET_NAMES.log);
  if (!sheet) return;
  var summary = "";
  if (d.role === "evaluation")
    summary = (d.church || "Unknown church") + " — Pastor " + (d.pastor || "");
  if (d.role === "expenses" && d.type === "Expense")
    summary = "Expense $" + (d.amount || "0") + " — " + (d.category || "");
  if (d.role === "expenses" && d.type === "Income")
    summary = "Income from " + (d.church || "Unknown") + " (" + ((d.entries || []).length) + " entries)";
  if (d.role === "socials")
    summary = "Photos uploaded — " + (d.location || "Unknown location");
  if (d.role === "infocard")
    summary = (d.firstName || "") + " " + (d.lastName || "") + " — " + (d.church || "");
  sheet.appendRow([new Date(), d.role, d.type || "", summary]);
}


// ============================================================
// RUN ONCE — Seeds the Ladies Tour schedule
function seedLadiesSchedule() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.schedule);
  if (!sheet) { Logger.log("Tour Schedule sheet not found. Run setupSheets first."); return; }
  var rows = [
    ["ladies", "May 16", "Saturday", "Travel", "", "Skandia, MI", "", "", "None provided", ""],
    ["ladies", "May 17", "Sunday", "AM & PM Service", "Heritage Baptist Church (AM) / Ontonagon Baptist Church (PM)", "Skandia, MI / Ontonagon, MI", "Pastor Virgil (PM)", "AM + 6:00 PM", "Sunday night provided", ""],
    ["ladies", "May 18", "Monday", "Travel", "", "", "", "", "", ""],
    ["ladies", "May 19", "Tuesday", "Rest Day", "", "", "", "", "", "Make sure girls get meals covered"],
    ["ladies", "May 20", "Wednesday", "Midweek Service", "Northstar Baptist Church", "McMillan, MI", "", "7:00 PM EDT", "2 motel rooms provided Wed night", "Bro Bartels will preach. Group invited to stay for school lunch & softball Thu"],
    ["ladies", "May 21", "Thursday", "Teen Activity", "Northstar Baptist Church", "McMillan, MI", "", "11:00 AM EDT", "", "Done by early afternoon. Then Tahquamenon Falls stop ($11/vehicle), travel home 4PM EDT"],
    ["ladies", "May 22", "Friday", "Special", "NWBBC - HS Graduation", "Elgin, IL", "", "7:00 PM", "PBC", "Girls meals need to be covered"],
    ["ladies", "May 23", "Saturday", "Travel", "", "Mercer, PA", "", "", "Lodging provided Sat night", ""],
    ["ladies", "May 24", "Sunday", "AM & PM Service", "Calvary Fellowship Chapel (AM) / Faith Baptist Church (PM)", "Mercer, PA / Osterburg, PA", "Pastor Henry (AM)", "9:30 AM EDT / 7:00 PM EDT", "2 motel rooms provided", "Bro Bartels will preach PM"],
    ["ladies", "May 25", "Monday", "Rest Day", "", "Osterburg, PA", "", "", "No lodging provided", ""],
    ["ladies", "May 26", "Tuesday", "Travel", "", "Coatesville, PA", "", "", "No lodging provided", ""],
    ["ladies", "May 27", "Wednesday", "Midweek Service", "Valley Baptist Church", "Coatesville, PA", "", "7:00 PM EDT", "Church provides lodging Wed night", ""],
    ["ladies", "May 28", "Thursday", "Rest/Activity", "", "Coatesville, PA", "", "", "No lodging provided", "Shopping + cheesesteak lunch"],
    ["ladies", "May 29", "Friday", "Travel", "", "Newcastle, PA", "", "", "No lodging provided", ""],
    ["ladies", "May 31", "Sunday", "PM Service", "Newcastle, PA", "Newcastle, PA", "", "5:00 PM EDT", "Church provides lodging Sun night", "Bro Bartels will preach"],
    ["ladies", "June 1", "Monday", "Travel + Camp", "PBC", "Elgin, IL", "", "", "PBC", "Volleyball Camp begins"],
    ["ladies", "June 2", "Tuesday", "Camp", "PBC", "Elgin, IL", "", "", "PBC", "Volleyball Camp"],
    ["ladies", "June 3", "Wednesday", "Camp", "PBC", "Elgin, IL", "", "", "PBC", "Volleyball Camp"],
    ["ladies", "June 4", "Thursday", "Camp", "PBC", "Elgin, IL", "", "", "PBC", "Volleyball Camp"],
    ["ladies", "June 5", "Friday", "Rest Day", "PBC", "Elgin, IL", "", "", "PBC", ""],
    ["ladies", "June 6", "Saturday", "Travel", "", "Burlington, IA", "", "", "Heritage Baptist provides lodging Sat night", ""],
    ["ladies", "June 7", "Sunday", "AM & PM Service", "Heritage Baptist Church (AM) / Pleasantville Baptist Church (PM)", "Burlington, IA / Pleasantville, IA", "Pastor Hall (AM)", "10:00 AM / 6:00 PM", "Church provides lodging Sun night", "Pastor Hall will preach AM"],
    ["ladies", "June 8", "Monday", "Rest Day", "", "", "", "", "", ""]
  ];
  rows.forEach(function(row) { sheet.appendRow(row); });
  Logger.log("Seeded " + rows.length + " Ladies schedule rows!");
}


// ============================================================
// RUN ONCE — Seeds the Men's Tour schedule structure
// ============================================================
function seedMensSchedule() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.schedule);
  if (!sheet) { Logger.log("Tour Schedule sheet not found. Run setupSheets first."); return; }
  
  var rows = [
    ["mens", "May 16", "Saturday", "Travel", "", "", "", "", "", "", ""],
    ["mens", "May 17", "Sunday", "AM Service", "", "", "", "", "", "", ""],
    ["mens", "May 17", "Sunday", "PM Service", "", "", "", "", "", "", ""],
    ["mens", "May 18", "Monday", "", "", "", "", "", "", "", ""],
    ["mens", "May 19", "Tuesday", "", "", "", "", "", "", "", ""],
    ["mens", "May 20", "Wednesday", "", "", "", "", "", "", "", ""],
    ["mens", "May 21", "Thursday", "", "", "", "", "", "", "", ""],
    ["mens", "May 22", "Friday", "", "", "", "", "", "", "", ""],
    ["mens", "May 23", "Saturday", "", "", "", "", "", "", "", ""],
    ["mens", "May 24", "Sunday", "AM Service", "", "", "", "", "", "", ""],
    ["mens", "May 24", "Sunday", "PM Service", "", "", "", "", "", "", ""],
    ["mens", "May 25", "Monday", "", "", "", "", "", "", "", ""],
    ["mens", "May 26", "Tuesday", "", "", "", "", "", "", "", ""],
    ["mens", "May 27", "Wednesday", "", "", "", "", "", "", "", ""],
    ["mens", "May 28", "Thursday", "", "", "", "", "", "", "", ""],
    ["mens", "May 29", "Friday", "", "", "", "", "", "", "", ""],
    ["mens", "May 30", "Saturday", "", "", "", "", "", "", "", ""],
    ["mens", "May 31", "Sunday", "AM Service", "", "", "", "", "", "", ""],
    ["mens", "May 31", "Sunday", "PM Service", "", "", "", "", "", "", ""],
    ["mens", "June 1", "Monday", "", "", "", "", "", "", "", ""],
    ["mens", "June 2", "Tuesday", "", "", "", "", "", "", "", ""],
    ["mens", "June 3", "Wednesday", "", "", "", "", "", "", "", ""],
    ["mens", "June 4", "Thursday", "", "", "", "", "", "", "", ""],
    ["mens", "June 5", "Friday", "", "", "", "", "", "", "", ""],
    ["mens", "June 6", "Saturday", "", "", "", "", "", "", "", ""],
    ["mens", "June 7", "Sunday", "AM Service", "", "", "", "", "", "", ""],
    ["mens", "June 7", "Sunday", "PM Service", "", "", "", "", "", "", ""],
    ["mens", "June 8", "Monday", "", "", "", "", "", "", "", ""],
    ["mens", "June 9", "Tuesday", "", "", "", "", "", "", "", ""],
    ["mens", "June 10", "Wednesday", "", "", "", "", "", "", "", ""]
  ];
  
  rows.forEach(function(row) {
    sheet.appendRow(row);
  });
  
  Logger.log("Seeded " + rows.length + " Men's schedule rows — fill in the details in the sheet!");
}


// ============================================================
// RUN ONCE — Seeds the Ladies Tour lodging data
// ============================================================
function seedLadiesLodging() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAMES.lodging);
  if (!sheet) { Logger.log("Tour Lodging sheet not found. Run setupSheets first."); return; }
  var rows = [
    ["ladies", "May 16", "None provided", "", "", ""],
    ["ladies", "May 17", "Sunday night provided", "", "Pastor Virgil", ""],
    ["ladies", "May 18", "", "", "", ""],
    ["ladies", "May 19", "", "", "", "Make sure the girls get their meals covered"],
    ["ladies", "May 20", "2 motel rooms provided", "", "", "Wed night"],
    ["ladies", "May 21", "No lodging", "", "", ""],
    ["ladies", "May 22", "PBC", "PBC", "", "Girls meals need to be covered this day"],
    ["ladies", "May 23", "Lodging provided", "", "", "Saturday night"],
    ["ladies", "May 24", "2 motel rooms", "", "", ""],
    ["ladies", "May 25", "No lodging provided", "", "", ""],
    ["ladies", "May 26", "No lodging provided", "", "", ""],
    ["ladies", "May 27", "Church provides lodging", "", "", "Wednesday night"],
    ["ladies", "May 28", "No lodging provided", "", "", ""],
    ["ladies", "May 29", "No lodging provided", "", "", ""],
    ["ladies", "May 31", "Church provides lodging", "", "", "Sunday night"],
    ["ladies", "June 1", "PBC", "PBC", "", ""],
    ["ladies", "June 2", "PBC", "PBC", "", ""],
    ["ladies", "June 3", "PBC", "PBC", "", ""],
    ["ladies", "June 4", "PBC", "PBC", "", ""],
    ["ladies", "June 5", "PBC", "PBC", "", ""],
    ["ladies", "June 6", "Heritage Baptist", "", "", "Saturday night provided"],
    ["ladies", "June 7", "Church provides lodging", "", "", "Sunday night"],
    ["ladies", "June 8", "", "", "", ""]
  ];
  rows.forEach(function(row) { sheet.appendRow(row); });
  Logger.log("Seeded " + rows.length + " Ladies lodging rows!");
}

// ============================================================
// LODGING FUNCTIONS
// ============================================================
function addLodging(ss, d) {
  var sheet = ss.getSheetByName(SHEET_NAMES.lodging);
  if (!sheet) return;
  sheet.appendRow([
    d.tourGroup      || "",
    d.date           || "",
    d.lodgingLocation|| "",
    d.address        || "",
    d.contact        || "",
    d.notes          || ""
  ]);
}

function updateLodging(ss, d) {
  var sheet = ss.getSheetByName(SHEET_NAMES.lodging);
  if (!sheet) return false;
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var groupCol = headers.indexOf("Tour Group");
  var dateCol = headers.indexOf("Date");
  var locationCol = headers.indexOf("Lodging Location");
  var addressCol = headers.indexOf("Address");
  var contactCol = headers.indexOf("Contact");
  var notesCol = headers.indexOf("Notes");
  var normalizedIncomingGroup = normalizeTourGroup(d.tourGroup);
  var normalizedIncomingDate = lodgingDateKey(d.date);
  for (var i = 1; i < data.length; i++) {
    if (lodgingRowMatches(data[i], headers, d.original) ||
      (normalizeTourGroup(String(data[i][groupCol])) === normalizedIncomingGroup && lodgingDateKey(data[i][dateCol]) === normalizedIncomingDate)) {
      sheet.getRange(i+1, locationCol+1).setValue(d.lodgingLocation||"");
      sheet.getRange(i+1, addressCol+1).setValue(d.address||"");
      sheet.getRange(i+1, contactCol+1).setValue(d.contact||"");
      sheet.getRange(i+1, notesCol+1).setValue(d.notes||"");
      return true;
    }
  }
  return false;
}

function lodgingRowMatches(row, headers, original) {
  if (!original) return false;
  var groupCol = headers.indexOf("Tour Group");
  var dateCol = headers.indexOf("Date");
  var locationCol = headers.indexOf("Lodging Location");
  var addressCol = headers.indexOf("Address");
  var contactCol = headers.indexOf("Contact");
  var notesCol = headers.indexOf("Notes");
  return normalizeTourGroup(String(row[groupCol] || "")) === normalizeTourGroup(String(original.tourGroup || "")) &&
    lodgingDateKey(row[dateCol]) === lodgingDateKey(original.date) &&
    String(row[locationCol] || "") === String(original.lodgingLocation || "") &&
    String(row[addressCol] || "") === String(original.address || "") &&
    String(row[contactCol] || "") === String(original.contact || "") &&
    String(row[notesCol] || "") === String(original.notes || "");
}

function normalizeTourGroup(val) {
  var g = String(val || "").toLowerCase().trim();
  if (g === "mens" || g === "men's" || g === "men") return "men";
  if (g === "ladies" || g === "ladies'") return "ladies";
  return g;
}

function lodgingDateKey(val) {
  if (val === null || val === undefined || val === "") return "";

  if (Object.prototype.toString.call(val) === "[object Date]" && !isNaN(val.getTime())) {
    return toIsoDateKey(val);
  }

  var s = String(val).trim();
  if (!s) return "";

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    return s;
  }

  var isoLike = new Date(s);
  if (!isNaN(isoLike.getTime())) {
    return toIsoDateKey(isoLike);
  }

  var withYear = new Date(s + ", 2026");
  if (!isNaN(withYear.getTime())) {
    return toIsoDateKey(withYear);
  }

  return s.toLowerCase();
}

function toIsoDateKey(dateObj) {
  var y = dateObj.getFullYear();
  var m = ("0" + (dateObj.getMonth() + 1)).slice(-2);
  var d = ("0" + dateObj.getDate()).slice(-2);
  return y + "-" + m + "-" + d;
}

function deleteLodging(ss, d) {
  var sheet = ss.getSheetByName(SHEET_NAMES.lodging);
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var groupCol = headers.indexOf("Tour Group");
  var dateCol = headers.indexOf("Date");
  for (var i = data.length-1; i >= 1; i--) {
    if (String(data[i][groupCol]) === d.tourGroup && String(data[i][dateCol]) === d.date) {
      sheet.deleteRow(i+1);
      return true;
    }
  }
  return false;
}

// ============================================================
// HELPER
// ============================================================
function sheetToObjects(sheet) {
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}
