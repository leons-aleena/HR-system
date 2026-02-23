function doGet(e) {
  try {
    return HtmlService.createTemplateFromFile('main')
      .evaluate()
      .setTitle('HR System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput(
      'Error loading app: ' + err.message
    );
  }
}

/**
 * Includes HTML files
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


function getCandidatePool() {
  try {
    console.log("=== SERVER: getCandidatePool started ===");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Candidate Pool"); // 🔴 change if needed

    if (!sheet) {
      return { error: "Candidate Pool sheet not found" };
    }

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return []; // ✅ return EMPTY ARRAY
    }

    const candidatePool = [];

    for (let i = 1; i < data.length; i++) {
      
      const row = data[i];
      if(!row[12]) continue;
      
      if(row[20]) continue;
      const profileId = row[0]?.toString().trim();
      if (!profileId) continue; // skip invalid rows

      candidatePool.push({
        // timestamp: row[0] instanceof Date ? row[0].toISOString() : (row[0] || ""),
        profileId: profileId,
        candidateName: row[2]?.toString() || "",
        phoneNumber: row[3]?.toString() || "",
        candidateEmail: row[4]?.toString() || "",
        gender: row[5]?.toString() || "",
        homeTown: row[6]?.toString() || "",
        experienced: row[7]?.toString() || "",
        yearsOfExperience: row[8]?.toString() || "",
        languages: row[9]?.toString() || "",
        highestQualification: row[10]?.toString() || "",

        qualifiedCompanies: row[12]
          ? row[12].toString().split(", ").map(x => x.trim())
          : [],
        
        scheduledCompanies: row[13]
          ? row[13].toString().split(", ").map(x => x.trim())
          : [],

        rescheduledCompanies: row[14]
          ? row[14].toString().split(", ").map(x => x.trim())
          : [],

        candidateRejected: row[15]
          ? row[15].toString().split(", ").map(x => x.trim())
          : [],
        
        ghostedCompanies: row[16]
          ? row[16].toString().split(", ").map(x => x.trim())
          : [],

        companyRejected: row[17]
          ? row[17].toString().split(", ").map(x => x.trim())
          : [],

        selectedCompanies: row[18]
          ? row[18].toString().split(", ").map(x => x.trim())
          : [],

        shortlistedWalkinCompanies: row[19]
          ? row[19].toString().split(", ").map(x => x.trim())
          : []

        // clientInterviewScheduledCount: parseInt(row[7], 10) || 0
      });
    }

    console.log("✔ Candidates returned:", candidatePool.length);

    return candidatePool; // ✅ ARRAY returned

  } catch (error) {
    console.error("❌ getCandidatePool error:", error);
    return { error: error.toString() };
  }
}


function getCV(profileId) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = "Candidate Pool";
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      console.log("Candidate Pool sheet not found");
      return {
        success: false,
        error: "Candidate Pool sheet not found"
      };
    }

    const data = sheet.getDataRange().getValues();

    // Check if there is data besides header
    if (data.length <= 1) {
      return {
        success: false,
        error: "No data"
      };
    }

    // Start from row 2 (index 1) to skip header
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const profile = row[0]; // Column A
      const CV = row[2];      // Column B

      if (profile === profileId) {
        console.log("CV found");
        return {
          success: true,
          CV: CV
        };
      }
    }

    return {
      success: false,
      error: "No CV found for the requested profile"
    };

  } catch (error) {
    return {
      success: false,
      error: error.toString()
    };
  }
}
function markNotInterestedBackend(profileId, employeeId, company, type, reason) {
  try {
    if (!['companyRejection', 'candidateRejection'].includes(type)) {
      throw new Error('Invalid rejection type');
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Candidate Pool'); 
    const data = sheet.getDataRange().getValues();

    const colA = 0;   // Profile ID
    const colM = 12;  // Active companies
    const colP = 15;  // Candidate rejected
    const colR = 17;  // Company rejected

    let found = false;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][colA]) === String(profileId)) {
        found = true;

        // Remove company from Column M
        const companiesM = data[i][colM] || '';
        const updatedM = companiesM
          .split(',')
          .map(c => c.trim())
          .filter(c => c && c !== company)
          .join(', ');
        sheet.getRange(i + 1, colM + 1).setValue(updatedM);

        const targetCol = type === 'companyRejection' ? colR : colP;
        const rejectionType = type === 'companyRejection'
          ? 'Company Rejected'
          : 'Candidate Rejected';

        // Add company without duplication
        let existing = data[i][targetCol] || '';
        let list = existing.split(',').map(e => e.trim()).filter(Boolean);

        if (!list.includes(company)) {
          list.push(company);
        }

        sheet.getRange(i + 1, targetCol + 1).setValue(list.join(', '));

        updateHistoricDataSheet(
          employeeId,
          profileId,
          company,
          rejectionType,
          reason
        );

        break;
      }
    }

    if (!found) {
      throw new Error('Profile ID not found in sheet.');
    }

    return { success: true, message: 'Updated successfully' };

  } catch (error) {
    console.error(error);
    return { success: false, message: error.message };
  }
}


// ==================Inteview ID===============================================
function generateInterviewId(type) {
  const prefix = type === "walkin" ? "IW-" : "IV-";

  // Generate random 8-character alphanumeric (uppercase)
  const randomPart = Utilities.getUuid()
    .replace(/-/g, "")
    .substring(0, 8)
    .toUpperCase();

  return prefix + randomPart;
}


function saveInterviewDetails(details, employeeId) {
  try {
    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('Historic Data');

    if (!sheet) {
      throw new Error("Sheet 'Historic Data' not found");
    }

    const timestamp = Utilities.formatDate(
      new Date(),
      'Asia/Kolkata',
      'dd-MM-yyyy HH:mm:ss'
    );

    // Generate Interview ID
    const interviewId = generateInterviewId(details.type);

    // ---------- SAVE INTERVIEW ----------
    const row =
      details.type === "walkin"
        ? [
            employeeId,
            details.profileId,
            details.company,
            "Scheduled",
            timestamp,
            "",
            "",
            details.type,
            details.date,
            details.time,
            details.address,
            details.mapsLink,
            interviewId
          ]
        : [
            employeeId,
            details.profileId,
            details.company,
            "Scheduled",
            timestamp,
            "",
            "",
            details.type,
            details.date,
            details.time,
            details.meetingLink,
            "",
            interviewId
          ];

    sheet.appendRow(row);

    // ---------- UPDATE CANDIDATE POOL ----------
    updateCandidatePool(details.profileId, details.company);

    Logger.log("Interview saved successfully: " + interviewId);
    return interviewId;

  } catch (error) {
    Logger.log("Error saving interview: " + error.message);
    throw new Error("Failed to save interview: " + error.message);
  }
}



function updateCandidatePool(profileId, company) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Candidate Pool");

  if (!sheet) throw new Error("Candidate Pool sheet not found");

  const data = sheet.getDataRange().getValues();

  const col_profile = 0; // A
  const col_pending = 12; // M
  const col_scheduled = 13; // N

  for (let i = 1; i < data.length; i++) {
    if (data[i][col_profile] == profileId) {

      // ✅ Add to Scheduled (Column N)
      let scheduled = data[i][col_scheduled] || "";
      scheduled = scheduled
        ? `${scheduled}, ${company}`
        : company;

      sheet.getRange(i + 1, col_scheduled + 1).setValue(scheduled);

      // ✅ Remove from Pending (Column M)
      let pending = data[i][col_pending] || "";
      const updatedPending = pending
        .split(",")
        .map(c => c.trim())
        .filter(c => c && c !== company)
        .join(", ");

      sheet.getRange(i + 1, col_pending + 1).setValue(updatedPending);

      return;
    }
  }

  throw new Error("Profile ID not found in Candidate Pool");
}

function updateHistoricDataSheet( employeeId, profileId, company, status, explanation = "") {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Historic Data");

    if (!sheet) {
      throw new Error("Sheet 'Historic Data' not found");
    }

    // IST timestamp
const timestamp = Utilities.formatDate(
  new Date(),
  'Asia/Kolkata',
  'dd-MM-yyyy HH:mm:ss'
);

    sheet.appendRow([
      employeeId,        // Employee ID
      profileId,         // Profile ID
      company,           // Company Name
      status,            // Status
      "",
      timestamp,         // Status Updated Timestamp(result updated time)
      explanation        // Explanation
    ]);

  } catch (error) {
    Logger.log("Error updating Historic Data: " + error.message);
    throw new Error("Failed to update Historic Data: " + error.message);
  }
}


function ghostedCompanyToReschedule(employeeId,profileId,company){

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("Candidate Pool");

  if (!sheet) throw new Error("Candidate Pool sheet not found");

  const data = sheet.getDataRange().getValues();

  const col_profile = 0; // A
  const col_ghosted = 16; // Q
  const col_reschedule = 14; // O


console.log(company,profileId)
  for (let i = 1; i < data.length; i++) {
    if (data[i][col_profile] == profileId) {

      // ✅ Add to ReScheduled (Column O)
      let scheduled = data[i][col_reschedule] || "";
      scheduled = scheduled
        ? `${scheduled}, ${company}`
        : company;

      sheet.getRange(i + 1, col_reschedule + 1).setValue(scheduled);

      // ✅ Remove from Pending (Column Q)
      let pending = data[i][col_ghosted] || "";
      const updatedPending = pending
        .split(",")
        .map(c => c.trim())
        .filter(c => c && c !== company)
        .join(", ");

      sheet.getRange(i + 1, col_ghosted + 1).setValue(updatedPending);


      updateHistoricDataSheet(employeeId,profileId,company,"Reschedule")

      return {
  status: "success",
  profileId,
  company
};
;
    }
  }

  throw new Error("error ",Error);


}

function fetchCompanyAddressAndLocation(companyName, position) {

  // Example: Spreadsheet lookup
  // Columns: Company | Position | Address | LocationLink
  const sheet = SpreadsheetApp.getActive().getSheetByName("Hiring Requirments");

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (
      data[i][0] === companyName &&
      data[i][1] === position
    ) {
      return {
        address: data[i][14],
        locationLink: data[i][15]
      };
    }
  }

  return { address: "", locationLink: "" };
}

function getHistoricData(employeeId, profileId, companyName, statusType) {
  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName('Historic Data');

  const data = sheet.getDataRange().getValues();

  let ghostedCount = 0;
  let rescheduleCount=0;
  let candidateRejectedReason = '';
  let companyRejectedReason = '';
  let scheduledSpan='';

  for (let i = 1; i < data.length; i++) {
    const empId   = data[i][0];
    const profId  = data[i][1];
    const company = data[i][2];
    const status  = data[i][3];
    const colG    = data[i][6]; // Reason
    const colH    = data[i][7]; // Interview Type
    const colI    = data[i][8]; // Interview Date
    const colJ    = data[i][9]; // Interview Time

    if (
      String(empId) === String(employeeId) &&
      String(profId) === String(profileId) &&
      String(company) === String(companyName)
    ) {
      const normalizedStatus = String(status).toLowerCase();
const normalizedType   = String(statusType).toLowerCase();

if (normalizedType === 'ghosted' && normalizedStatus === 'ghosted') {
  ghostedCount++;
}

if (normalizedType === 'reschedule' && normalizedStatus === 'reschedule') {
  rescheduleCount++;
}

// candidate rejected → reason
if (
  normalizedType === 'candidate rejected' &&
  normalizedStatus === 'candidate rejected'
) {
  candidateRejectedReason = colG || '';
}

// company rejected → reason
if (
  normalizedType === 'company rejected' &&
  normalizedStatus === 'company rejected'
) {
  companyRejectedReason = colG || '';
}

// scheduled → scheduledSpan
if (normalizedType === 'scheduled' && normalizedStatus === 'scheduled') {
  const tz = Session.getScriptTimeZone();

  const interviewType = colH ? String(colH).trim() : '';

  const interviewDate = colI instanceof Date
    ? Utilities.formatDate(colI, tz, 'dd MMM yyyy')
    : '';

  const interviewTime = colJ instanceof Date
    ? Utilities.formatDate(colJ, tz, 'h:mm a')
    : '';

  scheduledSpan = [interviewType, interviewDate, interviewTime]
    .filter(Boolean)
    .join(' | ');
}



    }
  }

  return {
    ghostedCount: ghostedCount,
    rescheduleCount: rescheduleCount,
    candidateRejectedReason: candidateRejectedReason,
    companyRejectedReason:companyRejectedReason,
    scheduledSpan:scheduledSpan
  };
}

function getInterviewScheduleCount(employeeId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Historic Data');
  const data = sheet.getDataRange().getValues();
  let count = 0;

  const today = new Date();
  today.setHours(0, 0, 0, 0); // normalize to midnight

  for (let i = 1; i < data.length; i++) {
    const empId = data[i][0];   // Col A
    const status = data[i][3]; // Col D
    const timestampCell = data[i][4]; // Col E

    if (
      String(empId) === String(employeeId) &&
      String(status).toLowerCase() === 'scheduled' &&
      timestampCell
    ) {
      const scheduledDate = new Date(timestampCell);

      if (isNaN(scheduledDate)) continue;

      scheduledDate.setHours(0, 0, 0, 0); // safe now

      if (scheduledDate.getTime() === today.getTime()) {
        count++;
      }
    }
  }

  Logger.log("Interview count for today: " + count);
  return count;
}

function getScheduledCandidates(){
  try {
    console.log("=== SERVER: getCandidatePool started ===");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Historic Data"); // 🔴 change if needed

    if (!sheet) {
      return { error: "Historic Data sheet not found" };
    }

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return []; // ✅ return EMPTY ARRAY
    }

    const scheduledCandidatesPool = [];

    for (let i = 1; i < data.length; i++) {
      
      const row = data[i];
      if(!row[12]) continue; //skip rows without intervviewId
      
      const profileId = row[1]?.toString().trim();
      if (!profileId) continue; // skip invalid rows

      const status = row[3]?.toString().trim();
      if (status!=="Scheduled") continue; // skip invalid rows

      scheduledCandidatesPool.push({
        // timestamp: row[0] instanceof Date ? row[0].toISOString() : (row[0] || ""),
        employeeId:row[0]?.toString(),
        profileId: profileId,
        scheduledCompanyName:row[2]?.toString(),
        status:status,
        interviewScheduledTimestamp:row[4] instanceof Date ? row[4].toISOString() : (row[4] || ""),
        interviewType:row[7]?.toString(),
        interviewDate:row[8]?.toString(),
        interviewTime:row[9]?.toString(),
        addressOrLink:row[10]?.toString(),
        mapLink:row[11]?.toString() || "",
        interviewId:row[12]?.toString()
      

        // clientInterviewScheduledCount: parseInt(row[7], 10) || 0
      });
    }

    console.log("✔ Scheduled Candidates returned:", scheduledCandidatesPool.length);

    return scheduledCandidatesPool; // ✅ ARRAY returned

  } catch (error) {
    console.error("❌ scheduledCandidatesPool error:", error);
    return { error: error.toString() };
  }
}



function checkIfInterviewReminded(interviewId) {
  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName("Historic Data");

  const data = sheet.getDataRange().getValues();
  const colors = sheet.getDataRange().getBackgrounds();

  for (let i = 1; i < data.length; i++) {
    if (data[i][12] == interviewId) {
      return {
        success: true,
        isReminded: colors[i][12] === "#d4f8d4"
      };
    }
  }

  return {
    success: false,
    isReminded: false
  };
}

function markInterviewReminded(interviewId) {
  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName("Historic Data");

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][12] == interviewId) {
      sheet
        .getRange(i + 1, 1, 1, sheet.getLastColumn())
        .setBackground("#d4f8d4");

      return { success: true };
    }
  }

  return { success: false };
}



function updateInteviewResultsToSheet(decision, employeeId, profileId, interviewId, company, reason = '') {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const timestamp = Utilities.formatDate(
    new Date(),
    'Asia/Kolkata',
    'dd-MM-yyyy HH:mm:ss'
  );

  // ==============================
  // 🔴 HISTORIC DATA
  // ==============================
  const sheet1 = ss.getSheetByName("Historic Data");
  const historicData = sheet1.getDataRange().getValues();

  for (let i = 1; i < historicData.length; i++) {

    const row = historicData[i];

    if (row[12] === interviewId) {

      let existingEmployees = row[0] ? row[0].toString().split(",") : [];

      if (!existingEmployees.includes(employeeId)) {
        existingEmployees.push(employeeId);
      }

      sheet1.getRange(i + 1, 1).setValue(existingEmployees.join(","));
      sheet1.getRange(i + 1, 4).setValue(decision);
      sheet1.getRange(i + 1, 6).setValue(timestamp);
      sheet1.getRange(i + 1, 7).setValue(reason);

      console.log("historic data updated")

      break;
    }
  }

  // ==============================
  // 🔴 CANDIDATE POOL
  // ==============================
  const sheet2 = ss.getSheetByName("Candidate Pool");
  const poolData = sheet2.getDataRange().getValues();
  

  for (let i = 1; i < poolData.length; i++) {

    const row = poolData[i];
    const decisionLower = decision.toLowerCase();



    if (row[0] === profileId) {

      // 🔁 compute updated schedule ONCE
      let schedule = row[13] || "";
      const updatedSchedule = schedule
        .split(",")
        .map(c => c.trim())
        .filter(c => c && c !== company)
        .join(", ");

      // ---------- RESCHEDULE ----------
      if (decisionLower === "reschedule") {

        let val = row[14] || "";
        val = val ? `${val}, ${company}` : company;
        sheet2.getRange(i + 1, 15).setValue(val);

        sheet2.getRange(i + 1, 14).setValue(updatedSchedule);
        
        console.log("candidate pool data updated : ",decision)
        return;
      }

      // ---------- GHOSTED ----------
      if (decisionLower === "ghosted (p1)" || decisionLower === "ghosted (p2)") {

        let val = row[16] || "";
        val = val ? `${val}, ${company}` : company;
        sheet2.getRange(i + 1, 17).setValue(val);

        sheet2.getRange(i + 1, 14).setValue(updatedSchedule);

        console.log("candidate pool data updated : ",decision)
        
        return;
      }

      // ---------- COMPANY REJECTED ----------
      if (decisionLower === "company rejected") {

        let val = row[17] || "";
        val = val ? `${val}, ${company}` : company;
        sheet2.getRange(i + 1, 18).setValue(val);

        sheet2.getRange(i + 1, 14).setValue(updatedSchedule);

        console.log("candidate pool data updated : ",decision)
        
        return;
      }

      // ---------- CANDIDATE REJECTED ----------
      if (decisionLower === "candidate rejected (p2)" || decisionLower === "candidate rejected (p3)") {

        let val = row[15] || "";
        val = val ? `${val}, ${company}` : company;
        sheet2.getRange(i + 1, 16).setValue(val);

        sheet2.getRange(i + 1, 14).setValue(updatedSchedule);
        
        console.log("candidate pool data updated : ",decision)
        
        return;
      }

      // ---------- SELECTED ----------
      if (decisionLower === "selected") {

        let val = row[18] || "";
        val = val ? `${val}, ${company}` : company;
        sheet2.getRange(i + 1, 19).setValue(val);

        sheet2.getRange(i + 1, 14).setValue(updatedSchedule);
        
        console.log("candidate pool data updated : ",decision)
        
        return;
      }
    }
  }
}


// ------------------------------------

function getSelectedCandidates() {
  try {
    console.log("=== SERVER: getCandidatePool started ===");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Historic Data"); // change if needed

    if (!sheet) {
      return { error: "Historic Data sheet not found" };
    }

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return []; // return empty array if no data
    }

    const selectedCandidatesPool = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      if (!row[12]) continue; // skip rows without interviewId

      const profileId = row[1]?.toString().trim();
      if (!profileId) continue;

      const status = row[3]?.toString().trim();
      if (status !== "Selected") continue;

      // 🔹 COLUMN G → salary + joining date parsing
      const salaryCell = row[6]?.toString().trim(); // column G
      let salary = "";
      let joiningDate = "";

      if (salaryCell) {
        const match = salaryCell.match(/^(.*?)\s*\((.*?)\)$/);
        if (match) {
          salary = match[1];        // e.g. "20000 per annum"
          joiningDate = match[2];   // e.g. "2026-02-24"
        } else {
          salary = salaryCell; // fallback if format differs
        }
      }

      selectedCandidatesPool.push({
        employeeId: row[0]?.toString(),
        profileId: profileId,
        selectedCompanyName: row[2]?.toString(),
        status: status,

        salary: salary,
        joiningDate: joiningDate,

        interviewSelectedTimestamp:
          row[4] instanceof Date ? row[4].toISOString() : (row[4] || ""),
        interviewType: row[7]?.toString(),
        interviewDate: row[8]?.toString(),
        interviewTime: row[9]?.toString(),
        addressOrLink: row[10]?.toString(),
        mapLink: row[11]?.toString() || "",
        interviewId: row[12]?.toString()
      });
    }

    console.log("✔ Selected Candidates returned:", selectedCandidatesPool.length);
    return selectedCandidatesPool;

  } catch (error) {
    console.error("❌ selectedCandidatesPool error:", error);
    return { error: error.toString() };
  }
}

function joiningdateReschedule(interviewId, newDateStr) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName("Historic Data");

    if (!sheet) return "Sheet not found";

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return "No data";

    for (let row = 2; row <= lastRow; row++) {
      const id = sheet.getRange(row, 13).getValue(); // Column M

      if (String(id).trim() === String(interviewId).trim()) {

        const cell = sheet.getRange(row, 7); // Column G
        let text = cell.getValue().toString();

        // replace date inside ()
        let newText = text.replace(/\(\d{4}-\d{2}-\d{2}\)/, `(${newDateStr})`);

        cell.setValue(newText);
      }
    }

    return "Done";

  } catch (err) {
    return err.toString();
  }
}


function updateJoinedCandidates(
  employeeId,
  profileId,
  interviewId,
  company,
  joiningDate,
  salary
) {
  try {
    console.log("=== moveToJoined started ===");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historicSheet = ss.getSheetByName("Historic Data");
    const retentionSheet = ss.getSheetByName("Retention Pool");
    const candidatePoolSheet = ss.getSheetByName("Candidate Pool");

    if (!historicSheet) throw new Error("Historic Data sheet missing");
    if (!retentionSheet) throw new Error("Retention Pool sheet missing");
    if (!candidatePoolSheet) throw new Error("Candidate Pool sheet missing");

    const data = historicSheet.getDataRange().getValues();

    let rowIndex = -1;
    let row = null;

    // 🔍 find interviewId row
    for (let i = 1; i < data.length; i++) {
      if (data[i][12]?.toString() === interviewId?.toString()) {
        rowIndex = i + 1;
        row = data[i];
        break;
      }
    }

    if (!row) return { error: "Interview ID not found" };

    const status = row[3]?.toString().trim();
    if (status !== "Selected") {
      return { error: "Candidate not Selected" };
    }

    // ======================================
    // 1️⃣ UPDATE STATUS → JOINED
    // ======================================
    historicSheet.getRange(rowIndex, 4).setValue("Joined");

    // ======================================
    // 2️⃣ UPDATE COL G (salary + joiningDate)
    // ======================================
    const colGValue = `joined on ${joiningDate} with ${salary}`;
    historicSheet.getRange(rowIndex, 7).setValue(colGValue);

    // ======================================
    // 3️⃣ EMPLOYEE ID MERGE LOGIC (COL A)
    // ======================================
    const existingEmp = row[0]?.toString().trim();

    if (existingEmp) {
      const ids = existingEmp.split(",").map(x => x.trim());

      if (!ids.includes(employeeId)) {
        const updated = existingEmp + ", " + employeeId;
        historicSheet.getRange(rowIndex, 1).setValue(updated);
      }
    } else {
      historicSheet.getRange(rowIndex, 1).setValue(employeeId);
    }

    // ======================================
    // 4️⃣ ADD TO RETENTION POOL
    // ======================================
    retentionSheet.appendRow([
      employeeId,
      profileId,
      company,
      "Joined","",
      joiningDate,
      salary
    ]);

    // ======================================
    // 5️⃣ UPDATE CANDIDATE POOL COL U
    // ======================================
    const cpData = candidatePoolSheet.getDataRange().getValues();

    for (let i = 1; i < cpData.length; i++) {
      if (cpData[i][0]?.toString() === profileId?.toString()) {
        const today = Utilities.formatDate(
          new Date(),
          Session.getScriptTimeZone(),
          "dd-MM-yyyy"
        );

        const value = `${company}-(${today}) ONGOING`;

        // col U = 21
        candidatePoolSheet.getRange(i + 1, 21).setValue(value);
        break;
      }
    }

    console.log("✔ Candidate moved to Joined & Retention");

    return { success: true };

  } catch (err) {
    console.error(err);
    return { error: err.toString() };
  }
}

// -------------------------------------------
function updateCompanyReschedule(employeeId, profileId, interviewId, company) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historicSheet = ss.getSheetByName("Historic data");
  const poolSheet = ss.getSheetByName("Candidate Pool");

  const historicData = historicSheet.getDataRange().getValues();
  const poolData = poolSheet.getDataRange().getValues();

  // Column indices (0-based)
  const COL_EMPLOYEE_HIST = 0;    // A
  const COL_PROFILE_HIST = 1;     // B
  const COL_COMPANY_HIST = 2;     // C
  const COL_STATUS_HIST = 3;      // D
  const COL_INTERVIEW_HIST = 12;  // M

  const COL_PROFILE_POOL = 0;      // A
  const COL_PENDING_POOL = 12;     // M
  const COL_SCHEDULED_POOL = 13;   // N
  const COL_RESCHEDULED_POOL = 14; // O

  let lastStatus = null;

  // Helper to handle comma-separated lists
  const updateList = (list, remove, add) => {
    let items = (list || "").split(",").map(c => c.trim()).filter(c => c);
    if (remove) items = items.filter(c => c !== remove);
    if (add && !items.includes(add)) items.push(add);
    return items.join(", ");
  };

  // 1️⃣ Process Historic Data
  for (let i = 1; i < historicData.length; i++) {
    const row = historicData[i];

    // Mark interviewId as Cancelled
    if (row[COL_INTERVIEW_HIST] == interviewId) {
      historicSheet.getRange(i + 1, COL_STATUS_HIST + 1).setValue("Cancelled");
    }

    // Track last status for profile + company
    if (row[COL_PROFILE_HIST] == profileId && row[COL_COMPANY_HIST] == company) {
      lastStatus = row[COL_STATUS_HIST];

      // Update employeeId column (comma-separated)
      let employees = (row[COL_EMPLOYEE_HIST] || "").split(",").map(e => e.trim()).filter(e => e);
      if (!employees.includes(employeeId)) {
        employees.push(employeeId);
        historicSheet.getRange(i + 1, COL_EMPLOYEE_HIST + 1)
          .setValue(employees.join(", "));
      }
    }
  }

  // 2️⃣ Process Candidate Pool
  for (let i = 1; i < poolData.length; i++) {
    if (poolData[i][COL_PROFILE_POOL] == profileId) {

      // Determine destination column based on last status
      const targetCol = lastStatus ? COL_RESCHEDULED_POOL : COL_SCHEDULED_POOL;

      // Remove company from all 3 columns
      let pending = updateList(poolData[i][COL_PENDING_POOL], company, null);
      let scheduled = updateList(poolData[i][COL_SCHEDULED_POOL], company, null);
      let rescheduled = updateList(poolData[i][COL_RESCHEDULED_POOL], company, null);

      // Add company to target column
      if (targetCol === COL_SCHEDULED_POOL) {
        scheduled = updateList(scheduled, null, company);
      } else {
        rescheduled = updateList(rescheduled, null, company);
      }

      // Update the sheet
      poolSheet.getRange(i + 1, COL_PENDING_POOL + 1).setValue(pending);
      poolSheet.getRange(i + 1, COL_SCHEDULED_POOL + 1).setValue(scheduled);
      poolSheet.getRange(i + 1, COL_RESCHEDULED_POOL + 1).setValue(rescheduled);

      break; // Stop after matching profile
    }
  }
}

function createMeetFromForm(date, time, company) {
  const start = new Date(`${date}T${time}:00`);
  const end = new Date(start.getTime() + 60 * 1000);

  const event = {
    summary: `Interview – ${company}`,
    description: `Company: ${company}`,
    start: {
      dateTime: start.toISOString(),
      timeZone: "Asia/Kolkata"
    },
    end: {
      dateTime: end.toISOString(),
      timeZone: "Asia/Kolkata"
    },
    conferenceData: {
      createRequest: {
        requestId: Utilities.getUuid(),
        conferenceSolutionKey: { type: "hangoutsMeet" }
      }
    }
  };

  const createdEvent = Calendar.Events.insert(
    event,
    "primary",
    { conferenceDataVersion: 1 }
  );

  const meetLink = createdEvent.conferenceData.entryPoints
    .find(e => e.entryPointType === "video")
    .uri;

  // 🔍 DEBUG LOG
  console.log("Generated Google Meet link:", meetLink);
  Logger.log("Generated Google Meet link: %s", meetLink);

  return { meetLink };
}


function scheduleMassWalkin(employeeId,data) {
  try {
    const profileIds    = data.profileIds || [];
    const company       = data.company || "";
    const interviewDate = data.date || "";
    const interviewTime = data.time || "";
    const address       = data.address || "";
    const locationLink  = data.mapsLink || "";

    if (!profileIds.length) {
      return { status: "error", message: "No profileIds provided" };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historicSheet = ss.getSheetByName("Historic Data");
    const poolSheet     = ss.getSheetByName("Candidate Pool");
    const timestamp = Utilities.formatDate(
      new Date(),
      'Asia/Kolkata',
      'dd-MM-yyyy HH:mm:ss'
    );
    // =========================
    // 1️⃣ APPEND HISTORIC DATA
    // =========================
    const rows = profileIds.map(profileId => {
      const interviewID = generateInterviewId("walkin");

      return [
        employeeId,
        profileId,          // profileId
        company,            // company
        "Scheduled",        // status
        '',
        timestamp,          // timestamp
        "",                 // empty col
        "walkin",           // type
        interviewDate,      // date
        interviewTime,      // time
        address,            // address
        locationLink,       // map link
        interviewID         // interview id
      ];
    });

    historicSheet
      .getRange(historicSheet.getLastRow() + 1, 1, rows.length, rows[0].length)
      .setValues(rows);

    // =========================
    // 2️⃣ UPDATE CANDIDATE POOL
    // =========================
    const dataRange = poolSheet.getDataRange().getValues();

    const profileCol = 0;    // A
    const qualifiedCol = 12; // M
    const scheduledCol = 13; // N
    const rescheduleCol = 14;// O

    const splitCell = (val) =>
      val ? val.split(",").map(s => s.trim()).filter(Boolean) : [];

    const joinCell = (arr) => arr.join(", ");

    profileIds.forEach(profileId => {

      for (let i = 1; i < dataRange.length; i++) {

        if (dataRange[i][profileCol] === profileId) {

          let qualifiedArr  = splitCell(dataRange[i][qualifiedCol]);
          let scheduledArr  = splitCell(dataRange[i][scheduledCol]);
          let rescheduleArr = splitCell(dataRange[i][rescheduleCol]);

          // remove from M or O
          if (qualifiedArr.includes(company)) {
            qualifiedArr = qualifiedArr.filter(c => c !== company);
          }

          if (rescheduleArr.includes(company)) {
            rescheduleArr = rescheduleArr.filter(c => c !== company);
          }

          // add to N
          if (!scheduledArr.includes(company)) {
            scheduledArr.push(company);
          }

          // write back
          poolSheet.getRange(i + 1, qualifiedCol + 1)
            .setValue(joinCell(qualifiedArr));

          poolSheet.getRange(i + 1, scheduledCol + 1)
            .setValue(joinCell(scheduledArr));

          poolSheet.getRange(i + 1, rescheduleCol + 1)
            .setValue(joinCell(rescheduleArr));

          break;
        }
      }
    });

    return {
      status: "success",
      inserted: profileIds.length
    };

  } catch (err) {
    return {
      status: "error",
      message: err.toString()
    };
  }
}


// ------------------highlights---------------------
function updateCandidateHighlights(profileId, reason, tags) {
  const sheet = SpreadsheetApp.getActive()
    .getSheetByName("Candidate pool");

  const data = sheet.getRange("A:A").getValues(); // Profile IDs
  const rowIndex = data.findIndex(row => row[0] == profileId);

  if (rowIndex === -1) {
    return { success: false, message: "Profile not found" };
  }

  const formattedValue = reason + " (" + tags.join(",") + ")";

  // Column AA = 27
  sheet.getRange(rowIndex + 1, 27).setValue(formattedValue);

  return { success: true };
}

function getCandidateHighlight(profileId) {
  const sheet = SpreadsheetApp.getActive()
    .getSheetByName("Candidate pool");

  const data = sheet.getRange("A:AA").getValues();
  const row = data.find(r => r[0] == profileId);

  if (!row) return null;

  const cellValue = row[26]; // Column AA

  if (!cellValue) return null;

  const match = cellValue.match(/(.*)\((.*)\)/);

  if (!match) {
    return {
      reason: cellValue,
      tags: []
    };
  }

  return {
    reason: match[1].trim(),
    tags: match[2].split(",").map(t => t.trim())
  };
}

// -------====================================Retention Pool------------------------------------
function normalizeDate(d) {
  const x = new Date(d);
  x.setHours(0, 0, 0, 0);
  return x;
}
function getRetentionPool() {
  try {
    console.log("=== SERVER: getRetentionPool started ===");

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const retentionSheet = ss.getSheetByName("Retention Pool");
    const candidateSheet = ss.getSheetByName("Candidate Pool");

    if (!retentionSheet) {
      return { error: "Retention Pool sheet not found" };
    }
    if (!candidateSheet) {
      return { error: "Candidate Pool sheet not found" };
    }

    // ---------- Candidate Pool lookup (profileId → name, phone) ----------
    const candidateData = candidateSheet.getDataRange().getValues();
    const candidateMap = {};

    for (let i = 1; i < candidateData.length; i++) {
      const row = candidateData[i];
      const profileId = row[0]?.toString().trim();
      if (!profileId) continue;

      candidateMap[profileId] = {
        candidateName: row[2]?.toString() || "",
        phoneNumber: row[3]?.toString() || ""
      };
    }

    // ---------- Retention Pool ----------
    const retentionData = retentionSheet.getDataRange().getValues();

    if (retentionData.length <= 1) {
      return [];
    }

    const retentionPool = [];

    for (let i = 1; i < retentionData.length; i++) {
      const row = retentionData[i];

      const employeeId = row[0]?.toString().trim();
      const profileId  = row[1]?.toString().trim();
      if (!employeeId || !profileId) continue;

      // Column C: "Company Name (Job Title)"
      const companyRaw = row[2]?.toString().trim() || "";
      let companyName = "";
      let jobTitle = "";

      const match = companyRaw.match(/^(.+?)\s*\((.+?)\)$/);
      if (match) {
        companyName = match[1].trim();
        jobTitle = match[2].trim();
      } else {
        companyName = companyRaw;
      }

      const candidateInfo = candidateMap[profileId] || {};

      // ---------- Retention completion logic ----------
      const rawStatus = row[3]?.toString().trim().toLowerCase() || "";
      const retentionDays = Number(row[4]); // Column E
      const rawJoiningDate = row[5];        // Column F

      let canComplete = false;

      //--status bar--
      let completedDays = 0;

      if (rawJoiningDate instanceof Date) {
        const start = normalizeDate(rawJoiningDate);
        const today = normalizeDate(new Date());

        completedDays = Math.max(
          0,
          Math.floor((today - start) / (1000 * 60 * 60 * 24))
        );
      }

      if (
        rawJoiningDate instanceof Date &&
        retentionDays &&
        rawStatus !== "completed"
      ) {
        const start = normalizeDate(rawJoiningDate);
        const today = normalizeDate(new Date());

        const diffDays = Math.floor(
          (today - start) / (1000 * 60 * 60 * 24)
        );

        if (diffDays >= retentionDays) {
          canComplete = true;
        }
      }

      retentionPool.push({
        employeeId,
        profileId,

        // candidate details
        candidateName: candidateInfo.candidateName || "",
        phoneNumber: candidateInfo.phoneNumber || "",

        // retention details
        companyName,
        jobTitle,
        status: rawStatus,
        retentionTimeFrame: retentionDays,
        completedDays,
        joiningDate: rawJoiningDate instanceof Date
          ? rawJoiningDate.toISOString()
          : "",
        salary: row[6] || "",

        // ⭐ THIS IS THE KEY
        canComplete,
      });
    }

    console.log("✔ Retention records returned:", retentionPool.length);
    return retentionPool;

  } catch (error) {
    console.error("❌ getRetentionPool error:", error);
    return { error: error.toString() };
  }
}

function markRetentionCompleted(employeeId, profileId) {
  const sheet = SpreadsheetApp
    .getActive()
    .getSheetByName("Retention Pool");

  if (!sheet) {
    throw new Error("Retention Pool sheet not found");
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rowEmployeeId = String(data[i][0]).trim(); // Column A
    const rowProfileId  = String(data[i][1]).trim(); // Column B

    if (
      rowEmployeeId === String(employeeId).trim() &&
      rowProfileId === String(profileId).trim()
    ) {
      // Column D = Status
      sheet.getRange(i + 1, 4).setValue("Completed");
      return true;
    }
  }

  throw new Error("Retention record not found");
}


function removeCompanyFromList(cellValue, company) {
  if (!cellValue || !company) return cellValue || "";

  const companyNorm = company.toLowerCase().trim();

  const list = cellValue
    .split(",")
    .map(v => v.trim())
    .filter(v => {
      const vNorm = v.toLowerCase();
      return !vNorm.includes(companyNorm);
    });

  return list.join(", ");
}

function addCompanyToList(cellValue, company) {
  if (!cellValue) return company;
  const list = cellValue
    .split(",")
    .map(v => v.trim());
  if (!list.includes(company)) list.push(company);
  return list.join(", ");
}

function markRetentionExit(
  employeeId,
  profileId,
  company,
  type,
  reason,
  lastWorkingDate
) {
  const ss = SpreadsheetApp.getActive();
  const retentionSheet = ss.getSheetByName("Retention Pool");
  const candidateSheet = ss.getSheetByName("Candidate Pool");

  let companyWithJob = ""; // ✅ DECLARED ONCE

  /* ================= Retention Pool ================= */
  const retentionData = retentionSheet.getDataRange().getValues();

  for (let i = 1; i < retentionData.length; i++) {
    const row = retentionData[i];

    if (row[0] == employeeId && row[1] == profileId) {

      companyWithJob = row[2]; // Column C

      const joiningDate = new Date(row[5]);
      const lastDate = new Date(lastWorkingDate);

      const completedDays = Math.max(
        0,
        Math.floor((lastDate - joiningDate) / (1000 * 60 * 60 * 24))
      );

      retentionSheet.getRange(i + 1, 4).setValue(type);
      retentionSheet.getRange(i + 1, 8).setValue(reason);
      retentionSheet.getRange(i + 1, 9).setValue(lastDate);
      retentionSheet.getRange(i + 1, 10).setValue(completedDays);

      break;
    }
  }

  /* ================= Candidate Pool ================= */
  const candData = candidateSheet.getDataRange().getValues();

  for (let i = 1; i < candData.length; i++) {
    if (candData[i][0] == profileId) {

      // U = Joined
      const joined = candData[i][20];
      const updatedJoined = removeCompanyFromList(joined, companyWithJob);
      candidateSheet.getRange(i + 1, 21).setValue(updatedJoined);

      if (type === "resignation") {
        // R = Candidate Rejected
        const rejected = candData[i][17];
        candidateSheet
          .getRange(i + 1, 18)
          .setValue(addCompanyToList(rejected, companyWithJob));
      } else {
        // S = Company Rejected
        const rejected = candData[i][18];
        candidateSheet
          .getRange(i + 1, 19)
          .setValue(addCompanyToList(rejected, companyWithJob));
      }

      break;
    }
  }
}









