function doGet(e) {
  const action = e.parameter.action;
  const body = e.parameter.data ? JSON.parse(e.parameter.data) : {};
  let result;

  if (action === "getCalendarData") {
    result = getCalendarData();
  }

  if (action === "getDashboardStats") {
    result = getDashboardStats();
  }

  if (action === "updateAssignedPerson") {
    result = updateAssignedPerson(body.rowId, body.people);
  }

  if (action === "updateActivityStatus") {
    result = updateActivityStatus(body.rowId, body.status, body.photos, body.gps);
  }

  if (action === "getUnavailableStaffForDate") {
    result = getUnavailableStaffForDate(body.rowId);
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.parameter.data);
    let result = {};

    if (body.action === "getCalendarData") {
      result = getCalendarData();
    }

    if (body.action === "getDashboardStats") {
      result = getDashboardStats();
    }

    if (body.action === "updateAssignedPerson") {
      result = updateAssignedPerson(body.rowId, body.people);
    }

    if (body.action === "updateActivityStatus") {
      result = updateActivityStatus(body.rowId, body.status, body.photos, body.gps);
    }

    if (body.action === "getUnavailableStaffForDate") {
      result = getUnavailableStaffForDate(body.rowId);
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/************************************************
 * LOCAL DEVELOPMENT TEST FUNCTION
 ************************************************/
function localTest() {
  Logger.log("LOCAL VERSION 12345");
}

/*************************************************************
 * REQUEST CALENDAR – UPDATED (TEMPLATE + PHOTO SUPPORT)
 *************************************************************/

const SHEET_NAME = "Form Responses 1";
const TEMPLATE_DOC_ID = "1pO4hODke9MuxLtFNWGHiadyhOOtOMz9r";
const SUMMARY_FOLDER_ID = "1Kok5861aFAQzhtvnNc8O1aLVZb7YGFYb";
const PHOTOS_FOLDER_ID = "1788VLYmDMyYzk7Rxdu1MOSAM12A6s80g";
const EVALUATION_FORM_LINK = "https://docs.google.com/spreadsheets/d/1BZ9gXDqmD7GZjZKzjlglwXYJ2lkXXR-ijeFKEE_JXao/edit?resourcekey=&gid=1157944124#gid=1157944124";

const GPS_JSON_COL = 20; // Column AB – stores municipality GPS JSON

/*************************************************************
 * 🔹 MAP RP NAME → RP EMAIL
 *************************************************************/
const RP_EMAILS = {
  "Katherine Andujare": "kayeandujare@gmail.com",
  "Rendell Lugtu": "rendelllugtu@gmail.com",
  "Krisha Cajucom": "kmcaj.workspace@gmail.com",
  "Emerene Pingol": "emaildocuments.1974@gmail.com",
  "Jodel Castillo": "mjrhcastillornd@gmail.com",
  "Camille Castañeda": "cascastaneda@ro4a.doh.gov.ph",
  "Cecille Lumbria": "lumbriaces.pmnp.doh@gmail.com",
  "Kay Legaspi": "kaydizon54@gmail.com",
  "Mark Reblora": "i.mark.reblora@gmail.com",
  "Mary Rose Comendador": "maryrose.lumbria.comendador@gmail.com",
  "MK Nolledo": "mariakathlynnn@gmail.com",
  "Jonathan Anat": "jo.anatflores13@gmail.com",
  "Paul Vicuña": "paulangelo.vicuna@yahoo.com",
  "Emmanuel Umali": "manuel.umali16@gmail.com",
  "Genella Moreno": "genellafsmoreno@gmail.com",
  "Cielo Cruz": "cielolaleicruz@gmail.com",
  "Lau Tamondong": "lauren.tamondong@gmail.com",
  "Kent Solibaga": "jksolibaga.rnd@gmail.com",
  "PDOHO Quezon": "pdohoquezon.nutrition@gmail.com"
};

/*************************************************************
 * GEOJSON AUTO-FILL (SAFE)
 *************************************************************/
const MUNICIPALITY_COORDS = {
  "Alabat":[14.1171,122.0269],"Buenavista":[13.7225,122.4241],
  "Burdeos":[14.9277,121.9587],"Calauag":[14.1157,122.2427],
  "Candelaria":[13.9287,121.4251],"Catanauan":[13.6545,122.3334],
  "General Nakar":[14.9068,121.4913],"Guinayangan":[13.8851,122.4241],
  "Gumaca":[13.9206,122.0989],"Jomalig":[14.6999,122.3674],
  "Macalelon":[13.7674,122.1860],"Panukulan":[14.9916,121.8790],
  "Patnanungan":[14.7896,122.1860],"Perez":[14.1886,121.9587],
  "Pitogo":[13.8095,122.1065],"Plaridel":[13.9569,122.0172],
  "Polillo":[14.7305,121.9701],"Quezon":[14.0628,122.1178],
  "San Andres":[13.3252,122.6504],"San Antonio":[13.8933,121.2897],
  "San Francisco":[13.2978,122.5599],"San Narciso":[13.5,122.5599],
  "Tagkawayan":[13.9684,122.5338],"Unisan":[13.8635,122.0155]
};

function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  const r = e.range.getRow();
  const c = e.range.getColumn();
  if (sh.getName() !== SHEET_NAME || r < 2) return;

  if (c === 5 && MUNICIPALITY_COORDS[e.range.getValue()]) {
  const [lat, lon] = MUNICIPALITY_COORDS[e.range.getValue()];
  sh.getRange(r, GPS_JSON_COL)
    .setValue(JSON.stringify({ lat, lon }));
  }
}

/*************************************************************
 * FETCH CALENDAR DATA
 *************************************************************/
function getCalendarData() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const data = sh.getDataRange().getValues();
  data.shift();

  const people = new Set([
    "Katherine Andujare", "Rendell Lugtu", "Krisha Cajucom", "Emerene Pingol",
    "Jodel Castillo", "Camille Castañeda", "Cecille Lumbria", "Mark Reblora",
    "Kay Legaspi", "Mary Rose Comendador", "MK Nolledo", "Jonathan Anat",
    "Paul Vicuña", "Emmanuel Umali", "Genella Moreno", "Cielo Cruz", "Lau Tamondong", 
    "Kent Solibaga", "PDOHO Quezon", "Refer", "Unassigned", "Deny"
  ]);

  const events = [];

  // Fetch the Staff Leave Sheet data
  const leaveSheet = SpreadsheetApp.getActive().getSheetByName("Leave");
  const leaveData = leaveSheet.getDataRange().getValues();
  leaveData.shift(); // Remove headers

  // Iterate through all data in the main sheet
  data.forEach((r, i) => {
    if (!r[10]) return;

    const start = new Date(r[10]);
    const end = r[12] ? new Date(r[12]) : new Date(r[10]);
    end.setDate(end.getDate() + 1);

    const assigned = r[17] || "Unassigned";
    const status = r[18] || "";  // 🔥 STATUS COLUMN (Column 19 in sheet)

    people.add(assigned);

    // 🔥 Default color (Assigned / Pending)
    let bgColor = "#3788d8"; // blue
    let titlePrefix = "";

    if (status === "Conducted") {
      bgColor = "#28a745"; // green
      titlePrefix = "✔ ";
    }

    if (status === "Denied") {
      bgColor = "#dc3545"; // red
    }

    if (status === "Referred") {
      bgColor = "#fd7e14"; // orange
    }

    const title = `${titlePrefix}${r[6]} — ${assigned} (${r[1]})`;

    // Check if assigned staff is on leave during this activity's time period
    let personIsOnLeave = false;

    leaveData.forEach((leaveRecord) => {
      const leaveName = leaveRecord[0];  // Staff Name from Leave sheet
      const leaveStart = new Date(leaveRecord[1]);
      const leaveEnd = leaveRecord[2] ? new Date(leaveRecord[2]) : new Date(leaveRecord[1]);
      leaveEnd.setDate(leaveEnd.getDate() + 1);

      // If assigned staff is on leave during the activity period
      if (assigned === leaveName && start <= leaveEnd && end >= leaveStart) {
        personIsOnLeave = true;
      }
    });

    events.push({
      id: i + 2,
      title: title,
      start: Utilities.formatDate(start, Session.getScriptTimeZone(), "yyyy-MM-dd"),
      end: Utilities.formatDate(end, Session.getScriptTimeZone(), "yyyy-MM-dd"),
      allDay: true,
      backgroundColor: bgColor,
      borderColor: bgColor,
      extendedProps: {
        row: i + 2,
        type: r[6],
        assigned: assigned,
        status: status,
        endUser: r[1],
        municipality: r[4],
        activityTitle: r[8],
        personIsOnLeave: personIsOnLeave  // Add the leave status to extendedProps
      }
    });
  });

  return { events, people: [...people].sort() };
}

/*************************************************************
 * UPDATE ASSIGNED PERSON
 *************************************************************/
function updateAssignedPerson(rowId, peopleStr) {
  const row = Number(rowId);
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);

  // ❌ DENY FLOW (no reason saved)
  if (peopleStr === "Deny") {
    sh.getRange(row, 18).setValue("Deny");     // Assigned RP
    sh.getRange(row, 19).setValue("Denied");   // Status

    // 📧 Notify requester (no reason)
    sendDenialEmailToRequester_(row);

    return "❌ Request denied and requester notified.";
  }

   // 🔁 REFER
  if (peopleStr === "Refer") {
    sh.getRange(row, 18).setValue("Refer");
    sh.getRange(row, 19).setValue("Referred");

    sendReferralEmail_(row);
    return "📨 Request referred to agency.";
  }

  // ✅ NORMAL ASSIGNMENT FLOW
  sh.getRange(row, 18).setValue(peopleStr);
  sh.getRange(row, 19).setValue("Assigned");
  SpreadsheetApp.flush();

  sendEmailToRequester_(row);
  sendEmailToAssignedRPs_(row);

  return "Assignment updated successfully";
}


/*************************************************************
 * 🔹 SEND EMAIL TO ASSIGNED RP(s)
 *************************************************************/
function sendEmailToAssignedRPs_(row) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);

  // 📍 Retrieve main activity details
  const municipality   = sh.getRange(row, 5).getDisplayValue();   // Column E - Municipality
  const requesterName  = sh.getRange(row, 2).getDisplayValue();   // Column C - Requester
  const activityTitle  = sh.getRange(row, 9).getDisplayValue();   // Column I - Activity / Type of Request
  const fileLink       = sh.getRange(row, 25).getDisplayValue();  // Column Y - Summary Link (if available)

  // 🗓️ Retrieve and format date(s)
  const activityStartRaw = sh.getRange(row, 11).getValue(); // Column K - Start Date
  const activityEndRaw   = sh.getRange(row, 13).getValue(); // Column M - End Date

  let activityDate = "";
  if (activityStartRaw) {
    const startFormatted = Utilities.formatDate(
      new Date(activityStartRaw),
      Session.getScriptTimeZone(),
      "MMM d, yyyy"
    );
    if (activityEndRaw) {
      const endFormatted = Utilities.formatDate(
        new Date(activityEndRaw),
        Session.getScriptTimeZone(),
        "MMM d, yyyy"
      );
      activityDate = `${startFormatted} - ${endFormatted}`;
    } else {
      activityDate = startFormatted;
    }
  } else {
    activityDate = "Not specified";
  }

  // 👥 Assigned Resource Persons
  const assignedNames = String(sh.getRange(row, 18).getValue())
    .split(",")
    .map(n => n.trim())
    .filter(Boolean);

  // ✉️ Send notification to each assigned RP
  assignedNames.forEach(rpName => {
    const rpEmail = RP_EMAILS[rpName];
    if (!rpEmail) return;

    const subject = "📌 You Have Been Assigned to a New Activity";

    const body = `
Hello ${rpName},

You have been assigned to handle the following activity request:

📍 Municipality: ${municipality}
📄 Activity: ${activityTitle}
📅 Activity Date: ${activityDate}
👤 Requested by: ${requesterName}

Please coordinate with the requester soon and update the activity status once completed.

Thank you,
PMNP RPMO CALABARZON
`;

    MailApp.sendEmail(rpEmail, subject, body);
  });
}

/*************************************************************
 * UPDATE ACTIVITY STATUS + PHOTO UPLOAD
 *************************************************************/
function updateActivityStatus(rowId, status, photos, gps) {
  const row = Number(rowId);
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);

  sh.getRange(row,19).setValue(status);

 if (gps && gps.lat && gps.lon) {
  verifyMunicipalityFromGPS_(row, gps);
}

  if (photos && photos.length) {
    const folder = DriveApp.getFolderById(PHOTOS_FOLDER_ID);
    photos.slice(0,3).forEach((p,i) => {
      const file = folder.createFile(
        Utilities.newBlob(
          Utilities.base64Decode(p.data),
          p.mimeType,
          `Row${row}_Photo${i+1}.jpg`
        )
      );
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      sh.getRange(row,21+i).setValue(file.getUrl());
    });
  }

 if (status.toLowerCase() === "conducted") {
  // Two Google Form links
  const EVAL_FORM_CSS1 = "https://tinyurl.com/CSS1TAME";
  const EVAL_FORM_CSS3 = "https://tinyurl.com/CSS3TAME";

  const cell = sh.getRange(row, 26); // ✅ Your CSS column

  // Create rich text containing two separate hyperlinks on separate lines
  const line1 = "View Evaluation CSS 1.0";
  const line2 = "View Evaluation CSS 3.0";
  const text = `${line1}\n${line2}`;

  const richText = SpreadsheetApp.newRichTextValue()
  .setText(text)
  .setLinkUrl(0, line1.length, EVAL_FORM_CSS1)
  .setLinkUrl(line1.length + 1, line1.length + 1 + line2.length, EVAL_FORM_CSS3)
  .build();


  cell.setRichTextValue(richText);
  cell.setVerticalAlignment("top");
  cell.setWrap(true); // ensure both lines show fully
}


  SpreadsheetApp.flush();
  return "Activity updated successfully";
}

/*************************************************************
 * 4️⃣ SEND EMAIL TO REQUESTER
 *************************************************************/
function sendEmailToRequester_(row) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);

  const requesterName  = sh.getRange(row, 3).getValue();  // Requester Name
  const requesterEmail = sh.getRange(row, 4).getValue();  // Requester Email
  const municipality   = sh.getRange(row, 5).getValue();  // Municipality
  const activityTitle  = sh.getRange(row, 9).getValue();  // Activity Title
  const fileLink       = sh.getRange(row,25).getDisplayValue();

  const assignedPeople = String(sh.getRange(row,18).getValue())
    .split(",")
    .map(n => n.trim())
    .filter(Boolean);

  if (!requesterEmail || assignedPeople.length === 0) return;

  const subject = "📌 Activity Assignment Update";
  const namesList = assignedPeople.join(", ");

  const body = `
Good Day ${requesterName},

Your request "${activityTitle}" in ${municipality} has been assigned to the following staff:
${namesList}

Please wait for the assigned staff to coordinate with you regarding this activity.

Thank you,
PMNP RPMO TEAM
  `;

  MailApp.sendEmail(requesterEmail, subject, body);
}

/*************************************************************
 * 🔹 DISTANCE CALCULATION (HAVERSINE FORMULA)
 *************************************************************/
function getDistanceKm(lat1, lon1, lat2, lon2) {
  const R = 6371; // Earth radius in km
  const dLat = deg2rad(lat2 - lat1);
  const dLon = deg2rad(lon2 - lon1);

  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(deg2rad(lat1)) *
      Math.cos(deg2rad(lat2)) *
      Math.sin(dLon / 2) *
      Math.sin(dLon / 2);

  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

function deg2rad(deg) {
  return deg * (Math.PI / 180);
}

/*************************************************************
 * 🔹 VERIFY GPS VS MUNICIPALITY
 *************************************************************/
const VERIFY_COL = 24; // 🔧 adjust if needed

function verifyMunicipalityFromGPS_(row, gps) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);

  // Verification (compare vs official coordinates)
  const municipality = sheet.getRange(row, 5).getValue();
  let verification = "Unknown";

  if (municipality && MUNICIPALITY_COORDS[municipality]) {
    const [refLat, refLon] = MUNICIPALITY_COORDS[municipality];
    const dist = getDistanceKm(gps.lat, gps.lon, refLat, refLon);
    verification = dist <= 5
      ? "✅ Within Municipality"
      : "⚠️ Outside Municipality";
  }

  sheet.getRange(row, VERIFY_COL).setValue(verification);
}

/*************************************************************
 * 📧 AUTO EMAIL NOTIFICATION ON FORM SUBMISSION
 * Sends email to a specific person (e.g. Jellie Palencia)
 * whenever a new form response is submitted
 *************************************************************/
function sendEmailOnFormSubmit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SHEET_NAME) return;

  const row = e.range.getRow();

  // 🔹 Basic details
  const name = sheet.getRange(row, 2).getDisplayValue();           // Column B or C: Requester Name
  const municipality = sheet.getRange(row, 5).getDisplayValue();   // Column E: Municipality
  const activityTitle = sheet.getRange(row, 9).getDisplayValue();  // Column I: Type of Request / Activity Title

  // 🗓️ Retrieve and format activity date(s)
  const startRaw = sheet.getRange(row, 11).getValue();  // Column K - Start Date
  const endRaw   = sheet.getRange(row, 13).getValue();  // Column M - End Date
  let activityDate = "";
  if (startRaw) {
    const startFmt = Utilities.formatDate(new Date(startRaw), Session.getScriptTimeZone(), "MMM d, yyyy");
    if (endRaw) {
      const endFmt = Utilities.formatDate(new Date(endRaw), Session.getScriptTimeZone(), "MMM d, yyyy");
      activityDate = `${startFmt} - ${endFmt}`;
    } else {
      activityDate = startFmt;
    }
  } else {
    activityDate = "Not specified";
  }

  // 📎 Request letter link (Column 15)
  const requestLetterLink = sheet.getRange(row, 15).getDisplayValue(); // Column O

  // 🔗 Short link to the main App Script / dashboard
  const appScriptLink = "https://pmnpivacalendarapp.web.app/";

  const email = "jelliepalencia@gmail.com"; // 🔧 Replace with her actual email
  const subject = "🆕 New Form Response Submitted – Please Assign Activity";

  const htmlBody = `
  <p>Hello Ma'am Jellie,</p>

  <p>
  A new form response has been submitted. Please review and assign the activity to a Resource Person.<br>
  Thank you so much.
  </p>

  <p><strong>📋 Request Details:</strong><br>
  • Name: ${name || "—"}<br>
  • Municipality: ${municipality || "—"}<br>
  • Activity / Type of Request: ${activityTitle || "—"}<br>
  • Activity Date: ${activityDate || "—"}<br>
  • Request Letter: ${requestLetterLink
    ? `<a href="${requestLetterLink}" target="_blank">View Request Letter</a>`
    : "(Not yet available)"}
  </p>

  <p>You can view it directly in our 
  <a href="${appScriptLink}" target="_blank">App Script link</a>.
  </p>

  <p>Thank you,<br>
  <strong>RPMO TAME Team</strong></p>
  `;

  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: htmlBody
  });
}

function sendDenialEmailToRequester_(row) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);

  const requesterName  = sh.getRange(row, 3).getDisplayValue(); // Name
  const requesterEmail = sh.getRange(row, 4).getDisplayValue(); // Email
  const municipality   = sh.getRange(row, 5).getDisplayValue();
  const activityTitle  = sh.getRange(row, 9).getDisplayValue();

  // 🗓 Activity date
  const startRaw = sh.getRange(row, 11).getValue();
  const endRaw   = sh.getRange(row, 13).getValue();

  let activityDate = "Not specified";
  if (startRaw) {
    const startFmt = Utilities.formatDate(new Date(startRaw), Session.getScriptTimeZone(), "MMM d, yyyy");
    activityDate = endRaw
      ? `${startFmt} - ${Utilities.formatDate(new Date(endRaw), Session.getScriptTimeZone(), "MMM d, yyyy")}`
      : startFmt;
  }

  if (!requesterEmail) return;

  const subject = "❌ Activity Request Update";

  const body = `
Good day ${requesterName},

We regret to inform you that your activity request conflicts with the team schedule.

📄 Request Details:
• Activity: ${activityTitle}
• Municipality: ${municipality}
• Proposed Date(s): ${activityDate}

You may submit a new request with a revised schedule or updated details. For further clarification, please contact your Project Associate, or the team may reach out to you regarding the activity.

Thank you for your understanding.

Sincerely,
RPMO TAME Team
`;

  MailApp.sendEmail(requesterEmail, subject, body);
}

/*************************************************************
 * 📧 SEND REFERRAL EMAIL TO AGENCY (WITH CC)
 *************************************************************/
function sendReferralEmail_(row) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);

  const requesterName  = sh.getRange(row, 3).getDisplayValue();
  const requesterEmail = sh.getRange(row, 4).getDisplayValue();
  const municipality   = sh.getRange(row, 5).getDisplayValue();
  const activityTitle  = sh.getRange(row, 9).getDisplayValue();

  // 🗓 Activity date
  const startRaw = sh.getRange(row, 11).getValue();
  const endRaw   = sh.getRange(row, 13).getValue();

  let activityDate = "Not specified";
  if (startRaw) {
    const startFmt = Utilities.formatDate(
      new Date(startRaw),
      Session.getScriptTimeZone(),
      "MMM d, yyyy"
    );
    activityDate = endRaw
      ? `${startFmt} - ${Utilities.formatDate(new Date(endRaw), Session.getScriptTimeZone(), "MMM d, yyyy")}`
      : startFmt;
  }

  // 📎 Request letter link uploaded via Google Form (Column 15)
  const requestLetterLink = sh.getRange(row, 15).getDisplayValue();

  // 🔧 CONFIGURE THESE
  const AGENCY_EMAIL = "calabarzon@nnc.gov.ph";  
  const CC_EMAILS    = "nutritionpmnp4a@gmail.com";

  const subject = "📨 Activity Request Referral – RPMO TAME";

  const body = `
Dear RNPC Bulante-Orongan,

This to forward the request Technical Assistance/ Capacity Building Activity shown in the details below:

📄 Activity: ${activityTitle}
📍 Municipality: ${municipality}
📅 Proposed Date(s): ${activityDate}
👤 Requested by: ${requesterName}

📎 Please see the attached Request Letter by the LGU:
${requestLetterLink || "(Not yet available)"}

For coordination, the requester may be contacted at:
📧 ${requesterEmail}

Thank you for your assistance and favorable response.

Sincerely,
PMNP RPMO CALABARZON
`;

  MailApp.sendEmail({
    to: AGENCY_EMAIL,
    cc: CC_EMAILS,
    subject: subject,
    body: body
  });
}

function getStaffOnLeaveForDate(rowId) {

  const row = Number(rowId);

  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);

  const activityStart = new Date(sh.getRange(row, 11).getValue());
  const activityEnd = sh.getRange(row, 13).getValue()
    ? new Date(sh.getRange(row, 13).getValue())
    : new Date(activityStart);

  const leaveSheet = SpreadsheetApp.getActive().getSheetByName("Leave"); // 🔥 match your tab name exactly
  if (!leaveSheet) return [];

  const leaveData = leaveSheet.getDataRange().getValues();
  leaveData.shift();

  const staffOnLeave = [];

  leaveData.forEach(r => {

    const leaveName = String(r[0]).trim().toLowerCase();
    const leaveStart = new Date(r[1]);
    const leaveEnd = r[2] ? new Date(r[2]) : new Date(r[1]);

    leaveEnd.setDate(leaveEnd.getDate() + 1);

    if (
      activityStart <= leaveEnd &&
      activityEnd >= leaveStart
    ) {
      staffOnLeave.push(leaveName);
    }

  });

  return staffOnLeave;
}

function getUnavailableStaffForDate(rowId) {

  const row = Number(rowId);
  const sh = SpreadsheetApp.getActive().getSheetByName("Form Responses 1");
  const data = sh.getDataRange().getValues();
  data.shift();

  const activityStart = new Date(sh.getRange(row, 11).getValue());
  const activityEnd = sh.getRange(row, 13).getValue()
    ? new Date(sh.getRange(row, 13).getValue())
    : new Date(activityStart);

  activityStart.setHours(0,0,0,0);
  activityEnd.setHours(23,59,59,999);

  const unavailable = new Set();

  // 🔥 CHECK OTHER ACTIVITIES
  data.forEach((r, i) => {

    const currentRow = i + 2;
    if (currentRow === row) return;

    const status = (r[18] || "").toLowerCase();
    if (status === "denied" || status === "referred") return;

    const assigned = String(r[17] || "").trim();
    if (!assigned) return;

    const otherStart = new Date(r[10]);
    const otherEnd = r[12] ? new Date(r[12]) : new Date(r[10]);

    otherStart.setHours(0,0,0,0);
    otherEnd.setHours(23,59,59,999);

    if (
      activityStart <= otherEnd &&
      activityEnd >= otherStart
    ) {
      assigned.split(",").forEach(name => {
        unavailable.add(name.trim().toLowerCase());
      });
    }

  });

  // 🔥 CHECK LEAVE
  const leaveSheet = SpreadsheetApp.getActive().getSheetByName("Leave");
  if (leaveSheet) {

    const leaveData = leaveSheet.getDataRange().getValues();
    leaveData.shift();

    leaveData.forEach(r => {

      const leaveName = String(r[0]).trim().toLowerCase();
      const leaveStart = new Date(r[1]);
      const leaveEnd = r[2] ? new Date(r[2]) : new Date(r[1]);

      leaveStart.setHours(0,0,0,0);
      leaveEnd.setHours(23,59,59,999);

      if (
        activityStart <= leaveEnd &&
        activityEnd >= leaveStart
      ) {
        unavailable.add(leaveName);
      }

    });
  }

  Logger.log("Unavailable: " + JSON.stringify([...unavailable]));

  return [...unavailable];
}

function getDashboardStats(){

const sh = SpreadsheetApp.getActive().getSheetByName("Form Responses 1");
const data = sh.getDataRange().getValues();
data.shift();

let stats={
total:0,
conducted:0,
assigned:0,
unassigned:0,
denied:0,
referred:0,
byType:{},
byMunicipality:{},
byStaff:{}
};

data.forEach(r=>{

stats.total++;

const municipality=(r[4]||"").toString().trim().toLowerCase();
const type=(r[6]||"").toString().trim().toLowerCase();
const assigned=(r[17]||"").toString().trim();
const status=(r[18]||"").toString().trim().toLowerCase();

/* STATUS COUNTS */

if(status==="conducted") stats.conducted++;

if(status==="denied") stats.denied++;

if(status==="referred") stats.referred++;

/* ASSIGNED / UNASSIGNED */

if(assigned && assigned!=="Unassigned"){
stats.assigned++;
}else{
stats.unassigned++;
}

/* ACTIVITY TYPE */

if(type){

stats.byType[type]=(stats.byType[type]||0)+1;

}

/* MUNICIPALITY */

if(municipality){

stats.byMunicipality[municipality]=(stats.byMunicipality[municipality]||0)+1;

}

/* STAFF */

if(assigned){

assigned.split(",").forEach(name=>{

const key=name.trim().toLowerCase();

if(!key) return;

stats.byStaff[key]=(stats.byStaff[key]||0)+1;

});

}

});

return stats;

}