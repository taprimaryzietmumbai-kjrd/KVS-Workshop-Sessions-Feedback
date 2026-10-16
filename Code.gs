/**
 * KVS Workshop Sessions Feedback — Google Apps Script backend
 * Writes responses to the "Session Feedback" tab in the existing spreadsheet.
 * Signatures and optional face photos are saved as private Drive files;
 * their links are recorded in the spreadsheet (not base64 image data).
 */
const SPREADSHEET_ID = "1Uuqoz2KFpTcw9d2HI02fLmFPopTaLnKWwZ9aR-uYCeY";
const SHEET_NAME = "Session Feedback";
const MEDIA_FOLDER_NAME = "ZIET Mumbai Workshop Feedback Media";

const HEADERS = [
  "Submitted At", "Participant Name", "KV Name", "Designation", "KVS Region",
  "Workshop", "Session Date", "Session Title", "Resource Person",
  "Relevance (1-5)", "Clarity (1-5)", "Subject Knowledge (1-5)",
  "Pedagogy (1-5)", "Engagement (1-5)", "Application (1-5)",
  "Materials (1-5)", "Overall (1-5)", "Key Learning", "What Worked Well",
  "Suggestions for Improvement", "Follow-up Support", "Classroom Implementation Plan",
  "Skill Acquisition", "Skills Learned / Strengthened", "Participant Signature File",
  "Participant Face Photo File"
];

function doGet() {
  return json_({ success: true, message: "ZIET Mumbai feedback endpoint is ready." });
}

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const required = [
      "participant", "kvName", "designation", "region", "workshop",
      "sessionDate", "session", "resourcePerson", "learning",
      "implementation", "skillAcquisition", "skillsDescription", "overall"
    ];
    const missing = required.filter(k => !String(data[k] == null ? "" : data[k]).trim());
    if (missing.length) throw new Error("Missing required fields: " + missing.join(", "));
    if (!data.ratings || typeof data.ratings !== "object") throw new Error("Session ratings are required.");
    if (!String(data.signature || "").startsWith("data:image/png;base64,")) {
      throw new Error("Participant signature is required.");
    }

    const ratingKeys = ["relevance", "clarity", "knowledge", "pedagogy", "engagement", "application", "materials"];
    ratingKeys.forEach(k => {
      const n = Number(data.ratings[k]);
      if (!Number.isInteger(n) || n < 1 || n > 5) throw new Error("Invalid rating: " + k);
    });
    const overall = Number(data.overall);
    if (!Number.isInteger(overall) || overall < 1 || overall > 5) throw new Error("Overall rating must be 1–5.");

    const mediaFolder = getMediaFolder_();
    const signatureUrl = saveImage_(data.signature, mediaFolder, "Signature", data.participant, "image/png");
    const photoUrl = data.facePhoto
      ? saveImage_(data.facePhoto, mediaFolder, "FacePhoto", data.participant, "")
      : "";

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
    ensureHeaders_(sheet);

    const r = data.ratings;
    const values = {
      "Submitted At": new Date(),
      "Participant Name": safeText_(data.participant),
      "KV Name": safeText_(data.kvName),
      "Designation": safeText_(data.designation),
      "KVS Region": safeText_(data.region),
      "Workshop": safeText_(data.workshop),
      "Session Date": safeText_(data.sessionDate),
      "Session Title": safeText_(data.session),
      "Resource Person": safeText_(data.resourcePerson),
      "Relevance (1-5)": Number(r.relevance),
      "Clarity (1-5)": Number(r.clarity),
      "Subject Knowledge (1-5)": Number(r.knowledge),
      "Pedagogy (1-5)": Number(r.pedagogy),
      "Engagement (1-5)": Number(r.engagement),
      "Application (1-5)": Number(r.application),
      "Materials (1-5)": Number(r.materials),
      "Overall (1-5)": overall,
      "Key Learning": safeText_(data.learning),
      "What Worked Well": safeText_(data.strengths),
      "Suggestions for Improvement": safeText_(data.improve),
      "Follow-up Support": safeText_(data.support),
      "Classroom Implementation Plan": safeText_(data.implementation),
      "Skill Acquisition": safeText_(data.skillAcquisition),
      "Skills Learned / Strengthened": safeText_(data.skillsDescription),
      "Participant Signature File": signatureUrl,
      "Participant Face Photo File": photoUrl
    };

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const row = headerRow.map(header => Object.prototype.hasOwnProperty.call(values, header) ? values[header] : "");
    sheet.appendRow(row);
    return json_({ success: true });
  } catch (err) {
    return json_({ success: false, message: String(err && err.message || err) });
  }
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  } else {
    const current = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    HEADERS.forEach(header => {
      if (current.indexOf(header) === -1) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
        current.push(header);
      }
    });
  }
  sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .setFontWeight("bold").setBackground("#073b66").setFontColor("#ffffff");
}

function getMediaFolder_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty("MEDIA_FOLDER_ID");
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (e) {}
  }
  const folder = DriveApp.createFolder(MEDIA_FOLDER_NAME);
  props.setProperty("MEDIA_FOLDER_ID", folder.getId());
  return folder;
}

function saveImage_(dataUrl, folder, label, participant, expectedMime) {
  const match = String(dataUrl || "").match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,([A-Za-z0-9+/=]+)$/);
  if (!match) throw new Error("Invalid " + label + " image data.");
  const mime = match[1];
  if (expectedMime && mime !== expectedMime) throw new Error("Invalid " + label + " image type.");
  if (!/^image\/(png|jpeg|jpg|webp)$/.test(mime)) throw new Error("Unsupported " + label + " image type.");
  const bytes = Utilities.base64Decode(match[2]);
  if (bytes.length > 5 * 1024 * 1024) throw new Error(label + " must be 5 MB or smaller.");
  const safeName = String(participant || "Participant").replace(/[^a-zA-Z0-9_-]/g, "_").slice(0, 60);
  const ext = mime === "image/png" ? "png" : (mime === "image/webp" ? "webp" : "jpg");
  const file = folder.createFile(Utilities.newBlob(bytes, mime, label + "_" + safeName + "_" + new Date().getTime() + "." + ext));
  // Keep files private in Drive; only users with appropriate Drive access can open them.
  return file.getUrl();
}

function safeText_(value) {
  const text = String(value == null ? "" : value).slice(0, 5000);
  return /^[=+@-]/.test(text) ? "'" + text : text;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
