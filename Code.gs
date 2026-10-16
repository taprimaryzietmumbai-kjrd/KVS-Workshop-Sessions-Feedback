/**
 * KVS Workshop Sessions Feedback — Google Apps Script backend
 * Uses the same Google spreadsheet as Arriving-Participant-s-Facilitation-Hub,
 * but writes feedback to a separate "Session Feedback" tab. The Hub's
 * existing "selectable" tab is left untouched.
 */
const SPREADSHEET_ID = "1Uuqoz2KFpTcw9d2HI02fLmFPopTaLnKWwZ9aR-uYCeY";
const SHEET_NAME = "Session Feedback";

function doGet() {
  return json_({ success: true, message: "ZIET Mumbai feedback endpoint is ready." });
}

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const required = ["participant", "kvName", "designation", "region", "workshop", "sessionDate", "session", "resourcePerson", "learning", "overall"];
    const missing = required.filter(k => !String(data[k] == null ? "" : data[k]).trim());
    if (missing.length) throw new Error("Missing required fields: " + missing.join(", "));
    if (!data.ratings || typeof data.ratings !== "object") throw new Error("Session ratings are required.");
    if (!String(data.signature || "").startsWith("data:image/png;base64,")) throw new Error("Participant signature is required.");
    const ratingKeys = ["relevance", "clarity", "knowledge", "pedagogy", "engagement", "application", "materials"];
    ratingKeys.forEach(k => {
      const n = Number(data.ratings[k]);
      if (!Number.isInteger(n) || n < 1 || n > 5) throw new Error("Invalid rating: " + k);
    });
    const overall = Number(data.overall);
    if (!Number.isInteger(overall) || overall < 1 || overall > 5) throw new Error("Overall rating must be 1–5.");
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
    const headers = ["Submitted At", "Participant Name", "KV Name", "Designation", "KVS Region", "Workshop", "Session Date", "Session Title", "Resource Person", "Relevance (1-5)", "Clarity (1-5)", "Subject Knowledge (1-5)", "Pedagogy (1-5)", "Engagement (1-5)", "Application (1-5)", "Materials (1-5)", "Overall (1-5)", "Key Learning", "What Worked Well", "Suggestions for Improvement", "Follow-up Support", "Participant Signature (PNG data URL)"];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#073b66").setFontColor("#ffffff");
    }
    const r = data.ratings;
    sheet.appendRow([
      new Date(), safeCell_(data.participant), safeCell_(data.kvName), safeCell_(data.designation), safeCell_(data.region),
      safeCell_(data.workshop), safeCell_(data.sessionDate), safeCell_(data.session), safeCell_(data.resourcePerson),
      Number(r.relevance), Number(r.clarity), Number(r.knowledge), Number(r.pedagogy), Number(r.engagement), Number(r.application), Number(r.materials),
      overall, safeCell_(data.learning), safeCell_(data.strengths), safeCell_(data.improve), safeCell_(data.support)
    ]);
    return json_({ success: true });
  } catch (err) {
    return json_({ success: false, message: String(err && err.message || err) });
  }
}

function safeCell_(value) {
  const text = String(value == null ? "" : value).slice(0, 5000);
  return /^[=+@]/.test(text) || /^-/.test(text) ? "'" + text : text;
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}