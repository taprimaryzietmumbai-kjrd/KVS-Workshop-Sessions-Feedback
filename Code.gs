/**
 * KVS Workshop Sessions Feedback — Google Apps Script backend
 * Create a Google Sheet, then set SPREADSHEET_ID below.
 * Deploy as a Web App. Keep the destination spreadsheet access restricted.
 */
const SPREADSHEET_ID = "PASTE_GOOGLE_SHEET_ID_HERE";
const SHEET_NAME = "Session Feedback";

function doGet() {
  return json_({ success: true, message: "ZIET Mumbai feedback endpoint is ready." });
}

function doPost(e) {
  try {
    if (!SPREADSHEET_ID || SPREADSHEET_ID === "PASTE_GOOGLE_SHEET_ID_HERE") {
      throw new Error("Spreadsheet ID is not configured.");
    }
    const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const required = ["participant", "vidyalaya", "region", "designation", "workshop", "sessionDate", "session", "resourcePerson", "learning", "overall"];
    const missing = required.filter(k => !String(data[k] == null ? "" : data[k]).trim());
    if (missing.length) throw new Error("Missing required fields: " + missing.join(", "));
    if (!data.ratings || typeof data.ratings !== "object") throw new Error("Session ratings are required.");
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
    const headers = ["Submitted At", "Participant", "Vidyalaya", "KVS Region", "Designation", "Workshop", "Session Date", "Session Title", "Resource Person", "Relevance (1-5)", "Clarity (1-5)", "Subject Knowledge (1-5)", "Pedagogy (1-5)", "Engagement (1-5)", "Application (1-5)", "Materials (1-5)", "Overall (1-5)", "Key Learning", "What Worked Well", "Suggestions for Improvement", "Follow-up Support"];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
    }
    const r = data.ratings;
    sheet.appendRow([
      new Date(), safeCell_(data.participant), safeCell_(data.vidyalaya), safeCell_(data.region), safeCell_(data.designation),
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
  // Prevent user-entered text from being interpreted as a spreadsheet formula.
  return /^[=+@]/.test(text) || /^-/.test(text) ? "'" + text : text;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
