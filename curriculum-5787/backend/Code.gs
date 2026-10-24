/**
 * דשבורד תכניות לימודים ה׳תשפ״ז – רענון מאוחד.
 * פועלים רק בפרויקט Apps Script אחד, מחשבון משרד החינוך שיש לו הרשאות ל־14 הקבצים.
 * אין צורך באישור IMPORTANGE או בשיתוף הקבצים עם הציבור.
 */
const MASTER_ID = "128s2DXRqOPJ2cCoaW8B5yUZ6zg6g-ZkambWPaLMs6Bo";
const DASHBOARD_ID = "1DcM1un_q4GxxWg5gAGT7ylClE1iMhEZSWYXMCrKuY7Q";
const EXPECTED_ROWS = 121;
const CACHE_KEY = "curriculum_5787_snapshot_v2";

/** מפעילים פעם אחת, מאשרים גישה, והוא ייצור רענון שעתי אוטומטי. */
function setupOnce() {
  const result = syncProgress();
  const old = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === "syncProgress");
  old.forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("syncProgress").timeBased().everyHours(1).create();
  Logger.log("הותקן רענון שעתי. עודכנו " + result + " תכניות.");
}

/** מושך רק סטטיסטיקת מילוי: לא עובר על פרטי תלמידים ולא מפרסם תכניות לימודים. */
function syncProgress() {
  const master = SpreadsheetApp.openById(MASTER_ID);
  const sourceSheet = master.getSheetByName("מקורות");
  const target = master.getSheetByName("נתונים מאוחדים");
  if (!sourceSheet || !target) throw new Error("חסרה לשונית מקורות או נתונים מאוחדים");
  const sourceList = sourceSheet.getRange(5, 1, 14, 3).getValues();
  const all = [];
  sourceList.forEach(([clazz, fileId, expected]) => {
    const n = Number(expected);
    if (!clazz || !fileId || !Number.isInteger(n) || n < 1) throw new Error("שורת מקור לא תקינה: " + clazz);
    const file = SpreadsheetApp.openById(String(fileId));
    const status = file.getSheetByName("סטטוס מילוי");
    if (!status) throw new Error("חסרה לשונית סטטוס מילוי בכיתה " + clazz);
    const values = status.getRange(5, 1, n, 10).getValues();
    values.forEach(row => {
      if (!row[1] || !row[2] || row[8] === "") throw new Error("נתוני מקור חסרים בכיתה " + clazz);
      const planned = Number(row[3]);
      const filled = Number(row[4]);
      if (!Number.isFinite(planned) || planned <= 0 || !Number.isFinite(filled) || filled < 0)
        throw new Error("ספירת שבועות לא תקינה בכיתה " + clazz);
      const fraction = Math.min(filled / planned, 1);
      const gid = String(row[8]);
      const source = String(fileId);
      const directLink = '=HYPERLINK("https://docs.google.com/spreadsheets/d/' + source + '/edit#gid=' + gid + '","פתח תכנית")';
      all.push([clazz, String(row[1]), String(row[2]), planned, filled, fraction,
                fraction === 1 ? "הושלם" : filled === 0 ? "טרם התחיל" : "בתהליך",
                directLink, gid, source]);
    });
  });
  // אם קובץ אחד לא נקרא כראוי, לא משנים אף גיליון ולא משדרים הצלחה.
  if (all.length !== EXPECTED_ROWS) throw new Error("נתקבלו " + all.length + " תכניות במקום " + EXPECTED_ROWS);
  const dashboard = SpreadsheetApp.openById(DASHBOARD_ID).getSheetByName("חיבור נתונים");
  if (!dashboard) throw new Error("חסרה לשונית חיבור נתונים בדשבורד");
  target.getRange(5, 1, EXPECTED_ROWS, 10).setValues(all);
  dashboard.getRange(6, 1, EXPECTED_ROWS, 10).setValues(all);
  dashboard.getRange("L6").setValue("מחובר");
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  target.getRange("A3").setValue("עודכן לאחרונה: " + stamp);
  dashboard.getRange("A3").setValue("נתונים עודכנו לאחרונה: " + stamp);
  CacheService.getScriptCache().remove(CACHE_KEY);
  return all.length;
}

/** שירות קריאה ציבורי: אך ורק סטטיסטיקת מילוי, ללא תוכן לימודי או נתוני תלמידים. */
function doGet(e) {
  const callback = String(e && e.parameter && e.parameter.callback || "");
  if (!/^[A-Za-z_$][0-9A-Za-z_$]{0,80}$/.test(callback))
    return ContentService.createTextOutput("Invalid callback").setMimeType(ContentService.MimeType.TEXT);
  const cache = CacheService.getScriptCache();
  let raw = cache.get(CACHE_KEY);
  if (!raw) {
    const rows = SpreadsheetApp.openById(MASTER_ID)
       .getSheetByName("נתונים מאוחדים").getRange(5, 1, EXPECTED_ROWS, 10).getValues();
    const safe = rows.map(r => ({
      clazz: String(r[0]), teacher: String(r[1]), subject: String(r[2]),
      total: Number(r[3]), filled: Number(r[4]),
      gid: String(r[8]), source: String(r[9])
    }));
    raw = JSON.stringify({updated:new Date().toISOString(),rows:safe});
    cache.put(CACHE_KEY, raw, 300);
  }
  return ContentService.createTextOutput(callback + "(" + raw + ");")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
