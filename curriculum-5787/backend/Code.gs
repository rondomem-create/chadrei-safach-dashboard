/**
 * Google Apps Script — backend for curriculum-5787 public progress site.
 * Deploy a NEW Web App: Execute as me; Who has access: Anyone.
 * Only aggregated metadata (teacher, class, subject, planned/done counts, links) is exposed.
 * No syllabus content or pupil information is ever returned.
 */
const SOURCE_BOOKS = [["א׳1","1i4vjdqxkbJi63i98TpWZmBJi35Fsno612zUtIpGdUjI",8],
  ["א׳3","12kgy0ps6-8dPeN3UrNax0A4kI03s7EXDglS3S2qOXD8",8],
  ["ב׳1","1izYT2E-yF9-mtorU2i20Kycv8-xpqmjvc3W56qWkE5I",8],
  ["ב׳3","1y9WwC4QEEDHq4L2jc7OlKGVbCASWL1xku0DjhBk96xU",8],
  ["ג׳1","1XdOxqZusJ5d-TA4kTGfaDuVHP8Z71TdTtqwmPttjlJQ",8],
  ["ג׳3","1zznIM5vAmnyO4XhlbmtxI3DNBYB2HOZ8gg0a2imNbwY",8],
  ["ד׳1","1WXjXdNQpm6d1AH8_F2mrpgl2KUoy55BcQ4_CQCvmnuY",9],
  ["ד׳2","14eXR6fvrkl911OirtxeLcvGha_4VaacZD9whAtgtRZE",9],
  ["ד׳3","1WpzWgkrvJaCGiEJKs2DeWzf9ptM-a9UsOzdBk-5Hsqg",10],
  ["ה׳1","1Lz0T2kp2wVvutrgNPuA-fcXhUTllWdYcuGBGCZUtZWE",10],
  ["ו׳1","1ufpslnysWQFFHY3BVYiYzs9uVM6w6EpTfUuSEOzClCE",10],
  ["ו׳2","1snVFWQnnJE95if4od5CD2rpWxjWv6L4wr_Eph_YspN0",9],
  ["ז׳1","1lFlG2PJFx5UyafpX3jFFKzAi0q9p4vAmeJUsgZCzvYk",10],
  ["ח׳1","1YWWjtZyCDD_gs3ihxgnEBweM9wmSxJAv7ad81lyfh5s",6]];
const CACHE_KEY = 'curriculum_5787_progress_v1';
const STATUS_SHEET = 'סטטוס מילוי';

function doGet(e) {
  const callback = String((e && e.parameter && e.parameter.callback) || '');
  if (!/^[A-Za-z_$][0-9A-Za-z_$]{0,100}$/.test(callback)) {
    return ContentService.createTextOutput('Invalid callback').setMimeType(ContentService.MimeType.TEXT);
  }
  const cache = CacheService.getScriptCache();
  let raw = cache.get(CACHE_KEY);
  if (!raw) {
    const all = [];
    SOURCE_BOOKS.forEach(function(item) {
      const clazz=item[0],fileId=item[1],count=item[2];
      const sheet=SpreadsheetApp.openById(fileId).getSheetByName(STATUS_SHEET);
      if(!sheet) throw new Error('No status sheet for '+clazz);
      const cells=sheet.getRange(5,1,count,10).getDisplayValues();
      cells.forEach(function(row) {
        if(!row[1]||!row[2])return;
        all.push({
          clazz:clazz,
          teacher:row[1],
          subject:row[2],
          total:Number(row[3])||48,
          filled:Number(row[4])||0,
          gid:String(row[8]),
          source:fileId
        });
      });
    });
    if(all.length!==121) throw new Error('Unexpected row count: '+all.length);
    raw=JSON.stringify({updated:new Date().toISOString(),rows:all});
    cache.put(CACHE_KEY,raw,300);
  }
  return ContentService.createTextOutput(callback+'('+raw+');')
                       .setMimeType(ContentService.MimeType.JAVASCRIPT);
}