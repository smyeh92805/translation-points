// ===== 皇家大衛中譯英集點站 ─ Google 試算表後端 =====
// 貼到「擴充功能 → Apps Script」後，部署為網頁應用程式（執行身分：我；存取權：所有人）。
// 更新程式後：部署 → 管理部署作業 → 鉛筆 → 版本「新版本」→ 部署（網址不變）。
//
//  * 「紀錄」工作表最底下固定一列「得點總計」（新紀錄會插在總計列上方）
//  * 自動建立「學生總表」工作表（每位學生的得點、獎勵點、總點數、完成單元數）
//  * 時間一律用台灣時間（Asia/Taipei）
//  * 「學生答案」欄保留學生每題的英文翻譯，老師可直接批閱

var TEACHER_KEY = 'DAVID2026';   // 要和 config.js 的 teacherCode 一樣
var SHEET_NAME = '紀錄';
var SUMMARY_NAME = '學生總表';
var TOTAL_LABEL = '得點總計';
var TZ = 'Asia/Taipei';
var HEADERS = ['時間','姓名','班級代碼','班級','單元代碼','單元','命中關鍵','關鍵總數','得點','獎勵點','題數','學生答案'];

function fmt_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm:ss');
  return String(v);
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  if (sh.getLastRow() === 0 || String(sh.getRange(1,1).getValue()).trim() === '') {
    sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(12, 420);
  }
  if (sh.getName() !== SHEET_NAME) sh.setName(SHEET_NAME);
  ensureTotalRow_(sh);
  ensureSummary_(ss);
  return sh;
}

function ensureTotalRow_(sh) {
  var last = sh.getLastRow();
  if (last >= 2 && String(sh.getRange(last,1).getValue()) === TOTAL_LABEL) return last;
  var row = last + 1;
  sh.getRange(row,1,1,HEADERS.length).setValues([[
    TOTAL_LABEL,'','','','','=COUNTA(E2:E' + (row-1) + ')','=SUM(G2:G' + (row-1) + ')','=SUM(H2:H' + (row-1) + ')',
    '=SUM(I2:I' + (row-1) + ')','=SUM(J2:J' + (row-1) + ')','=SUM(K2:K' + (row-1) + ')','=I' + row + '+J' + row
  ]]);
  sh.getRange(row,1,1,HEADERS.length).setFontWeight('bold').setBackground('#fdf3d0');
  sh.getRange(row,6).setNote('完成單元次數');
  sh.getRange(row,12).setNote('得點＋獎勵點 合計');
  return row;
}

function ensureSummary_(ss) {
  if (ss.getSheetByName(SUMMARY_NAME)) return;
  var s = ss.insertSheet(SUMMARY_NAME);
  s.getRange('A1').setFormula(
    '=QUERY(紀錄!A:K,"select B, C, D, sum(I), sum(J), count(E) where E is not null and A <> \'' + TOTAL_LABEL + '\' and A <> \'時間\' group by B, C, D ' +
    'label B \'姓名\', C \'班級代碼\', D \'班級\', sum(I) \'得點\', sum(J) \'獎勵點\', count(E) \'完成單元\'",1)');
  s.getRange('G1').setValue('總點數');
  s.getRange('G2').setFormula('=ARRAYFORMULA(IF(A2:A="","",D2:D+E2:E))');
  s.getRange('A1:G1').setFontWeight('bold').setBackground('#e8eefb');
  s.setFrozenRows(1);
}

function records_() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 3) return [];
  var v = sh.getRange(2,1,last-2,HEADERS.length).getValues();
  return v.filter(function(r){ return String(r[0]) !== TOTAL_LABEL && r[4] !== ''; }).map(function(r){
    return {time: fmt_(r[0]), name: String(r[1]), code: String(r[2]), className: String(r[3]), unitId: String(r[4]),
            unitTitle: String(r[5]), score: Number(r[6]), max: Number(r[7]), points: Number(r[8]) + Number(r[9]),
            base: Number(r[8]), bonus: Number(r[9]), qn: Number(r[10])};
  });
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var d = JSON.parse(e.postData.contents);
    if (d.action !== 'submit') return out_({ok:false, error:'unknown action'});
    var dup = records_().some(function(r){ return r.name === d.name && r.code.toUpperCase() === String(d.code).toUpperCase() && r.unitId === d.unitId; });
    if (!dup) {
      var sh = getSheet_();
      var totalRow = ensureTotalRow_(sh);
      sh.insertRowBefore(totalRow);
      var bonus = Number(d.bonus || 0), base = Number(d.points) - bonus;
      sh.getRange(totalRow,1,1,HEADERS.length).setValues([[
        Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'), d.name, d.code, d.className, d.unitId, d.unitTitle,
        d.score, d.max, base, bonus, d.qn, d.answers
      ]]).setFontWeight('normal').setBackground(null);
      sh.getRange(totalRow,12).setWrap(true);
      fixTotals_(sh);
    }
    return out_({ok:true, duplicate:dup});
  } catch (err) {
    return out_({ok:false, error:String(err)});
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// 插入新列後，總計列的公式範圍自動延伸
function fixTotals_(sh) {
  var row = sh.getLastRow();
  if (String(sh.getRange(row,1).getValue()) !== TOTAL_LABEL) return;
  var r = row - 1;
  sh.getRange(row,6,1,7).setFormulas([[
    '=COUNTA(E2:E' + r + ')','=SUM(G2:G' + r + ')','=SUM(H2:H' + r + ')','=SUM(I2:I' + r + ')',
    '=SUM(J2:J' + r + ')','=SUM(K2:K' + r + ')','=I' + row + '+J' + row
  ]]);
}

function doGet(e) {
  var p = e.parameter || {};
  try {
    if (p.action === 'student') {
      var code = String(p.code || '').toUpperCase();
      return out_({ok:true, records: records_().filter(function(r){ return r.name === p.name && r.code.toUpperCase() === code; })});
    }
    if (p.action === 'all') {
      if (String(p.key || '').toUpperCase() !== TEACHER_KEY.toUpperCase()) return out_({ok:false, error:'老師驗證碼錯誤'});
      return out_({ok:true, records: records_()});
    }
    return out_({ok:true, msg:'皇家大衛中譯英集點站後端運作中'});
  } catch (err) {
    return out_({ok:false, error:String(err)});
  }
}
