// ===== 皇家大衛中譯英集點站 設定檔（老師可自行修改） =====
window.SITE_CONFIG = {
  siteName: "皇家大衛中譯英集點站",
  siteNameEn: "Royal David Translation Points",
  // 班級代碼：學生登入時輸入「姓名＋班級代碼」。左邊是代碼（學生輸入），右邊是班級名稱。
  classCodes: {
    "RD2026": "皇家大衛補習班",
    "GEPT1": "英檢初級班",
    "GEPT2": "英檢中級班",
    "GEPT3": "英檢中高級班",
    "JCEE": "學測衝刺班"
  },
  // 老師驗證碼：在登入頁輸入此代碼可進入「老師後台」，看全班點數紀錄。
  teacherCode: "DAVID2026",
  // Google 試算表後端（Apps Script 網頁應用程式網址，以 /exec 結尾）。
  // 部署 apps-script/Code.gs 後把網址貼在這裡。留空則紀錄只存在學生瀏覽器。
  sheetUrl: "https://script.google.com/macros/s/AKfycbwoc4X_Ncj0rJ3cgYpgFUOeAXVIcPpMmq8gc6EdbFAaK0K0n_JSnBM_ycApqU8vCyxq/exec",
  // 點數規則
  pointsPerKey: 1,   // 每個「得分關鍵」寫對得 1 點
  perfectBonus: 3    // 一個單元的得分關鍵全部寫對，再加 3 點
};
