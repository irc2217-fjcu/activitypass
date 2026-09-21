/**
 * =========================================================================
 * 區資護照 (北區區域原住民族學生資源中心 研習時數查詢系統)
 * Google Apps Script (GAS) 後端核心程式碼
 * 
 * 試算表 ID: 12lHpaiWIYeuBzTum8KNEfUPKvTm0PDpMrXODrf2sHa8
 * 
 * 工作表架構說明：
 * 1. 工作表1 (時數總表)
 *    Col A (0): 姓名
 *    Col B (1): 身分證字號
 *    Col C (2): 開課單位 / 服務單位
 *    Col D (3): 課程編號 / 課程分類
 *    Col E (4): 課程名稱
 *    Col F (5): 活動開始時間
 *    Col G (6): 活動結束時間
 *    Col H (7): 時數
 *    Col I (8): 審核欄位 (待審核 / 通過 / 不通過)
 *    Col J (9): 佐證圖片連結
 *    Col K (10): 紀錄ID
 *    Col L (11): 管理員標記
 * 
 * 2. 學員名冊 (身分驗證與名冊表)
 *    Col A (0): 姓名
 *    Col B (1): 身分證字號
 *    Col C (2): 出生年月日 (西元 YYYYMMDD 或 日期)
 *    Col D (3): 服務單位
 *    Col E (4): 管理員欄位 (若值為 'V' 則具備管理權限)
 * =========================================================================
 */

const SPREADSHEET_ID = "12lHpaiWIYeuBzTum8KNEfUPKvTm0PDpMrXODrf2sHa8";
const SHEET_RECORDS = "工作表1";
const SHEET_MEMBERS = "學員名冊";

/**
 * POST 請求分派入口
 */
function doPost(e) {
  try {
    const contents = e.postData ? e.postData.contents : "{}";
    const req = JSON.parse(contents);
    const action = req.action;

    if (action === "verifyUser") {
      return jsonResponse(verifyUser(req.idNumber, req.birthDate));
    } else if (action === "getDashboard") {
      return jsonResponse(getDashboard(req.idNumber, req.birthDate));
    } else if (action === "queryData") {
      return jsonResponse(queryData(req.idNumber, req.year));
    }

    return jsonResponse({ status: "error", message: "未知的請求動作: " + action });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

/**
 * GET 請求探測入口
 */
function doGet(e) {
  return jsonResponse({
    status: "online",
    system: "區資護照 研習時數查詢系統 API",
    version: "1.0.0",
    time: Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss")
  });
}

// =========================================================================
// 輔助函式
// =========================================================================

function cleanId(id) {
  return id ? id.toString().trim().toUpperCase() : "";
}

function normalizeBirthDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Taipei", "yyyyMMdd");
  }
  return val.toString().trim().replace(/[-/\s.]/g, "");
}

function isVMark(val) {
  if (!val) return false;
  const s = val.toString().trim().toUpperCase();
  return s === "V" || s === "TRUE" || s === "YES" || s === "Y";
}

function extractCleanUrl(val, formula) {
  const f = formula ? formula.toString().trim() : "";
  if (f) {
    const m = f.match(/https?:\/\/[^\s"'\)]+/i);
    if (m) return m[0];
  }
  const v = val ? val.toString().trim() : "";
  if (v) {
    const m = v.match(/https?:\/\/[^\s"'\)]+/i);
    if (m) return m[0];
  }
  return "";
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * 從「學員名冊」驗證身分並回傳學員資訊
 */
function findMember(idNumber, birthDate) {
  const targetId = cleanId(idNumber);
  const targetDob = normalizeBirthDate(birthDate);

  if (!targetId || !targetDob) return null;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_MEMBERS);
  if (!sheet) throw new Error("找不到工作表：" + SHEET_MEMBERS);

  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowId = cleanId(row[1]); // Col B
    const rowDob = normalizeBirthDate(row[2]); // Col C

    if (rowId === targetId && rowDob === targetDob) {
      const isAdmin = isVMark(row[4]) || isVMark(row[11]);
      return {
        rowIndex: i + 1,
        name: row[0] ? row[0].toString().trim() : "",
        idNumber: rowId,
        unit: row[3] ? row[3].toString().trim() : "北區原資中心",
        isAdmin: isAdmin
      };
    }
  }

  return null;
}

// =========================================================================
// 主要查詢功能
// =========================================================================

/**
 * 身分核對
 */
function verifyUser(idNumber, birthDate) {
  const user = findMember(idNumber, birthDate);
  if (!user) {
    return { status: "error", message: "身分證字號或出生年月日核對不符！" };
  }
  return { status: "success", user: user };
}

/**
 * 取得學員區資時數儀表板資料 (依開課單位整合)
 */
function getDashboard(idNumber, birthDate) {
  const user = findMember(idNumber, birthDate);
  if (!user) {
    return { status: "error", message: "身分驗證失敗，請確認身分證與出生日期！" };
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_RECORDS);
  if (!sheet) throw new Error("找不到工作表：" + SHEET_RECORDS);

  const values = sheet.getDataRange().getValues();
  const formulas = sheet.getDataRange().getFormulas();
  const records = [];
  let totalHours = 0;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowId = cleanId(row[1]); // Col B: 身分證
    if (rowId !== user.idNumber) continue;

    const courseUnit = row[2] ? row[2].toString().trim() : (user.unit || "北區原資中心"); // Col C: 開課單位
    const courseClass = row[3] ? row[3].toString().trim() : ""; // Col D: 課程分類
    const courseName = row[4] ? row[4].toString().trim() : ""; // Col E: 課程名稱

    let sTime = row[5];
    if (sTime instanceof Date) {
      sTime = Utilities.formatDate(sTime, "Asia/Taipei", "yyyy-MM-dd HH:mm");
    } else {
      sTime = sTime ? sTime.toString().trim() : "";
    }

    let eTime = row[6];
    if (eTime instanceof Date) {
      eTime = Utilities.formatDate(eTime, "Asia/Taipei", "yyyy-MM-dd HH:mm");
    } else {
      eTime = eTime ? eTime.toString().trim() : "";
    }

    const h = parseFloat(row[7]) || 0;
    const status = row[8] ? row[8].toString().trim() : "待審核";
    const fileUrl = extractCleanUrl(row[9], formulas[i] ? formulas[i][9] : "");
    const recId = row[10] ? row[10].toString().trim() : ("ROW_" + (i + 1));

    if (status === "通過" || status === "已審核") {
      totalHours += h;
    }

    records.push({
      recordId: recId,
      courseName: courseName,
      courseUnit: courseUnit,
      mainCategory: courseUnit, // 相容欄位
      subCategory: courseClass,
      startTime: sTime,
      endTime: eTime,
      hours: h,
      status: status || "待審核",
      fileUrl: fileUrl
    });
  }

  // 排序：依活動時間由新至舊
  records.sort((a, b) => (b.startTime || "").localeCompare(a.startTime || ""));

  return {
    status: "success",
    user: user,
    totalHours: Math.round(totalHours * 10) / 10,
    records: records
  };
}

/**
 * 相容式單一身分證查詢 (供外部純身分證查詢使用)
 */
function queryData(idNumber, year) {
  const targetId = cleanId(idNumber);
  if (!targetId) return { status: "error", message: "身分證字號不可為空！" };

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const memSheet = ss.getSheetByName(SHEET_MEMBERS);
  let userName = "";
  let userUnit = "北區原資中心";

  if (memSheet) {
    const memValues = memSheet.getDataRange().getValues();
    for (let i = 1; i < memValues.length; i++) {
      if (cleanId(memValues[i][1]) === targetId) {
        userName = memValues[i][0] ? memValues[i][0].toString().trim() : "";
        userUnit = memValues[i][3] ? memValues[i][3].toString().trim() : userUnit;
        break;
      }
    }
  }

  const recSheet = ss.getSheetByName(SHEET_RECORDS);
  if (!recSheet) return { status: "error", message: "找不到時數記錄表！" };

  const values = recSheet.getDataRange().getValues();
  const grouped = {};
  let totalHours = 0;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (cleanId(row[1]) !== targetId) continue;

    if (!userName && row[0]) {
      userName = row[0].toString().trim();
    }

    const sTime = row[5] ? row[5].toString() : "";
    if (year && !sTime.startsWith(year)) continue;

    const status = row[8] ? row[8].toString().trim() : "待審核";
    if (status !== "通過" && status !== "已審核") continue;

    const unit = row[2] ? row[2].toString().trim() : userUnit;
    const courseName = row[4] ? row[4].toString().trim() : "研習課程";
    const h = parseFloat(row[7]) || 0;

    if (!grouped[unit]) {
      grouped[unit] = { subTotal: 0, courses: [] };
    }
    grouped[unit].courses.push({ name: courseName, hours: h });
    grouped[unit].subTotal += h;
    totalHours += h;
  }

  if (totalHours === 0 && Object.keys(grouped).length === 0 && !userName) {
    return { status: "not_found", message: "查無此身分證字號的時數資料" };
  }

  for (const u in grouped) {
    grouped[u].subTotal = Math.round(grouped[u].subTotal * 10) / 10;
  }

  return {
    status: "success",
    name: userName,
    idNumber: targetId,
    unit: userUnit,
    filteredYear: year,
    totalHours: Math.round(totalHours * 10) / 10,
    groupedCourses: grouped
  };
}
