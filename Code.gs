const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
const DATA_SHEET_NAME = "CheckinData";
const TEST_DATA_SHEET_NAME = "CheckinData_Test";
const CONFIG_SHEET_NAME = "Config";
const SYSTEM_SECRET_TOKEN = "EVT_SECURE_2026"; // Khóa chống spam cho trang khách (nằm lộ trong index.html nên KHÔNG dùng để khóa admin)
// Mật khẩu admin KHÔNG nằm trong code này: cấu hình trong Apps Script > Project Settings > Script Properties, key = ADMIN_TOKEN.
// Chưa cấu hình ADMIN_TOKEN thì mọi thao tác admin (lưu cấu hình / xóa dữ liệu) sẽ bị từ chối.

function doGet(e) {
  try {
    const cfgSheet = SPREADSHEET.getSheetByName(CONFIG_SHEET_NAME);
    if (!cfgSheet) return createJsonResponse({});
    const configValues = cfgSheet.getDataRange().getValues();
    const config = {};
    configValues.forEach((row, idx) => {
      if (idx > 0 && row[0]) config[row[0]] = row[1];
    });
    return createJsonResponse(config);
  } catch (err) {
    return createJsonResponse({});
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ status: "error", message: "Thiếu dữ liệu gửi lên" });
    }
    const postData = JSON.parse(e.postData.contents);

    // Chặn request bên ngoài nếu thiếu mã bí mật
    if (postData.token !== SYSTEM_SECRET_TOKEN) {
      return createJsonResponse({ status: "error", message: "Truy cập không hợp lệ!" });
    }

    // Thao tác admin bắt buộc mật khẩu riêng (adminToken), KHÔNG chấp nhận token trang khách.
    if (postData.action === "updateSettings" || postData.action === "resetData" || postData.action === "resetTestData") {
      const storedAdminToken = PropertiesService.getScriptProperties().getProperty("ADMIN_TOKEN");
      if (!storedAdminToken) {
        return createJsonResponse({ status: "error", message: "Admin chưa được cấu hình: thiếu ADMIN_TOKEN trong Script Properties của Apps Script." });
      }
      if (postData.adminToken !== storedAdminToken) {
        return createJsonResponse({ status: "error", message: "Mật khẩu admin không đúng!" });
      }
    }

    if (postData.action === "updateSettings") return handleUpdateSettings(postData);
    if (postData.action === "resetData") return handleResetData();
    if (postData.action === "resetTestData") return handleResetTestData();

    return handleCheckinBulletproof(postData);
  } catch (err) {
    return createJsonResponse({ status: "error", message: "Lỗi xử lý: " + err.toString() });
  }
}

function handleCheckinBulletproof(data) {
  const name = (data.name || "").trim();
  const phone = (data.phone || "").trim();
  const eventCode = (data.event || "DEFAULT").trim();
  const customData = data.customData || {};

  if (!name || !phone) {
    return createJsonResponse({ status: "error", message: "Vui lòng nhập đủ Họ tên và Số điện thoại!" });
  }

  // Khóa đa luồng 20 giây xử lý hàng đợi tuần tự
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (e) {
    return createJsonResponse({ status: "error", message: "Hệ thống đang tiếp nhận đông khách, vui lòng thử lại sau 2 giây!" });
  }

  try {
    // Stress test (event STRESS_TEST) ghi vào sheet test riêng, KHÔNG chiếm số thật.
    const targetSheetName = (eventCode === "STRESS_TEST") ? TEST_DATA_SHEET_NAME : DATA_SHEET_NAME;
    let sheet = SPREADSHEET.getSheetByName(targetSheetName);
    if (!sheet) {
      sheet = SPREADSHEET.insertSheet(targetSheetName);
      sheet.getRange("A1:E1").setValues([["STT", "Họ và tên", "Số điện thoại", "Sự kiện", "Thời gian"]]);
      SpreadsheetApp.flush();
    }

    // 1. Tự động phục hồi vé cũ: khớp SĐT đã chuẩn hóa + ĐÚNG event (trước đây chỉ khớp SĐT, sai event vẫn trả vé cũ).
    const lastRowForSearch = sheet.getLastRow();
    if (lastRowForSearch > 1) {
      const rows = sheet.getRange(2, 1, lastRowForSearch - 1, 4).getValues();
      const wantedPhone = normalizePhone(phone);
      for (let i = 0; i < rows.length; i++) {
        const rowPhone = rows[i][2] ? rows[i][2].toString() : "";
        const rowEvent = rows[i][3] ? rows[i][3].toString() : "";
        if (rowPhone && normalizePhone(rowPhone) === wantedPhone && rowEvent === eventCode) {
          return createJsonResponse({
            status: "success",
            id: String(rows[i][0]).padStart(3, '0'),
            name: rows[i][1] ? rows[i][1].toString() : "",
            recovered: true
          });
        }
      }
    }

    // 2. Tự động tạo thêm cột mới cho các trường tùy biến
    const lastCol = Math.max(sheet.getLastColumn(), 5);
    let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].filter(h => h !== "");
    if (headers.length === 0) headers = ["STT", "Họ và tên", "Số điện thoại", "Sự kiện", "Thời gian"];

    let headersChanged = false;
    Object.keys(customData).forEach(k => {
      const trimKey = safeCellValue(k.trim());
      if (trimKey && !headers.includes(trimKey)) {
        headers.push(trimKey);
        headersChanged = true;
      }
    });

    if (headersChanged) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    // 3. Cấp số thứ tự tịnh tiến & Ép ghi vật lý xuống đĩa cứng (Flush)
    const lastRow = sheet.getLastRow();
    const formattedId = String(lastRow).padStart(3, '0');
    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");

    const newRow = headers.map(h => {
      if (h === "STT") return formattedId;
      if (h === "Họ và tên") return safeCellValue(name);
      if (h === "Số điện thoại") return safeCellValue(phone);
      if (h === "Sự kiện") return safeCellValue(eventCode);
      if (h === "Thời gian") return timestamp;
      const originalKey = Object.keys(customData).find(k => safeCellValue(k.trim()) === h);
      return originalKey !== undefined ? safeCellValue(String(customData[originalKey])) : "";
    });

    sheet.getRange(lastRow + 1, 1, 1, headers.length).setValues([newRow]);
    SpreadsheetApp.flush(); // Ép lưu ngay lập tức chống trùng số

    return createJsonResponse({
      status: "success",
      id: formattedId,
      name: name
    });

  } finally {
    lock.releaseLock();
  }
}

// Chuẩn hóa SĐT để so khớp: chỉ giữ chữ số, bỏ đầu 84 (VN) hoặc 0 ở đầu.
// VD: "+84 912345678", "0912345678", "84912345678" đều về "912345678".
function normalizePhone(raw) {
  let digits = (raw || "").toString().replace(/\D/g, "");
  if (digits.startsWith("84") && digits.length >= 11) digits = digits.slice(2);
  else if (digits.startsWith("0") && digits.length > 1) digits = digits.slice(1);
  return digits;
}

// Chặn Sheets hiểu nhầm chuỗi thành công thức (VD: "=1+1", "+cmd") khi ghi bằng setValues.
function safeCellValue(v) {
  const s = (v === undefined || v === null) ? "" : v.toString();
  if (/^[=+\-@\t\r]/.test(s)) return "'" + s;
  return s;
}

function handleUpdateSettings(data) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    return createJsonResponse({ status: "error", message: "Hệ thống đang bận, vui lòng thử lại!" });
  }
  try {
  let cfgSheet = SPREADSHEET.getSheetByName(CONFIG_SHEET_NAME);
  if (!cfgSheet) {
    cfgSheet = SPREADSHEET.insertSheet(CONFIG_SHEET_NAME);
    cfgSheet.getRange("A1:B1").setValues([["Key", "Value"]]);
  }

  const payloadMap = [
    ["eventName", data.eventName || ""],
    ["eventCode", data.eventCode || "EVENT_01"],
    ["language", data.language || "vi"],
    ["primaryColor", data.primaryColor || "#2563EB"],
    ["fontFamily", data.fontFamily || "Inter"],
    ["logoUrl", data.logoUrl || ""],
    ["bgType", data.bgType || "solid"],
    ["bgValue", data.bgValue || "#0f172a"]
  ];

  cfgSheet.getRange(2, 1, payloadMap.length, 2).setValues(payloadMap);
  SpreadsheetApp.flush();
  return createJsonResponse({ status: "success", message: "Lưu cấu hình thành công!" });
  } finally {
    lock.releaseLock();
  }
}

function handleResetData() {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    return createJsonResponse({ status: "error", message: "Hệ thống đang bận, vui lòng thử lại!" });
  }
  try {
    const sheet = SPREADSHEET.getSheetByName(DATA_SHEET_NAME);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
      SpreadsheetApp.flush();
    }
    return createJsonResponse({ status: "success", message: "Đã reset bảng số về 001!" });
  } finally {
    lock.releaseLock();
  }
}

function handleResetTestData() {
  const sheet = SPREADSHEET.getSheetByName(TEST_DATA_SHEET_NAME);
  if (sheet) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
    SpreadsheetApp.flush();
  }
  return createJsonResponse({ status: "success", message: "Đã xóa dữ liệu test!" });
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
