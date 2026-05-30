/**
 * 家計簿アプリ GAS実装テンプレート
 * - スプレッドシート2シート構成: users / transactions
 * - action: register / getUsers / upload / getList
 *
 * 使い方:
 * 1) このコードをApps Scriptに貼り付け
 * 2) SPREADSHEET_ID を対象スプレッドシートIDに変更
 * 3) 画像保存先フォルダIDは「スクリプトのプロパティ」にのみ設定する（Gitに載せない）
 *    - エディタ左「プロジェクトの設定」（歯車）→「スクリプトのプロパティ」→「スクリプトのプロパティを表示」
 *    - プロパティ: IMAGE_DRIVE_FOLDER_ID / 値: GoogleドライブのフォルダURL末尾のID
 *    未設定の場合は、マイドライブ直下の IMAGE_FOLDER_NAME（既定「画像」）を検索し、無ければ作成して保存する
 * 4) Webアプリとしてデプロイ（実行ユーザー: 自分、アクセス: 全員）
 */

const SPREADSHEET_ID = "YOUR_SPREADSHEET_ID";
/** スクリプトのプロパティ名（値にドライブの画像フォルダIDを設定） */
const SCRIPT_PROP_IMAGE_FOLDER_ID = "IMAGE_DRIVE_FOLDER_ID";
/** スクリプトプロパティ未設定時に使うフォルダ名（マイドライブ直下） */
const IMAGE_FOLDER_NAME = "画像";
const USERS_SHEET = "users";
const TX_SHEET = "transactions";
const TZ = "Asia/Tokyo";

const USERS_HEADERS = ["lineId", "name", "createdAt", "updatedAt"];
const TX_HEADERS = [
  "id",
  "lineId",
  "userName",
  "type",
  "amount",
  "category",
  "shop",
  "date",
  "imageUrlsJson",
  "paymentMethod",
  "createdAt",
  "updatedAt",
];

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const action = body.action;
    const data = body.data || {};

    ensureSheets_();

    switch (action) {
      case "register":
        return jsonOutput_(registerUser_(data));
      case "getUsers":
        return jsonOutput_({ status: "success", result: getUsers_() });
      case "upload":
        return jsonOutput_(uploadTransaction_(data));
      case "getList":
        return jsonOutput_({ status: "success", result: getTransactions_() });
      default:
        return jsonOutput_({ status: "error", message: "Unknown action" });
    }
  } catch (err) {
    return jsonOutput_({
      status: "error",
      message: err && err.message ? err.message : "Unexpected error",
    });
  }
}

function registerUser_(data) {
  const lineId = safeTrim_(data.userId);
  const name = safeTrim_(data.displayName);
  if (!lineId || !name) return { status: "error", message: "userId/displayName is required" };

  const sh = getSheet_(USERS_SHEET, USERS_HEADERS);
  const values = readDataRows_(sh, USERS_HEADERS.length);
  const now = isoNow_();

  const index = values.findIndex((r) => r[0] === lineId);
  if (index >= 0) {
    const row = index + 2;
    sh.getRange(row, 2).setValue(name);
    sh.getRange(row, 4).setValue(now);
  } else {
    sh.appendRow([lineId, name, now, now]);
  }

  return { status: "success" };
}

function getUsers_() {
  const sh = getSheet_(USERS_SHEET, USERS_HEADERS);
  const rows = readDataRows_(sh, USERS_HEADERS.length);
  return rows
    .filter((r) => r[0])
    .map((r) => ({
      lineId: r[0],
      name: r[1] || "",
    }));
}

function uploadTransaction_(data) {
  const lineId = safeTrim_(data.lineId);
  const userName = safeTrim_(data.userName);
  const type = safeTrim_(data.type);
  const amount = Number(data.amount);
  const category = safeTrim_(data.category);
  const shop = safeTrim_(data.shop) || "未設定";
  const date = safeTrim_(data.date);
  const imageBlobs = Array.isArray(data.imageBlobs) ? data.imageBlobs : [];
  const paymentMethod = safeTrim_(data.paymentMethod);

  if (!lineId || !userName || !type || !category || !date) {
    return { status: "error", message: "Missing required fields" };
  }
  if (type !== "入金" && type !== "出金") return { status: "error", message: "type must be 入金 or 出金" };
  if (!Number.isFinite(amount) || amount <= 0) return { status: "error", message: "amount must be > 0" };
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { status: "error", message: "date format must be YYYY-MM-DD" };
  if (imageBlobs.length > 3) return { status: "error", message: "image max is 3" };

  const imageUrls = saveImages_(imageBlobs, lineId);
  const sh = getSheet_(TX_SHEET, TX_HEADERS);
  const now = isoNow_();
  const id = "txn_" + Utilities.formatDate(new Date(), TZ, "yyyyMMdd_HHmmss_SSS");

  sh.appendRow([
    id,
    lineId,
    userName,
    type,
    amount,
    category,
    shop,
    date,
    JSON.stringify(imageUrls),
    paymentMethod,
    now,
    now,
  ]);

  return { status: "success", id: id };
}

function getTransactions_() {
  const sh = getSheet_(TX_SHEET, TX_HEADERS);
  const rows = readDataRows_(sh, TX_HEADERS.length);

  return rows
    .filter((r) => r[0])
    .map((r) => {
      let imageUrls = [];
      try {
        imageUrls = r[8] ? JSON.parse(r[8]) : [];
      } catch (_) {
        imageUrls = [];
      }
      return {
        lineId: r[1] || "",
        userName: r[2] || "",
        type: r[3] || "",
        amount: Number(r[4]) || 0,
        category: r[5] || "",
        shop: r[6] || "",
        date: r[7] || "",
        imageUrls: Array.isArray(imageUrls) ? imageUrls : [],
        paymentMethod: r[9] || "",
      };
    });
}

/**
 * スクリプトのプロパティ IMAGE_DRIVE_FOLDER_ID に設定したフォルダIDを返す（未設定は空文字）。
 */
function getImageDriveFolderIdFromProps_() {
  const raw = PropertiesService.getScriptProperties().getProperty(SCRIPT_PROP_IMAGE_FOLDER_ID);
  return safeTrim_(raw).replace(/^["']|["']$/g, "");
}

/**
 * レシート画像の保存先フォルダを返す。
 * - スクリプトプロパティ IMAGE_DRIVE_FOLDER_ID があればそのフォルダ
 * - 未設定ならマイドライブ直下の IMAGE_FOLDER_NAME を検索、無ければ作成
 */
function getImageParentFolder_() {
  const id = getImageDriveFolderIdFromProps_();
  if (id) {
    return DriveApp.getFolderById(id);
  }
  const name = safeTrim_(IMAGE_FOLDER_NAME) || "画像";
  const root = DriveApp.getRootFolder();
  const it = root.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return root.createFolder(name);
}

function saveImages_(imageBlobs, lineId) {
  if (!imageBlobs.length) return [];

  const parent = getImageParentFolder_();
  const root = DriveApp.getRootFolder();
  const urls = [];
  imageBlobs.forEach((dataUrl, i) => {
    const parsed = parseDataUrl_(dataUrl);
    if (!parsed) return;

    const ext = mimeToExt_(parsed.mimeType);
    const fileName = `${lineId}_${Utilities.formatDate(new Date(), TZ, "yyyyMMdd_HHmmss")}_${i + 1}.${ext}`;
    const blob = Utilities.newBlob(parsed.bytes, parsed.mimeType, fileName);
    // マイドライブ直下に一度作成し、目的フォルダへ移す（環境によって createFile の置き場所がずれる対策）
    const file = DriveApp.createFile(blob);
    parent.addFile(file);
    try {
      root.removeFile(file);
    } catch (e) {
      // 既にルート外のみにある場合などは無視
    }
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    urls.push(file.getUrl());
  });
  return urls;
}

function ensureSheets_() {
  getSheet_(USERS_SHEET, USERS_HEADERS);
  getSheet_(TX_SHEET, TX_HEADERS);
}

function getSheet_(name, headers) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const currentHeaders = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    const mismatch = headers.some((h, i) => currentHeaders[i] !== h);
    if (mismatch) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  return sh;
}

function readDataRows_(sh, width) {
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return [];
  return sh.getRange(2, 1, lastRow - 1, width).getValues();
}

function parseDataUrl_(dataUrl) {
  const s = dataUrl || "";
  const marker = ";base64,";
  const i = s.indexOf(marker);
  if (i < 0 || !/^data:/i.test(s)) return null;
  const header = s.substring(5, i); // after "data:"
  const mimeType = header.split(";")[0].trim();
  const b64 = s.substring(i + marker.length).replace(/\s/g, "");
  return {
    mimeType: mimeType || "image/jpeg",
    bytes: Utilities.base64Decode(b64),
  };
}

function mimeToExt_(mime) {
  if (mime === "image/png") return "png";
  if (mime === "image/webp") return "webp";
  return "jpg";
}

function isoNow_() {
  return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function safeTrim_(v) {
  return String(v == null ? "" : v).trim();
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
