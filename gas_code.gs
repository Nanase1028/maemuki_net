/**
 * ネット前向き研究2026 データ受信用 Google Apps Script
 * ------------------------------------------------------
 * 使い方:
 * 1. 新規Googleスプレッドシートを作成し、URLの /d/ と /edit の間のIDをコピー
 * 2. 下の SPREADSHEET_ID に貼り付け
 * 3. setupSheets() を一度実行（シートとヘッダーを自動作成）
 * 4. デプロイ > 新しいデプロイ > ウェブアプリ
 *    - 次のユーザーとして実行: 自分
 *    - アクセスできるユーザー: 全員
 * 5. 発行されたURLを index.html の CONFIG.GAS_URL に貼り付け
 *
 * ※修正時は「デプロイを管理」から既存デプロイを更新すること
 *   （新規デプロイを作るとURLが変わります）
 */

const SPREADSHEET_ID = "1MWb5VJm0Cpkdnyi9FiIs6mSJhd4mnJeq8V_zxkHWhFE";

const RAW_SHEET = "回答RAW";
const SCORE_SHEET = "得点サマリ";

const META_KEYS = ["timestamp_server", "submitted_at", "participant_id", "age", "sex", "form_type"];

// 得点サマリシートに載せる得点キー（index.html の computeScores と対応）
const SCORE_KEYS = [
  "cius_total", "dq_total", "iat_total", "sas_total",
  "igdt_total", "igdt_rawsum", "igds_total", "games_total",
  "socrates_recognition", "socrates_ambivalence", "socrates_takingsteps",
  "hyperfocus_total",
  "bis11_total", "bis11_attention", "bis11_motor", "bis11_nonplanning",
  "bisbas_bis", "bisbas_bas_total", "bisbas_bas_drive", "bisbas_bas_fun", "bisbas_bas_reward",
  "audit_total", "ftnd_total", "sogs_total", "sogs_experience",
  "iri_pt", "iri_fs", "iri_ec", "iri_pd",
  "tas_total", "tas_dif", "tas_ddf", "tas_eot",
  "cfs_total", "rei_rational", "rei_experiential",
  "qol_physical", "qol_psych", "qol_social", "qol_environ", "qol_overall", "qol_mean26"
];

/* ============ 受信 ============ */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const meta = {
      timestamp_server: new Date(),
      submitted_at: payload.submitted_at || "",
      participant_id: payload.participant_id || "",
      age: payload.age || "",
      sex: payload.sex || "",
      form_type: payload.form_type || ""
    };

    // 回答RAW: メタ + 全回答 + 全得点
    const rawFlat = Object.assign({}, meta, payload.answers || {}, payload.scores || {});
    appendByHeader(getOrCreateSheet(ss, RAW_SHEET), rawFlat);

    // 得点サマリ: メタ + 得点のみ
    const scoreFlat = Object.assign({}, meta);
    SCORE_KEYS.forEach(function (k) {
      scoreFlat[k] = (payload.scores && payload.scores[k] !== undefined) ? payload.scores[k] : "";
    });
    appendByHeader(getOrCreateSheet(ss, SCORE_SHEET), scoreFlat);

    return jsonOut({ status: "ok" });
  } catch (err) {
    return jsonOut({ status: "error", message: String(err) });
  }
}

/* ヘッダー行に沿って1行追記。未知のキーは列を自動追加 */
function appendByHeader(sheet, flatObj) {
  const keys = Object.keys(flatObj);
  let lastCol = sheet.getLastColumn();
  let headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];

  // ヘッダーが空なら作成
  if (headers.length === 0 || headers.every(function (h) { return h === ""; })) {
    headers = keys.slice();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else {
    // 足りない列を追加
    const missing = keys.filter(function (k) { return headers.indexOf(k) === -1; });
    if (missing.length > 0) {
      sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
      headers = headers.concat(missing);
    }
  }

  const row = headers.map(function (h) {
    return flatObj[h] !== undefined ? flatObj[h] : "";
  });
  sheet.appendRow(row);
}

function getOrCreateSheet(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============ 初期セットアップ ============ */
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  getOrCreateSheet(ss, RAW_SHEET);
  const sc = getOrCreateSheet(ss, SCORE_SHEET);
  // 得点サマリのヘッダーを先に作っておく
  if (sc.getLastRow() === 0) {
    const headers = META_KEYS.concat(SCORE_KEYS);
    sc.getRange(1, 1, 1, headers.length).setValues([headers]);
    sc.setFrozenRows(1);
  }
  Logger.log("セットアップ完了: " + ss.getName());
}

/* ============ 動作テスト ============ */
function testPost() {
  const sample = {
    form_type: "net_addiction_battery_2026",
    submitted_at: new Date().toISOString(),
    participant_id: "TEST001",
    age: "30",
    sex: "その他・回答しない",
    answers: { cius_1: "2", cius_2: "3", dq_1: "1", iat_1: "4" },
    scores: { cius_total: 30, dq_total: 5, iat_total: 55 }
  };
  const e = { postData: { contents: JSON.stringify(sample) } };
  const result = doPost(e);
  Logger.log(result.getContent());
}
