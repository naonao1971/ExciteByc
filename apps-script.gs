/**
 * ExciteByc ランキング用 Google Apps Script (Web App)
 *
 * セットアップ手順:
 * 1. Google スプレッドシートを新規作成する
 * 2. シート名を "ranking" にする（1行目はヘッダー行として自動生成されます）
 * 3. 拡張機能 → Apps Script を開き、このファイルの内容を貼り付けて保存
 * 4. 「デプロイ」→「新しいデプロイ」→種類「ウェブアプリ」を選択
 *    - 実行するユーザー: 自分
 *    - アクセスできるユーザー: 全員
 * 5. デプロイ後に発行される URL（https://script.google.com/macros/s/.../exec）を
 *    index.html の GAS_URL 定数にセットする
 *
 * データ形式（1レコード = 1行）:
 *   nickname | xid | score | created
 */

const SHEET_NAME = "ranking";
const MAX_RECORDS_RETURNED = 100; // 取得件数の上限（フロント側でTOP10に絞り込む）

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["nickname", "xid", "score", "created"]);
  }
  return sheet;
}

function doGet(e) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const rows = values.slice(1); // ヘッダー行を除く

  const records = rows
    .filter(function (r) {
      return r[0] !== "" && r[0] != null;
    })
    .map(function (r) {
      return {
        nickname: String(r[0] || ""),
        xid: String(r[1] || ""),
        score: Number(r[2] || 0),
        created: r[3] ? new Date(r[3]).getTime() : null
      };
    })
    .sort(function (a, b) {
      return b.score - a.score;
    })
    .slice(0, MAX_RECORDS_RETURNED);

  return ContentService
    .createTextOutput(JSON.stringify({ records: records }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.action === "add") {
      const rec = body.record || {};
      const nickname = String(rec.nickname || "").trim().slice(0, 20);
      const xid = String(rec.xid || "").replace(/^@/, "").trim().slice(0, 20);
      const score = Number(rec.score || 0);
      const created = rec.created || Date.now();

      if (!nickname) {
        return jsonOut_({ success: false, error: "nickname is required" });
      }

      const sheet = getSheet_();
      sheet.appendRow([nickname, xid, score, new Date(created)]);
      return jsonOut_({ success: true });
    }

    return jsonOut_({ success: false, error: "unknown action" });
  } catch (err) {
    return jsonOut_({ success: false, error: String(err) });
  }
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
