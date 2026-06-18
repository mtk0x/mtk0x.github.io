/**
 * パターン認識クイズ — 回答受信用 Apps Script Web App
 *
 * 役割: クイズページから送られてくる「生の回答データ」を、
 *       既存のフォーム回答シート（タブ）の末尾に1行追記する。
 *       採点（スコア・点数）はシート側の既存の関数に任せる。
 *       → B（スコア）・D（点数）・K（削除06）の列は一切書き込まない。
 *         （関数を上書きして壊さないため。appendRow は使わず、担当列だけ個別に書く）
 *
 * 書き込む列:
 *   A(1)  タイムスタンプ        ← 送信時刻
 *   C(3)  氏名を記入してください ← 氏名
 *   E(5)  01(5)                ← 第1問で選んだ番号
 *   F(6)  02(2)                ← 第2問
 *   G(7)  03(10)               ← 第3問
 *   H(8)  04(6)                ← 第4問
 *   I(9)  05(9)                ← 第5問
 *   J(10) 06(12)               ← 第6問
 *   L(12) 職種                  ← 職種（この列に「職種」見出しを用意しておく）
 *   ※ B(2)スコア / D(4)点数 / K(11)削除06 は触らない
 *
 * 設置方法: 対象スプレッドシートを開く → 拡張機能 → Apps Script →
 *           このコードを貼り付け → デプロイ → ウェブアプリ
 *           （実行ユーザー=自分 / アクセスできるユーザー=全員）。
 *           発行された URL を quiz/index.html の ENDPOINT に設定する。
 */

var RESPONSE_SHEET_GID = 1642761423;  // 追記先タブ（フォーム回答シート）の gid

function doPost(e) {
  // 同時アクセスでの追記競合を防ぐ（最大30秒待機）
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var data = JSON.parse(e.postData.contents);

    var sheet = getSheetByGid(RESPONSE_SHEET_GID);
    if (!sheet) throw new Error('gid ' + RESPONSE_SHEET_GID + ' のシートが見つかりません');

    // 回答を id でひけるように
    var ans = {};
    (data.answers || []).forEach(function (a) { ans[a.id] = a; });
    function sel(id) { return (ans[id] && ans[id].selected != null) ? ans[id].selected : ''; }

    // 既存データの最終行は「A列（タイムスタンプ）」基準で判定する。
    // （B/D が ARRAYFORMULA だと getLastRow() が膨らむため、A列の中身で数える）
    var maxRows = sheet.getMaxRows();
    var colA = sheet.getRange(1, 1, maxRows, 1).getValues();
    var lastDataRow = 0;
    for (var i = 0; i < colA.length; i++) {
      if (colA[i][0] !== '' && colA[i][0] !== null) lastDataRow = i + 1;
    }
    var row = lastDataRow + 1;

    // 担当列だけ個別に書き込む（B/D/K は触らない）
    sheet.getRange(row, 1).setValue(data.submittedAt ? new Date(data.submittedAt) : new Date()); // A タイムスタンプ
    sheet.getRange(row, 3).setValue(data.name || '');                                            // C 氏名
    sheet.getRange(row, 5, 1, 6).setValues([[                                                    // E〜J 第1〜6問
      sel('q1'), sel('q2'), sel('q3'), sel('q4'), sel('q5'), sel('q6')
    ]]);
    sheet.getRange(row, 12).setValue(data.jobType || '');                                        // L 職種

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, row: row }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function getSheetByGid(gid) {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === gid) return sheets[i];
  }
  return null;
}

// 動作確認用（ブラウザでWebアプリURLを開いたとき）
function doGet() {
  return ContentService.createTextOutput('OK: クイズ回答受信エンドポイントです。');
}
