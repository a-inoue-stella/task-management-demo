// =================================================================
// 株式会社クオーレ様向け タスク管理デモ v1.0
// 目的: スプレッドシート上のボタンからタスクを管理する
// 作成日: 2025/11/18
// ドキュメント: task_manager_design_doc_outline.md
// =================================================================

// --- グローバル設定 ---
// TODO: 1.1で取得した貴社（ステラリープ社）のテスト用Webhook URLを以下に設定してください
const WEBHOOK_URL = "https://chat.googleapis.com/v1/spaces/XXXXX/messages?key=XXXXX&token=XXXXX";
const SHEET_NAME_TASKS = "タスク管理";
const SHEET_NAME_ARCHIVE = "完了タスク";
const SHEET_NAME_MASTER = "マスタ";

// 通知を何日前に送るか (0 = 当日, 1 = 1日前)
const DAYS_BEFORE_REMIND = 1; 

// --- 1. メイン機能（ボタン割り当て用） ---

/**
 * [ボタンA: リマインド通知]
 * 期限切れ・期限直前のタスクを検知し、Googleチャットに通知カードを送信します。
 * (要件 FR-002 準拠)
 */
function checkDeadlines() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME_TASKS);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("エラー: 'タスク管理'シートが見つかりません。");
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert("タスクがありません。");
    return;
  }

  // A列(タスク名), B列(担当者), C列(優先度), D列(ステータス), E列(期限) を取得
  const dataValues = sheet.getRange(2, 1, lastRow - 1, 5).getValues(); 
  const today = new Date();
  today.setHours(0, 0, 0, 0); // 時刻をリセットして日付のみで比較

  let notificationCount = 0;

  for (let i = 0; i < dataValues.length; i++) {
    const rowData = dataValues[i];
    const taskName = rowData[0];
    const assignee = rowData[1];
    const priority = rowData[2];
    const status = rowData[3];
    const dueDateValue = rowData[4];

    // 要件: 未完了かつ期限が設定されている場合のみチェック
    if (status !== "完了" && dueDateValue instanceof Date) {
      const dueDate = new Date(dueDateValue);
      dueDate.setHours(0, 0, 0, 0);
      
      const diffTime = dueDate.getTime() - today.getTime();
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

      let title = "";
      let icon = "";

      if (diffDays < 0) {
        // 期限切れ
        title = "🚨【警告：期限切れ！】";
        icon = "WARNING";
      } else if (diffDays <= DAYS_BEFORE_REMIND) {
        // 期限直前 (当日含む)
        title = "⏰【リマインド：対応期限です】";
        icon = "CLOCK";
      }

      // 通知対象ならカードを送信
      if (title !== "") {
        // 該当行へのリンクを生成 (FR-002-03)
        const rowLink = ss.getUrl() + "#gid=" + sheet.getSheetId() + "&range=A" + (i + 2);
        
        const payload = createChatCard(title, taskName, assignee, priority, rowLink, icon);
        sendToChat(payload);
        notificationCount++;
        Utilities.sleep(500); // 連続送信によるAPI制限を回避
      }
    }
  }

  // 実行結果をポップアップ (FR-003-03 の思想を流用)
  if (notificationCount > 0) {
    SpreadsheetApp.getUi().alert(notificationCount + "件のリマインドを送信しました。");
  } else {
    SpreadsheetApp.getUi().alert("リマインド対象のタスクはありませんでした。");
  }
}

/**
 * [ボタンB: 完了タスク整理]
 * 「完了」ステータスの行を一括でアーカイブシートへ移動します。
 * (要件 FR-003 準拠)
 */
function archiveCompletedTasks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SHEET_NAME_TASKS);
  let targetSheet = ss.getSheetByName(SHEET_NAME_ARCHIVE);

  // アーカイブシートがなければ作成 (ヘッダーコピー)
  if (!targetSheet) {
    targetSheet = ss.insertSheet(ARCHIVE_SHEET_NAME);
    sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).copyTo(targetSheet.getRange(1, 1));
  }

  const lastRow = sourceSheet.getLastRow();
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert("タスクがありません。");
    return;
  }

  const range = sourceSheet.getRange(2, 1, lastRow - 1, 5); // A-E列
  const values = range.getValues();
  
  const rowsToArchive = [];
  const rowsToDelete = []; // 削除する行番号(インデックスではない)

  // ループは下から順に行う (行削除時のインデックスずれを防ぐため)
  for (let i = values.length - 1; i >= 0; i--) {
    const statusColIndex = 3; // D列 (0始まり)
    if (values[i][statusColIndex] === "完了") {
      rowsToArchive.unshift(values[i]); // アーカイブ配列に追加
      rowsToDelete.push(i + 2); // 行番号(1始まり + ヘッダー行)を追加
    }
  }

  if (rowsToArchive.length === 0) {
    SpreadsheetApp.getUi().alert("完了済みのタスクはありませんでした。");
    return;
  }

  // 1. アーカイブシートへ一括書き込み (FR-003-02)
  targetSheet.getRange(
    targetSheet.getLastRow() + 1,
    1,
    rowsToArchive.length,
    rowsToArchive[0].length
  ).setValues(rowsToArchive);

  // 2. 元シートから行を削除 (下から順に削除するためインデックスずれなし)
  rowsToDelete.forEach(function(rowIndex) {
    sourceSheet.deleteRow(rowIndex);
  });

  // 3. 完了メッセージ (FR-003-03)
  SpreadsheetApp.getUi().alert(rowsToArchive.length + "件のタスクをアーカイブしました。\nお疲れ様でした！");
}

// --- 2. ヘルパー関数 ---

/**
 * Google Chat カード (v2) のJSONペイロードを生成します。
 * (設計書 2.5 準拠)
 * @param {string} headerTitle - カードのヘッダータイトル
 * @param {string} taskName - タスク名
 * @param {string} assignee - 担当者
 * @param {string} priority - 優先度
 * @param {string} link - 該当行へのURL
 * @param {string} iconType - "WARNING" または "CLOCK"
 * @return {object} Google Chat Card v2 JSON object
 */
function createChatCard(headerTitle, taskName, assignee, priority, link, iconType) {
  return {
    "cardsV2": [{
      "cardId": "task-reminder-" + new Date().getTime(), // 簡易的なユニークID
      "card": {
        "header": {
          "title": headerTitle,
          "subtitle": "タスク管理Botより",
          "imageUrl": (iconType === "WARNING") 
            ? "https://www.gstatic.com/images/icons/material/system/2x/warning_amber_black_48dp.png" 
            : "https://www.gstatic.com/images/icons/material/system/2x/alarm_black_48dp.png",
          "imageType": "CIRCLE"
        },
        "sections": [{
          "widgets": [
            { "decoratedText": { "startIcon": { "knownIcon": "DESCRIPTION" }, "text": "<b>タスク:</b> " + (taskName || "(未設定)") } },
            { "decoratedText": { "startIcon": { "knownIcon": "PERSON" }, "text": "<b>担当:</b> " + (assignee || "(未設定)") } },
            { "decoratedText": { "startIcon": { "knownIcon": "TICKET" }, "text": "<b>優先度:</b> " + (priority || "(未設定)") } },
            { "buttonList": { "buttons": [{ "text": "シートを開く", "onClick": { "openLink": { "url": link } } }] } }
          ]
        }]
      }
    }]
  };
}

/**
 * Google Chat Webhookにペイロードを送信します。
 * @param {object} payload - Card v2 JSON object
 */
function sendToChat(payload) {
  const options = {
    "method": "POST",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };
  try {
    UrlFetchApp.fetch(WEBHOOK_URL, options);
  } catch (e) {
    Logger.log("Google Chatへの通知に失敗しました: " + e);
    // デモ中はアラートを出すと親切
    SpreadsheetApp.getUi().alert("Chat通知の送信に失敗しました。\nWebhook URLが正しいか確認してください。");
  }
}