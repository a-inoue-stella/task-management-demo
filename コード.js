/**
 * 【設定エリア】
 */
const CONFIG = {
  SHEET_TASK: 'タスク管理',
  SHEET_SETTING: '設定',
  SHEET_LOG: 'ログ',
  // 列番号
  COL_TASK_NAME: 2,
  COL_ASSIGNEE: 3,
  COL_DEADLINE: 5,
  COL_STATUS: 6,
  COL_TRIGGER: 7,
  // 設定シート位置
  CELL_WEBHOOK: 'C2',
  RANGE_USER_MAP: 'A2:B20'
};

/**
 * 0. メニューバーの作成 (onOpen)
 * シートを開いた時に自動実行され、メニューバーにカスタムメニューを追加します。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡️ タスク管理デモ') // メニュー名
    .addItem('🔔 リマインドを実行', 'sendReminders') // 項目名, 実行する関数名
    .addToUi();
}

/* --- 1. トリガー制御 --- */
function handleEdit(e) { // 関数名は手動トリガー設定に合わせてください
  const range = e.range;
  const sheet = range.getSheet();

  if (sheet.getName() !== CONFIG.SHEET_TASK) return;
  if (range.getColumn() !== CONFIG.COL_TRIGGER) return;
  if (e.value !== "TRUE") return;

  processNotification(sheet, range.getRow());
}

/* --- 2. 通知処理実行 --- */
function processNotification(sheet, rowIndex) {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(10000)) {
    try {
      const data = sheet.getRange(rowIndex, 1, 1, 10).getValues()[0];
      const taskName = data[CONFIG.COL_TASK_NAME - 1];
      const assignee = data[CONFIG.COL_ASSIGNEE - 1];
      const deadline = data[CONFIG.COL_DEADLINE - 1]; // 日付オブジェクト
      const status   = data[CONFIG.COL_STATUS - 1];
      
      // カードペイロードの生成
      const payload = createCardPayload(taskName, assignee, deadline, status);
      
      const webhookUrl = getWebhookUrl();
      if(webhookUrl) {
        sendCard(webhookUrl, payload);
        writeLog(taskName, status, assignee, "送信成功");
      } else {
        Browser.msgBox("エラー：Webhook URL未設定");
      }

      sheet.getRange(rowIndex, CONFIG.COL_TRIGGER).setValue(false);

    } catch (e) {
      console.error(e);
      sheet.getRange(rowIndex, CONFIG.COL_TRIGGER).setValue(false);
    } finally {
      lock.releaseLock();
    }
  }
}

/**
 * ★修正版：大きなアイコン付きのカードを作る関数
 */
function createCardPayload(taskName, assigneeName, deadlineObj, status) {
  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const deadlineStr = deadlineObj ? Utilities.formatDate(deadlineObj, 'JST', 'yyyy/MM/dd') : '未設定';

  // デフォルト設定（通常通知：ベル）
  let headerTitle = "【通知】タスク更新";
  let headerSubtitle = "タスク管理Botより";
  let headerIcon = "https://www.gstatic.com/images/icons/material/system/2x/notifications_black_48dp.png";
  let headerStyle = "SQUARE"; 

  // ステータスに応じたデザイン切り替え
  if (status === "🟡 確認待ち") {
    headerTitle = "🟡 【確認依頼】承認をお願いします";
    // 人型アイコン
    headerIcon = "https://www.gstatic.com/images/icons/material/system/2x/account_circle_black_48dp.png";
  } else if (status === "🟢 完了") {
    headerTitle = "🟢 【完了】タスクが完了しました";
    // チェックマーク
    headerIcon = "https://www.gstatic.com/images/icons/material/system/2x/check_circle_black_48dp.png";
  }

  const card = {
    "cardsV2": [
      {
        "cardId": "unique-card-id",
        "card": {
          "header": {
            "title": headerTitle,
            "subtitle": headerSubtitle,
            "imageUrl": headerIcon,
            "imageType": headerStyle
          },
          "sections": [
            {
              "widgets": [
                {
                  "decoratedText": {
                    "startIcon": { "knownIcon": "DESCRIPTION" },
                    "topLabel": "タスク",
                    "text": `<b>${taskName}</b>`,
                    "wrapText": true
                  }
                },
                {
                  "decoratedText": {
                    "startIcon": { "knownIcon": "PERSON" },
                    "topLabel": "担当",
                    "text": `<b>${assigneeName}</b>`
                  }
                },
                {
                  "decoratedText": {
                    "startIcon": { "knownIcon": "BOOKMARK" },
                    "topLabel": "ステータス",
                    "text": `<b>${status}</b>`
                  }
                },
                {
                  "decoratedText": {
                    "startIcon": { "knownIcon": "CLOCK" },
                    "topLabel": "期限日",
                    "text": `<b>${deadlineStr}</b>`
                  }
                }
              ]
            },
            {
              "widgets": [
                {
                  "buttonList": {
                    "buttons": [
                      {
                        "text": "シートを開く",
                        "onClick": {
                          "openLink": {
                            "url": sheetUrl
                          }
                        }
                      }
                    ]
                  }
                }
              ]
            }
          ]
        }
      }
    ]
  };

  return card;
}

/**
 * ★変更点：リマインドもカードで送る
 */
/**
 * 4. リマインド実行 (修正版：期限切れ・今日・明日を区別)
 */
function sendReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_TASK);
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 2) {
    Browser.msgBox("データがありません");
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  
  const today = new Date();
  today.setHours(0,0,0,0);
  
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);
  
  let alertCount = 0;
  const webhookUrl = getWebhookUrl();

  if (!webhookUrl) {
    Browser.msgBox("Webhook URL未設定");
    return;
  }

  data.forEach(row => {
    const taskName = row[CONFIG.COL_TASK_NAME - 1];
    const deadlineStr = row[CONFIG.COL_DEADLINE - 1];
    const status   = row[CONFIG.COL_STATUS - 1];
    const assignee = row[CONFIG.COL_ASSIGNEE - 1];

    if (status === "🟢 完了" || !taskName || !deadlineStr) return;

    const deadline = new Date(deadlineStr);
    deadline.setHours(0,0,0,0);

    let title = "";
    let iconUrl = "";
    let isTarget = false;

    if (deadline.getTime() < today.getTime()) {
      // ① 期限切れ（ビックリマーク）※ご提示いただいたURL
      title = "🔥 【遅延】期限が過ぎています！";
      iconUrl = "https://www.gstatic.com/images/icons/material/system/2x/warning_amber_black_48dp.png";
      isTarget = true;
    } else if (deadline.getTime() === today.getTime()) {
      // ② 今日が期限（時計）※ご提示いただいたURL
      title = "⏰ 【今日】本日が対応期限です";
      iconUrl = "https://www.gstatic.com/images/icons/material/system/2x/alarm_black_48dp.png";
      isTarget = true;
    } else if (deadline.getTime() === tomorrow.getTime()) {
      // ③ 明日が期限（カレンダー）
      title = "⚠️ 【明日】明日が期限です";
      iconUrl = "https://www.gstatic.com/images/icons/material/system/2x/event_black_48dp.png";
      isTarget = true;
    }

    if (isTarget) {
       let payload = createCardPayload(taskName, assignee, deadline, status);
       
       // ヘッダーをアラート用に上書き
       payload.cardsV2[0].card.header.title = title;
       payload.cardsV2[0].card.header.imageUrl = iconUrl;
       payload.cardsV2[0].card.header.imageType = "SQUARE"; // ここもSQUAREにします
       
       sendCard(webhookUrl, payload);
       alertCount++;
       Utilities.sleep(500); 
    }
  });

  if(alertCount > 0) {
    Browser.msgBox(`送信完了：${alertCount}件のリマインドを送信しました`);
  } else {
    Browser.msgBox("リマインド対象はありません");
  }
}

/* --- ユーティリティ --- */

// カード送信関数（JSONをそのまま送る）
function sendCard(url, payload) {
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };
  UrlFetchApp.fetch(url, options);
}

function getWebhookUrl() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SETTING);
  return sheet.getRange(CONFIG.CELL_WEBHOOK).getValue();
}

function getUserMap() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SETTING);
  const data = sheet.getRange(CONFIG.RANGE_USER_MAP).getValues();
  let map = {};
  data.forEach(row => { if(row[0] && row[1]) map[row[0]] = row[1]; });
  return map;
}

function writeLog(task, status, user, result) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_LOG);
  const date = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss');
  sheet.appendRow([date, task, status, user, result]);
}