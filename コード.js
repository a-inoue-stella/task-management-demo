/**
 * 【設定エリア】（再掲：ここがないと動きません）
 */
const CONFIG = {
  SHEET_TASK: 'タスク管理',
  SHEET_SETTING: '設定',
  SHEET_LOG: 'ログ',
  COL_TASK_NAME: 2,
  COL_ASSIGNEE: 3,
  COL_DEADLINE: 5,
  COL_STATUS: 6,
  COL_TRIGGER: 7,
  CELL_WEBHOOK: 'C2',
  RANGE_USER_MAP: 'A2:B20'
};

/**
 * 0. メニューバーの作成 (onOpen)
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡️ タスク管理デモ')
    .addItem('🔔 リマインドを実行', 'sendReminders')
    .addSeparator() // 区切り線
    .addItem('🤖 AIプラン取り込み', 'importAiPlan') // ★追加
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
      if (webhookUrl) {
        const res = sendCard(webhookUrl, payload, { task: taskName, status: status, user: assignee, context: 'processNotification:row' + rowIndex });
        if (res && res.success) {
          writeLog(taskName, status, assignee, "送信成功", 'processNotification:row' + rowIndex);
        } else {
          writeLog(taskName, status, assignee, "送信失敗: " + (res && res.error ? res.error : 'Unknown'), 'processNotification:row' + rowIndex);
        }
      } else {
        // 自動処理ではダイアログを表示しない。代わりにログを書く。
        writeLog(taskName, status, assignee, "送信失敗: Webhook URL未設定", 'processNotification:row' + rowIndex);
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
  const deadlineStr = deadlineObj ? formatDate(deadlineObj, 'yyyy/MM/dd') : '未設定';

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

      const res = sendCard(webhookUrl, payload, { task: taskName, status: status, user: assignee, context: 'sendReminders' });
      if (res && res.success) {
        writeLog(taskName, status, assignee, '送信成功', 'sendReminders');
      } else {
        writeLog(taskName, status, assignee, '送信失敗: ' + (res && res.error ? res.error : 'Unknown'), 'sendReminders');
      }

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


/**
 * 5. AIプラン取り込み機能（HTMLモーダル版・デバッグ強化）
 */
function importAiPlan() {
  console.log("【Client Debug】importAiPlan関数が起動しました"); // ログ1
  
  const htmlString = `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: sans-serif; padding: 15px; color: #333; }
          h3 { margin-top: 0; color: #202124; }
          textarea { 
            width: 100%; height: 300px; margin-bottom: 15px; 
            font-family: monospace; font-size: 12px; border: 1px solid #dadce0; 
            border-radius: 4px; padding: 8px; box-sizing: border-box;
          }
          button { 
            padding: 10px 24px; background-color: #1a73e8; color: white; 
            border: none; border-radius: 4px; cursor: pointer; font-weight: bold;
          }
          #status { margin-top: 15px; font-weight: bold; font-size: 13px; white-space: pre-wrap; }
          .error { color: #d93025; }
        </style>
      </head>
      <body>
        <h3>🤖 AIプラン取り込み（Debug）</h3>
        <p>Gemが出力したJSONコードを貼り付けてください。</p>
        <textarea id="jsonInput" placeholder='[ ... ]'></textarea>
        <br>
        <button onclick="submitJson()" id="submitBtn">取り込み実行</button>
        <div id="status"></div>

        <script>
          function submitJson() {
            const input = document.getElementById('jsonInput').value;
            const statusDiv = document.getElementById('status');
            const btn = document.getElementById('submitBtn');

            if (!input.trim()) {
              statusDiv.innerText = "⚠️ テキストを入力してください";
              return;
            }
            
            statusDiv.innerText = "🔄 GASへ送信中...";
            btn.disabled = true;
            btn.innerText = "処理中...";

            // サーバー側関数の呼び出し
            google.script.run
              .withSuccessHandler(onSuccess)
              .withFailureHandler(onFailure)
              .processPlanJson(input);
          }

          function onSuccess(resultMsg) {
            document.getElementById('status').innerText = resultMsg;
            document.getElementById('submitBtn').innerText = "完了";
            // 成功しても閉じずに結果を見せる
          }

          function onFailure(err) {
            const statusDiv = document.getElementById('status');
            statusDiv.className = "error";
            statusDiv.innerText = "❌ エラー:\\n" + err.message;
            document.getElementById('submitBtn').disabled = false;
            document.getElementById('submitBtn').innerText = "再試行";
          }
        </script>
      </body>
    </html>
  `;

  const htmlOutput = HtmlService.createHtmlOutput(htmlString)
    .setWidth(600)
    .setHeight(550);
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'AIプロジェクト・アーキテクト連携');
}

/**
 * バックエンド処理（連番ID対応版）
 */
function processPlanJson(input) {
  console.log("【Server Debug】processPlanJsonが呼び出されました");

  try {
    // 1. JSON抽出
    const firstBracket = input.indexOf("[");
    const lastBracket = input.lastIndexOf("]");

    if (firstBracket === -1 || lastBracket === -1 || firstBracket >= lastBracket) {
      throw new Error("JSON配列（[...]）が見つかりませんでした。");
    }

    const jsonString = input.substring(firstBracket, lastBracket + 1);
    
    // 2. パース
    let tasks;
    try {
      tasks = JSON.parse(jsonString);
    } catch (e) {
      throw new Error("JSON形式が不正です。\n" + e.message);
    }
    
    if (!Array.isArray(tasks) || tasks.length === 0) {
      throw new Error("タスクが含まれていません。");
    }

    // 3. ID採番の準備（既存IDの最大値を取得）
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_TASK);
    if (!sheet) throw new Error(`シート「${CONFIG.SHEET_TASK}」が見つかりません`);

    const existingIds = sheet.getRange("A:A").getValues().flat();
    let maxIdNum = 0;

    existingIds.forEach(id => {
      // "TASK-" で始まり、後ろが数字であるものを抽出
      if (typeof id === 'string' && id.startsWith('TASK-')) {
        const numPart = parseInt(id.replace('TASK-', ''), 10);
        if (!isNaN(numPart) && numPart > maxIdNum) {
          maxIdNum = numPart;
        }
      }
    });

    console.log(`【Server Debug】現在の最大ID番号: ${maxIdNum}`);

    // 4. データ生成（連番ID付与）
    const newRows = tasks.map((t, index) => {
      const start = t.start_date ? new Date(t.start_date) : new Date();
      const due   = t.due_date   ? new Date(t.due_date)   : new Date();
      
      // 連番生成: 最大値 + インデックス + 1
      // ('000' + num).slice(-3) で3桁埋め（001, 010, 100）
      const nextNum = maxIdNum + index + 1;
      const newId = 'TASK-' + ('000' + nextNum).slice(-3);

      return [
        newId,                        // A列: 連番ID (TASK-XXX)
        t.task_name || "名称未定",      // B列
        t.assignee_name || "",        // C列
        start,                        // D列
        due,                          // E列
        "⚪️ 未着手",                   // F列
        false,                        // G列
        t.description || ""           // H列
      ];
    });

    // 5. 書き込み位置の特定（A列基準）
    // チェックボックス(G列)に惑わされないよう、A列の最終行を探す
    const columnA = sheet.getRange("A:A").getValues();
    let lastRow = 0;

    for (let i = columnA.length - 1; i >= 0; i--) {
      if (columnA[i][0] !== "" && columnA[i][0] != null) {
        lastRow = i + 1;
        break;
      }
    }
    if (lastRow < 1) lastRow = 1; // ヘッダー行考慮

    console.log(`【Server Debug】書き込み開始行: ${lastRow + 1}`);
    
    // 6. 書き込み実行
    sheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length).setValues(newRows);

    return `✅ 成功！\n${newRows.length}件のタスクを追加しました。\n(ID: TASK-${('000' + (maxIdNum + 1)).slice(-3)} 〜)`;

  } catch (e) {
    console.error("【Server Error】", e);
    throw e;
  }
}

/* --- ユーティリティ --- */

// カード送信関数（JSONをそのまま送る）
function sendCard(url, payload) {
  // 第3引数 meta: { task, status, user, context }
  // 戻り値: { success: boolean, error?: string }
  const meta = arguments.length >= 3 ? arguments[2] : null;
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };

  const maxAttempts = 3;
  let attempt = 0;
  while (attempt < maxAttempts) {
    try {
      attempt++;
      const resp = UrlFetchApp.fetch(url, options);
      const code = resp.getResponseCode ? resp.getResponseCode() : 200;
      if (code >= 200 && code < 300) {
        return { success: true };
      } else {
        const body = resp.getContentText ? resp.getContentText() : '';
        const err = `HTTP ${code} ${body}`;
        if (attempt >= maxAttempts) return { success: false, error: err };
        Utilities.sleep(500 * attempt);
      }
    } catch (e) {
      const errMsg = e && e.message ? e.message : String(e);
      if (attempt >= maxAttempts) return { success: false, error: errMsg };
      Utilities.sleep(500 * attempt);
    }
  }
  return { success: false, error: 'Unknown' };
}

function getWebhookUrl() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_SETTING);
  if (!sheet) {
    Logger.log('設定シートが見つかりません: ' + CONFIG.SHEET_SETTING);
    return null;
  }
  const val = sheet.getRange(CONFIG.CELL_WEBHOOK).getValue();
  if (!val) return null;
  return String(val).trim();
}

function getUserMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_SETTING);
  if (!sheet) {
    Logger.log('設定シートが見つかりません: ' + CONFIG.SHEET_SETTING);
    return {};
  }
  const data = sheet.getRange(CONFIG.RANGE_USER_MAP).getValues();
  let map = {};
  data.forEach(row => { if (row[0] && row[1]) map[row[0]] = row[1]; });
  return map;
}

function writeLog(task, status, user, result, context) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_LOG);
  const date = formatDate(new Date(), 'yyyy/MM/dd HH:mm:ss');
  try {
    if (!sheet) {
      Logger.log(`[writeLog] ログシートが見つかりません。${date} ${task} ${status} ${user} ${result} ${context || ''}`);
      return;
    }
    sheet.appendRow([date, task, status, user, result, context || '']);
  } catch (e) {
    Logger.log('[writeLog] 例外: ' + e && e.message ? e.message : String(e));
  }
}

/**
 * タイムゾーンに基づいて日付文字列を返すヘルパー
 * @param {Date} d
 * @param {string} fmt
 */
function getTimeZone() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return (ss && ss.getSpreadsheetTimeZone) ? ss.getSpreadsheetTimeZone() : Session.getScriptTimeZone();
}

function formatDate(d, fmt) {
  if (!d) return '';
  const tz = getTimeZone() || 'JST';
  try {
    return Utilities.formatDate(d, tz, fmt || 'yyyy/MM/dd');
  } catch (e) {
    // fallback
    return Utilities.formatDate(d, 'JST', fmt || 'yyyy/MM/dd');
  }
}