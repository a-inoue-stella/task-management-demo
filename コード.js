/**
 * 【設定エリア】
 * シートの列番号が変わった場合はここを修正してください。
 */
const CONFIG = {
  SHEET_TASK: 'タスク管理',
  SHEET_SETTING: '設定',
  SHEET_LOG: 'ログ',
  // 列番号（A列=1, B列=2...）
  COL_TASK_NAME: 2,   // B列: タスク名
  COL_ASSIGNEE: 3,    // C列: 担当者
  COL_DEADLINE: 5,    // E列: 期限日
  COL_STATUS: 6,      // F列: ステータス
  COL_TRIGGER: 7,     // G列: 通知送信（チェックボックス）
  // 設定シートのセル位置
  CELL_WEBHOOK: 'C2',     // Webhook URL
  RANGE_USER_MAP: 'A2:B20' // 担当者マスタ範囲
};

/**
 * 1. トリガー関数 (onEdit)
 * ユーザーが操作した瞬間に動く関数です。
 * 負荷対策のため「タスク管理シートのG列がチェックされた時」以外は即終了させます。
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();

  // ガード節：無関係な編集は無視して負荷を下げる
  if (sheet.getName() !== CONFIG.SHEET_TASK) return;
  if (range.getColumn() !== CONFIG.COL_TRIGGER) return;
  if (e.value !== "TRUE") return; // チェックON以外（OFFにした時など）は無視

  // 通知処理を実行
  processNotification(sheet, range.getRow());
}

/**
 * 2. 通知処理の実行 (排他制御付き)
 * 複数人が同時にチェックしてもバッティングしないよう制御します。
 */
function processNotification(sheet, rowIndex) {
  const lock = LockService.getScriptLock();
  
  // ロック取得（最大10秒待機）
  if (lock.tryLock(10000)) {
    try {
      // 必要なデータを一行分取得
      // getRange(行, 列, 行数, 列数) -> 1行目のデータ全体を取得
      const data = sheet.getRange(rowIndex, 1, 1, 10).getValues()[0];
      
      const taskName = data[CONFIG.COL_TASK_NAME - 1];
      const assignee = data[CONFIG.COL_ASSIGNEE - 1];
      const status   = data[CONFIG.COL_STATUS - 1];
      
      // 1. メッセージを作る
      const message = createMessage(taskName, assignee, status);
      
      // 2. チャットに送る
      const webhookUrl = getWebhookUrl();
      if(webhookUrl) {
        sendChat(webhookUrl, message);
        writeLog(taskName, status, assignee, "送信成功");
      } else {
        Browser.msgBox("エラー：設定シート(C2)にWebhook URLが設定されていません");
        writeLog(taskName, status, assignee, "エラー：URL未設定");
      }

      // 3. チェックボックスをOFFに戻す（処理完了の合図）
      sheet.getRange(rowIndex, CONFIG.COL_TRIGGER).setValue(false);

      // 4. 完了トーストを表示（画面右下に小さく出る）
      SpreadsheetApp.getActiveSpreadsheet().toast(`「${taskName}」の通知を送信しました`, "完了");

    } catch (e) {
      console.error(e);
      writeLog("システムエラー", "エラー", "不明", e.message);
      // エラーでもチェックは戻す
      sheet.getRange(rowIndex, CONFIG.COL_TRIGGER).setValue(false);
      SpreadsheetApp.getActiveSpreadsheet().toast("エラーが発生しました", "失敗");
    } finally {
      lock.releaseLock();
    }
  }
}

/**
 * 3. メッセージ生成ロジック
 * ステータスに応じて文面とアイコンを変えます。
 */
function createMessage(taskName, assigneeName, status) {
  const userMap = getUserMap();
  const email = userMap[assigneeName];
  
  // Emailがあればメンション化、なければ名前だけ
  const mention = email ? `<users/${email}>` : assigneeName;
  const sheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

  let header = "";
  let body = "";
  
  if (status === "🟡 確認待ち") {
    // 確認待ちは目立つように
    header = `*🟡 【確認依頼】タスクの確認をお願いします*`;
    body = `担当者：${assigneeName} さんより\nステータスが「確認待ち」になりました。`;
  } else if (status === "🟢 完了") {
    // 完了はポジティブに
    header = `*🟢 【完了】タスクが完了しました！*`;
    body = `担当者：${mention} お疲れ様でした！`;
  } else if (status === "🔵 進行中") {
    header = `*🔵 【着手】タスクを開始しました*`;
    body = `担当者：${mention}`;
  } else {
    // その他
    header = `*🔄 【更新】タスク状況が変わりました*`;
    body = `担当者：${mention}\n現在：${status}`;
  }

  // 統合メッセージ
  const text = `${header}\n` +
               `タスク：*${taskName}*\n` +
               `${body}\n` +
               `──────────────\n` +
               `<${sheetUrl}|📂 スプレッドシートを開く>`;
  
  return text;
}

/**
 * 4. リマインド機能（デモボタン用）
 * 期限切れタスクを吸い上げて通知します。
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
  today.setHours(0,0,0,0); // 時間をリセットして日付比較
  
  let alertTasks = [];

  data.forEach(row => {
    const taskName = row[CONFIG.COL_TASK_NAME - 1];
    const deadlineStr = row[CONFIG.COL_DEADLINE - 1];
    const status   = row[CONFIG.COL_STATUS - 1];

    // 完了済みと空行は無視
    if (status === "🟢 完了" || !taskName) return;

    const deadline = new Date(deadlineStr);
    
    // 期限切れチェック (期限 < 今日)
    if (deadline < today && deadlineStr) {
      const dateStr = Utilities.formatDate(deadline, 'JST', 'MM/dd');
      alertTasks.push(`・🔥 ${taskName} (期限: ${dateStr}) -> ${status}`);
    }
  });

  if (alertTasks.length > 0) {
    const webhookUrl = getWebhookUrl();
    if (!webhookUrl) {
      Browser.msgBox("Webhook URLが設定されていません");
      return;
    }
    
    const msg = `*🔴 【期限アラート】以下のタスクが遅延しています*\n` + 
                alertTasks.join("\n") + 
                `\n\n<${ss.getUrl()}|📂 至急確認してください>`;
    
    sendChat(webhookUrl, msg);
    Browser.msgBox(`送信完了：${alertTasks.length}件の遅延タスクを通知しました。`);
  } else {
    Browser.msgBox("現在、期限切れのタスクはありません。優秀です！");
  }
}

/* --- 以下、ユーティリティ関数 --- */

// Chat送信
function sendChat(url, text) {
  const payload = { text: text };
  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload)
  };
  UrlFetchApp.fetch(url, options);
}

// Webhook URL取得
function getWebhookUrl() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SETTING);
  return sheet.getRange(CONFIG.CELL_WEBHOOK).getValue();
}

// ユーザーマスタ取得
function getUserMap() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SETTING);
  const data = sheet.getRange(CONFIG.RANGE_USER_MAP).getValues();
  let map = {};
  data.forEach(row => {
    if(row[0] && row[1]) map[row[0]] = row[1];
  });
  return map;
}

// ログ書き込み
function writeLog(task, status, user, result) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_LOG);
  const date = Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss');
  sheet.appendRow([date, task, status, user, result]);
}