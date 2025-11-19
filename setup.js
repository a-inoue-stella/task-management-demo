/**
 * 【環境構築用スクリプト】
 * この関数を実行すると、設計書通りのシート構造、入力規則、条件付き書式が一括で設定されます。
 * ※既存のデータがある場合、シートが上書きされる可能性があるため、新規シートで実行してください。
 */
function setupEnvironment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. シートの作成・取得
  const sheetTask = getOrCreateSheet(ss, 'タスク管理');
  const sheetConfig = getOrCreateSheet(ss, '設定');
  const sheetLog = getOrCreateSheet(ss, 'ログ');

  // 2. 「設定」シートの構築
  setupConfigSheet(sheetConfig);

  // 3. 「タスク管理」シートの構築
  setupTaskSheet(sheetTask, sheetConfig);

  // 4. 「ログ」シートの構築
  setupLogSheet(sheetLog);

  // 5. 初期シート（シート1等）の削除処理（任意）
  const defaultSheet = ss.getSheetByName('シート1');
  if (defaultSheet) ss.deleteSheet(defaultSheet);

  Browser.msgBox("環境構築が完了しました！");
}

/**
 * シートがあれば取得、なければ作成するユーティリティ
 */
function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * 「設定」シートの中身を作成
 */
function setupConfigSheet(sheet) {
  sheet.clear(); // 初期化
  
  // ヘッダー設定
  const headers = [["担当者名", "Email", "Webhook URL", "", "ステータス定義"]];
  sheet.getRange("A1:E1").setValues(headers).setFontWeight("bold").setBackground("#efefef");
  
  // ステータス定義（マスタデータ）の投入
  const statuses = [
    ["⚪️ 未着手"],
    ["🔵 進行中"],
    ["🟡 確認待ち"],
    ["🟢 完了"]
  ];
  sheet.getRange("E2:E5").setValues(statuses);

  // 列幅調整
  sheet.setColumnWidth(2, 200); // Email列
  sheet.setColumnWidth(3, 300); // Webhook URL列
}

/**
 * 「タスク管理」シートの中身を作成（UI、入力規則、条件付き書式）
 */
function setupTaskSheet(sheet, configSheet) {
  sheet.clear(); // 初期化
  
  // 1. ヘッダー設定
  // I列以降はガントチャート用の日付を入れる（デモ用に30日分）
  let headers = ["task_id", "タスク名", "担当者", "開始日", "期限日", "ステータス", "通知送信", "メモ"];
  
  // 日付ヘッダー生成（今日から30日分）
  const today = new Date();
  for (let i = 0; i < 30; i++) {
    let d = new Date(today);
    d.setDate(today.getDate() + i);
    headers.push(Utilities.formatDate(d, 'JST', 'MM/dd'));
  }
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight("bold")
       .setBackground("#4c8bf5")
       .setFontColor("white")
       .setHorizontalAlignment("center");

  // 列幅調整
  sheet.setColumnWidth(1, 1);  // ID列はほぼ隠す
  sheet.setColumnWidth(2, 250); // タスク名
  sheet.setColumnWidth(7, 60);  // 通知チェックボックス
  // ガントチャートエリア（I列以降）を細くする
  sheet.setColumnWidths(9, 30, 25); 

  // 固定行・列
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  // 2. 入力規則（プルダウン・チェックボックス）の設定
  const maxRow = 100; // 設定範囲

  // C列：担当者（設定シートA列参照）
  const ruleAssignee = SpreadsheetApp.newDataValidation()
    .requireValueInRange(configSheet.getRange("A2:A"))
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 3, maxRow, 1).setDataValidation(ruleAssignee);

  // F列：ステータス（設定シートE列参照）
  const ruleStatus = SpreadsheetApp.newDataValidation()
    .requireValueInRange(configSheet.getRange("E2:E5"))
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 6, maxRow, 1).setDataValidation(ruleStatus);

  // G列：通知送信（チェックボックス）
  const ruleCheck = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();
  sheet.getRange(2, 7, maxRow, 1).setDataValidation(ruleCheck);

  // D, E列：日付
  const ruleDate = SpreadsheetApp.newDataValidation()
    .requireDate()
    .build();
  sheet.getRange(2, 4, maxRow, 2).setDataValidation(ruleDate);


  // 3. 条件付き書式の設定
  const rules = [];
  const rangeAll = sheet.getRange("A2:Z100");
  const rangeGantt = sheet.getRange("I2:AL100"); // ガントチャートエリア

  // ① 完了行のグレーアウト
  // 数式: =$F2="🟢 完了"
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="🟢 完了"')
    .setBackground("#eeeeee")
    .setFontColor("#aaaaaa")
    .setRanges([rangeAll])
    .build());

  // ② 確認待ちのハイライト
  // 数式: =$F2="🟡 確認待ち"
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$F2="🟡 確認待ち"')
    .setBackground("#fff9c4") // 薄い黄色
    .setRanges([rangeAll])
    .build());

  // ③ 遅延アラート（赤）
  // 数式: =AND($F2<>"🟢 完了", $E2 < TODAY(), $E2<>"")
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($F2<>"🟢 完了", $E2 < TODAY(), $E2<>"")')
    .setBackground("#ffcdd2") // 薄い赤
    .setFontColor("#c62828")
    .setRanges([rangeAll])
    .build());

  // ④ ガントチャートのバー表示（青）
  // 数式: =AND(I$1>=$D2, I$1<=$E2)
  // ※GASで設定する場合、R1C1形式の方が安定するためR1C1で記述
  //   I$1 -> R1C[0] (相対列の1行目)
  //   $D2 -> RC4 (固定D列の相対行)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND(R1C[0]>=RC4, R1C[0]<=RC5)')
    .setBackground("#4285f4") // Google Blue
    .setRanges([rangeGantt])
    .build());

  sheet.setConditionalFormatRules(rules);
}

/**
 * 「ログ」シートの中身を作成
 */
function setupLogSheet(sheet) {
  sheet.clear();
  const headers = [["日時", "タスク名", "ステータス", "実行者", "結果"]];
  sheet.getRange("A1:E1").setValues(headers).setFontWeight("bold").setBackground("#efefef");
}