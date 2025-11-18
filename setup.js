/**
 * [初期セットアップ用 - 追加機能]
 * デモ用の「ダッシュボード」シートを作成し、
 * 「タスク管理」シートのデータを自動集計する関数とグラフを配置します。
 *
 * @param {string} taskSheetName 集計対象のシート名（デフォルト: "タスク管理"）
 */
function setupDashboard(taskSheetName = "タスク管理") {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. ダッシュボードシートの作成・初期化
  let dashboardSheet = ss.getSheetByName("ダッシュボード");
  if (!dashboardSheet) {
    dashboardSheet = ss.insertSheet("ダッシュボード", 0); // 先頭にシートを挿入
  }
  dashboardSheet.clear(); // 既存の内容をクリア
  dashboardSheet.setFrozenRows(1); // 1行目（タイトル行）を固定

  // -------------------------------------------------------
  // 2. タイトル設定
  // -------------------------------------------------------
  dashboardSheet.getRange("A1").setValue("プロジェクト・ヘルスダッシュボード")
    .setFontWeight("bold")
    .setFontSize(18)
    .setFontFamily("Arial");
  // A1:G1セルを結合
  dashboardSheet.getRange("A1:G1").merge();

  // -------------------------------------------------------
  // 3. ウィジェット①: 全体進捗サマリー
  // -------------------------------------------------------
  dashboardSheet.getRange("A3").setValue("① 全体進捗サマリー").setFontWeight("bold");
  
  // 集計表の作成
  const summaryData = [
    ["ステータス", "件数"],
    ["完了",        `=COUNTIF('${taskSheetName}'!D:D, "完了")`],
    ["作業中",      `=COUNTIF('${taskSheetName}'!D:D, "作業中")`],
    ["未着手",      `=COUNTIF('${taskSheetName}'!D:D, "未着手")`],
    ["(リスク) 期限切れ", `=COUNTIFS('${taskSheetName}'!D:D, "<>完了", '${taskSheetName}'!E:E, "<"&TODAY())`]
  ];
  
  const summaryRange = dashboardSheet.getRange("A4:B8");
  summaryRange.setValues(summaryData);
  dashboardSheet.getRange("A4:B4").setFontWeight("bold").setBackground("#eeeeee");
  dashboardSheet.getRange("A8:B8").setFontWeight("bold").setFontColor("red");

  // 円グラフの作成
  const pieChartRange = dashboardSheet.getRange("A5:B7"); // 完了, 作業中, 未着手のみ
  const pieChart = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(pieChartRange)
    .setOption("title", "タスクステータス（全体）")
    .setPosition(3, 3, 0, 0) // C3セルを基点に配置
    .build();
  dashboardSheet.insertChart(pieChart);

  // -------------------------------------------------------
  // 4. ウィジェット②: 担当者別 負荷状況（最重要）
  // -------------------------------------------------------
  dashboardSheet.getRange("E3").setValue("② 担当者別 負荷状況 (未完了タスク)").setFontWeight("bold");
  
  // QUERY関数でデータソースを作成
  const queryFormulaLoad = `=QUERY('${taskSheetName}'!A:E, 
    "SELECT B, COUNT(B) 
     WHERE D <> '完了' AND B IS NOT NULL 
     GROUP BY B 
     ORDER BY COUNT(B) DESC 
     LABEL B '担当者', COUNT(B) '未完了タスク数'", 
    1)`;
  
  const queryRangeLoad = dashboardSheet.getRange("E4");
  queryRangeLoad.setFormula(queryFormulaLoad);

  // 横棒グラフの作成
  const barChart = dashboardSheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(queryRangeLoad) // QUERY結果を動的に参照
    .setOption("title", "担当者別 未完了タスク数")
    .setOption("legend", { position: "none" })
    .setOption("hAxis", { title: "件数" })
    .setPosition(4, 5, 0, 0) // E4セルを基点に配置
    .build();
  dashboardSheet.insertChart(barChart);

  // -------------------------------------------------------
  // 5. ウィジェット③: リスク管理
  // -------------------------------------------------------
  dashboardSheet.getRange("A10").setValue("③ リスク管理 (優先度別 期限切れタスク)").setFontWeight("bold");

  // QUERY関数でデータソースを作成
  const queryFormulaRisk = `=QUERY('${taskSheetName}'!A:E, 
    "SELECT C, COUNT(C) 
     WHERE D <> '完了' AND E < TODAY() AND C IS NOT NULL 
     GROUP BY C 
     ORDER BY C ASC 
     LABEL C '優先度', COUNT(C) '期限切れ件数'", 
    1)`;
  
  const queryRangeRisk = dashboardSheet.getRange("A11");
  queryRangeRisk.setFormula(queryFormulaRisk);
  
  // リスクテーブルの書式設定（グラフの代わり）
  dashboardSheet.getRange("A11:B11").setFontWeight("bold").setBackground("#eeeeee");
  
  // -------------------------------------------------------
  // 6. 完了通知
  // -------------------------------------------------------
  SpreadsheetApp.getUi().alert("ダッシュボードシートの構築が完了しました！");
}