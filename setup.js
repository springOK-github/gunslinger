/**
 * 初期設定とメニュー関連の関数
 */

/**
 * スプレッドシートを開いたときにカスタムメニューを作成します。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🃏 ポケモンマッチング')
    .addItem('シートの初期設定', 'setupSheets')
    .addSeparator()
    .addItem('新しいプレイヤーの登録', 'registerPlayer')
    .addItem('プレイヤーのドロップアウト', 'handleDropout')
    .addSeparator()
    .addItem('対戦結果の入力', 'promptAndRecordResult')
    .addToUi();
}

/**
 * スプレッドシートを初期化し、必要なシートとヘッダーを作成します。
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. プレイヤーシート
  let playerSheet = ss.getSheetByName(SHEET_PLAYERS);
  if (!playerSheet) {
    playerSheet = ss.insertSheet(SHEET_PLAYERS);
  }
  playerSheet.clear();
  playerSheet.getRange("A1:F1").setValues([
    ["プレイヤーID", "勝数", "敗数", "消化試合数", "参加状況", "最終対戦日時"]
  ]).setFontWeight("bold").setBackground("#c9daf8").setHorizontalAlignment("center");
  // 幅の調整
  playerSheet.setColumnWidth(1, 100);
  playerSheet.setColumnWidth(5, 100);
  playerSheet.setColumnWidth(6, 150);

  // 2. 対戦履歴シート
  let historySheet = ss.getSheetByName(SHEET_HISTORY);
  if (!historySheet) {
    historySheet = ss.insertSheet(SHEET_HISTORY);
  }
  historySheet.clear();
  historySheet.getRange("A1:E1").setValues([
    ["日時", "プレイヤー1 ID", "プレイヤー2 ID", "勝者ID", "対戦ID"]
  ]).setFontWeight("bold").setBackground("#fce5cd").setHorizontalAlignment("center");
  historySheet.setColumnWidth(1, 150);

  // 3. 対戦中シート
  let inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
  if (!inProgressSheet) {
    inProgressSheet = ss.insertSheet(SHEET_IN_PROGRESS);
  }
  inProgressSheet.clear();
  inProgressSheet.getRange("A1:B1").setValues([
    ["プレイヤー1 ID", "プレイヤー2 ID"]
  ]).setFontWeight("bold").setBackground("#d9ead3").setHorizontalAlignment("center");
  inProgressSheet.setColumnWidth(3, 80);

  Logger.log("シートの初期設定が完了しました。");
}