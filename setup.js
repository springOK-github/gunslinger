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
    .addItem('プレイヤーのドロップアウト', 'dropoutPlayer')
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
  const playerHeaders = REQUIRED_HEADERS[SHEET_PLAYERS];
  playerSheet.getRange(1, 1, 1, playerHeaders.length).setValues([playerHeaders])
    .setFontWeight("bold").setBackground("#c9daf8").setHorizontalAlignment("center");
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
  const historyHeaders = REQUIRED_HEADERS[SHEET_HISTORY];
  historySheet.getRange(1, 1, 1, historyHeaders.length).setValues([historyHeaders])
    .setFontWeight("bold").setBackground("#fce5cd").setHorizontalAlignment("center");
  historySheet.setColumnWidth(1, 150);

  // 3. 対戦中シート
  let inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
  if (!inProgressSheet) {
    inProgressSheet = ss.insertSheet(SHEET_IN_PROGRESS);
  }
  inProgressSheet.clear();
  const inProgressHeaders = REQUIRED_HEADERS[SHEET_IN_PROGRESS];
  inProgressSheet.getRange(1, 1, 1, inProgressHeaders.length).setValues([inProgressHeaders])
    .setFontWeight("bold").setBackground("#d9ead3").setHorizontalAlignment("center");
  inProgressSheet.setColumnWidth(3, 80);

  Logger.log("シートの初期設定が完了しました。");
}