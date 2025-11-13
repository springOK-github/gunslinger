/**
 * ガンスリンガーバトル用マッチングシステム
 * @fileoverview アプリケーション層 - 初期化・設定・排他制御
 * @author springOK
 */

// =========================================
// システム初期化・メニュー
// =========================================

/**
 * スプレッドシートを開いたときにカスタムメニューを作成します。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🃏 ガンスリンガーマッチング")
    .addItem("⚙️ シートの初期設定", "setupSheets")
    .addSeparator()
    .addItem("⏱️ 経過時間更新の開始", "setupMatchTimeUpdaterTrigger")
    .addItem("⏹️ 経過時間更新の停止", "deleteMatchTimeUpdaterTrigger")
    .addSeparator()
    .addItem("➕ プレイヤーを追加する", "registerPlayer")
    .addItem("☕ プレイヤーを休憩にする", "restPlayer")
    .addItem("↩️ 休憩から復帰させる", "returnPlayerFromResting")
    .addItem("❌ プレイヤーをドロップアウトさせる", "dropoutPlayer")
    .addSeparator()
    .addItem("✅ 対戦結果の記録", "promptAndRecordResult")
    .addItem("🔧 対戦結果の修正", "correctMatchResult")
    .addSeparator()
    .addItem("⚙️ 最大卓数の設定", "configureMaxTables")
    .addToUi();
}

/**
 * スプレッドシートを初期化し、必要なシートとヘッダーを作成します。
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 確認メッセージを表示
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert("シートの初期設定", "シートの初期設定を行いますか？\n\n既存のデータはすべて削除されます。", ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    ui.alert("シートの初期設定をキャンセルしました。");
    return;
  }

  // タイムゾーンを東京に設定
  ss.setSpreadsheetTimeZone("Asia/Tokyo");

  // 対戦時間計測トリガーを削除
  deleteMatchTimeUpdaterTrigger(false);

  // 1. プレイヤーシート
  let playerSheet = ss.getSheetByName(SHEET_PLAYERS);
  if (!playerSheet) {
    playerSheet = ss.insertSheet(SHEET_PLAYERS);
  }
  playerSheet.clear();
  const playerHeaders = REQUIRED_HEADERS[SHEET_PLAYERS];
  playerSheet
    .getRange(1, 1, 1, playerHeaders.length)
    .setValues([playerHeaders])
    .setFontWeight("bold")
    .setBackground("#c9daf8")
    .setHorizontalAlignment("center");

  // 2. 対戦履歴シート
  let historySheet = ss.getSheetByName(SHEET_HISTORY);
  if (!historySheet) {
    historySheet = ss.insertSheet(SHEET_HISTORY);
  }
  historySheet.clear();
  const historyHeaders = REQUIRED_HEADERS[SHEET_HISTORY];
  historySheet
    .getRange(1, 1, 1, historyHeaders.length)
    .setValues([historyHeaders])
    .setFontWeight("bold")
    .setBackground("#fce5cd")
    .setHorizontalAlignment("center");

  // 3. マッチングシート
  let inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
  if (!inProgressSheet) {
    inProgressSheet = ss.insertSheet(SHEET_IN_PROGRESS);
  }
  inProgressSheet.clear();
  const inProgressHeaders = REQUIRED_HEADERS[SHEET_IN_PROGRESS];
  inProgressSheet
    .getRange(1, 1, 1, inProgressHeaders.length)
    .setValues([inProgressHeaders])
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center");

  Logger.log("シートの初期設定が完了しました。");
}

// =========================================
// システム設定管理
// =========================================

/**
 * 現在の最大卓数を取得します。
 * PropertiesServiceに保存されている値、なければデフォルト値を返します。
 * @returns {number} 最大卓数
 */
function getMaxTables() {
  const properties = PropertiesService.getDocumentProperties();
  const savedMaxTables = properties.getProperty("MAX_TABLES");

  if (savedMaxTables) {
    return parseInt(savedMaxTables, 10);
  }

  // デフォルト値
  return TABLE_CONFIG.MAX_TABLES;
}

/**
 * 最大卓数を設定します。
 * @param {number} maxTables - 設定する最大卓数
 */
function setMaxTables(maxTables) {
  const properties = PropertiesService.getDocumentProperties();
  properties.setProperty("MAX_TABLES", maxTables.toString());
  Logger.log(`最大卓数を ${maxTables} に設定しました。`);
}

/**
 * 最大卓数の設定をユーザーに促すダイアログを表示します。
 */
function configureMaxTables() {
  const ui = SpreadsheetApp.getUi();
  const currentMaxTables = getMaxTables();

  const response = ui.prompt(
    "最大卓数の設定",
    `現在の最大卓数: ${currentMaxTables}卓\n\n` + `新しい最大卓数を入力してください（1～200）：`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert("設定をキャンセルしました。");
    return;
  }

  const input = response.getResponseText().trim();

  // 入力検証
  if (!/^\d+$/.test(input)) {
    ui.alert("エラー", "数字のみで入力してください。", ui.ButtonSet.OK);
    return;
  }

  const newMaxTables = parseInt(input, 10);

  // 範囲検証
  if (newMaxTables < 1 || newMaxTables > 200) {
    ui.alert("エラー", "最大卓数は1～200の範囲で入力してください。", ui.ButtonSet.OK);
    return;
  }

  // 使用中の卓がある場合、それより小さい値には減らせない
  const maxUsedTable = getMaxUsedTableNumber();
  if (newMaxTables < maxUsedTable) {
    ui.alert(
      "エラー",
      `現在、卓番号 ${maxUsedTable} まで使用中です。\n\n` + `使用中の卓番号より小さい値には減らせません。\n` + `最小値: ${maxUsedTable}卓`,
      ui.ButtonSet.OK
    );
    return;
  }

  // 確認ダイアログ
  const confirmResponse = ui.alert(
    "設定の確認",
    `最大卓数を ${currentMaxTables}卓 → ${newMaxTables}卓 に変更します。\n\n` + "よろしいですか？",
    ui.ButtonSet.YES_NO
  );

  if (confirmResponse !== ui.Button.YES) {
    ui.alert("設定をキャンセルしました。");
    return;
  }

  // 設定を保存
  setMaxTables(newMaxTables);

  ui.alert("設定完了", `最大卓数を ${newMaxTables}卓 に設定しました。`, ui.ButtonSet.OK);
}

/**
 * updateAllMatchTimesを1分周期でGASトリガーと仕掛けます。
 */

function setupMatchTimeUpdaterTrigger() {
  // 確認するダイアログを表示
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert("対戦時間計測タイマーの開始", "対戦時間計測タイマーを開始しますか？", ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    ui.alert("タイマーの開始をキャンセルしました。");
    return;
  }

  // 既存のトリガーを削除
  deleteMatchTimeUpdaterTrigger(false);

  // 新しいトリガーを作成（1分ごと）
  ScriptApp.newTrigger("updateAllMatchTimes").timeBased().everyMinutes(1).create();

  // 初回実行
  updateAllMatchTimes();
}

/**
 * トリガーを削除します
 * @param {boolean} showAlert - ユーザーに完了メッセージを表示するかどうか
 */

function deleteMatchTimeUpdaterTrigger(showAlert = true) {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  if (triggers.length === 0 && showAlert) {
    const ui = SpreadsheetApp.getUi();
    ui.alert("タイマーは既に停止されています。", ui.ButtonSet.OK);
    return;
  }
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === "updateAllMatchTimes") {
      if (showAlert) {
        // 確認するダイアログを表示
        const ui = SpreadsheetApp.getUi();
        const response = ui.alert("対戦時間計測タイマーの停止", "対戦時間計測タイマーを停止しますか？", ui.ButtonSet.YES_NO);

        if (response !== ui.Button.YES) {
          ui.alert("タイマーの停止をキャンセルしました。");
          return;
        }
      }
      showAlert = false; // 一度表示したら表示しない

      ScriptApp.deleteTrigger(trigger);
    } else {
      // メッセージ
      if (showAlert) {
        const ui = SpreadsheetApp.getUi();
        ui.alert("タイマーは既に停止されています。", ui.ButtonSet.OK);
      }
    }
  }
}

// =========================================
// 排他制御
// =========================================

// ロックの最大待機時間（ミリ秒）
const LOCK_TIMEOUT = 30000; // 30秒

/**
 * スプレッドシートの排他ロックを取得します。
 * @param {string} lockName - ロックの名前（操作の種類を識別）
 * @returns {LockService.Lock} 取得したロック
 * @throws {Error} ロックが取得できない場合
 */
function acquireLock(lockName) {
  const lock = LockService.getScriptLock();
  const success = lock.tryLock(LOCK_TIMEOUT);

  if (!success) {
    throw new Error("他のユーザーが操作中です。\n" + "しばらく待ってから再度お試しください。\n" + `(${lockName})`);
  }

  return lock;
}

/**
 * ロックを解放します。
 * @param {LockService.Lock} lock - 解放するロック
 */
function releaseLock(lock) {
  if (lock) {
    try {
      lock.releaseLock();
    } catch (e) {
      Logger.log("ロックの解放に失敗: " + e.toString());
    }
  }
}
