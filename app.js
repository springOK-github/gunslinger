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
    .addItem("🏁 大会開始", "startTournament")
    .addItem("🏁 大会終了", "endTournament")
    .addSeparator()
    .addItem("➕ プレイヤーを追加する", "registerPlayer")
    .addItem("🖊️ プレイヤー名を編集", "editPlayerName")
    .addItem("☕ プレイヤーを休憩させる", "restPlayer")
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
function startTournament() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const autoMatchingGate = withAutoMatchingGate("大会開始", (gate) => gate);

  if (!autoMatchingGate) {
    return;
  }

  // 確認メッセージを表示
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert("大会開始", "大会を開始します。\n\n既存のデータはすべて削除されます。大会を開始してもよろしいですか？", ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    ui.alert("大会開始をキャンセルしました。");
    return;
  }

  if (!autoMatchingGate.assertUnchanged()) {
    return;
  }

  // タイムゾーンを東京に設定
  ss.setSpreadsheetTimeZone("Asia/Tokyo");

  // 対戦時間計測トリガーを追加
  setupMatchTimeUpdaterTrigger(false);

  // 1/2/3. 必要なシートを作成してヘッダーを初期化
  const playerSheet = ensureAndInitSheet(ss, SHEET_PLAYERS);
  const historySheet = ensureAndInitSheet(ss, SHEET_HISTORY);
  const inProgressSheet = ensureAndInitSheet(ss, SHEET_IN_PROGRESS);

  Logger.log("大会を開始します。Ready to go!!");
}

/**
 * 大会終了処理: 対戦履歴シートをコピーして日時付きでリネーム（バックアップ）します。
 * コピー元のデータはそのまま残します（運用で必要ならコピー後にクリアも可能）。
 */
function endTournament() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const autoMatchingGate = withAutoMatchingGate("大会終了", (gate) => gate);

  if (!autoMatchingGate) {
    return;
  }

  const response = ui.alert("大会終了", "大会を終了して対戦履歴をバックアップします。よろしいですか？", ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) {
    ui.alert("大会終了をキャンセルしました。");
    return;
  }

  if (!autoMatchingGate.assertUnchanged()) {
    return;
  }

  let lock = null;
  try {
    lock = acquireLock("大会終了");

    // メンテナンスモードを有効化して自動マッチングを抑止
    try {
      PropertiesService.getDocumentProperties().setProperty("MAINTENANCE_MODE", "1");
    } catch (e) {
      Logger.log("MAINTENANCE_MODE の設定に失敗: " + e && e.toString());
    }

    // 対戦時間計測タイマーの停止
    deleteMatchTimeUpdaterTrigger(false);

    const historySheet = ss.getSheetByName(SHEET_HISTORY);
    if (!historySheet) {
      ui.alert("エラー", `シートが見つかりません: ${SHEET_HISTORY}`, ui.ButtonSet.OK);
      return;
    }

    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    if (!playerSheet) {
      ui.alert("エラー", `シートが見つかりません: ${SHEET_PLAYERS}`, ui.ButtonSet.OK);
      return;
    }

    // 進行中の対戦があるか確認し、ユーザーの確認後に別関数で強制終了処理を実行する
    const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
    if (inProgressSheet) {
      const { indices: inIdx, data: inData } = getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);
      let activeCount = 0;
      for (let i = 1; i < inData.length; i++) {
        const row = inData[i];
        const id1 = row[inIdx["ID1"]];
        const id2 = row[inIdx["ID2"]];
        if (id1 && id2) activeCount++;
      }

      if (activeCount > 0) {
        const confirm = ui.alert(
          "対戦中の卓があります",
          `現在 ${activeCount} 件の対戦が進行中です。大会終了の前にこれらを強制終了してバックアップしますか？\n\n(はいを選ぶと、進行中の対戦を対戦履歴に『大会終了』として記録し、選手を待機状態に戻します。)`,
          ui.ButtonSet.YES_NO
        );

        if (confirm !== ui.Button.YES) {
          ui.alert("大会終了をキャンセルしました。");
          return;
        }

        // 実際の強制終了処理は別関数に切り出し
        endAllActiveMatches();
      }
    }

    // 日時を取得（スプレッドシートのタイムゾーンを使用）
    const tz = ss.getSpreadsheetTimeZone() || "Asia/Tokyo";
    const timestamp = Utilities.formatDate(new Date(), tz, "yyyyMMdd_HHmmss");
    const baseName = `${SHEET_HISTORY}_${timestamp}`;

    const backupHistoryName = createSheetBackup(ss, historySheet, baseName);
    const backupPlayerName = createSheetBackup(ss, playerSheet, `${SHEET_PLAYERS}_${timestamp}`);
    Logger.log(`大会終了: 対戦履歴とプレイヤーをバックアップしました -> ${backupHistoryName}, ${backupPlayerName}`);

    // バックアップ後にオリジナルのシートを初期化（ヘッダー再作成）
    try {
      ensureAndInitSheet(ss, SHEET_HISTORY);
    } catch (e) {
      Logger.log("endTournament: シート初期化に失敗しました: " + (e && e.toString()));
    }
  } catch (e) {
    ui.alert("エラー", `大会終了に失敗しました: ${e.message}`, ui.ButtonSet.OK);
    Logger.log("endTournament エラー: " + e.toString());
  } finally {
    // メンテナンスモードを解除
    try {
      PropertiesService.getDocumentProperties().deleteProperty("MAINTENANCE_MODE");
    } catch (e) {
      Logger.log("MAINTENANCE_MODE の解除に失敗: " + e && e.toString());
    }

    releaseLock(lock);
  }
}

/**
 * 進行中の全対戦を強制終了して対戦履歴に追記し、プレイヤーを待機状態に戻す処理。
 * この関数はロックを保持した状態で呼び出すことを想定しています（呼び出し元で acquireLock を行ってください）。
 * @returns {number} 終了させた対戦の件数
 */
function endAllActiveMatches() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  let ownedMaintenanceFlag = false;

  // もし MAINTENANCE_MODE が未設定ならここで設定して、終了時に解除する
  try {
    const current = props.getProperty("MAINTENANCE_MODE");
    if (current !== "1") {
      props.setProperty("MAINTENANCE_MODE", "1");
      ownedMaintenanceFlag = true;
    }
  } catch (e) {
    Logger.log("endAllActiveMatches: MAINTENANCE_MODE 操作に失敗: " + (e && e.toString()));
  }

  const historySheet = ss.getSheetByName(SHEET_HISTORY);
  const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);

  if (!historySheet || !inProgressSheet) return 0;

  const { indices: inIdx, data: inData } = getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);
  const { indices: histIdx, data: histData } = getSheetStructure(historySheet, SHEET_HISTORY);

  // 活動中の対戦行を収集
  const activeRows = [];
  for (let i = 1; i < inData.length; i++) {
    const row = inData[i];
    const id1 = row[inIdx["ID1"]];
    const id2 = row[inIdx["ID2"]];
    if (id1 && id2) activeRows.push({ rowIndex: i + 1, row });
  }

  if (activeRows.length === 0) return 0;

  // 対戦ID の最大を計算
  let maxNum = 0;
  for (let i = 1; i < histData.length; i++) {
    const vid = histData[i][histIdx["対戦ID"]];
    if (typeof vid === "string" && vid.startsWith("T")) {
      const num = parseInt(vid.substring(1), 10);
      if (!isNaN(num) && num > maxNum) maxNum = num;
    }
  }

  const tz = ss.getSpreadsheetTimeZone() || "Asia/Tokyo";
  const now = new Date();
  const endTimeStr = Utilities.formatDate(now, tz, "yyyy/MM/dd HH:mm:ss");

  const rowsToAppend = [];
  for (const item of activeRows) {
    const r = item.row;
    const tableNumber = r[inIdx["卓番号"]] || "";
    const id1 = r[inIdx["ID1"]] || "";
    const name1 = r[inIdx["プレイヤー1"]] || id1;
    const id2 = r[inIdx["ID2"]] || "";
    const name2 = r[inIdx["プレイヤー2"]] || id2;

    maxNum++;
    const matchId = "T" + Utilities.formatString("%04d", maxNum);
    const winnerName = "大会終了";
    const matchTime = "";

    const newRow = [];
    newRow[histIdx["対戦ID"]] = matchId;
    newRow[histIdx["卓番号"]] = tableNumber;
    newRow[histIdx["ID1"]] = id1;
    newRow[histIdx["プレイヤー1"]] = name1;
    newRow[histIdx["ID2"]] = id2;
    newRow[histIdx["プレイヤー2"]] = name2;
    newRow[histIdx["勝者名"]] = winnerName;
    newRow[histIdx["対戦終了時刻"]] = endTimeStr;
    newRow[histIdx["対戦時間"]] = matchTime;

    rowsToAppend.push(newRow);

    // プレイヤー状態を待機に戻す
    try {
      updatePlayerState({
        targetPlayerId: id1,
        newStatus: PLAYER_STATUS.WAITING,
        opponentNewStatus: PLAYER_STATUS.WAITING,
        recordResult: false,
        isTargetWinner: false,
      });
    } catch (e) {
      Logger.log("updatePlayerState error for %s: %s", id1, e && e.toString());
    }
    try {
      updatePlayerState({
        targetPlayerId: id2,
        newStatus: PLAYER_STATUS.WAITING,
        opponentNewStatus: PLAYER_STATUS.WAITING,
        recordResult: false,
        isTargetWinner: false,
      });
    } catch (e) {
      Logger.log("updatePlayerState error for %s: %s", id2, e && e.toString());
    }
  }

  // 追記
  const appendStartRow = historySheet.getLastRow() + 1;
  for (let i = 0; i < rowsToAppend.length; i++) {
    const rowVals = [];
    const headers = REQUIRED_HEADERS[SHEET_HISTORY];
    for (let j = 0; j < headers.length; j++) {
      rowVals.push(rowsToAppend[i][j] || "");
    }
    historySheet.getRange(appendStartRow + i, 1, 1, rowVals.length).setValues([rowVals]);
  }

  // 進行中シートの該当行はクリアする
  for (const item of activeRows) {
    const rIdx = item.rowIndex;
    const clearCols = [inIdx["ID1"], inIdx["プレイヤー1"], inIdx["ID2"], inIdx["プレイヤー2"], inIdx["対戦開始時刻"], inIdx["経過時間"]];
    for (const c of clearCols) {
      inProgressSheet.getRange(rIdx, c + 1).setValue("");
    }
  }

  // MAINTENANCE_MODE をこの関数で立てた場合は解除する
  try {
    if (ownedMaintenanceFlag) {
      props.deleteProperty("MAINTENANCE_MODE");
    }
  } catch (e) {
    Logger.log("endAllActiveMatches: MAINTENANCE_MODE の解除に失敗: " + (e && e.toString()));
  }

  return activeRows.length;
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
  const autoMatchingGate = withAutoMatchingGate("最大卓数の設定", (gate) => gate);

  if (!autoMatchingGate) {
    return;
  }

  const response = ui.prompt(
    "最大卓数の設定",
    `現在の最大卓数: ${currentMaxTables}卓\n\n` + `新しい最大卓数を入力してください（1～200）：`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert("設定をキャンセルしました。");
    return;
  }

  if (!autoMatchingGate.assertUnchanged()) {
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

function setupMatchTimeUpdaterTrigger(showAlert = true) {
  // 確認するダイアログを表示
  if (showAlert) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert("対戦時間計測タイマーの開始", "対戦時間計測タイマーを開始しますか？", ui.ButtonSet.YES_NO);

    if (response !== ui.Button.YES) {
      ui.alert("タイマーの開始をキャンセルしました。");
      return;
    }
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

/**
 * シートをコピーし、一意の名前で末尾に配置します。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシート本体
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - コピー対象のシート
 * @param {string} baseName - 付与したい基本名（例: プレイヤー_20241128）
 * @returns {string} コピー後のシート名
 */
function createSheetBackup(ss, sheet, baseName) {
  const copied = sheet.copyTo(ss);
  let newName = baseName;
  let suffix = 1;
  while (ss.getSheetByName(newName)) {
    newName = `${baseName}_${suffix}`;
    suffix++;
  }

  copied.setName(newName);
  ss.setActiveSheet(copied);
  ss.moveActiveSheet(ss.getNumSheets());
  return newName;
}

/**
 * 指定したシート名のシートを取得（なければ作成）し、ヘッダー行をクリアして再作成します。
 * ヘッダーの内容は `REQUIRED_HEADERS` から取得し、シート名ごとに背景色を設定します。
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetName
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} 初期化されたシート
 */
function ensureAndInitSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  // ヘッダーをクリアして再作成
  sheet.clear();
  const headers = REQUIRED_HEADERS[sheetName] || [];

  // シートごとの背景色
  const bgMap = {};
  bgMap[SHEET_PLAYERS] = "#c9daf8";
  bgMap[SHEET_HISTORY] = "#fce5cd";
  bgMap[SHEET_IN_PROGRESS] = "#d9ead3";

  const bg = bgMap[sheetName] || null;

  if (headers.length > 0) {
    const range = sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    if (bg) range.setBackground(bg);
    range.setHorizontalAlignment("center");
  }

  return sheet;
}

// =========================================
// 排他制御
// =========================================

// ロックの最大待機時間（ミリ秒）
const LOCK_TIMEOUT = 30000; // 30秒

/**
 * スプレッドシートの排他ロックを取得します。
 * @param {string} lockName - ロックの名前（操作の種類を識別）
 * @returns {Object} 取得したロック
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
 * @param {Object} lock - 解放するロック
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
