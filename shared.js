/**
 * ガンスリンガーバトル用マッチングシステム
 * @fileoverview 共有ユーティリティ - シート操作とUI共通処理
 * @author springOK
 */

// =========================================
// シート操作ユーティリティ
// =========================================

/**
 * シートの構造を取得し、ヘッダー行の検証を行います。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 対象のシート
 * @param {string} sheetName シート名（定数から取得）
 * @returns {Object} { headers: 配列, indices: オブジェクト, data: 2次元配列 }
 */

function getSheetStructure(sheet, sheetName) {
  if (!sheet) {
    throw new Error(`シート「${sheetName}」が見つかりません。`);
  }

  const data = sheet.getDataRange().getValues();
  if (!data || data.length === 0) {
    throw new Error(`シート「${sheetName}」にデータがありません。`);
  }

  const headers = data[0].map((h) => String(h).trim());
  const indices = {};
  const missing = [];

  const requiredHeaders = REQUIRED_HEADERS[sheetName];
  if (!requiredHeaders) {
    throw new Error(`シート「${sheetName}」の必須ヘッダー定義が見つかりません。`);
  }

  requiredHeaders.forEach((required) => {
    const idx = headers.indexOf(required);
    if (idx === -1) {
      missing.push(required);
    } else {
      indices[required] = idx;
    }
  });

  if (missing.length > 0) {
    throw new Error(`シート「${sheetName}」に必須ヘッダーが不足しています: ${missing.join(", ")}`);
  }

  return { headers, indices, data };
}

/**
 * プレイヤーIDから名前を取得します
 * @param {string} playerId プレイヤーID
 * @returns {string} プレイヤー名。見つからない場合はIDをそのまま返します
 */
function getPlayerName(playerId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
  if (!playerSheet) {
    Logger.log("エラー: プレイヤーシートが見つかりません。");
    return playerId;
  }
  const { indices, data } = getSheetStructure(playerSheet, SHEET_PLAYERS);

  for (let i = 1; i < data.length; i++) {
    if (data[i][indices["プレイヤーID"]] === playerId) {
      return data[i][indices["プレイヤー名"]] || playerId;
    }
  }
  return playerId;
}

/**
 * 卓番号が有効かどうかを検証します
 * @param {number} tableNumber 検証する卓番号
 * @returns {{isValid: boolean, message: string}} 検証結果とメッセージ
 */
function validateTableNumber(tableNumber) {
  const maxTables = getMaxTables(); // 動的に取得

  if (!Number.isInteger(tableNumber)) {
    return { isValid: false, message: "卓番号は整数である必要があります。" };
  }

  if (tableNumber < TABLE_CONFIG.MIN_TABLE_NUMBER) {
    return { isValid: false, message: `卓番号は${TABLE_CONFIG.MIN_TABLE_NUMBER}以上である必要があります。` };
  }

  if (tableNumber > maxTables) {
    return { isValid: false, message: `卓番号は${maxTables}以下である必要があります。` };
  }

  return { isValid: true, message: "有効な卓番号です。" };
}

/**
 * 現在使用中の最大卓番号を取得します
 * @returns {number} 使用中の最大卓番号（使用中の卓がない場合は0）
 */
function getMaxUsedTableNumber() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);

  if (!inProgressSheet) {
    return 0;
  }

  const { indices, data } = getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);
  let maxUsed = 0;

  // 現在使用中の卓番号から最大値を取得
  for (let i = 1; i < data.length; i++) {
    const tableNumber = data[i][indices["卓番号"]];
    if (tableNumber && tableNumber > maxUsed) {
      maxUsed = tableNumber;
    }
  }

  return maxUsed;
}

/**
 * 使用可能な次の卓番号を取得します
 * @param {GoogleAppsScript.Spreadsheet.Sheet} inProgressSheet マッチングシート
 * @returns {number} 使用可能な次の卓番号
 */
function getNextAvailableTableNumber(inProgressSheet) {
  const { indices, data } = getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);
  const maxTables = getMaxTables();
  const usedNumbers = new Set();

  // 現在使用中の卓番号を収集
  for (let i = 1; i < data.length; i++) {
    const tableNumber = data[i][indices["卓番号"]];
    if (tableNumber) {
      usedNumbers.add(tableNumber);
    }
  }

  // 1から順に空いている番号を探す
  for (let i = TABLE_CONFIG.MIN_TABLE_NUMBER; i <= maxTables; i++) {
    if (!usedNumbers.has(i)) {
      return i;
    }
  }

  throw new Error(`使用可能な卓番号がありません。最大${maxTables}卓まで設定可能です。`);
}

/**
 * プレイヤーが前回使用した卓番号を取得します
 * @param {string} playerId プレイヤーID
 * @returns {number|null} 卓番号。見つからない場合はnull
 */
function getLastTableNumber(playerId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName(SHEET_HISTORY);
  if (!historySheet) {
    Logger.log("エラー: 履歴シートが見つかりません。");
    return null;
  }
  const { indices: historyIndices, data: historyData } = getSheetStructure(historySheet, SHEET_HISTORY);

  // 最新の対戦履歴を探す
  for (let i = historyData.length - 1; i > 0; i--) {
    const row = historyData[i];
    const winnerId = row[historyIndices["ID1"]];
    const tableNumber = row[historyIndices["卓番号"]];
    const duration = row[historyIndices["対戦時間"]];

    // 対戦時間が空の行は未決着扱い（大会終了など）としてスキップ
    if (winnerId === playerId && duration) {
      return tableNumber;
    }
  }
  return null;
}

/**
 * 経過ミリ秒を HH:mm:ss 形式の文字列に変換する。
 */
function formatElapsedMs(elapsedMs) {
  if (!Number.isFinite(elapsedMs) || elapsedMs < 0) {
    return "00:00:00";
  }
  const totalSeconds = Math.floor(elapsedMs / 1000);
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;

  const hh = Utilities.formatString("%02d", hours);
  const mm = Utilities.formatString("%02d", minutes);
  const ss = Utilities.formatString("%02d", seconds);
  return `${hh}:${mm}:${ss}`;
}

// =========================================
// UI共通ユーティリティ
// =========================================

/**
 * プレイヤーIDの入力を受け付ける共通関数
 * @param {string} title - プロンプトのタイトル
 * @param {string} message - プロンプトのメッセージ
 * @returns {string|null} 整形されたプレイヤーID、キャンセル時はnull
 */
function promptPlayerId(title, message) {
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert("処理をキャンセルしました。");
    return null;
  }

  const rawId = response.getResponseText().trim();

  if (!/^\d+$/.test(rawId)) {
    ui.alert("エラー: IDは数字のみで入力してください。");
    return null;
  }

  return PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, parseInt(rawId, 10));
}

/**
 * プレイヤーの状態を変更する共通処理
 * @param {Object} config - 設定オブジェクト
 * @param {string} config.actionName - アクション名（例: "ドロップアウト"）
 * @param {string} config.promptMessage - 入力プロンプトのメッセージ
 * @param {string} config.confirmMessage - 確認ダイアログのメッセージ
 * @param {string} config.newStatus - 新しい状態
 */
function changePlayerStatus(config) {
  const ui = SpreadsheetApp.getUi();

  const playerId = promptPlayerId(config.actionName, config.promptMessage);
  if (!playerId) return;

  // プレイヤー名を取得
  const playerName = getPlayerName(playerId);

  const confirmResponse = ui.alert(
    config.actionName + "の確認",
    `プレイヤー名: ${playerName}\nプレイヤーID: ${playerId}\n\n` + config.confirmMessage + "\n\nよろしいですか？",
    ui.ButtonSet.YES_NO
  );

  if (confirmResponse !== ui.Button.YES) {
    ui.alert("処理をキャンセルしました。");
    return;
  }

  // 共通処理を呼び出し
  const result = updatePlayerState({
    targetPlayerId: playerId,
    newStatus: config.newStatus,
    opponentNewStatus: PLAYER_STATUS.WAITING,
    recordResult: false,
    isTargetWinner: false,
  });

  if (!result.success) {
    ui.alert("エラー", result.message, ui.ButtonSet.OK);
    return;
  }
}
