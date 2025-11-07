/**
 * ポケモンカード・ガンスリンガーバトル用マッチングシステム
 * @fileoverview 共有ユーティリティ - シート操作とUI共通処理
 * @author SpringOK
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

  const headers = data[0].map(h => String(h).trim());
  const indices = {};
  const missing = [];

  const requiredHeaders = REQUIRED_HEADERS[sheetName];
  if (!requiredHeaders) {
    throw new Error(`シート「${sheetName}」の必須ヘッダー定義が見つかりません。`);
  }

  requiredHeaders.forEach(required => {
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
  const { indices, data } = getSheetStructure(playerSheet, SHEET_PLAYERS);

  for (let i = 1; i < data.length; i++) {
    if (data[i][indices["プレイヤーID"]] === playerId) {
      return data[i][indices["プレイヤー名"]] || playerId;
    }
  }
  return playerId;
}

/**
 * 「マッチング」シート内の空行の処理。
 * 卓番号制の導入により、空行は削除せず維持します。
 */
function cleanUpInProgressSheet() {
  // 卓番号制の導入により、意図的に何もしない
  Logger.log("卓番号制により、マッチングシートの行は維持されます。");
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
  const { indices: historyIndices, data: historyData } = getSheetStructure(historySheet, SHEET_HISTORY);

  // 最新の対戦履歴を探す
  for (let i = historyData.length - 1; i > 0; i--) {
    const row = historyData[i];
    const id1 = row[historyIndices["ID1"]];
    const id2 = row[historyIndices["ID2"]];
    const winner = row[historyIndices["勝者名"]];
    const tableNumber = row[historyIndices["卓番号"]];

    // 勝者のプレイヤー名から勝者IDを特定
    if (getPlayerName(playerId) === winner && (id1 === playerId || id2 === playerId)) {
      return tableNumber;
    }
  }
  return null;
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
    ui.alert('処理をキャンセルしました。');
    return null;
  }

  const rawId = response.getResponseText().trim();

  if (!/^\d+$/.test(rawId)) {
    ui.alert('エラー: IDは数字のみで入力してください。');
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

  const confirmResponse = ui.alert(
    config.actionName + 'の確認',
    `プレイヤー ${playerId} \n` + config.confirmMessage + '\n\nよろしいですか？',
    ui.ButtonSet.YES_NO
  );

  if (confirmResponse !== ui.Button.YES) {
    ui.alert('処理をキャンセルしました。');
    return;
  }

  // 共通処理を呼び出し
  const result = updatePlayerState({
    targetPlayerId: playerId,
    newStatus: config.newStatus,
    opponentNewStatus: PLAYER_STATUS.WAITING,
    recordResult: false
  });

  if (!result.success) {
    ui.alert('エラー', result.message, ui.ButtonSet.OK);
    return;
  }

}
