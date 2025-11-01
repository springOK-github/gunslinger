/**
 * ポケモンカード・ガンスリンガーバトル用マッチングシステム
 * Google Apps Script (GAS) とスプレッドシートで動作します。
 *
 * プレイヤーの状態遷移:
 * 1. 待機: マッチング待ち
 * 2. 対戦中: 対戦進行中
 * 3. 終了: ドロップアウト済み
 *
 * 機能一覧:
 * - プレイヤー登録・管理
 * - 自動マッチング（再戦回避）
 * - 対戦結果記録
 * - プレイヤードロップアウト処理
 */

// --- 定数定義 ---
const SHEET_PLAYERS = "プレイヤー";
const SHEET_HISTORY = "対戦履歴";
const SHEET_IN_PROGRESS = "対戦中";
const PLAYER_ID_PREFIX = "P";
const ID_DIGITS = 3; // IDの数字部分の桁数 (例: P001なら3)
const REQUIRED_HEADERS = {
  [SHEET_PLAYERS]: ["プレイヤーID", "勝数", "敗数", "消化試合数", "参加状況", "最終対戦日時"],
  [SHEET_HISTORY]: ["日時", "プレイヤー1 ID", "プレイヤー2 ID", "勝者ID", "対戦ID"],
  [SHEET_IN_PROGRESS]: ["プレイヤー1 ID", "プレイヤー2 ID"]
};

// ----------------------------------------------------------------------
// --- 初期設定・メニュー関連 ---
// ----------------------------------------------------------------------

/**
 * シートのヘッダーを検証し、列インデックスを返します。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 検証対象のシート
 * @param {string} sheetName - シート名（SHEET_PLAYERS等の定数）
 * @returns {{headers: string[], indices: Object.<string, number>, data: any[][]}} ヘッダー情報と全データ
 * @throws {Error} 必須ヘッダーが不足している場合
 */
function validateHeaders(sheet, sheetName) {
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
  playerSheet.getRange("A1:F1").setValues([ // F列まで拡張
    ["プレイヤーID", "勝数", "敗数", "消化試合数", "参加状況", "最終対戦日時"]
  ]).setFontWeight("bold").setBackground("#c9daf8").setHorizontalAlignment("center");
  // 幅の調整
  playerSheet.setColumnWidth(1, 100);
  playerSheet.setColumnWidth(5, 100);
  playerSheet.setColumnWidth(6, 150); // 最終対戦日時

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



/**
 * 待機中のプレイヤーを抽出し、再戦履歴を厳格に考慮してマッチングを行います。
 * 過去に対戦した相手しかいない場合、マッチングを成立させずに待機させます。
 */
function matchPlayers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);

  try {
    // シートヘッダーの検証
    validateHeaders(inProgressSheet, SHEET_IN_PROGRESS);
    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    const { indices: playerIndices, data: playerData } = validateHeaders(playerSheet, SHEET_PLAYERS);

    // 1. 待機プレイヤーの取得とマッチング
    const waitingPlayers = getWaitingPlayers();
    if (waitingPlayers.length < 2) {
      Logger.log(`警告: 現在待機中のプレイヤーは ${waitingPlayers.length} 人です。2人以上必要です。`);
      return;
    }

    let matches = [];
    let availablePlayers = [...waitingPlayers];
    let skippedPlayers = [];

    Logger.log("--- 厳格な再戦回避マッチング開始 (勝者優先) ---");
    while (availablePlayers.length >= 2) {
      const p1 = availablePlayers.shift();
      const p1Id = p1[playerIndices["プレイヤーID"]];
      const p1BlackList = getPastOpponents(p1Id);

      let p2Index = -1;
      for (let i = 0; i < availablePlayers.length; i++) {
        const p2Id = availablePlayers[i][playerIndices["プレイヤーID"]];
        if (!p1BlackList.includes(p2Id)) {
          p2Index = i;
          break;
        }
      }

      if (p2Index !== -1) {
        const p2 = availablePlayers.splice(p2Index, 1)[0];
        matches.push([p1Id, p2[playerIndices["プレイヤーID"]]]);
        Logger.log(`マッチング成立 (再戦なし): ${p1Id} vs ${p2[playerIndices["プレイヤーID"]]}`);
      } else {
        skippedPlayers.push(p1);
      }
    }

    skippedPlayers.push(...availablePlayers);

    if (skippedPlayers.length > 0) {
      Logger.log(`警告: ${skippedPlayers.length} 人のプレイヤーは適切な相手が見つからなかったため、待機を継続します。`);
    }

    // 2. マッチング結果の反映
    if (matches.length > 0) {
      const playerIdsToUpdate = matches.flat();
      
      for (let i = 1; i < playerData.length; i++) {
        const row = playerData[i];
        const playerId = row[playerIndices["プレイヤーID"]];
        if (playerIdsToUpdate.includes(playerId)) {
          playerSheet.getRange(i + 1, playerIndices["参加状況"] + 1)
            .setValue("対戦中");
        }
      }

      const lastRow = inProgressSheet.getLastRow();
      if (matches.length > 0) {
        inProgressSheet.getRange(lastRow + 1, 1, matches.length, 2)
          .setValues(matches);
      }

      Logger.log(`マッチングが ${matches.length} 件成立しました。「対戦中」シートを確認してください。`);
      return matches.length;
    } else {
      Logger.log("警告: 新しいマッチングは成立しませんでした。");
      return 0;
    }

  } catch (e) {
    Logger.log("matchPlayers エラー: " + e.message);
    return 0;
  }
}

/**
 * カスタムメニューから実行するためのラッパー関数。
 */
function promptAndRecordResult() {
  const ui = SpreadsheetApp.getUi();

  // 勝者IDの数字部分を尋ねる (入力が必要なため維持)
  const winnerResponse = ui.prompt(
    '対戦結果の記録',
    '勝者のプレイヤーIDの**数字部分のみ**を入力してください (例: P001なら「1」)。\n敗者は「対戦中」シートから自動特定されます。',
    ui.ButtonSet.OK_CANCEL
  );

  if (winnerResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert('処理をキャンセルしました。');
    return;
  }

  const rawId = winnerResponse.getResponseText().trim();

  // 数字入力チェックとP00X形式への変換
  if (!/^\d+$/.test(rawId)) {
    ui.alert('エラー: IDは数字のみで入力してください。');
    return;
  }

  // 自動でプレフィックスとゼロパディングを付与
  const formattedWinnerId = PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, parseInt(rawId, 10));

  // 対戦中シートから敗者IDを特定して確認
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
  
  try {
    const { indices, data } = validateHeaders(inProgressSheet, SHEET_IN_PROGRESS);
    let loserId = null;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const p1 = row[indices["プレイヤー1 ID"]];
      const p2 = row[indices["プレイヤー2 ID"]];

      if (p1 === formattedWinnerId) {
        loserId = p2;
        break;
      } else if (p2 === formattedWinnerId) {
        loserId = p1;
        break;
      }
    }

    if (loserId === null) {
      ui.alert(`エラー: 勝者ID (${formattedWinnerId}) は「対戦中」シートに見つかりませんでした。\n入力IDが間違っているか、対戦が記録されていません。`);
      return;
    }

    // 勝者と敗者の確認
    const confirmResponse = ui.alert(
      '対戦結果の確認',
      `以下の内容で記録してよろしいですか？\n\n` +
      `勝者: ${formattedWinnerId}\n` +
      `敗者: ${loserId}`,
      ui.ButtonSet.YES_NO
    );

    if (confirmResponse !== ui.Button.YES) {
      ui.alert('処理をキャンセルしました。');
      return;
    }

    recordResult(formattedWinnerId);
  } catch (e) {
    ui.alert("エラーが発生しました: " + e.toString());
    Logger.log("promptAndRecordResult エラー: " + e.toString());
  }
}


/**
 * 対戦結果を記録し、プレイヤーの統計情報とステータスを更新し、自動で次をマッチングします。
 */
function recordResult(winnerId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  if (!winnerId) {
    ui.alert("勝者IDを入力してください。");
    return;
  }

  try {
    // 1. 対戦中シートの検証と敗者ID特定
    const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
    const { indices: inProgressIndices, data: inProgressData } = 
      validateHeaders(inProgressSheet, SHEET_IN_PROGRESS);

    let loserId = null;
    let rowToClear = -1;

    for (let i = 1; i < inProgressData.length; i++) {
      const row = inProgressData[i];
      const p1 = row[inProgressIndices["プレイヤー1 ID"]];
      const p2 = row[inProgressIndices["プレイヤー2 ID"]];

      if (p1 === winnerId) {
        loserId = p2;
        rowToClear = i + 1;
        break;
      } else if (p2 === winnerId) {
        loserId = p1;
        rowToClear = i + 1;
        break;
      }
    }

    if (loserId === null) {
      ui.alert(`エラー: 勝者ID (${winnerId}) は「対戦中」シートに見つかりませんでした。\n入力IDが間違っているか、対戦が記録されていません。`);
      return;
    }

    const currentTime = new Date();

    // 2. 対戦履歴シートの検証と記録
    const historySheet = ss.getSheetByName(SHEET_HISTORY);
    validateHeaders(historySheet, SHEET_HISTORY);
    const newId = "T" + Utilities.formatString("%04d", historySheet.getLastRow());

    historySheet.appendRow([
      currentTime,
      winnerId,
      loserId,
      winnerId,
      newId
    ]);

    // 3. プレイヤー統計更新
    updatePlayerStats(winnerId, true, currentTime);
    updatePlayerStats(loserId, false, currentTime);

    // 4. 対戦中シートのクリア
    if (rowToClear !== -1) {
      inProgressSheet.getRange(rowToClear, 1, 1, 2).clearContent();
    }

    // 5. プレイヤーシートの検証と参加状況更新
    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    const { indices: playerIndices, data: playerData } = 
      validateHeaders(playerSheet, SHEET_PLAYERS);

    for (let i = 1; i < playerData.length; i++) {
      const row = playerData[i];
      const playerId = row[playerIndices["プレイヤーID"]];
      if (playerId === winnerId || playerId === loserId) {
        playerSheet.getRange(i + 1, playerIndices["参加状況"] + 1)
          .setValue("待機");
      }
    }

    Logger.log(`対戦結果が記録されました。勝者: ${winnerId}, 敗者: ${loserId}。両プレイヤーは待機状態に戻りました。`);

    // 6. 対戦中シートを整理
    cleanUpInProgressSheet();

    // 7. 待機プレイヤーが2人以上いれば自動マッチング
    const waitingPlayersCount = getWaitingPlayers().length;
    if (waitingPlayersCount >= 2) {
      Logger.log(`待機プレイヤーが ${waitingPlayersCount} 人いるため、自動でマッチングを開始します。`);
      matchPlayers();
    } else {
      Logger.log(`待機プレイヤーが ${waitingPlayersCount} 人です。自動マッチングはスキップされました。`);
    }

  } catch (e) {
    ui.alert("エラーが発生しました: " + e.toString());
    Logger.log("recordResult エラー: " + e.toString());
  }
}


// ----------------------------------------------------------------------
// --- メイン機能（マッチング・結果記録・ドロップアウト） ---
// ----------------------------------------------------------------------


// ----------------------------------------------------------------------
// --- シート操作ヘルパー関数 ---
// ----------------------------------------------------------------------

/**
 * シートのヘッダーを検証し、列インデックスを返します。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 検証対象のシート
 * @param {string} sheetName - シート名（SHEET_PLAYERS等の定数）
 * @returns {{headers: string[], indices: Object.<string, number>, data: any[][]}} ヘッダー情報と全データ
 * @throws {Error} 必須ヘッダーが不足している場合
 */
function validateHeaders(sheet, sheetName) {
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
 * 「対戦中」シート内の空行（対戦が終了し、コンテンツがクリアされた行）を削除し、
 * シート内のデータを上詰めして整理します。
 */
function cleanUpInProgressSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);

  try {
    validateHeaders(inProgressSheet, SHEET_IN_PROGRESS);

    const lastRow = inProgressSheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log("「対戦中」シートにデータがないため、整理は不要です。");
      return;
    }

    let deletedCount = 0;
    for (let i = lastRow; i >= 2; i--) {
      const cellA = inProgressSheet.getRange(i, 1).getValue();
      if (cellA === "") {
        inProgressSheet.deleteRow(i);
        deletedCount++;
      }
    }

    if (deletedCount > 0) {
      Logger.log(`対戦中シートの整理 (自動実行) が完了しました。${deletedCount} 行の空行を削除しました。`);
    }
  } catch (e) {
    Logger.log("cleanUpInProgressSheet エラー: " + e.message);
  }
}


// ----------------------------------------------------------------------
// --- プレイヤー管理ヘルパー関数 ---
// ----------------------------------------------------------------------

/**
 * 待機中のプレイヤーを抽出し、以下の優先順位でソートして返します。
 * 1. 勝数（降順）
 * 2. 最終対戦日時（降順 = 最近待機に戻った人優先 = 直近の勝者優先）
 */
function getWaitingPlayers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);

  try {
    const { indices, data } = validateHeaders(playerSheet, SHEET_PLAYERS);
    if (data.length <= 1) return [];

    const waiting = data.slice(1).filter(row => 
      row[indices["参加状況"]] === "待機" && row[indices["参加状況"]] !== "終了"
    );

    waiting.sort((a, b) => {
      const winsDiff = b[indices["勝数"]] - a[indices["勝数"]];
      if (winsDiff !== 0) return winsDiff;

      const dateA = a[indices["最終対戦日時"]] instanceof Date ? a[indices["最終対戦日時"]].getTime() : 0;
      const dateB = b[indices["最終対戦日時"]] instanceof Date ? b[indices["最終対戦日時"]].getTime() : 0;
      return dateB - dateA;
    });

    return waiting;
  } catch (e) {
    Logger.log("getWaitingPlayers エラー: " + e.message);
    return [];
  }
}

/**
 * 特定プレイヤーの過去の対戦相手のIDリスト（ブラックリスト）を取得します。
 */
function getPastOpponents(playerId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName(SHEET_HISTORY);

  try {
    const { indices, data } = validateHeaders(historySheet, SHEET_HISTORY);
    if (data.length <= 1) return [];

    const p1Col = indices["プレイヤー1 ID"];
    const p2Col = indices["プレイヤー2 ID"];
    const opponents = new Set();

    data.slice(1).forEach(row => {
      if (row[p1Col] === playerId) {
        opponents.add(row[p2Col]);
      } else if (row[p2Col] === playerId) {
        opponents.add(row[p1Col]);
      }
    });

    return Array.from(opponents);
  } catch (e) {
    Logger.log("getPastOpponents エラー: " + e.message);
    return [];
  }
}

/**
 * プレイヤーの統計情報 (勝数, 敗数, 消化試合数) と最終対戦日時を更新します。
 */
function updatePlayerStats(playerId, isWinner, timestamp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);

  try {
    const { indices, data } = validateHeaders(playerSheet, SHEET_PLAYERS);
    if (data.length <= 1) return;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[indices["プレイヤーID"]] === playerId) {
        const rowNum = i + 1;
        const currentWins = parseInt(row[indices["勝数"]]) || 0;
        const currentLosses = parseInt(row[indices["敗数"]]) || 0;
        const currentTotal = parseInt(row[indices["消化試合数"]]) || 0;

        playerSheet.getRange(rowNum, indices["勝数"] + 1)
          .setValue(currentWins + (isWinner ? 1 : 0));
        playerSheet.getRange(rowNum, indices["敗数"] + 1)
          .setValue(currentLosses + (isWinner ? 0 : 1));
        playerSheet.getRange(rowNum, indices["消化試合数"] + 1)
          .setValue(currentTotal + 1);
        playerSheet.getRange(rowNum, indices["最終対戦日時"] + 1)
          .setValue(timestamp);

        return;
      }
    }
    Logger.log(`エラー: プレイヤーID ${playerId} が見つかりません。`);
  } catch (e) {
    Logger.log("updatePlayerStats エラー: " + e.message);
  }
}


// ----------------------------------------------------------------------
// --- テスト・管理用関数 ---
// ----------------------------------------------------------------------

/**
 * プレイヤーのドロップアウトを処理します。
 * 指定されたプレイヤーの参加状況を「終了」に変更し、進行中の対戦を無効にします。
 */
function handleDropout() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // プレイヤーIDの数字部分を尋ねる
  const response = ui.prompt(
    'プレイヤーのドロップアウト',
    'ドロップアウトするプレイヤーIDの**数字部分のみ**を入力してください (例: P001なら「1」)。',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert('処理をキャンセルしました。');
    return;
  }

  const rawId = response.getResponseText().trim();

  // 数字入力チェックとP00X形式への変換
  if (!/^\d+$/.test(rawId)) {
    ui.alert('エラー: IDは数字のみで入力してください。');
    return;
  }

  const playerId = PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, parseInt(rawId, 10));

  // 確認ダイアログ
  const confirmResponse = ui.alert(
    'ドロップアウトの確認',
    `プレイヤー ${playerId} をドロップアウトさせ、参加状況を「終了」に変更します。\n` +
    '進行中の対戦がある場合は無効となります。\n\n' +
    'よろしいですか？',
    ui.ButtonSet.YES_NO
  );

  if (confirmResponse !== ui.Button.YES) {
    ui.alert('処理をキャンセルしました。');
    return;
  }

  try {
    // プレイヤーシートの更新
    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    const { indices: playerIndices, data: playerData } = validateHeaders(playerSheet, SHEET_PLAYERS);
    
    let playerFound = false;
    for (let i = 1; i < playerData.length; i++) {
      const row = playerData[i];
      if (row[playerIndices["プレイヤーID"]] === playerId) {
        playerSheet.getRange(i + 1, playerIndices["参加状況"] + 1).setValue("終了");
        playerFound = true;
        break;
      }
    }

    if (!playerFound) {
      ui.alert(`エラー: プレイヤーID ${playerId} が見つかりません。`);
      return;
    }

    // 対戦中シートのチェックと更新
    const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
    const { indices: inProgressIndices, data: inProgressData } = validateHeaders(inProgressSheet, SHEET_IN_PROGRESS);

    let matchCancelled = false;
    let opponentId = null;
    let matchRow = -1;

    for (let i = 1; i < inProgressData.length; i++) {
      const row = inProgressData[i];
      if (row[inProgressIndices["プレイヤー1 ID"]] === playerId) {
        opponentId = row[inProgressIndices["プレイヤー2 ID"]];
        matchRow = i + 1;
        break;
      } else if (row[inProgressIndices["プレイヤー2 ID"]] === playerId) {
        opponentId = row[inProgressIndices["プレイヤー1 ID"]];
        matchRow = i + 1;
        break;
      }
    }

    if (matchRow !== -1) {
      // 対戦中の試合を無効化
      inProgressSheet.getRange(matchRow, 1, 1, 2).clearContent();
      matchCancelled = true;

      // 対戦相手を待機状態に戻す
      for (let i = 1; i < playerData.length; i++) {
        const row = playerData[i];
        if (row[playerIndices["プレイヤーID"]] === opponentId) {
          playerSheet.getRange(i + 1, playerIndices["参加状況"] + 1).setValue("待機");
          break;
        }
      }
    }

    // 結果の通知
    let message = `プレイヤー ${playerId} のドロップアウトを処理しました。\n参加状況を「終了」に変更しました。`;
    if (matchCancelled) {
      message += `\n\n進行中の対戦を無効とし、対戦相手（${opponentId}）を待機状態に戻しました。`;
    }
    ui.alert('完了', message, ui.ButtonSet.OK);

    // 対戦中シートの整理
    cleanUpInProgressSheet();

    // 新しいマッチングの実行
    const waitingPlayersCount = getWaitingPlayers().length;
    if (waitingPlayersCount >= 2) {
      matchPlayers();
    }

  } catch (e) {
    ui.alert("エラーが発生しました: " + e.toString());
    Logger.log("handleDropout エラー: " + e.toString());
  }
}

/**
 * 新しいプレイヤーを登録します。（本番・運営用）
 * 実行すると、次のID（例: P009）が自動で採番され、シートに追加されます。
 */
function registerPlayer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
  const ui = SpreadsheetApp.getUi();

  try {
    validateHeaders(playerSheet, SHEET_PLAYERS);

    const lastRow = playerSheet.getLastRow();
    const newIdNumber = lastRow;
    const newId = PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, newIdNumber);
    const currentTime = new Date();

    playerSheet.appendRow([newId, 0, 0, 0, "待機", currentTime]);
    Logger.log(`プレイヤー ${newId} を登録しました。`);

    const waitingPlayersCount = getWaitingPlayers().length;
    if (waitingPlayersCount >= 2) {
      Logger.log(`プレイヤー登録後、待機プレイヤーが ${waitingPlayersCount} 人いるため、自動でマッチングを開始します。`);
      matchPlayers();
    } else {
      Logger.log(`プレイヤー登録後、待機プレイヤーが ${waitingPlayersCount} 人です。自動マッチングはスキップされました。`);
    }
  } catch (e) {
    ui.alert("エラーが発生しました: " + e.toString());
    Logger.log("registerPlayer エラー: " + e.toString());
  }
}

/**
 * テスト用のプレイヤーを一括登録します。
 */
function registerTestPlayers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);

  // シートをクリア
  if (playerSheet.getLastRow() > 1) {
    playerSheet.getRange(2, 1, playerSheet.getLastRow() - 1, playerSheet.getLastColumn()).clearContent();
  }

  // P001からP008まで、8人分を登録
  const numTestPlayers = 8;
  for (let i = 0; i < numTestPlayers; i++) {
    // P001, P002, ... P008を直接登録
    const newIdNumber = i + 1;
    const newId = PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, newIdNumber);
    // 最終対戦日時を初期化時も設定
    playerSheet.appendRow([newId, 0, 0, 0, "待機", new Date()]);
  }

  // 最終的にテストプレイヤーが揃った後に、一度マッチングを試みる
  const waitingPlayersCount = getWaitingPlayers().length;
  if (waitingPlayersCount >= 2) {
    Logger.log("テストプレイヤー登録完了。自動で初回マッチングを開始します。");
    matchPlayers();
  } else {
    Logger.log("テストプレイヤーの登録が完了しました。マッチングには2人以上が必要です。");
  }
}
