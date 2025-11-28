/**
 * ガンスリンガーバトル用マッチングシステム
 * @fileoverview 対戦ドメイン - マッチング管理と対戦結果の記録・修正
 * @author springOK
 */

// =========================================
// マッチング管理
// =========================================

/**
 * プレイヤーID→名前のマップを構築します
 * @param {Array[]} playerData プレイヤーシートのデータ
 * @param {Object} playerIndices 列インデックス
 * @returns {Map<string, string>} プレイヤーID→名前のMap
 */
function buildPlayerNameMap(playerData, playerIndices) {
  const playerNameMap = new Map();
  for (let i = 1; i < playerData.length; i++) {
    const row = playerData[i];
    const playerId = row[playerIndices["プレイヤーID"]];
    const playerName = row[playerIndices["プレイヤー名"]];
    if (!playerId) continue;
    playerNameMap.set(playerId, playerName);
  }
  return playerNameMap;
}

/**
 * 過去の対戦相手マップを構築します（再戦回避用）
 * @param {Array[]} historyData 対戦履歴シートのデータ
 * @param {Object} historyIndices 列インデックス
 * @returns {Map<string, Set<string>>} プレイヤーID→対戦相手IDのSetのMap
 */
function buildOpponentsMap(historyData, historyIndices) {
  const opponentsMap = new Map();
  const p1Col = historyIndices["ID1"];
  const p2Col = historyIndices["ID2"];

  for (let i = 1; i < historyData.length; i++) {
    const row = historyData[i];
    const p1 = row[p1Col];
    const p2 = row[p2Col];

    if (!p1 || !p2) continue;

    if (!opponentsMap.has(p1)) opponentsMap.set(p1, new Set());
    if (!opponentsMap.has(p2)) opponentsMap.set(p2, new Set());

    opponentsMap.get(p1).add(p2);
    opponentsMap.get(p2).add(p1);
  }
  return opponentsMap;
}

/**
 * 待機中プレイヤーを抽出してソートします
 * @param {Array[]} playerData プレイヤーシートのデータ
 * @param {Object} playerIndices 列インデックス
 * @returns {Array[]} ソート済み待機プレイヤーの配列
 */
function getAndSortWaitingPlayers(playerData, playerIndices) {
  return playerData
    .slice(1)
    .filter((row) => row[playerIndices["参加状況"]] === PLAYER_STATUS.WAITING)
    .sort((a, b) => {
      const winsDiff = b[playerIndices["勝数"]] - a[playerIndices["勝数"]];
      if (winsDiff !== 0) return winsDiff;

      const dateA = a[playerIndices["最終対戦時刻"]] instanceof Date ? a[playerIndices["最終対戦時刻"]].getTime() : 0;
      const dateB = b[playerIndices["最終対戦時刻"]] instanceof Date ? b[playerIndices["最終対戦時刻"]].getTime() : 0;
      return dateA - dateB;
    });
}

/**
 * 再戦回避を考慮してマッチングペアを決定します
 * @param {Array[]} waitingPlayers 待機プレイヤーの配列
 * @param {Object} playerIndices 列インデックス
 * @param {Map<string, Set<string>>} opponentsMap 過去対戦相手マップ
 * @returns {{matches: Array, skippedPlayers: Array}} マッチング結果
 */
function findMatchPairs(waitingPlayers, playerIndices, opponentsMap) {
  const matches = [];
  const availablePlayers = [...waitingPlayers];
  const skippedPlayers = [];

  while (availablePlayers.length >= 2) {
    const p1 = availablePlayers.splice(0, 1)[0];
    if (!p1) break;

    const p1Id = p1[playerIndices["プレイヤーID"]];
    const p1Opponents = opponentsMap.get(p1Id) || new Set();

    let p2Index = -1;
    for (let i = 0; i < availablePlayers.length; i++) {
      const p2Id = availablePlayers[i][playerIndices["プレイヤーID"]];
      if (!p1Opponents.has(p2Id)) {
        p2Index = i;
        break;
      }
    }

    if (p2Index !== -1) {
      const p2 = availablePlayers.splice(p2Index, 1)[0];
      const matchPair = [p1Id, p2[playerIndices["プレイヤーID"]]];
      matches.push(matchPair);
      Logger.log(`マッチング成立 (再戦なし): ${p1Id} vs ${p2[playerIndices["プレイヤーID"]]}`);
    } else {
      skippedPlayers.push(p1);
    }
  }

  skippedPlayers.push(...availablePlayers);
  return { matches, skippedPlayers };
}

/**
 * マッチングシートから卓の使用状況を取得します
 * @param {Array[]} matchData マッチングシートのデータ
 * @param {Object} matchIndices 列インデックス
 * @returns {{availableTables: Array, usedTables: Set<number>}} 卓の状態
 */
function getTableStatus(matchData, matchIndices) {
  const availableTables = [];
  const usedTables = new Set();

  for (let i = 1; i < matchData.length; i++) {
    const row = matchData[i];
    const tableNumber = row[matchIndices["卓番号"]];
    if (!tableNumber) continue;

    if (!row[matchIndices["ID1"]]) {
      availableTables.push({ row: i, tableNumber: tableNumber });
    } else {
      usedTables.add(tableNumber);
    }
  }

  return { availableTables, usedTables };
}

/**
 * マッチング結果をシートに書き込みます
 * @param {Object} params 書き込みパラメータ
 */
function writeMatchResults(params) {
  const { actualMatches, playerSheet, playerData, playerIndices, inProgressSheet, matchData, playerNameMap, availableTables, usedTables } = params;

  const playerIdsToUpdate = new Set(actualMatches.flat());

  // プレイヤー状態を更新
  for (let i = 1; i < playerData.length; i++) {
    const row = playerData[i];
    const playerId = row[playerIndices["プレイヤーID"]];
    if (playerIdsToUpdate.has(playerId)) {
      playerSheet.getRange(i + 1, playerIndices["参加状況"] + 1).setValue(PLAYER_STATUS.IN_PROGRESS);
    }
  }

  let nextNewRow = matchData.length;

  for (const match of actualMatches) {
    const [p1Id, p2Id] = match;
    let targetRow = null;
    let tableNumber = null;

    // 勝者の前回使用した卓を確認
    const lastTable = getLastTableNumber(p1Id);
    if (lastTable) {
      const validation = validateTableNumber(lastTable);
      if (validation.isValid && !usedTables.has(lastTable)) {
        const availableTableIndex = availableTables.findIndex((t) => t.tableNumber === lastTable);
        if (availableTableIndex !== -1) {
          const table = availableTables.splice(availableTableIndex, 1)[0];
          targetRow = table.row;
          tableNumber = table.tableNumber;
          usedTables.add(tableNumber);
        }
      }
    }

    if (targetRow === null) {
      if (availableTables.length > 0) {
        const table = availableTables.shift();
        targetRow = table?.row;
        tableNumber = table?.tableNumber;
      } else {
        const newTableNumber = getNextAvailableTableNumber(inProgressSheet);
        tableNumber = newTableNumber;
        targetRow = nextNewRow;
        nextNewRow++;
        inProgressSheet.getRange(targetRow + 1, 1).setValue(newTableNumber);
      }
      usedTables.add(tableNumber);
    }

    inProgressSheet
      .getRange(targetRow + 1, 2, 1, 6)
      .setValues([
        [
          p1Id,
          playerNameMap.get(p1Id) || p1Id,
          p2Id,
          playerNameMap.get(p2Id) || p2Id,
          Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss"),
          ,
        ],
      ]);
  }
}

/**
 * 待機中のプレイヤーを抽出し、再戦履歴を厳格に考慮してマッチングを行います。
 * 過去に対戦した相手しかいない場合、マッチングを成立させずに待機させます。
 */
function matchPlayers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // メンテナンスモードチェック
  try {
    const props = PropertiesService.getDocumentProperties();
    if (props.getProperty("MAINTENANCE_MODE") === "1") {
      Logger.log("matchPlayers: MAINTENANCE_MODE のためマッチングをスキップします。");
      return 0;
    }
  } catch (e) {
    Logger.log("matchPlayers: PropertiesService チェックでエラー: " + e?.toString());
  }

  let lock = null;

  try {
    lock = acquireLock("マッチング実行");

    // シート取得
    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    const historySheet = ss.getSheetByName(SHEET_HISTORY);
    const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
    if (!playerSheet || !historySheet || !inProgressSheet) {
      Logger.log("エラー: 必要なシートが見つかりません。");
      return 0;
    }

    // データ取得
    const { indices: playerIndices, data: playerData } = getSheetStructure(playerSheet, SHEET_PLAYERS);
    const { indices: historyIndices, data: historyData } = getSheetStructure(historySheet, SHEET_HISTORY);
    const { indices: matchIndices, data: matchData } = getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);

    // マップ構築
    const playerNameMap = buildPlayerNameMap(playerData, playerIndices);
    const opponentsMap = buildOpponentsMap(historyData, historyIndices);

    // 待機プレイヤー取得
    const waitingPlayers = getAndSortWaitingPlayers(playerData, playerIndices);
    if (waitingPlayers.length < 2) {
      Logger.log(`警告: 現在待機中のプレイヤーは ${waitingPlayers.length} 人です。2人以上必要です。`);
      return 0;
    }

    // マッチングペア決定
    Logger.log("--- 厳格な再戦回避マッチング開始 (勝者優先) ---");
    const { matches, skippedPlayers } = findMatchPairs(waitingPlayers, playerIndices, opponentsMap);

    if (skippedPlayers.length > 0) {
      Logger.log(`警告: ${skippedPlayers.length} 人のプレイヤーは適切な相手が見つからなかったため、待機を継続します。`);
    }

    if (matches.length === 0) {
      Logger.log("警告: 新しいマッチングは成立しませんでした。");
      return 0;
    }

    // 卓の状態を取得
    const { availableTables, usedTables } = getTableStatus(matchData, matchIndices);

    // 卓数制限の計算
    const maxTables = getMaxTables();
    const totalExistingTables = availableTables.length + usedTables.size;
    const maxNewTables = Math.max(0, maxTables - totalExistingTables);
    const totalAvailableSlots = availableTables.length + maxNewTables;

    Logger.log(
      `[デバッグ] 卓数情報: 最大=${maxTables}, 空き=${availableTables.length}, 使用中=${usedTables.size}, 新規作成可能=${maxNewTables}, 利用可能スロット=${totalAvailableSlots}`
    );

    // 卓数制限を適用
    const actualMatches = matches.slice(0, totalAvailableSlots);
    const skippedMatches = matches.slice(totalAvailableSlots);

    if (skippedMatches.length > 0) {
      Logger.log(`警告: 卓数上限により ${skippedMatches.length} 組のマッチングを見送りました。`);
      skippedMatches.forEach(([p1Id, p2Id]) => Logger.log(`見送り: ${p1Id} vs ${p2Id}`));
    }

    // シートに書き込み
    writeMatchResults({
      actualMatches,
      playerSheet,
      playerData,
      playerIndices,
      inProgressSheet,
      matchData,
      playerNameMap,
      availableTables,
      usedTables,
    });

    Logger.log(`マッチングが ${actualMatches.length} 件成立しました。「${SHEET_IN_PROGRESS}」シートを確認してください。`);
    return actualMatches.length;
  } catch (e) {
    Logger.log("matchPlayers エラー: " + e.message);
    return 0;
  } finally {
    releaseLock(lock);
  }
}

// =========================================
// 対戦結果記録
// =========================================

/**
 * カスタムメニューから実行するためのラッパー関数。
 */
function promptAndRecordResult() {
  const ui = SpreadsheetApp.getUi();

  // 最初にユーザー入力を受け付け
  const winnerResponse = ui.prompt("対戦結果の記録", `勝者のプレイヤーIDの**数字部分のみ**を入力してください (例: P001なら「1」)。`, ui.ButtonSet.OK_CANCEL);

  if (winnerResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert("処理をキャンセルしました。");
    return;
  }

  const rawId = winnerResponse.getResponseText().trim();

  if (!/^\d+$/.test(rawId)) {
    ui.alert("エラー: IDは数字のみで入力してください。");
    return;
  }

  const formattedWinnerId = PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, parseInt(rawId, 10));

  // この時点でロックは取得しない
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
  if (!inProgressSheet) {
    ui.alert("エラー: 対戦中シートが見つかりません。");
    return;
  }
  const { indices, data } = getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);
  let loserId = null;

  // 対戦相手の確認（ロック不要）
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const p1 = row[indices["ID1"]];
    const p2 = row[indices["ID2"]];

    if (p1 === formattedWinnerId) {
      loserId = p2;
      break;
    } else if (p2 === formattedWinnerId) {
      loserId = p1;
      break;
    }
  }

  if (loserId === null) {
    ui.alert(
      `エラー: 勝者ID (${formattedWinnerId}) は「${SHEET_IN_PROGRESS}」シートに見つかりませんでした。\n入力IDが間違っているか、対戦が記録されていません。`
    );
    return;
  }

  // ユーザーに確認
  const confirmResponse = ui.alert(
    "対戦結果の確認",
    `以下の内容で記録してよろしいですか？\n\n` + `勝者: ${getPlayerName(formattedWinnerId)}\n` + `敗者: ${getPlayerName(loserId)}`,
    ui.ButtonSet.YES_NO
  );

  if (confirmResponse !== ui.Button.YES) {
    ui.alert("処理をキャンセルしました。");
    return;
  }

  try {
    // recordResult内でロックを取得するため、ここではロック不要
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
  const ui = SpreadsheetApp.getUi();

  if (!winnerId) {
    ui.alert("勝者IDを入力してください。");
    return;
  }

  // 共通処理を呼び出し
  const result = updatePlayerState({
    targetPlayerId: winnerId,
    newStatus: PLAYER_STATUS.WAITING,
    opponentNewStatus: PLAYER_STATUS.WAITING,
    recordResult: true,
    isTargetWinner: true,
  });

  if (!result.success) {
    ui.alert("エラー", result.message, ui.ButtonSet.OK);
    return;
  }

  Logger.log(`対戦結果が記録されました。勝者: ${winnerId}, 敗者: ${result.opponentId}。両プレイヤーは待機状態に戻りました。`);
}

// =========================================
// 対戦結果修正系
// =========================================

/**
 * 対戦結果の勝敗を修正します。
 * 対戦IDを指定して、勝者と敗者を入れ替えます。
 * 両プレイヤーの統計情報（勝数・敗数）も自動的に調整されます。
 */
function correctMatchResult() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let lock = null;

  try {
    // 1. 対戦IDの入力
    const response = ui.prompt("対戦結果の修正", "修正する対戦IDの**数字部分のみ**を入力してください (例: T0001なら「1」)。", ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() !== ui.Button.OK) {
      ui.alert("処理をキャンセルしました。");
      return;
    }

    const rawId = response.getResponseText().trim();
    if (!/^\d+$/.test(rawId)) {
      ui.alert("エラー", "IDは数字のみで入力してください。", ui.ButtonSet.OK);
      return;
    }

    const matchId = "T" + Utilities.formatString("%04d", parseInt(rawId, 10));

    lock = acquireLock("対戦結果の修正");

    // 2. 対戦履歴から該当の対戦を検索
    const historySheet = ss.getSheetByName(SHEET_HISTORY);
    if (!historySheet) {
      ui.alert("エラー", `シート「${SHEET_HISTORY}」が見つかりません。`, ui.ButtonSet.OK);
      return;
    }
    const { indices: historyIndices, data: historyData } = getSheetStructure(historySheet, SHEET_HISTORY);

    let matchRow = -1;
    let matchData = null;

    for (let i = 1; i < historyData.length; i++) {
      const row = historyData[i];
      if (row[historyIndices["対戦ID"]] === matchId) {
        matchRow = i + 1;
        matchData = row;
        break;
      }
    }

    if (matchRow === -1) {
      ui.alert("エラー", `対戦ID ${matchId} が見つかりません。`, ui.ButtonSet.OK);
      return;
    }

    // 3. 現在の勝者と敗者を取得
    const currentWinnerId = matchData[historyIndices["ID1"]];
    const currentWinnerName = matchData[historyIndices["プレイヤー1"]];
    const currentLoserId = matchData[historyIndices["ID2"]];
    const currentLoserName = matchData[historyIndices["プレイヤー2"]];

    // 4. 修正の確認
    const confirmResponse = ui.alert(
      "勝敗修正の確認",
      `対戦ID: ${matchId}\n\n` +
        `【現在】\n` +
        `勝者: ${currentWinnerName} (${currentWinnerId})\n` +
        `敗者: ${currentLoserName} (${currentLoserId})\n\n` +
        `【修正後】\n` +
        `勝者: ${currentLoserName} (${currentLoserId})\n` +
        `敗者: ${currentWinnerName} (${currentWinnerId})\n\n` +
        "勝敗を入れ替えますか？",
      ui.ButtonSet.YES_NO
    );

    if (confirmResponse !== ui.Button.YES) {
      ui.alert("処理をキャンセルしました。");
      return;
    }

    // 5. 対戦履歴を更新（勝者と敗者を入れ替え）
    historySheet.getRange(matchRow, historyIndices["ID1"] + 1).setValue(currentLoserId);
    historySheet.getRange(matchRow, historyIndices["プレイヤー1"] + 1).setValue(currentLoserName);
    historySheet.getRange(matchRow, historyIndices["ID2"] + 1).setValue(currentWinnerId);
    historySheet.getRange(matchRow, historyIndices["プレイヤー2"] + 1).setValue(currentWinnerName);
    historySheet.getRange(matchRow, historyIndices["勝者名"] + 1).setValue(currentLoserName);

    // 6. プレイヤーの統計を修正
    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    if (!playerSheet) {
      ui.alert("エラー", `シート「${SHEET_PLAYERS}」が見つかりません。`, ui.ButtonSet.OK);
      return;
    }
    const { indices: playerIndices, data: playerData } = getSheetStructure(playerSheet, SHEET_PLAYERS);

    for (let i = 1; i < playerData.length; i++) {
      const row = playerData[i];
      const playerId = row[playerIndices["プレイヤーID"]];
      const rowNum = i + 1;

      if (playerId === currentWinnerId) {
        // 元の勝者: 勝数-1、敗数+1
        const currentWins = parseInt(row[playerIndices["勝数"]]) || 0;
        const currentLosses = parseInt(row[playerIndices["敗数"]]) || 0;
        playerSheet.getRange(rowNum, playerIndices["勝数"] + 1).setValue(Math.max(0, currentWins - 1));
        playerSheet.getRange(rowNum, playerIndices["敗数"] + 1).setValue(currentLosses + 1);
      } else if (playerId === currentLoserId) {
        // 元の敗者: 勝数+1、敗数-1
        const currentWins = parseInt(row[playerIndices["勝数"]]) || 0;
        const currentLosses = parseInt(row[playerIndices["敗数"]]) || 0;
        playerSheet.getRange(rowNum, playerIndices["勝数"] + 1).setValue(currentWins + 1);
        playerSheet.getRange(rowNum, playerIndices["敗数"] + 1).setValue(Math.max(0, currentLosses - 1));
      }
    }

    Logger.log(`対戦結果修正完了: ${matchId}, 新勝者: ${currentLoserId}, 新敗者: ${currentWinnerId}`);
  } catch (e) {
    ui.alert("エラーが発生しました: " + e.toString());
    Logger.log("correctMatchResult エラー: " + e.toString());
  } finally {
    releaseLock(lock);
  }
}

// =========================================
// 対戦時間更新
// =========================================

/**
 * 全卓の対戦時間を現在時刻に基づいて更新します。
 */
function updateAllMatchTimes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let updatedCount = 0;
  const currentTime = new Date();
  let matchLock = null;

  try {
    const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
    if (!inProgressSheet) {
      Logger.log(`エラー: シート「${SHEET_IN_PROGRESS}」が見つかりません。`);
      return;
    }
    const { indices, data } = getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);
    matchLock = acquireLock("対戦結果の記録");
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const startTime = new Date(row[indices["対戦開始時刻"]]);
      // 対戦開始時刻が存在する場合のみ更新
      if (startTime.toString() !== "Invalid Date") {
        const elapsedMs = currentTime.getTime() - startTime.getTime();
        const elapsedDate = new Date(elapsedMs);
        const formattedElapsed = Utilities.formatDate(elapsedDate, "UTC", "HH:mm:ss");
        inProgressSheet.getRange(i + 1, indices["経過時間"] + 1).setValue(formattedElapsed);
        updatedCount++;
      }
    }
    Logger.log(`対戦時間が ${updatedCount} 卓分更新されました。`);
  } finally {
    releaseLock(matchLock);
  }
}
