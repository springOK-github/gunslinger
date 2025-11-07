/**
 * ポケモンカード・ガンスリンガーバトル用マッチングシステム
 * @fileoverview 対戦ドメイン - マッチング管理と対戦結果の記録・修正
 * @author SpringOK
 */

// =========================================
// マッチング管理
// =========================================

/**
 * 待機中のプレイヤーを抽出し、再戦履歴を厳格に考慮してマッチングを行います。
 * 過去に対戦した相手しかいない場合、マッチングを成立させずに待機させます。
 */
function matchPlayers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
  let lock = null;

  try {
    lock = acquireLock('マッチング実行');
    getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);
    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    const { indices: playerIndices, data: playerData } = getSheetStructure(playerSheet, SHEET_PLAYERS);

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

    // マッチング結果の反映
    if (matches.length > 0) {
      // プレイヤーの状態を更新
      const playerIdsToUpdate = matches.flat();
      for (let i = 1; i < playerData.length; i++) {
        const row = playerData[i];
        const playerId = row[playerIndices["プレイヤーID"]];
        if (playerIdsToUpdate.includes(playerId)) {
          playerSheet.getRange(i + 1, playerIndices["参加状況"] + 1)
            .setValue(PLAYER_STATUS.IN_PROGRESS);
        }
      }

      // マッチングシートの現在の状態を取得
      const { indices: matchIndices, data: matchData } = getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);

      // 既存の卓の状態を確認
      const availableTables = [];
      const usedTables = new Set();

      for (let i = 1; i < matchData.length; i++) {
        const row = matchData[i];
        const tableNumber = row[matchIndices["卓番号"]];
        if (!tableNumber) continue;

        if (!row[matchIndices["ID1"]]) {
          // 空いている卓
          availableTables.push({ row: i, tableNumber: tableNumber });
        } else {
          usedTables.add(tableNumber);
        }
      }

      // マッチを卓に割り当て
      for (const match of matches) {
        const [p1Id, p2Id] = match;
        let targetRow = null;
        let tableNumber = null;

        // 勝者の前回使用した卓を確認
        const lastTable = getLastTableNumber(p1Id);
        if (lastTable) {
          const validation = validateTableNumber(lastTable);
          if (validation.isValid && !usedTables.has(lastTable)) {
            // 前回の卓が有効で空いている場合、その卓を使用
            const availableTableIndex = availableTables.findIndex(t => t.tableNumber === lastTable);
            if (availableTableIndex !== -1) {
              const table = availableTables.splice(availableTableIndex, 1)[0];
              targetRow = table.row;
              tableNumber = table.tableNumber;
              usedTables.add(tableNumber);
            }
          }
        }

        // 前回の卓が使えない場合は新しい卓を割り当て
        if (targetRow === null) {
          if (availableTables.length > 0) {
            // 既存の空き卓から割り当て
            const table = availableTables.shift();
            targetRow = table.row;
            tableNumber = table.tableNumber;
          } else {
            // 新しい卓を作成
            const newTableNumber = getNextAvailableTableNumber(inProgressSheet);
            tableNumber = newTableNumber;
            targetRow = matchData.length;
            // 新しい卓番号を設定
            inProgressSheet.getRange(targetRow + 1, 1).setValue(tableNumber);
          }
          usedTables.add(tableNumber);
        }

        // マッチを卓に割り当て
        inProgressSheet.getRange(targetRow + 1, 2, 1, 4).setValues([[
          p1Id,
          getPlayerName(p1Id),
          p2Id,
          getPlayerName(p2Id)
        ]]);
      }

      Logger.log(`マッチングが ${matches.length} 件成立しました。「${SHEET_IN_PROGRESS}」シートを確認してください。`);
      return matches.length;
    } else {
      Logger.log("警告: 新しいマッチングは成立しませんでした。");
      return 0;
    }

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
  const winnerResponse = ui.prompt(
    '対戦結果の記録',
    `勝者のプレイヤーIDの**数字部分のみ**を入力してください (例: P001なら「1」)。`,
    ui.ButtonSet.OK_CANCEL
  );

  if (winnerResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert('処理をキャンセルしました。');
    return;
  }

  const rawId = winnerResponse.getResponseText().trim();

  if (!/^\d+$/.test(rawId)) {
    ui.alert('エラー: IDは数字のみで入力してください。');
    return;
  }

  const formattedWinnerId = PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, parseInt(rawId, 10));

  // この時点でロックは取得しない
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
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
    ui.alert(`エラー: 勝者ID (${formattedWinnerId}) は「${SHEET_IN_PROGRESS}」シートに見つかりませんでした。\n入力IDが間違っているか、対戦が記録されていません。`);
    return;
  }

  // ユーザーに確認
  const confirmResponse = ui.alert(
    '対戦結果の確認',
    `以下の内容で記録してよろしいですか？\n\n` +
    `勝者: ${getPlayerName(formattedWinnerId)}\n` +
    `敗者: ${getPlayerName(loserId)}`,
    ui.ButtonSet.YES_NO
  );

  if (confirmResponse !== ui.Button.YES) {
    ui.alert('処理をキャンセルしました。');
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
    isTargetWinner: true
  });

  if (!result.success) {
    ui.alert('エラー', result.message, ui.ButtonSet.OK);
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
    const response = ui.prompt(
      '対戦結果の修正',
      '修正する対戦IDの**数字部分のみ**を入力してください (例: T0001なら「1」)。',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) {
      ui.alert('処理をキャンセルしました。');
      return;
    }

    const rawId = response.getResponseText().trim();
    if (!/^\d+$/.test(rawId)) {
      ui.alert('エラー', 'IDは数字のみで入力してください。', ui.ButtonSet.OK);
      return;
    }

    const matchId = "T" + Utilities.formatString("%04d", parseInt(rawId, 10));

    lock = acquireLock('対戦結果の修正');

    // 2. 対戦履歴から該当の対戦を検索
    const historySheet = ss.getSheetByName(SHEET_HISTORY);
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
      ui.alert('エラー', `対戦ID ${matchId} が見つかりません。`, ui.ButtonSet.OK);
      return;
    }

    // 3. 現在の勝者と敗者を取得
    const currentWinnerId = matchData[historyIndices["ID1"]];
    const currentWinnerName = matchData[historyIndices["プレイヤー1"]];
    const currentLoserId = matchData[historyIndices["ID2"]];
    const currentLoserName = matchData[historyIndices["プレイヤー2"]];

    // 4. 修正の確認
    const confirmResponse = ui.alert(
      '勝敗修正の確認',
      `対戦ID: ${matchId}\n\n` +
      `【現在】\n` +
      `勝者: ${currentWinnerName} (${currentWinnerId})\n` +
      `敗者: ${currentLoserName} (${currentLoserId})\n\n` +
      `【修正後】\n` +
      `勝者: ${currentLoserName} (${currentLoserId})\n` +
      `敗者: ${currentWinnerName} (${currentWinnerId})\n\n` +
      '勝敗を入れ替えますか？',
      ui.ButtonSet.YES_NO
    );

    if (confirmResponse !== ui.Button.YES) {
      ui.alert('処理をキャンセルしました。');
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

    ui.alert(
      '修正完了',
      `対戦ID ${matchId} の勝敗を修正しました。\n\n` +
      `新しい勝者: ${currentLoserName} (${currentLoserId})\n` +
      `新しい敗者: ${currentWinnerName} (${currentWinnerId})`,
      ui.ButtonSet.OK
    );

    Logger.log(`対戦結果修正完了: ${matchId}, 新勝者: ${currentLoserId}, 新敗者: ${currentWinnerId}`);

  } catch (e) {
    ui.alert("エラーが発生しました: " + e.toString());
    Logger.log("correctMatchResult エラー: " + e.toString());
  } finally {
    releaseLock(lock);
  }
}
