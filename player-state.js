/**
 * ポケモンカード・ガンスリンガーバトル用マッチングシステム
 * @fileoverview プレイヤーの状態遷移に関する共通処理
 * @author SpringOK
 */

/**
 * 対戦中のプレイヤーの状態を変更します。
 * @param {Object} options - 状態変更のオプション
 * @param {string} options.targetPlayerId - 状態を変更するプレイヤーのID
 * @param {string} options.newStatus - 対象プレイヤーの新しい状態
 * @param {string} options.opponentNewStatus - 対戦相手の新しい状態
 * @param {boolean} options.recordResult - 結果を記録するかどうか
 * @param {boolean} options.isTargetWinner - 対象プレイヤーが勝者かどうか（結果記録時のみ使用）
 * @returns {Object} 処理結果 { success: boolean, message: string, opponentId?: string }
 */
function updatePlayerState(options) {
  const {
    targetPlayerId,
    newStatus,
    opponentNewStatus,
    recordResult = false,
    isTargetWinner = false
  } = options;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let matchLock = null;
  let stateLock = null;

  try {
    // 両方のロックを取得（順序を固定して デッドロック防止）
    stateLock = acquireLock('プレイヤー状態変更');
    matchLock = acquireLock('対戦結果の記録');

    // 1. プレイヤーの現在の状態を確認
    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    const { indices: playerIndices, data: playerData } = getSheetStructure(playerSheet, SHEET_PLAYERS);
    
    let targetFound = false;
    let currentStatus = null;

    for (let i = 1; i < playerData.length; i++) {
      const row = playerData[i];
      const playerId = row[playerIndices["プレイヤーID"]];
      
      if (playerId === targetPlayerId) {
        targetFound = true;
        currentStatus = row[playerIndices["参加状況"]];
        break;
      }
    }

    if (!targetFound) {
      return { success: false, message: `プレイヤー ${targetPlayerId} が見つかりません。` };
    }

    if (currentStatus === PLAYER_STATUS.DROPPED) {
      return { success: false, message: `プレイヤー ${targetPlayerId} はすでにドロップアウトしています。` };
    }

    // 3. 対戦中の場合のみ、対戦相手の処理
    let opponentId = null;
    let matchRow = -1;

    if (currentStatus === PLAYER_STATUS.IN_PROGRESS) {
      const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
      const { indices: matchIndices, data: matchData } = getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);

      for (let i = 1; i < matchData.length; i++) {
        const row = matchData[i];
        const p1 = row[matchIndices["ID1"]];
        const p2 = row[matchIndices["ID2"]];

        if (p1 === targetPlayerId) {
          opponentId = p2;
          matchRow = i + 1;
          break;
        } else if (p2 === targetPlayerId) {
          opponentId = p1;
          matchRow = i + 1;
          break;
        }
      }

      if (!opponentId) {
        return { success: false, message: `データ不整合: 対戦中のはずのプレイヤーID ${targetPlayerId} の対戦相手が見つかりません。` };
      }

      // 対戦相手の状態確認
      let opponentDropped = false;
      for (let i = 1; i < playerData.length; i++) {
        const row = playerData[i];
        if (row[playerIndices["プレイヤーID"]] === opponentId) {
          opponentDropped = row[playerIndices["参加状況"]] === PLAYER_STATUS.DROPPED;
          break;
        }
      }

      if (opponentDropped && opponentNewStatus !== PLAYER_STATUS.DROPPED) {
        return { success: false, message: `対戦相手はすでにドロップアウトしています。` };
      }
    }

    // 4. 結果の記録（必要な場合）
    if (recordResult) {
      const currentTime = new Date();
      const formattedTime = Utilities.formatDate(currentTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
      const historySheet = ss.getSheetByName(SHEET_HISTORY);
      getSheetStructure(historySheet, SHEET_HISTORY);
      const newId = "T" + Utilities.formatString("%04d", historySheet.getLastRow());

      const winner = isTargetWinner ? targetPlayerId : opponentId;
      const loser = isTargetWinner ? opponentId : targetPlayerId;
      const winnerName = getPlayerName(winner);
      const loserName = getPlayerName(loser);

      // マッチング中の卓番号を取得
      const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
      const { indices: matchIndices } = getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);
      const matchTableNumber = inProgressSheet.getRange(matchRow, 1).getValue();

      historySheet.appendRow([
        formattedTime,
        matchTableNumber,
        winner,
        winnerName,
        loser,
        loserName,
        winnerName,
        newId
      ]);

      updatePlayerMatchStats(winner, true, formattedTime);
      updatePlayerMatchStats(loser, false, formattedTime);
    }

    // 5. 対戦情報をクリア（対戦中の場合のみ）。卓番号は残す
    if (currentStatus === PLAYER_STATUS.IN_PROGRESS && matchRow !== -1) {
      const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
      inProgressSheet.getRange(matchRow, 2, 1, 4).clearContent(); // ID1からID2までをクリア
    }

    // 6. プレイヤーの状態を更新
    for (let i = 1; i < playerData.length; i++) {
      const row = playerData[i];
      const playerId = row[playerIndices["プレイヤーID"]];
      if (playerId === targetPlayerId) {
        playerSheet.getRange(i + 1, playerIndices["参加状況"] + 1)
          .setValue(newStatus);
      } else if (playerId === opponentId) {
        playerSheet.getRange(i + 1, playerIndices["参加状況"] + 1)
          .setValue(opponentNewStatus);
      }
    }

    // 7. シートのクリーンアップ
    cleanUpInProgressSheet();

    // 8. 必要に応じて次のマッチング
    const waitingPlayersCount = getWaitingPlayers().length;
    if (waitingPlayersCount >= 2) {
      matchPlayers();
    }

    return {
      success: true,
      message: "状態変更が完了しました。",
      opponentId
    };

  } catch (e) {
    Logger.log("handleMatchStateChange エラー: " + e.message);
    return {
      success: false,
      message: "エラーが発生しました: " + e.toString()
    };
  } finally {
    releaseLock(matchLock);
    releaseLock(stateLock);
  }
}