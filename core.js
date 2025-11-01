/**
 * ポケモンカード・ガンスリンガーバトル用マッチングシステム
 * @fileoverview マッチングと対戦結果記録のメインロジック
 * @author SpringOK
 */

// =========================================
// マッチング系
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
    validateHeaders(inProgressSheet, SHEET_IN_PROGRESS);
    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    const { indices: playerIndices, data: playerData } = validateHeaders(playerSheet, SHEET_PLAYERS);

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
      const playerIdsToUpdate = matches.flat();
      
      for (let i = 1; i < playerData.length; i++) {
        const row = playerData[i];
        const playerId = row[playerIndices["プレイヤーID"]];
        if (playerIdsToUpdate.includes(playerId)) {
          playerSheet.getRange(i + 1, playerIndices["参加状況"] + 1)
            .setValue(PLAYER_STATUS.IN_PROGRESS);
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
  } finally {
    releaseLock(lock);
  }
}

// =========================================
// 対戦結果系
// =========================================

/**
 * カスタムメニューから実行するためのラッパー関数。
 */
function promptAndRecordResult() {
  const ui = SpreadsheetApp.getUi();

  // 最初にユーザー入力を受け付け
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

  if (!/^\d+$/.test(rawId)) {
    ui.alert('エラー: IDは数字のみで入力してください。');
    return;
  }

  const formattedWinnerId = PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, parseInt(rawId, 10));

  // この時点でロックは取得しない
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
  const { indices, data } = validateHeaders(inProgressSheet, SHEET_IN_PROGRESS);
  let loserId = null;

  // 対戦相手の確認（ロック不要）
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

  // ユーザーに確認
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let lock = null;

  if (!winnerId) {
    ui.alert("勝者IDを入力してください。");
    return;
  }

  try {
    lock = acquireLock('対戦結果の記録');

    // まず両プレイヤーの状態を確認
    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    const { indices: playerIndices, data: playerData } = validateHeaders(playerSheet, SHEET_PLAYERS);

    let winnerFound = false;
    let winnerDropped = false;
    let loserDropped = false;

    const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
    const { indices: inProgressIndices, data: inProgressData } = 
      validateHeaders(inProgressSheet, SHEET_IN_PROGRESS);
    
    let tempLoserId = null;
    // まず敗者を特定
    for (let i = 1; i < inProgressData.length; i++) {
      const row = inProgressData[i];
      const p1 = row[inProgressIndices["プレイヤー1 ID"]];
      const p2 = row[inProgressIndices["プレイヤー2 ID"]];

      if (p1 === winnerId) {
        tempLoserId = p2;
        break;
      } else if (p2 === winnerId) {
        tempLoserId = p1;
        break;
      }
    }

    if (tempLoserId === null) {
      ui.alert(`エラー: 勝者ID (${winnerId}) は「対戦中」シートに見つかりませんでした。\n入力IDが間違っているか、対戦が記録されていません。`);
      return;
    }

    // 両プレイヤーの状態を確認
    for (let i = 1; i < playerData.length; i++) {
      const row = playerData[i];
      const playerId = row[playerIndices["プレイヤーID"]];
      const status = row[playerIndices["参加状況"]];
      
      if (playerId === winnerId) {
        winnerFound = true;
        winnerDropped = status === PLAYER_STATUS.DROPPED;
      } else if (playerId === tempLoserId) {
        loserDropped = status === PLAYER_STATUS.DROPPED;
      }
    }

    if (!winnerFound) {
      ui.alert('エラー: 指定された勝者が見つかりません。');
      return;
    }

    if (winnerDropped) {
      ui.alert('エラー: 勝者としてマークされたプレイヤーはすでにドロップアウトしています。');
      return;
    }

    if (loserDropped) {
      ui.alert('エラー: 敗者としてマークされたプレイヤーはすでにドロップアウトしています。');
      return;
    }

    // tempLoserIdをloserIdとして使用
    const loserId = tempLoserId;
    let rowToClear = -1;

    // 対戦中の行を特定
    for (let i = 1; i < inProgressData.length; i++) {
      const row = inProgressData[i];
      const p1 = row[inProgressIndices["プレイヤー1 ID"]];
      const p2 = row[inProgressIndices["プレイヤー2 ID"]];

      if ((p1 === winnerId && p2 === loserId) || (p2 === winnerId && p1 === loserId)) {
        rowToClear = i + 1;
        break;
      }
    }

    const currentTime = new Date();

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

    updatePlayerStats(winnerId, true, currentTime);
    updatePlayerStats(loserId, false, currentTime);

    if (rowToClear !== -1) {
      inProgressSheet.getRange(rowToClear, 1, 1, 2).clearContent();
    }

    // プレイヤーの状態を待機に戻す
    for (let i = 1; i < playerData.length; i++) {
      const row = playerData[i];
      const playerId = row[playerIndices["プレイヤーID"]];
      if (playerId === winnerId || playerId === loserId) {
        playerSheet.getRange(i + 1, playerIndices["参加状況"] + 1)
          .setValue(PLAYER_STATUS.WAITING);
      }
    }

    Logger.log(`対戦結果が記録されました。勝者: ${winnerId}, 敗者: ${loserId}。両プレイヤーは待機状態に戻りました。`);

    cleanUpInProgressSheet();

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
  } finally {
    releaseLock(lock);
  }
}