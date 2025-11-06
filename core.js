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
        const matchesWithNames = matches.map(match => [
          match[0],
          getPlayerName(match[0]),
          match[1],
          getPlayerName(match[1])
        ]);
        inProgressSheet.getRange(lastRow + 1, 1, matches.length, 4)
          .setValues(matchesWithNames);
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