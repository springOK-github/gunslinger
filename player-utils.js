/**
 * プレイヤー管理ヘルパー関数
 */

/**
 * プレイヤーのドロップアウトを処理します。
 * 指定されたプレイヤーの参加状況を「終了」に変更し、進行中の対戦を無効にします。
 */
function handleDropout() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

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

  if (!/^\d+$/.test(rawId)) {
    ui.alert('エラー: IDは数字のみで入力してください。');
    return;
  }

  const playerId = PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, parseInt(rawId, 10));

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
    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    const { indices: playerIndices, data: playerData } = validateHeaders(playerSheet, SHEET_PLAYERS);
    
    let playerFound = false;
    for (let i = 1; i < playerData.length; i++) {
      const row = playerData[i];
      if (row[playerIndices["プレイヤーID"]] === playerId) {
        playerSheet.getRange(i + 1, playerIndices["参加状況"] + 1).setValue(PLAYER_STATUS.DROPPED);
        playerFound = true;
        break;
      }
    }

    if (!playerFound) {
      ui.alert(`エラー: プレイヤーID ${playerId} が見つかりません。`);
      return;
    }

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
      inProgressSheet.getRange(matchRow, 1, 1, 2).clearContent();
      matchCancelled = true;

      for (let i = 1; i < playerData.length; i++) {
        const row = playerData[i];
        if (row[playerIndices["プレイヤーID"]] === opponentId) {
          playerSheet.getRange(i + 1, playerIndices["参加状況"] + 1).setValue(PLAYER_STATUS.WAITING);
          break;
        }
      }
    }

    let message = `プレイヤー ${playerId} のドロップアウトを処理しました。\n参加状況を「終了」に変更しました。`;
    if (matchCancelled) {
      message += `\n\n進行中の対戦を無効とし、対戦相手（${opponentId}）を待機状態に戻しました。`;
    }
    ui.alert('完了', message, ui.ButtonSet.OK);

    cleanUpInProgressSheet();

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

    playerSheet.appendRow([newId, 0, 0, 0, PLAYER_STATUS.WAITING, currentTime]);
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
      row[indices["参加状況"]] === PLAYER_STATUS.WAITING && row[indices["参加状況"]] !== PLAYER_STATUS.DROPPED
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