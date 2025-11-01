/**
 * テストと管理用関数
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
 * テスト用のプレイヤーを一括登録します。
 */
function registerTestPlayers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);

  if (playerSheet.getLastRow() > 1) {
    playerSheet.getRange(2, 1, playerSheet.getLastRow() - 1, playerSheet.getLastColumn()).clearContent();
  }

  const numTestPlayers = 8;
  for (let i = 0; i < numTestPlayers; i++) {
    const newIdNumber = i + 1;
    const newId = PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, newIdNumber);
    playerSheet.appendRow([newId, 0, 0, 0, PLAYER_STATUS.WAITING, new Date()]);
  }

  const waitingPlayersCount = getWaitingPlayers().length;
  if (waitingPlayersCount >= 2) {
    Logger.log("テストプレイヤー登録完了。自動で初回マッチングを開始します。");
    matchPlayers();
  } else {
    Logger.log("テストプレイヤーの登録が完了しました。マッチングには2人以上が必要です。");
  }
}