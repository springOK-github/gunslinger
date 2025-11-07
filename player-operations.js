/**
 * ポケモンカード・ガンスリンガーバトル用マッチングシステム
 * @fileoverview プレイヤーのUI操作（登録、ドロップアウト、休憩管理）
 * @author SpringOK
 */

/**
 * 新しいプレイヤーを登録します。（本番・運営用）
 * 実行すると、次のID（例: P009）が自動で採番され、シートに追加されます。
 */
function registerPlayer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
  const ui = SpreadsheetApp.getUi();
  let lock = null;
  
  try {
    lock = acquireLock('プレイヤー登録');
    getSheetStructure(playerSheet, SHEET_PLAYERS);

    const response = ui.prompt(
      'プレイヤー登録',
      'プレイヤー名を入力してください：',
      ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() == ui.Button.CANCEL) {
      Logger.log('プレイヤー登録がキャンセルされました。');
      return;
    }

    const playerName = response.getResponseText().trim();
    if (!playerName) {
      ui.alert('エラー', 'プレイヤー名を入力してください。', ui.ButtonSet.OK);
      return;
    }

    const lastRow = playerSheet.getLastRow();
    const newIdNumber = lastRow;
    const currentTime = new Date();
    const newId = PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, newIdNumber);
    const formattedTime = Utilities.formatDate(currentTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

    playerSheet.appendRow([newId, playerName, 0, 0, 0, PLAYER_STATUS.WAITING, formattedTime]);
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
  finally {
    releaseLock(lock);
  }
}

/**
 * プレイヤーを大会からドロップアウトさせます。
 * 参加状況を「終了」に変更し、進行中の対戦がある場合は無効にします。
 */
function dropoutPlayer() {
  changePlayerStatus({
    actionName: 'プレイヤーのドロップアウト',
    promptMessage: 'ドロップアウトするプレイヤーIDの**数字部分のみ**を入力してください (例: P001なら「1」)。',
    confirmMessage: 'をドロップアウトさせます。\n進行中の対戦がある場合は無効となります。',
    newStatus: PLAYER_STATUS.DROPPED
  });
}

/**
 * プレイヤーを休憩状態にします。
 * 待機中または対戦中から休憩に遷移できます。
 */
function setPlayerResting() {
  changePlayerStatus({
    actionName: 'プレイヤーの休憩',
    promptMessage: '休憩するプレイヤーIDの**数字部分のみ**を入力してください (例: P001なら「1」)。',
    confirmMessage: 'を休憩状態にします。\n進行中の対戦がある場合は無効となります。',
    newStatus: PLAYER_STATUS.RESTING,
  });
}

/**
 * 休憩中のプレイヤーを待機状態に復帰させます。
 */
function returnPlayerFromResting() {
  const ui = SpreadsheetApp.getUi();

  const playerId = promptPlayerId(
    '休憩からの復帰',
    '復帰するプレイヤーIDの**数字部分のみ**を入力してください (例: P001なら「1」)。'
  );
  if (!playerId) return;

  // プレイヤーの現在の状態を確認
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
  let lock = null;

  try {
    lock = acquireLock('休憩からの復帰');
    const { indices, data } = getSheetStructure(playerSheet, SHEET_PLAYERS);

    let found = false;
    let currentStatus = null;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[indices["プレイヤーID"]] === playerId) {
        found = true;
        currentStatus = row[indices["参加状況"]];
        break;
      }
    }

    if (!found) {
      ui.alert('エラー', `プレイヤー ${playerId} が見つかりません。`, ui.ButtonSet.OK);
      return;
    }

    if (currentStatus !== PLAYER_STATUS.RESTING) {
      ui.alert('エラー', `プレイヤー ${playerId} は休憩中ではありません（現在: ${currentStatus}）。`, ui.ButtonSet.OK);
      return;
    }

    const confirmResponse = ui.alert(
      '復帰の確認',
      `プレイヤー ${playerId} を休憩から待機状態に復帰させます。\n\n` +
      'よろしいですか？',
      ui.ButtonSet.YES_NO
    );

    if (confirmResponse !== ui.Button.YES) {
      ui.alert('処理をキャンセルしました。');
      return;
    }

    // 状態を待機に変更
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[indices["プレイヤーID"]] === playerId) {
        playerSheet.getRange(i + 1, indices["参加状況"] + 1)
          .setValue(PLAYER_STATUS.WAITING);
        break;
      }
    }

    ui.alert('完了', `プレイヤー ${playerId} を待機状態に復帰させました。`, ui.ButtonSet.OK);

    // 待機者が2人以上いれば自動マッチング
    const waitingPlayersCount = getWaitingPlayers().length;
    if (waitingPlayersCount >= 2) {
      Logger.log(`復帰後、待機プレイヤーが ${waitingPlayersCount} 人いるため、自動でマッチングを開始します。`);
      matchPlayers();
    }

  } catch (e) {
    ui.alert("エラーが発生しました: " + e.toString());
    Logger.log("returnPlayerFromResting エラー: " + e.toString());
  } finally {
    releaseLock(lock);
  }
}
