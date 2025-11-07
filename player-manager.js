/**
 * ポケモンカード・ガンスリンガーバトル用マッチングシステム
 * @fileoverview プレイヤーのライフサイクル管理（登録、状態管理、統計、ドロップアウト）
 * @author SpringOK
 */

// =========================================
// プレイヤー登録・管理系
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

// =========================================
// プレイヤーの状態取得・更新系
// =========================================

/**
 * 待機中のプレイヤーを抽出し、以下の優先順位でソートして返します。
 * 1. 勝数（降順）
 * 2. 最終対戦日時（降順 = 最近待機に戻った人優先 = 直近の勝者優先）
 */
function getWaitingPlayers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);

  try {
    const { indices, data } = getSheetStructure(playerSheet, SHEET_PLAYERS);
    if (data.length <= 1) return [];

    const waiting = data.slice(1).filter(row => 
      row[indices["参加状況"]] === PLAYER_STATUS.WAITING
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

// =========================================
// 対戦履歴・統計系
// =========================================

/**
 * 特定プレイヤーの過去の対戦相手のIDリスト（ブラックリスト）を取得します。
 */
function getPastOpponents(playerId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName(SHEET_HISTORY);

  try {
    const { indices, data } = getSheetStructure(historySheet, SHEET_HISTORY);
    if (data.length <= 1) return [];

    const p1Col = indices["ID1"];
    const p2Col = indices["ID2"];
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
function updatePlayerMatchStats(playerId, isWinner, timestamp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);

  try {
    const { indices, data } = getSheetStructure(playerSheet, SHEET_PLAYERS);
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
    Logger.log(`エラー: プレイヤー ${playerId} が見つかりません。`);
  } catch (e) {
    Logger.log("updatePlayerStats エラー: " + e.message);
  }
}