/**
 * ポケモンカード・ガンスリンガーバトル用マッチングシステム
 * @fileoverview テスト用のユーティリティ関数
 * @author SpringOK
 */

// テスト設定
const TEST_CONFIG = {
  NUM_PLAYERS: 8  // 登録するテストプレイヤー数（この値を変更して調整）
};

/**
 * テスト用のプレイヤーを一括登録します。
 * registerPlayer()と同じロジックを使用し、名前はプレイヤーIDに固定。
 */
function registerTestPlayers() {
  const numPlayers = TEST_CONFIG.NUM_PLAYERS;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
  let lock = null;

  try {
    lock = acquireLock('テストプレイヤー登録');
    const { indices, data } = getSheetStructure(playerSheet, SHEET_PLAYERS);

    // 既存の最大ID番号を取得（本物のregisterPlayer()と同じロジック）
    let maxIdNumber = 0;
    for (let i = 1; i < data.length; i++) {
      const playerId = data[i][indices["プレイヤーID"]];
      if (playerId && playerId.startsWith(PLAYER_ID_PREFIX)) {
        const idNumber = parseInt(playerId.substring(PLAYER_ID_PREFIX.length), 10);
        if (!isNaN(idNumber) && idNumber > maxIdNumber) {
          maxIdNumber = idNumber;
        }
      }
    }

    Logger.log(`テストプレイヤー ${numPlayers} 人の登録を開始します。`);

    // 各プレイヤーを登録
    for (let i = 0; i < numPlayers; i++) {
      const newIdNumber = maxIdNumber + i + 1;
      const newId = PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, newIdNumber);
      const playerName = newId;  // 名前はIDと同じ
      const currentTime = new Date();
      const formattedTime = Utilities.formatDate(currentTime, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

      playerSheet.appendRow([newId, playerName, 0, 0, 0, PLAYER_STATUS.WAITING, formattedTime]);
      Logger.log(`プレイヤー ${playerName} (${newId}) を登録しました。`);
    }

    // 自動マッチング
    const waitingPlayersCount = getWaitingPlayers().length;
    if (waitingPlayersCount >= 2) {
      Logger.log(`テストプレイヤー登録完了。待機プレイヤーが ${waitingPlayersCount} 人いるため、自動でマッチングを開始します。`);
      matchPlayers();
    } else {
      Logger.log(`テストプレイヤー登録完了。待機プレイヤーが ${waitingPlayersCount} 人です。`);
    }

  } catch (e) {
    Logger.log("registerTestPlayers エラー: " + e.toString());
  } finally {
    releaseLock(lock);
  }
}