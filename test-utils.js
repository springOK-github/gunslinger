/**
 * ポケモンカード・ガンスリンガーバトル用マッチングシステム
 * @fileoverview テスト用のユーティリティ関数
 * @author SpringOK
 */

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
  const baseDate = new Date();
  for (let i = 0; i < numTestPlayers; i++) {
    const newIdNumber = i + 1;
    const newId = PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, newIdNumber);
    // 登録順を保証するため、登録日時を1秒ずつずらす
    const registrationDate = new Date(baseDate.getTime() + i * 1000);
    const formattedTime = Utilities.formatDate(registrationDate, 'Asia/Tokyo', 'MM/dd HH:mm');
    playerSheet.appendRow([newId, newId, 0, 0, 0, PLAYER_STATUS.WAITING, formattedTime]);
  }

  const waitingPlayersCount = getWaitingPlayers().length;
  if (waitingPlayersCount >= 2) {
    Logger.log("テストプレイヤー登録完了。自動で初回マッチングを開始します。");
    matchPlayers();
  } else {
    Logger.log("テストプレイヤーの登録が完了しました。マッチングには2人以上が必要です。");
  }
}