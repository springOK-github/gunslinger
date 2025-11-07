/**
 * ポケモンカード・ガンスリンガーバトル用マッチングシステム
 * @fileoverview プレイヤー統計情報の更新
 * @author SpringOK
 */

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
