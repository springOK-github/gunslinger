/**
 * ポケモンカード・ガンスリンガーバトル用マッチングシステム
 * @fileoverview プレイヤーデータの取得・検索
 * @author SpringOK
 */

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
