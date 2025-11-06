/**
 * ポケモンカード・ガンスリンガーバトル用マッチングシステム
 * @fileoverview システム全体で使用する定数の定義
 * @author SpringOK
 */

const SHEET_PLAYERS = "プレイヤー";
const SHEET_HISTORY = "対戦履歴";
const SHEET_IN_PROGRESS = "マッチング";
const PLAYER_ID_PREFIX = "P";
const ID_DIGITS = 3; // IDの数字部分の桁数 (例: P001なら3)
const PLAYER_STATUS = {
  WAITING: "待機",
  IN_PROGRESS: "対戦中",
  DROPPED: "終了"
};
const REQUIRED_HEADERS = {
  [SHEET_PLAYERS]: ["プレイヤーID", "勝数", "敗数", "消化試合数", "参加状況", "最終対戦日時"],
  [SHEET_HISTORY]: ["日時", "プレイヤー1 ID", "プレイヤー2 ID", "勝者ID", "対戦ID"],
  [SHEET_IN_PROGRESS]: ["プレイヤー1 ID", "プレイヤー2 ID"]
};