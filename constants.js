/**
 * ガンスリンガーバトル用マッチングシステム
 * @fileoverview システム全体で使用する定数の定義
 * @author springOK
 */

const SHEET_PLAYERS = "プレイヤー";
const SHEET_HISTORY = "対戦履歴";
const SHEET_IN_PROGRESS = "マッチング";
const PLAYER_ID_PREFIX = "P";
const ID_DIGITS = 3; // IDの数字部分の桁数 (例: P001なら3)
const PLAYER_STATUS = {
  WAITING: "待機",
  IN_PROGRESS: "対戦中",
  RESTING: "休憩",
  DROPPED: "終了"
};
// 卓に関する設定
const TABLE_CONFIG = {
  MAX_TABLES: 50,      // デフォルトの最大卓数（PropertiesServiceで上書き可能、範囲: 1-200）
  MIN_TABLE_NUMBER: 1  // 最小卓番号
};

// 各シートの必須ヘッダー定義
const REQUIRED_HEADERS = {
  [SHEET_PLAYERS]: ["プレイヤーID", "プレイヤー名", "勝数", "敗数", "試合数", "参加状況", "最終対戦日時"],
  [SHEET_HISTORY]: ["対戦ID", "日時", "卓番号", "ID1", "プレイヤー1", "ID2", "プレイヤー2", "勝者名"],
  [SHEET_IN_PROGRESS]: ["卓番号", "ID1", "プレイヤー1", "ID2", "プレイヤー2"]
};