/**
 * ポケモンカード・ガンスリンガーバトル用マッチングシステム
 * @fileoverview UI操作の共通ユーティリティ関数
 * @author SpringOK
 */

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
