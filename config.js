/**
 * ポケモンカード・ガンスリンガーバトル用マッチングシステム
 * @fileoverview システム設定の管理（PropertiesServiceを使用した永続化）
 * @author SpringOK
 */

/**
 * 現在の最大卓数を取得します。
 * PropertiesServiceに保存されている値、なければデフォルト値を返します。
 * @returns {number} 最大卓数
 */
function getMaxTables() {
  const properties = PropertiesService.getDocumentProperties();
  const savedMaxTables = properties.getProperty('MAX_TABLES');
  
  if (savedMaxTables) {
    return parseInt(savedMaxTables, 10);
  }
  
  // デフォルト値
  return TABLE_CONFIG.MAX_TABLES;
}

/**
 * 最大卓数を設定します。
 * @param {number} maxTables - 設定する最大卓数
 */
function setMaxTables(maxTables) {
  const properties = PropertiesService.getDocumentProperties();
  properties.setProperty('MAX_TABLES', maxTables.toString());
  Logger.log(`最大卓数を ${maxTables} に設定しました。`);
}

/**
 * 最大卓数の設定をユーザーに促すダイアログを表示します。
 */
function configureMaxTables() {
  const ui = SpreadsheetApp.getUi();
  const currentMaxTables = getMaxTables();
  
  const response = ui.prompt(
    '最大卓数の設定',
    `現在の最大卓数: ${currentMaxTables}卓\n\n` +
    `新しい最大卓数を入力してください（1～200）：`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert('設定をキャンセルしました。');
    return;
  }

  const input = response.getResponseText().trim();

  // 入力検証
  if (!/^\d+$/.test(input)) {
    ui.alert('エラー', '数字のみで入力してください。', ui.ButtonSet.OK);
    return;
  }

  const newMaxTables = parseInt(input, 10);

  // 範囲検証
  if (newMaxTables < 1 || newMaxTables > 200) {
    ui.alert('エラー', '最大卓数は1～200の範囲で入力してください。', ui.ButtonSet.OK);
    return;
  }

  // 確認ダイアログ
  const confirmResponse = ui.alert(
    '設定の確認',
    `最大卓数を ${currentMaxTables}卓 → ${newMaxTables}卓 に変更します。\n\n` +
    'よろしいですか？',
    ui.ButtonSet.YES_NO
  );

  if (confirmResponse !== ui.Button.YES) {
    ui.alert('設定をキャンセルしました。');
    return;
  }

  // 設定を保存
  setMaxTables(newMaxTables);
  
  ui.alert(
    '設定完了',
    `最大卓数を ${newMaxTables}卓 に設定しました。`,
    ui.ButtonSet.OK
  );
}
