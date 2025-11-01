/**
 * ポケモンカード・ガンスリンガーバトル用マッチングシステム
 * Google Apps Script (GAS) とスプレッドシートで動作します。
 *
 * 【変更点】
 * - プレイヤーシートに「最終対戦日時」列を追加。
 * - 待機プレイヤーのソート順を「勝数（降順）」と「最終対戦日時（降順=最近待機になった人優先）」に変更し、
 * 直近の勝者が優先的にマッチングされるようにしました。
 */

// --- 設定 ---
const SHEET_PLAYERS = "プレイヤー";
const SHEET_HISTORY = "対戦履歴";
const SHEET_IN_PROGRESS = "対戦中";
const PLAYER_ID_PREFIX = "P";
const ID_DIGITS = 3; // IDの数字部分の桁数 (例: P001なら3)

/**
 * スプレッドシートを開いたときにカスタムメニューを作成します。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🃏 ポケモンマッチング')
    .addItem('① シートの初期設定', 'setupSheets')
    .addSeparator()
    .addItem('② 新しいプレイヤーの登録 (自動マッチング実行)', 'registerPlayer')
    .addItem('②-B テストプレイヤー登録 (初期登録用)', 'registerTestPlayers')
    .addSeparator()
    .addItem('④ 対戦結果の記録 (自動マッチング実行)', 'promptAndRecordResult')
    .addToUi();
}


/**
 * スプレッドシートを初期化し、必要なシートとヘッダーを作成します。
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. プレイヤーシート
  let playerSheet = ss.getSheetByName(SHEET_PLAYERS);
  if (!playerSheet) {
    playerSheet = ss.insertSheet(SHEET_PLAYERS);
  }
  playerSheet.clear();
  playerSheet.getRange("A1:F1").setValues([ // F列まで拡張
    ["プレイヤーID", "勝数", "敗数", "消化試合数", "参加状況", "最終対戦日時"]
  ]).setFontWeight("bold").setBackground("#c9daf8").setHorizontalAlignment("center");
  // 幅の調整
  playerSheet.setColumnWidth(1, 100);
  playerSheet.setColumnWidth(5, 100);
  playerSheet.setColumnWidth(6, 150); // 最終対戦日時

  // 2. 対戦履歴シート
  let historySheet = ss.getSheetByName(SHEET_HISTORY);
  if (!historySheet) {
    historySheet = ss.insertSheet(SHEET_HISTORY);
  }
  historySheet.clear();
  historySheet.getRange("A1:E1").setValues([
    ["日時", "プレイヤー1 ID", "プレイヤー2 ID", "勝者ID", "対戦ID"]
  ]).setFontWeight("bold").setBackground("#fce5cd").setHorizontalAlignment("center");
  historySheet.setColumnWidth(1, 150);

  // 3. 対戦中シート
  let inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
  if (!inProgressSheet) {
    inProgressSheet = ss.insertSheet(SHEET_IN_PROGRESS);
  }
  inProgressSheet.clear();
  inProgressSheet.getRange("A1:B1").setValues([
    ["プレイヤー1 ID", "プレイヤー2 ID"]
  ]).setFontWeight("bold").setBackground("#d9ead3").setHorizontalAlignment("center");
  inProgressSheet.setColumnWidth(3, 80);

  Logger.log("シートの初期設定が完了しました。");
}

// ----------------------------------------------------------------------
// --- メイン関数 ---
// ----------------------------------------------------------------------

/**
 * 待機中のプレイヤーを抽出し、再戦履歴を厳格に考慮してマッチングを行います。
 * 過去に対戦した相手しかいない場合、マッチングを成立させずに待機させます。
 */
function matchPlayers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);

  // 1. 待機中のプレイヤーリスト（勝数順 and 最終対戦日時順）を取得
  const waitingPlayers = getWaitingPlayers();

  if (waitingPlayers.length < 2) {
    Logger.log(`警告: 現在待機中のプレイヤーは ${waitingPlayers.length} 人です。2人以上必要です。`);
    return;
  }

  // 2. マッチングを実行 (再戦回避のみ)
  let matches = [];
  let availablePlayers = [...waitingPlayers]; // 操作用のリスト
  let skippedPlayers = []; // マッチングできなかったプレイヤー

  Logger.log("--- 厳格な再戦回避マッチング開始 (勝者優先) ---");
  while (availablePlayers.length >= 2) {
    const p1 = availablePlayers.shift();
    const p1Id = p1[0];
    const p1BlackList = getPastOpponents(p1Id);

    let p2Index = -1;

    // 再戦なしの相手を探す
    for (let i = 0; i < availablePlayers.length; i++) {
      const p2Id = availablePlayers[i][0];
      if (!p1BlackList.includes(p2Id)) {
        p2Index = i;
        break;
      }
    }

    if (p2Index !== -1) {
      // 再戦なしでマッチング成立
      const p2 = availablePlayers.splice(p2Index, 1)[0];
      matches.push([p1Id, p2[0]]);
      Logger.log(`マッチング成立 (再戦なし): ${p1Id} vs ${p2[0]}`);
    } else {
      // 適切な相手が見つからなかった場合、スキップして待機リストに残す
      skippedPlayers.push(p1);
    }
  }

  // 最後に availablePlayers に残ったプレイヤー（奇数で余ったプレイヤー、またはマッチング不可のプレイヤー）もスキップ扱い
  skippedPlayers.push(...availablePlayers);

  if (skippedPlayers.length > 0) {
    Logger.log(`警告: ${skippedPlayers.length} 人のプレイヤーは適切な相手が見つからなかったため、待機を継続します。`);
  }

  // 3. シートの更新
  if (matches.length > 0) {
    // プレイヤーシートの「参加状況」を更新（待機 -> 対戦中）
    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    const playerIdsToUpdate = matches.flat();

    const data = playerSheet.getDataRange().getValues();
    const headers = data[0];
    const statusCol = headers.indexOf("参加状況");
    const idCol = headers.indexOf("プレイヤーID");

    let inProgressData = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const playerId = row[idCol];
      if (playerIdsToUpdate.includes(playerId)) {
        playerSheet.getRange(i + 1, statusCol + 1).setValue("対戦中");
      }
    }

    // --- 対戦中シートへの追記処理 ---
    const lastRow = inProgressSheet.getLastRow();
    let startRow = lastRow + 1;

    for (const match of matches) {
      // プレイヤーIDのペアのみを配列に追加
      inProgressData.push([match[0], match[1]]);
    }

    if (inProgressData.length > 0) {
      // B列まで(2列)にデータを追記する
      inProgressSheet.getRange(startRow, 1, inProgressData.length, 2).setValues(inProgressData);
    }

    Logger.log(`マッチングが ${matches.length} 件成立しました。「対戦中」シートを確認してください。`);
    return matches.length; // 成立したマッチング数を返す
  } else {
    Logger.log("警告: 新しいマッチングは成立しませんでした。");
    return 0;
  }
}

/**
 * カスタムメニューから実行するためのラッパー関数。
 */
function promptAndRecordResult() {
  const ui = SpreadsheetApp.getUi();

  // 勝者IDの数字部分を尋ねる (入力が必要なため維持)
  const winnerResponse = ui.prompt(
    '対戦結果の記録',
    '勝者のプレイヤーIDの**数字部分のみ**を入力してください (例: P001なら「1」)。\n敗者は「対戦中」シートから自動特定されます。',
    ui.ButtonSet.OK_CANCEL
  );

  if (winnerResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert('処理をキャンセルしました。');
    return;
  }

  const rawId = winnerResponse.getResponseText().trim();

  // 数字入力チェックとP00X形式への変換
  if (!/^\d+$/.test(rawId)) {
    ui.alert('エラー: IDは数字のみで入力してください。');
    return;
  }

  // 自動でプレフィックスとゼロパディングを付与
  const formattedWinnerId = PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, parseInt(rawId, 10));

  recordResult(formattedWinnerId);
}


/**
 * 対戦結果を記録し、プレイヤーの統計情報とステータスを更新し、自動で次をマッチングします。
 */
function recordResult(winnerId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // winnerIdは既にP00X形式にフォーマットされていることを前提とする
  if (!winnerId) {
    ui.alert("勝者IDを入力してください。");
    return;
  }

  // 1. 対戦中シートから敗者IDを特定
  const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
  const data = inProgressSheet.getDataRange().getValues();

  let loserId = null;
  let rowToClear = -1; // クリア対象の行番号

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const p1 = row[0];
    const p2 = row[1];

    // シート上のID (P00X形式) と入力されたID (P00X形式) を比較
    if (p1 === winnerId) {
      loserId = p2;
      rowToClear = i + 1; // シートの行番号
      break;
    } else if (p2 === winnerId) {
      loserId = p1;
      rowToClear = i + 1; // シートの行番号
      break;
    }
  }

  if (loserId === null) {
    ui.alert(`エラー: 勝者ID (${winnerId}) は「対戦中」シートに見つかりませんでした。\n入力IDが間違っているか、対戦が記録されていません。`);
    return;
  }

  const currentTime = new Date(); // 現在時刻を取得

  // 2. 対戦履歴に記録
  try {
    const historySheet = ss.getSheetByName(SHEET_HISTORY);
    const newId = "T" + Utilities.formatString("%04d", historySheet.getLastRow());

    historySheet.appendRow([
      currentTime, // 履歴シートには処理時刻を記録
      winnerId,
      loserId,
      winnerId,
      newId
    ]);

    // 3. プレイヤーの統計情報とステータスを更新
    updatePlayerStats(winnerId, true, currentTime); // 勝者の統計と最終対戦日時を更新
    updatePlayerStats(loserId, false, currentTime); // 敗者の統計と最終対戦日時を更新

    // 4. 「対戦中」シートから終了した対戦のコンテンツをクリア
    if (rowToClear !== -1) {
      // A列とB列 (2列) のみをクリア
      inProgressSheet.getRange(rowToClear, 1, 1, 2).clearContent();
    }

    // 5. 参加状況を「待機」に更新
    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    const dataRange = playerSheet.getDataRange();
    const values = dataRange.getValues();
    const headers = values[0];
    const statusCol = headers.indexOf("参加状況");
    const idCol = headers.indexOf("プレイヤーID");

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const playerId = row[idCol];
      if (playerId === winnerId || playerId === loserId) {
        // updatePlayerStatsですでに日時を記録しているため、ここではステータス更新のみ
        playerSheet.getRange(i + 1, statusCol + 1).setValue("待機");
      }
    }

    Logger.log(`対戦結果が記録されました。勝者: ${winnerId}, 敗者: ${loserId}。両プレイヤーは待機状態に戻りました。`);

    // 6. 対戦中シートを自動で整理
    cleanUpInProgressSheet();

    // 7. 待機プレイヤーが2人以上いれば、自動でマッチングを実行
    const waitingPlayersCount = getWaitingPlayers().length;
    if (waitingPlayersCount >= 2) {
      Logger.log(`待機プレイヤーが ${waitingPlayersCount} 人いるため、自動でマッチングを開始します。`);
      matchPlayers();
    } else {
      Logger.log(`待機プレイヤーが ${waitingPlayersCount} 人です。自動マッチングはスキップされました。`);
    }

  } catch (e) {
    ui.alert("エラーが発生しました: " + e.toString());
    Logger.log("エラー: " + e.toString());
  }
}

/**
 * 「対戦中」シート内の空行（対戦が終了し、コンテンツがクリアされた行）を削除し、
 * シート内のデータを上詰めして整理します。
 */
function cleanUpInProgressSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);

  const lastRow = inProgressSheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log("「対戦中」シートにデータがないため、整理は不要です。");
    return;
  }

  // データの最終行から2行目まで逆順にチェック
  // 逆順にすることで、行を削除してもインデックスが狂わない
  let deletedCount = 0;
  for (let i = lastRow; i >= 2; i--) {
    const cellA = inProgressSheet.getRange(i, 1).getValue(); // A列の値

    // A列が空（対戦が終了しクリアされた行）であれば、行を削除
    if (cellA === "") {
      inProgressSheet.deleteRow(i);
      deletedCount++;
    }
  }

  if (deletedCount > 0) {
    Logger.log(`対戦中シートの整理 (自動実行) が完了しました。${deletedCount} 行の空行を削除しました。`);
  } else {
    // 頻繁に実行されるため、特にログは出力しない
  }
}


// ----------------------------------------------------------------------
// --- ヘルパー関数 ---
// ----------------------------------------------------------------------

/**
 * 待機中のプレイヤーを抽出し、以下の優先順位でソートして返します。
 * 1. 勝数（降順）
 * 2. 最終対戦日時（降順 = 最近待機に戻った人優先 = 直近の勝者優先）
 */
function getWaitingPlayers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);

  const data = playerSheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const winCol = headers.indexOf("勝数");
  const statusCol = headers.indexOf("参加状況");
  const lastPlayedCol = headers.indexOf("最終対戦日時");

  const waiting = data.slice(1).filter(row => row[statusCol] === "待機");

  // ソート処理
  waiting.sort((a, b) => {
    // 1. 勝数で比較 (b > a ならbが先)
    if (b[winCol] !== a[winCol]) {
      return b[winCol] - a[winCol];
    }

    // 2. 勝数が同じ場合、最終対戦日時で比較 (b > a ならbが先 = 新しい日時が先)
    const dateA = a[lastPlayedCol] instanceof Date ? a[lastPlayedCol].getTime() : 0;
    const dateB = b[lastPlayedCol] instanceof Date ? b[lastPlayedCol].getTime() : 0;

    return dateB - dateA;
  });

  return waiting;
}

/**
 * 特定プレイヤーの過去の対戦相手のIDリスト（ブラックリスト）を取得します。
 */
function getPastOpponents(playerId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = ss.getSheetByName(SHEET_HISTORY);

  const data = historySheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const p1Col = headers.indexOf("プレイヤー1 ID");
  const p2Col = headers.indexOf("プレイヤー2 ID");

  const opponents = new Set();

  data.slice(1).forEach(row => {
    if (row[p1Col] === playerId) {
      opponents.add(row[p2Col]);
    } else if (row[p2Col] === playerId) {
      opponents.add(row[p1Col]);
    }
  });

  return Array.from(opponents);
}

/**
 * プレイヤーの統計情報 (勝数, 敗数, 消化試合数) と最終対戦日時を更新します。
 */
function updatePlayerStats(playerId, isWinner, timestamp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);

  const data = playerSheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const headers = data[0];
  const idCol = headers.indexOf("プレイヤーID");
  const winCol = headers.indexOf("勝数");
  const lossCol = headers.indexOf("敗数");
  const totalCol = headers.indexOf("消化試合数");
  const lastPlayedCol = headers.indexOf("最終対戦日時");

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[idCol] === playerId) {
      const rowNum = i + 1;

      const currentWins = parseInt(row[winCol]) || 0;
      const currentLosses = parseInt(row[lossCol]) || 0;
      const currentTotal = parseInt(row[totalCol]) || 0;

      playerSheet.getRange(rowNum, winCol + 1).setValue(currentWins + (isWinner ? 1 : 0));
      playerSheet.getRange(rowNum, lossCol + 1).setValue(currentLosses + (isWinner ? 0 : 1));
      playerSheet.getRange(rowNum, totalCol + 1).setValue(currentTotal + 1);

      // 最終対戦日時を更新
      playerSheet.getRange(rowNum, lastPlayedCol + 1).setValue(timestamp);

      return;
    }
  }
  Logger.log(`エラー: プレイヤーID ${playerId} が見つかりません。`);
}

// ----------------------------------------------------------------------
// --- テスト・管理用関数 ---
// ----------------------------------------------------------------------

/**
 * 新しいプレイヤーを登録します。（本番・運営用）
 * 実行すると、次のID（例: P009）が自動で採番され、シートに追加されます。
 */
function registerPlayer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
  const ui = SpreadsheetApp.getUi();

  if (!playerSheet) {
    ui.alert("先に `setupSheets` を実行してシートを初期化してください。");
    return;
  }

  const lastRow = playerSheet.getLastRow();
  const newIdNumber = lastRow;
  const newId = PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, newIdNumber);

  const currentTime = new Date();
  // 新規プレイヤーは初期時点で最終対戦日時 = 現在時刻とする
  playerSheet.appendRow([newId, 0, 0, 0, "待機", currentTime]);

  Logger.log(`プレイヤー ${newId} を登録しました。`);

  // ★★★ 追記: プレイヤー登録後の自動マッチング ★★★
  const waitingPlayersCount = getWaitingPlayers().length;
  if (waitingPlayersCount >= 2) {
    Logger.log(`プレイヤー登録後、待機プレイヤーが ${waitingPlayersCount} 人いるため、自動でマッチングを開始します。`);
    matchPlayers();
  } else {
    Logger.log(`プレイヤー登録後、待機プレイヤーが ${waitingPlayersCount} 人です。自動マッチングはスキップされました。`);
  }
}

/**
 * テスト用のプレイヤーを一括登録します。
 */
function registerTestPlayers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);

  // シートをクリア
  if (playerSheet.getLastRow() > 1) {
    playerSheet.getRange(2, 1, playerSheet.getLastRow() - 1, playerSheet.getLastColumn()).clearContent();
  }

  // P001からP008まで、8人分を登録
  const numTestPlayers = 8;
  for (let i = 0; i < numTestPlayers; i++) {
    // P001, P002, ... P008を直接登録
    const newIdNumber = i + 1;
    const newId = PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, newIdNumber);
    // 最終対戦日時を初期化時も設定
    playerSheet.appendRow([newId, 0, 0, 0, "待機", new Date()]);
  }

  // 最終的にテストプレイヤーが揃った後に、一度マッチングを試みる
  const waitingPlayersCount = getWaitingPlayers().length;
  if (waitingPlayersCount >= 2) {
    Logger.log("テストプレイヤー登録完了。自動で初回マッチングを開始します。");
    matchPlayers();
  } else {
    Logger.log("テストプレイヤーの登録が完了しました。マッチングには2人以上が必要です。");
  }
}
