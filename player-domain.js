/**
 * ガンスリンガーバトル用マッチングシステム
 * @fileoverview プレイヤードメイン - プレイヤー操作・統計・状態管理
 * @author springOK
 */

// =========================================
// プレイヤー操作（UI層）
// =========================================

/**
 * 新しいプレイヤーを登録します。（本番・運営用）
 * 実行すると、次のID（例: P009）が自動で採番され、シートに追加されます。
 */
function registerPlayer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let lock = null;

  // ロック取得前にUI入力を完了させる
  const response = ui.prompt("プレイヤー登録", "プレイヤー名を入力してください：", ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.CANCEL) {
    Logger.log("プレイヤー登録がキャンセルされました。");
    return;
  }

  const inputPlayerName = response.getResponseText().trim();

  try {
    lock = acquireLock("プレイヤー登録");

    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    if (!playerSheet) {
      ui.alert("エラー: プレイヤーシートが見つかりません。");
      return;
    }

    // 既存の最大ID番号を取得
    const { indices, data } = getSheetStructure(playerSheet, SHEET_PLAYERS);
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

    const newIdNumber = maxIdNumber + 1;
    const newId = PLAYER_ID_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, newIdNumber);

    // プレイヤー名が空の場合はIDを使用
    let playerName = inputPlayerName;
    if (!playerName) {
      playerName = PLAYER_NAME_PREFIX + Utilities.formatString(`%0${ID_DIGITS}d`, newIdNumber);
    }

    // 確認ダイアログを表示
    const confirmResponse = ui.alert("登録確認", `プレイヤー名: ${playerName}\nプレイヤーID: ${newId}\n\nこの内容で登録しますか？`, ui.ButtonSet.YES_NO);

    if (confirmResponse == ui.Button.NO) {
      Logger.log("プレイヤー登録が確認段階でキャンセルされました。");
      return;
    }

    const currentTime = new Date();
    const formattedTime = Utilities.formatDate(currentTime, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");

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
  } finally {
    releaseLock(lock);
  }
}

/**
 * プレイヤーを大会からドロップアウトさせます。
 * 参加状況を「終了」に変更し、進行中の対戦がある場合は無効にします。
 */
function dropoutPlayer() {
  changePlayerStatus({
    actionName: "プレイヤーのドロップアウト",
    promptMessage: "ドロップアウトするプレイヤーIDの**数字部分のみ**を入力してください (例: P001なら「1」)。",
    confirmMessage: "をドロップアウトさせます。\n進行中の対戦がある場合は無効となります。",
    newStatus: PLAYER_STATUS.DROPPED,
  });
}

/**
 * プレイヤーを休憩状態にします。
 * 待機中または対戦中から休憩に遷移できます。
 */
function restPlayer() {
  changePlayerStatus({
    actionName: "プレイヤーの休憩",
    promptMessage: "休憩するプレイヤーIDの**数字部分のみ**を入力してください (例: P001なら「1」)。",
    confirmMessage: "を休憩状態にします。\n進行中の対戦がある場合は無効となります。",
    newStatus: PLAYER_STATUS.RESTING,
  });
}

/**
 * 休憩中のプレイヤーを待機状態に復帰させます。
 */
function returnPlayerFromResting() {
  const ui = SpreadsheetApp.getUi();

  const playerId = promptPlayerId("休憩からの復帰", "復帰するプレイヤーIDの**数字部分のみ**を入力してください (例: P001なら「1」)。");
  if (!playerId) return;

  // プレイヤーの現在の状態を確認（ロックなしで読み取り）
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let lock = null;

  // まず状態確認と確認ダイアログ（ロック不要）
  const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
  if (!playerSheet) {
    ui.alert("エラー: プレイヤーシートが見つかりません。");
    return;
  }

  const { indices, data } = getSheetStructure(playerSheet, SHEET_PLAYERS);

  let found = false;
  let currentStatus = null;
  let playerName = playerId;
  let targetRowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[indices["プレイヤーID"]] === playerId) {
      found = true;
      currentStatus = row[indices["参加状況"]];
      playerName = row[indices["プレイヤー名"]];
      targetRowIndex = i + 1;
      break;
    }
  }

  if (!found) {
    ui.alert("エラー", `プレイヤー ${playerId} が見つかりません。`, ui.ButtonSet.OK);
    return;
  }

  if (currentStatus !== PLAYER_STATUS.RESTING) {
    ui.alert("エラー", `プレイヤー名: ${playerName}\nプレイヤーID: ${playerId}\n\n休憩中ではありません（現在: ${currentStatus}）。`, ui.ButtonSet.OK);
    return;
  }

  const confirmResponse = ui.alert(
    "復帰の確認",
    `プレイヤー名: ${playerName}\nプレイヤーID: ${playerId}\n\n休憩から待機状態に復帰させます。\n\nよろしいですか？`,
    ui.ButtonSet.YES_NO
  );

  if (confirmResponse !== ui.Button.YES) {
    ui.alert("処理をキャンセルしました。");
    return;
  }

  try {
    // 確認後にロックを取得して状態変更
    lock = acquireLock("休憩からの復帰");

    // 状態を待機に変更（ロック取得後に再度シートを取得して最新状態で更新）
    const freshPlayerSheet = ss.getSheetByName(SHEET_PLAYERS);
    if (!freshPlayerSheet) {
      ui.alert("エラー: プレイヤーシートが見つかりません。");
      return;
    }
    const { indices: freshIndices } = getSheetStructure(freshPlayerSheet, SHEET_PLAYERS);
    freshPlayerSheet.getRange(targetRowIndex, freshIndices["参加状況"] + 1).setValue(PLAYER_STATUS.WAITING);

    // 待機者が2人以上いれば自動マッチング
    const waitingPlayersCount = getWaitingPlayers().length;
    if (waitingPlayersCount >= 2) {
      Logger.log(`復帰後、待機プレイヤーが ${waitingPlayersCount} 人いるため、自動でマッチングを開始します。`);
      matchPlayers();
    }

    Logger.log(`プレイヤー ${playerId} を休憩から復帰させました。`);
  } catch (e) {
    ui.alert("エラーが発生しました: " + e.toString());
    Logger.log("returnPlayerFromResting エラー: " + e.toString());
  } finally {
    releaseLock(lock);
  }
}

// =========================================
// プレイヤーデータ取得・検索
// =========================================

/**
 * 待機中のプレイヤーを抽出し、以下の優先順位でソートして返します。
 * 1. 勝数（降順）
 * 2. 最終対戦時刻（降順 = 最近待機に戻った人優先 = 直近の勝者優先）
 */
function getWaitingPlayers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    if (!playerSheet) {
      Logger.log("エラー: プレイヤーシートが見つかりません。");
      return [];
    }
    const { indices, data } = getSheetStructure(playerSheet, SHEET_PLAYERS);
    if (data.length <= 1) return [];

    const waiting = data.slice(1).filter((row) => row[indices["参加状況"]] === PLAYER_STATUS.WAITING);

    waiting.sort((a, b) => {
      const winsDiff = b[indices["勝数"]] - a[indices["勝数"]];
      if (winsDiff !== 0) return winsDiff;

      const dateA = a[indices["最終対戦時刻"]] instanceof Date ? a[indices["最終対戦時刻"]].getTime() : 0;
      const dateB = b[indices["最終対戦時刻"]] instanceof Date ? b[indices["最終対戦時刻"]].getTime() : 0;
      return dateA - dateB; // 古い時刻が先（登録順・先着順）
    });

    return waiting;
  } catch (e) {
    Logger.log("getWaitingPlayers エラー: " + e.message);
    return [];
  }
}

// =========================================
// プレイヤー統計更新
// =========================================

/**
 * プレイヤーの統計情報 (勝数, 敗数, 試合数) と最終対戦時刻を更新します。
 */
function updatePlayerMatchStats(playerId, isWinner, timestamp) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    if (!playerSheet) {
      Logger.log("エラー: プレイヤーシートが見つかりません。");
      return;
    }
    const { indices, data } = getSheetStructure(playerSheet, SHEET_PLAYERS);
    if (data.length <= 1) return;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[indices["プレイヤーID"]] === playerId) {
        const rowNum = i + 1;
        const currentWins = parseInt(row[indices["勝数"]]) || 0;
        const currentLosses = parseInt(row[indices["敗数"]]) || 0;
        const currentTotal = parseInt(row[indices["試合数"]]) || 0;

        playerSheet.getRange(rowNum, indices["勝数"] + 1).setValue(currentWins + (isWinner ? 1 : 0));
        playerSheet.getRange(rowNum, indices["敗数"] + 1).setValue(currentLosses + (isWinner ? 0 : 1));
        playerSheet.getRange(rowNum, indices["試合数"] + 1).setValue(currentTotal + 1);
        playerSheet.getRange(rowNum, indices["最終対戦時刻"] + 1).setValue(timestamp);

        return;
      }
    }
    Logger.log(`エラー: プレイヤー ${playerId} が見つかりません。`);
  } catch (e) {
    Logger.log("updatePlayerStats エラー: " + e.message);
  }
}

// =========================================
// プレイヤー状態遷移
// =========================================

/**
 * 対戦中のプレイヤーの状態を変更します。
 * @param {Object} options - 状態変更のオプション
 * @param {string} options.targetPlayerId - 状態を変更するプレイヤーのID
 * @param {string} options.newStatus - 対象プレイヤーの新しい状態
 * @param {string} options.opponentNewStatus - 対戦相手の新しい状態
 * @param {boolean} options.recordResult - 結果を記録するかどうか
 * @param {boolean} options.isTargetWinner - 対象プレイヤーが勝者かどうか（結果記録時のみ使用）
 * @returns {Object} 処理結果 { success: boolean, message: string, opponentId?: string }
 */
function updatePlayerState(options) {
  const { targetPlayerId, newStatus, opponentNewStatus, recordResult = false, isTargetWinner = false } = options;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let matchLock = null;
  let stateLock = null;

  try {
    // 両方のロックを取得（順序を固定して デッドロック防止）
    stateLock = acquireLock("プレイヤー状態変更");
    matchLock = acquireLock("対戦結果の記録");

    // 1. プレイヤーの現在の状態を確認
    const playerSheet = ss.getSheetByName(SHEET_PLAYERS);
    if (!playerSheet) {
      Logger.log("エラー: プレイヤーシートが見つかりません。");
      return { success: false, message: "プレイヤーシートが見つかりません。" };
    }
    const { indices: playerIndices, data: playerData } = getSheetStructure(playerSheet, SHEET_PLAYERS);

    let targetFound = false;
    let currentStatus = null;

    for (let i = 1; i < playerData.length; i++) {
      const row = playerData[i];
      const playerId = row[playerIndices["プレイヤーID"]];

      if (playerId === targetPlayerId) {
        targetFound = true;
        currentStatus = row[playerIndices["参加状況"]];
        break;
      }
    }

    if (!targetFound) {
      return { success: false, message: `プレイヤー ${targetPlayerId} が見つかりません。` };
    }

    if (currentStatus === PLAYER_STATUS.DROPPED) {
      return { success: false, message: `プレイヤー ${targetPlayerId} はすでにドロップアウトしています。` };
    }

    // 3. 対戦中の場合のみ、対戦相手の処理
    let opponentId = null;
    let matchRow = -1;

    if (currentStatus === PLAYER_STATUS.IN_PROGRESS) {
      const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
      if (!inProgressSheet) {
        Logger.log("エラー: 対戦中シートが見つかりません。");
        return { success: false, message: "対戦中シートが見つかりません。" };
      }
      const { indices: matchIndices, data: matchData } = getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);

      for (let i = 1; i < matchData.length; i++) {
        const row = matchData[i];
        const p1 = row[matchIndices["ID1"]];
        const p2 = row[matchIndices["ID2"]];

        if (p1 === targetPlayerId) {
          opponentId = p2;
          matchRow = i + 1;
          break;
        } else if (p2 === targetPlayerId) {
          opponentId = p1;
          matchRow = i + 1;
          break;
        }
      }

      if (!opponentId) {
        return { success: false, message: `データ不整合: 対戦中のはずのプレイヤーID ${targetPlayerId} の対戦相手が見つかりません。` };
      }

      // 対戦相手の状態確認
      let opponentDropped = false;
      for (let i = 1; i < playerData.length; i++) {
        const row = playerData[i];
        if (row[playerIndices["プレイヤーID"]] === opponentId) {
          opponentDropped = row[playerIndices["参加状況"]] === PLAYER_STATUS.DROPPED;
          break;
        }
      }

      if (opponentDropped && opponentNewStatus !== PLAYER_STATUS.DROPPED) {
        return { success: false, message: `対戦相手はすでにドロップアウトしています。` };
      }
    }

    // 4. 結果の記録（必要な場合）
    if (recordResult) {
      const currentTime = new Date();
      const formattedTime = Utilities.formatDate(currentTime, "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
      const historySheet = ss.getSheetByName(SHEET_HISTORY);
      if (!historySheet) {
        Logger.log("エラー: 履歴シートが見つかりません。");
        return { success: false, message: "履歴シートが見つかりません。" };
      }
      getSheetStructure(historySheet, SHEET_HISTORY);
      const newId = "T" + Utilities.formatString("%04d", historySheet.getLastRow());

      const winner = isTargetWinner ? targetPlayerId : opponentId;
      const loser = isTargetWinner ? opponentId : targetPlayerId;
      const winnerName = getPlayerName(winner);
      const loserName = getPlayerName(loser);

      // マッチング中の卓番号を取得
      const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
      if (!inProgressSheet) {
        Logger.log("エラー: 対戦中シートが見つかりません。");
        return { success: false, message: "対戦中シートが見つかりません。" };
      }
      const { indices: matchIndices } = getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);
      const matchTableNumber = inProgressSheet.getRange(matchRow, matchIndices["卓番号"] + 1).getValue();
      const startTime = inProgressSheet.getRange(matchRow, matchIndices["対戦開始時刻"] + 1).getValue();
      const diffTime = Utilities.formatDate(new Date(currentTime.getTime() - startTime.getTime()), "UTC", "HH:mm:ss");

      historySheet.appendRow([newId, matchTableNumber, winner, winnerName, loser, loserName, formattedTime, diffTime]);

      updatePlayerMatchStats(winner, true, formattedTime);
      updatePlayerMatchStats(loser, false, formattedTime);
    }

    // 5. 対戦情報をクリア（対戦中の場合のみ）。卓番号は残す
    if (currentStatus === PLAYER_STATUS.IN_PROGRESS && matchRow !== -1) {
      const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);
      if (!inProgressSheet) {
        Logger.log("エラー: 対戦中シートが見つかりません。");
        return { success: false, message: "対戦中シートが見つかりません。" };
      }
      inProgressSheet.getRange(matchRow, 2, 1, 6).clearContent(); // ID1から対戦時間までをクリア
    }

    // 6. プレイヤーの状態を更新
    for (let i = 1; i < playerData.length; i++) {
      const row = playerData[i];
      const playerId = row[playerIndices["プレイヤーID"]];
      if (playerId === targetPlayerId) {
        playerSheet.getRange(i + 1, playerIndices["参加状況"] + 1).setValue(newStatus);
      } else if (playerId === opponentId) {
        playerSheet.getRange(i + 1, playerIndices["参加状況"] + 1).setValue(opponentNewStatus);
      }
    }

    // 7. 必要に応じて次のマッチング
    const waitingPlayersCount = getWaitingPlayers().length;
    if (waitingPlayersCount >= 2) {
      matchPlayers();
    }

    return {
      success: true,
      message: "状態変更が完了しました。",
      opponentId,
    };
  } catch (e) {
    Logger.log("handleMatchStateChange エラー: " + e.message);
    return {
      success: false,
      message: "エラーが発生しました: " + e.toString(),
    };
  } finally {
    releaseLock(matchLock);
    releaseLock(stateLock);
  }
}
