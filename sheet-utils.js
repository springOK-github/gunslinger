/**
 * ポケモンカード・ガンスリンガーバトル用マッチングシステム
 * @fileoverview スプレッドシートの操作に関するユーティリティ関数
 * @author SpringOK
 */

/**
 * シートのヘッダーを検証し、列インデックスを返します。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 検証対象のシート
 * @param {string} sheetName - シート名（SHEET_PLAYERS等の定数）
 * @returns {{headers: string[], indices: Object.<string, number>, data: any[][]}} ヘッダー情報と全データ
 * @throws {Error} 必須ヘッダーが不足している場合
 */
function getSheetStructure(sheet, sheetName) {
  if (!sheet) {
    throw new Error(`シート「${sheetName}」が見つかりません。`);
  }

  const data = sheet.getDataRange().getValues();
  if (!data || data.length === 0) {
    throw new Error(`シート「${sheetName}」にデータがありません。`);
  }

  const headers = data[0].map(h => String(h).trim());
  const indices = {};
  const missing = [];
  
  const requiredHeaders = REQUIRED_HEADERS[sheetName];
  if (!requiredHeaders) {
    throw new Error(`シート「${sheetName}」の必須ヘッダー定義が見つかりません。`);
  }

  requiredHeaders.forEach(required => {
    const idx = headers.indexOf(required);
    if (idx === -1) {
      missing.push(required);
    } else {
      indices[required] = idx;
    }
  });

  if (missing.length > 0) {
    throw new Error(`シート「${sheetName}」に必須ヘッダーが不足しています: ${missing.join(", ")}`);
  }

  return { headers, indices, data };
}

/**
 * 「マッチング」シート内の空行（対戦が終了し、コンテンツがクリアされた行）を削除し、
 * シート内のデータを上詰めして整理します。
 */
function cleanUpInProgressSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inProgressSheet = ss.getSheetByName(SHEET_IN_PROGRESS);

  try {
    getSheetStructure(inProgressSheet, SHEET_IN_PROGRESS);

    const lastRow = inProgressSheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log(`「${SHEET_IN_PROGRESS}」シートにデータがないため、整理は不要です。`);
      return;
    }

    let deletedCount = 0;
    for (let i = lastRow; i >= 2; i--) {
      const cellA = inProgressSheet.getRange(i, 1).getValue();
      if (cellA === "") {
        inProgressSheet.deleteRow(i);
        deletedCount++;
      }
    }

    if (deletedCount > 0) {
      Logger.log(`「${SHEET_IN_PROGRESS}」シートの整理 (自動実行) が完了しました。${deletedCount} 行の空行を削除しました。`);
    }
  } catch (e) {
    Logger.log("cleanUpInProgressSheet エラー: " + e.message);
  }
}