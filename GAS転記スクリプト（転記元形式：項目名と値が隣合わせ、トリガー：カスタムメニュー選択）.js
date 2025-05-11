/*
 * ==========================================================================
 * File: GAS転記スクリプト（転記元形式：項目名と値が隣合わせ、トリガー：カスタムメニュー選択）.js
 * Created: 2025-05-11
 * Description: 
 * 転記用汎用スクリプト（SpreadSheet -> SpreadSheet）
 *  ・転記処理はカスタムメニューの選択から実行される
 *  ・転記元形式：項目名と値が隣合わせ、関連データが1シートにまとまっている
 *  ・転記先形式：項目名が横並び、関連データが1行ずつまとまっている
 *  ・TODOコメントを参照の上、PJに合わせて変更を行い、使用する
 *  ・renameSheetsByLookupKeyはシート名を検索キーの値に設定する関数（リポジトリ参照）
 *  ・listSheetNamesはシート名を一覧化し、各種処理のためチェックボックスを設けた関数（リポジトリ参照）
 *     ※GASでは別ファイルに記載した関数もimport文等書かずに使用できる
 *  ・GAS上での拡張子は.gs
 * ==========================================================================
 */
function transferData() {
  const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheets = sourceSpreadsheet.getSheets();

  // 転記先情報
  const destinationSpreadsheetId = '転記先URL'; 
  const destinationSpreadsheet = SpreadsheetApp.openById(destinationSpreadsheetId);
  const destinationSheet = destinationSpreadsheet.getSheetByName('転記先シート名'); 
  
  // 転記元項目名：転記先項目名マッピング（TODO: PJに合わせて適宜変更）
  const mapping = {
    "企業名": "会社名",
    "担当者名": "担当者名",
    "TEL": "電話番号",
  };

  const destData = destinationSheet.getDataRange().getValues();
  const headers = destData[0];

  sourceSheets.forEach(sheet => {
    if (sheet.getName() === destinationSheet.getName()) return;

    const valuesMap = getItemValueMap(sheet);
    // TODO: 検索キーはPJに合わせて適宜変更　※ここでは企業名（転記元項目名）
    const lookupKey = valuesMap["企業名"];
    if (!lookupKey) return;

    let targetRow = -1;

    for (let i = 1; i < destData.length; i++) {
      // TODO: 検索キーはPJに合わせて適宜変更　※ここでは会社名（転記先項目名）
      if (destData[i][headers.indexOf("会社名")] === lookupKey) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) {
      SpreadsheetApp.getUi().alert(`「${lookupKey}」が転記先に存在しません。スキップします。`);
      return;
    }

    for (let sourceKey in mapping) {
      const destKey = mapping[sourceKey];
      const colIndex = headers.indexOf(destKey);
      if (colIndex === -1) continue;

      const currentValue = destinationSheet.getRange(targetRow, colIndex + 1).getValue();
      const newValue = valuesMap[sourceKey];

      if (!currentValue && newValue !== "" && newValue !== null && newValue !== undefined) {
        destinationSheet.getRange(targetRow, colIndex + 1).setValue(newValue);
      }
    }
  });
  
  // シート名の変更（別ファイルに記載: 必要なければ削除）
  // renameSheetsByLookupKey();
  // シート名の一覧作成（別ファイルに記載: 必要なければ削除）
  // listSheetNames();
}

function getItemValueMap(sheet) {
  const values = sheet.getDataRange().getValues();
  let map = {};

  for (let row = 0; row < values.length; row++) {
    for (let col = 0; col < values[row].length - 1; col++) {
      const key = values[row][col];
      const value = values[row][col + 1];
      if (key && value && !(key in map)) {
        map[key] = value;
      }
    }
  }

  return map;
}