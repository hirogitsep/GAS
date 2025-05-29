function deleteCheckedSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("シート1");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("削除対象がありません。");
    return;
  }

  // シート名はC列、チェックボックスはD列にある
  const names = sheet.getRange(2, 3, lastRow - 1, 1).getValues();
  const checks = sheet.getRange(2, 4, lastRow - 1, 1).getValues();

  // チェックされたシートを削除
  for (let i = 0; i < names.length; i++) {
    const sheetName = names[i][0];
    const checked = checks[i][0];

    // チェックされていて、削除対象のシート名が「シート1」でない場合に削除
    if (checked === true && sheetName !== "シート1") {
      const targetSheet = spreadsheet.getSheetByName(sheetName);
      if (targetSheet) {
        spreadsheet.deleteSheet(targetSheet);
      }
    }
  }

  SpreadsheetApp.getUi().alert("チェックされたシートを削除しました。");
  listSheetNames();
}
