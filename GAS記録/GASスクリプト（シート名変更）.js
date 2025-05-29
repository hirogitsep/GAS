function renameSheetsByLookupKey() {

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const usedNames = new Set(sheets.map(s => s.getName()));

  for (const sheet of sheets) {
    const currentName = sheet.getName();
    // シート1は手動で任意のシート名に設定する
    if (currentName === "シート1") continue;

    const valuesMap = getItemValueMap(sheet);
    // TODO: PJに合わせて検索キーを変更 ※ここでは会社名
    const lookupKey = valuesMap["会社名"];
    if (!lookupKey || lookupKey === currentName) continue;

    let newSheetName = lookupKey.trim().substring(0, 30);
    newSheetName = newSheetName.replace(/[\\/?*[\]:]/g, '');

    let baseName = newSheetName;
    let suffix = 1;
    while (usedNames.has(newSheetName)) {
      newSheetName = `${baseName}_${suffix++}`;
    }

    try {
      sheet.setName(newSheetName);
      usedNames.add(newSheetName);
      Logger.log(`シート「${currentName}」を「${newSheetName}」に変更しました。`);
    } catch (e) {
      Logger.log(`シート名変更失敗: ${e}`);
    }
  }
}
