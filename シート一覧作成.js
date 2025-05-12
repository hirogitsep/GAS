function listSheetNames() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("シート1");
  const sheets = spreadsheet.getSheets();

  if (!sheet) {
    SpreadsheetApp.getUi().alert("「シート1」が存在しません。作成してください。");
    return;
  }

  // C列とD列の内容、書式、チェックボックスをクリア
  const range = sheet.getRange("C:D");
  range.clearContent();
  range.clearFormat();
  range.clearDataValidations();

  // ヘッダー設定
  sheet.getRange("C2").setValue("シート一覧");
  sheet.getRange("C2:D2").setBackground("#1E88E5").setFontColor("white").setFontWeight("bold");

  // 「すべて選択」のチェックボックスを追加（D2セル）
  const allCheckBoxCell = sheet.getRange("D2");
  allCheckBoxCell.insertCheckboxes();
  allCheckBoxCell.setValue(false);

  let row = 3;

  // シート名をリスト化し、チェックボックスを挿入
  for (const s of sheets) {
    const sName = s.getName();
    if (sName === "シート1") continue;

    sheet.getRange(row, 3).setValue(sName);
    sheet.getRange(row, 4).insertCheckboxes();
    row++;
  }

  // C列に入力したシート名の数を正確に取得（空白行を除外）
  const sheetNames = sheet.getRange("C3:C" + (row - 1)).getValues()
    .map(r => r[0])
    .filter(v => v !== "");

  const lastDataRow = sheetNames.length + 2;

  // 列幅を調整
  sheet.setColumnWidth(3, 200);

  // 枠線をC2:D(実データ範囲)に設定
  const dataRange = sheet.getRange(2, 3, sheetNames.length + 1, 2);
  dataRange.setBorder(true, true, true, true, false, false);
}

// すべて選択/解除を行う関数
function toggleAllCheckboxes() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName("シート1");

  // 「すべて選択」チェックボックスの状態を取得
  const allChecked = sheet.getRange("D2").getValue();

  // C列のデータ数から対象行数を決定
  const sheetNames = sheet.getRange("C3:C").getValues()
    .map(r => r[0])
    .filter(v => v !== "");
  const count = sheetNames.length;

  if (count === 0) return;

  // D列のチェックボックスに一括適用
  const checkboxesRange = sheet.getRange(3, 4, count, 1);
  const checkboxes = Array.from({ length: count }, () => [allChecked]);
  checkboxesRange.setValues(checkboxes);
}

// 「すべて選択」のチェックボックスが変更されたときに実行
function onEdit(e) {
  const sheet = e.source.getSheetByName("シート1");
  if (!sheet) return;

  if (e.range.getA1Notation() === "D2") {
    toggleAllCheckboxes();
  }
}