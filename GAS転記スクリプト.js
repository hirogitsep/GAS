/*
 * ==========================================================================
 * File: GAS転記スクリプト.js
 * Created: 2025-05-10
 * Description: 
 * 転記用汎用スクリプト（SpreadSheet -> SpreadSheet）
 *  ・転記タイミングは手動実行またはGASのトリガー設定から行う
 *  ・TODOコメントを参照の上、PJに合わせて変更を行い、使用する
 *  ・初回のみsetProperties関数のコメントを外し、実行する
 *    （scriptPropertiesを現在に設定することで過去のデータが再転記されないようにする）
 * ==========================================================================
 */

// 初回実行時のみ下記3行のコメントアウトを外す(行選択 → Ctrl + /)
// function setProperties() {
//   PropertiesService.getScriptProperties().setScriptProperties(new Date());
//   return;
// }

function transferFormData() {
  const scriptProperties = PropertiesService.getScriptProperties();

  // 前回実行時刻を取得
  const lastRunStored = scriptProperties.getProperty('lastRunTime');  
  const lastRun = lastRunStored ? new Date(lastRunStored) : null;

  // 現在時刻を取得
  const now = new Date();

  // 転記元情報（TODO: PJに合わせて適宜変更）
  const sourceSpreadsheet = SpreadsheetApp.openById('転記元URL');
  const sourceSheet = sourceSpreadsheet.getSheetByName('転記元シート名');
  const sourceData = sourceSheet.getDataRange().getValues();
  const sourceHeaders = sourceData[0];

  // 転記先情報（TODO: PJに合わせて適宜変更）
  const targetSpreadsheet = SpreadsheetApp.openById('転記元URL');
  const targetSheet = targetSpreadsheet.getSheetByName('転記先シート名');
  const targetHeaders = targetSheet.getDataRange().getValues()[0];

  // 転記元項目名：転記先項目名マッピング（TODO: PJに合わせて適宜変更）
  const headerMapping = {
    'Timestamp': '入力日時',
    '名前': '会社名',
    '電話番号': '担当連絡先',
  };

  // 新規追加行
  const newRows = [];

  // 転記データを配列に格納
  for (let i = 1; i < sourceData.length; i++) {
    const row = sourceData[i];

    // 入力日時を取得
    const timestamp = lastRun ? new Date(row[sourceHeaders.indexOf('Timestamp')]) : null;

    // 新規追加データかどうかを判定
    const isNew = lastRun ? (timestamp > lastRun && timestamp <= now) : false;

    // 新規追加データであれば処理実行
    if (isNew) {
      const newRow = [];
      for (let j = 0; j < targetHeaders.length; j++) {
        const targetHeader = targetHeaders[j];
        const sourceKey = Object.keys(headerMapping).find(key => headerMapping[key] === targetHeader);

        if (sourceKey) {
          const sourceColIndex = sourceHeaders.indexOf(sourceKey);
          newRow[j] = sourceColIndex !== -1 ? row[sourceColIndex] : '';
        } else {
          newRow[j] = '';
        }
      }
      newRows.push(newRow);
    }
  }
  
  // 転記処理
  if (newRows.length > 0) {
    targetSheet.getRange(targetSheet.getLastRow() + 1, 1, newRows.length, targetHeaders.length).setValues(newRows);
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheetUrl = sheet.getUrl();

    // メール通知：転記データの件数を通知（TODO: メールアドレスを追加する）
    MailApp.sendEmail({
      to: '任意のメールアドレス',
      subject: `GAS 転記処理 実行結果 (${newRows.length} 件)`,
      body: `${newRows.length} 件の新規登録データを転記しました。\n\n${sheetUrl}\n\n実行時刻: ${now.toLocaleString()}\n前回実行: ${lastRun ? lastRun.toLocaleString() : '初回実行'}`
    });
  } else {
    // 新規データがない場合の通知（TODO: メールアドレスを追加する）
    MailApp.sendEmail({
      to: '任意のメールアドレス',
      subject: `GAS 転記処理 実行結果 (0 件)`,
      body: `新しく登録されたデータはありませんでした。\n\n実行時刻: ${now.toLocaleString()}\n前回実行: ${lastRun ? lastRun.toLocaleString() : '初回実行'}`
    });
  }

  // 処理完了後、lastRunTime を現在時刻に更新
  scriptProperties.setProperty('lastRunTime', now.toISOString());
}