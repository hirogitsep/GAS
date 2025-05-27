function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('GドライブGrep検索アプリ')
    .setWidth(600)
    .setHeight(400);
}

function searchFiles(folderIds, keyword, mode, outputToSpreadsheet) {
  const results = [];
  const timestamp = new Date();
  let ss, resultSheet;

  if (outputToSpreadsheet) {
    ss = SpreadsheetApp.create('Grep検索結果_' + new Date().toISOString());
    resultSheet = ss.getActiveSheet();
    resultSheet.setName('Grep検索結果');
    resultSheet.appendRow([
      '検索日時', 'キーワード', '検索モード', 'ファイル名',
      'ファイル種類', 'フォルダID', 'フォルダ名', 'ファイルURL', 'スニペット'
    ]);
  }

  function isMatch(text) {
    if (!text) return false;
    try {
      switch (mode) {
        case 'exact': return text === keyword;
        case 'partial': return text.includes(keyword);
        case 'regex': return new RegExp(keyword).test(text);
        default: return false;
      }
    } catch (e) {
      return false;
    }
  }

  for (const folderId of folderIds) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        const mime = file.getMimeType();
        const fileName = file.getName();

        if (!(mime === MimeType.GOOGLE_DOCS || mime === MimeType.GOOGLE_SHEETS)) continue;

        let content = '';
        if (mime === MimeType.GOOGLE_DOCS) {
          try {
            content = DocumentApp.openById(file.getId()).getBody().getText();
          } catch (e) {
            continue;
          }
        } else {
          content = getSheetText(file.getId());
        }

        if (isMatch(fileName) || isMatch(content)) {
          const snippet = makeSnippet(content, keyword);
          const result = {
            name: fileName,
            url: file.getUrl(),
            snippet
          };
          results.push(result);

          if (outputToSpreadsheet) {
            resultSheet.appendRow([
              timestamp, keyword, mode, fileName,
              (mime === MimeType.GOOGLE_DOCS ? 'Googleドキュメント' : 'Googleスプレッドシート'),
              folderId, folder.getName(), file.getUrl(), snippet
            ]);
          }
        }
      }
    } catch (e) {
      Logger.log('フォルダ検索エラー: ' + e.message);
    }
  }

  return outputToSpreadsheet
    ? { results, spreadsheetUrl: ss.getUrl() }
    : { results };
}

function getSheetText(fileId) {
  try {
    const ss = SpreadsheetApp.openById(fileId);
    return ss.getSheets()
      .map(sheet => sheet.getDataRange().getDisplayValues()
        .map(row => row.join(' ')).join(' '))
      .join(' ');
  } catch (e) {
    return '';
  }
}

function makeSnippet(text, keyword) {
  const index = text.indexOf(keyword);
  if (index === -1) return '';
  const start = Math.max(0, index - 30);
  const end = Math.min(text.length, index + keyword.length + 30);
  return text.substring(start, end).replace(/\n/g, ' ');
}