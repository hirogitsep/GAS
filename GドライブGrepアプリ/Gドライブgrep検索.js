// Webアプリのエントリポイント
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('GドライブGrep検索アプリ')
    .setWidth(600)
    .setHeight(400);
}

// 検索関数：フォルダID群、検索キーワード、モード、結果をスプレッドシートに出すか
function searchFiles(folderIds, keyword, mode, outputToSpreadsheet) {
  const results = [];
  const timestamp = new Date();
  const { matcher, regex } = createMatcher(keyword, mode);
  let ss, resultSheet;

  if (outputToSpreadsheet) {
    ss = SpreadsheetApp.create(`Grep検索結果_${timestamp.toISOString()}`);
    resultSheet = ss.getActiveSheet();
    resultSheet.setName('Grep検索結果');
    resultSheet.appendRow([
      '検索日時', 'キーワード', '検索モード', 'ファイル名',
      'ファイル種類', 'フォルダID', 'フォルダ名', 'ファイルURL', 'スニペット'
    ]);
  }

  for (const folderId of folderIds) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        const mime = file.getMimeType();

        if (![MimeType.GOOGLE_DOCS, MimeType.GOOGLE_SHEETS].includes(mime)) continue;

        const fileId = file.getId();
        const fileName = file.getName();
        const fileUrl = file.getUrl();
        let content = '', matchedText = null;

        if (mime === MimeType.GOOGLE_DOCS) {
          try {
            content = DocumentApp.openById(fileId).getBody().getText();
          } catch (e) {
            continue;  // 読み込み失敗はスキップ
          }
        } else {  // Google Sheets
          matchedText = findMatchInSheet(fileId, matcher);
          content = extractSheetText(fileId);
        }

        // ファイル名 OR シート内テキスト OR (Google Docs本文)でマッチしたら
        if (matcher(fileName) || matchedText !== null || matcher(content)) {
          const snippet = matchedText || extractSnippet(content, keyword, regex);

          results.push({ name: fileName, url: fileUrl, snippet });

          if (outputToSpreadsheet) {
            resultSheet.appendRow([
              timestamp, keyword, mode, fileName,
              mime === MimeType.GOOGLE_DOCS ? 'Googleドキュメント' : 'Googleスプレッドシート',
              folderId, folder.getName(), fileUrl, snippet
            ]);
          }
        }
      }
    } catch (e) {
      Logger.log(`フォルダ検索エラー: ${e.message}`);
    }
  }

  return outputToSpreadsheet
    ? { results, spreadsheetUrl: ss.getUrl() }
    : { results };
}

function createMatcher(keyword, mode) {
  if (mode === 'regex') {
    let pattern = keyword;
    let flags = 'u';

    const regexInput = keyword.match(/^\/(.+)\/([gimsuy]*)$/);
    if (regexInput) {
      pattern = regexInput[1];
      flags = regexInput[2] || 'u';
    }

    try {
      const regex = new RegExp(pattern, flags);
      return {
        matcher: text => typeof text === 'string' && regex.test(text),
        regex
      };
    } catch (e) {
      Logger.log(`正規表現エラー: ${e.message}`);
      return { matcher: () => false, regex: null };
    }
  }

  if (mode === 'exact') {
    return { matcher: text => text === keyword, regex: null };
  }

  if (mode === 'partial') {
    return { matcher: text => typeof text === 'string' && text.includes(keyword), regex: null };
  }

  return { matcher: () => false, regex: null };
}



// シート内のテキストをすべて結合して取得（スニペット作成用）
function extractSheetText(fileId) {
  try {
    return SpreadsheetApp.openById(fileId).getSheets()
      .map(sheet => sheet.getDataRange().getDisplayValues()
        .flat().join(' ')).join(' ');
  } catch (e) {
    return '';
  }
}

// マッチしたキーワードの前後30文字ずつを切り出してスニペットを作成
// 正規表現オブジェクト(regex)があれば使い、なければkeyword文字列で検索
function extractSnippet(text, keyword, regex = null) {
  try {
    if (!text) return '';
    let index = -1;
    let matchLength = keyword.length;

    if (regex) {
      const match = text.match(regex);
      if (match && match.index !== undefined) {
        index = match.index;
        matchLength = match[0].length;
      }
    } else {
      index = text.indexOf(keyword);
    }

    if (index === -1) return '';
    const start = Math.max(0, index - 30);
    const end = Math.min(text.length, index + matchLength + 30);
    return text.substring(start, end).replace(/\n/g, ' ');
  } catch (e) {
    return '';
  }
}

function findMatchInSheet(fileId, matcher) {
  try {
    const sheets = SpreadsheetApp.openById(fileId).getSheets();
    for (const sheet of sheets) {
      const data = sheet.getDataRange().getDisplayValues();
      for (const row of data) {
        for (const cell of row) {
          if (matcher(cell)) {
            Logger.log(`マッチ: ${cell}`);
            return cell;
          }
        }
      }
    }
  } catch (e) {
    Logger.log(`シートマッチエラー: ${e.message}`);
  }
  return null;
}