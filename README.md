// ==========================================
// Google Apps Script (程式碼.gs)
// ==========================================

// 1. 處理網頁讀取與自動排序的 doGet 函數
function doGet(e) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var records = [];
    
    if (data.length > 1) {
      for (var i = 1; i < data.length; i++) {
        if (data[i][0] || data[i][1] || data[i][2]) {
          records.push({
            id: String(data[i][0] || i),
            publisher: String(data[i][1] || ''),
            bookName: String(data[i][2] || ''),
            songNameZh: String(data[i][3] || ''),
            songNameEn: String(data[i][4] || ''),
            composer: String(data[i][5] || '')
          });
        }
      }
      
      // 自動排序邏輯：優先依「出版社」排序，其次「書名」，再其次「中文歌名」
      // 英文/數字按 A-Z 排序，中文按筆劃由少到多排序 (zh-HK / zh-TW 筆劃順序)
      records.sort(function(a, b) {
        var compPublisher = compareStrings(a.publisher, b.publisher);
        if (compPublisher !== 0) return compPublisher;
        
        var compBook = compareStrings(a.bookName, b.bookName);
        if (compBook !== 0) return compBook;
        
        return compareStrings(a.songNameZh || a.songNameEn, b.songNameZh || b.songNameEn);
      });
    }
    
    return ContentService.createTextOutput(JSON.stringify(records))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 輔助比較函數：自動判斷語言並進行 A-Z 或筆劃排序
function compareStrings(str1, str2) {
  if (!str1) str1 = "";
  if (!str2) str2 = "";
  return str1.localeCompare(str2, ['zh-HK', 'zh-TW', 'en'], { numeric: true, sensitivity: 'base' });
}

// 2. 處理新增、修改、刪除的 doPost 函數
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var contents = JSON.parse(e.postData.contents);
    var action = contents.action;

    if (action === 'add') {
      var newId = new Date().getTime().toString();
      sheet.appendRow([
        newId,
        contents.publisher || '',
        contents.bookName || '',
        contents.songNameZh || '',
        contents.songNameEn || '',
        contents.composer || ''
      ]);
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', id: newId }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'edit') {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(contents.id)) {
          sheet.getRange(i + 1, 2, 1, 5).setValues([[
            contents.publisher || '',
            contents.bookName || '',
            contents.songNameZh || '',
            contents.songNameEn || '',
            contents.composer || ''
          ]]);
          break;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'delete') {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(contents.id)) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ status: 'success' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
