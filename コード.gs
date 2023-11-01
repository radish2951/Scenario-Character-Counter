const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const sheet = spreadsheet.getSheetByName("文字数");
const fileList = spreadsheet.getSheetByName("対象ファイルID").getDataRange().getValues().flat();

const START_ROW = 3;
const START_COLUMN = 1;
const COLUMNS_PER_FILE = 3;



function execute() {
  let updated = false;

  fileList.forEach((fileID, i) => {
    const column = START_COLUMN + i * COLUMNS_PER_FILE;
    const currentData = fetchCurrent(column);
    const updatedData = listFileRevisions(fileID, column, currentData);
    updated = updateSheet(column, updatedData, currentData) || updated;
  });

  if (updated) mergeRevisions();
}

function fetchCurrent(column) {
  let currentData = new Map();
  const values = sheet.getRange(START_ROW, column, sheet.getLastRow(), COLUMNS_PER_FILE).getValues();
  values.forEach(row => {
    const [revisionID, date, len] = row;
    if (revisionID) currentData.set(Number(revisionID), [new Date(date).getTime(), len]);
  });
  return currentData;
}

function listFileRevisions(fileID, column, currentData) {
  const revisions = Drive.Revisions.list(fileID).items;
  const updatedData = new Map(currentData);

  revisions.forEach(revision => {
    const revisionID = Number(revision.id);
    if (currentData.has(revisionID)) return;

    const doc = Drive.Revisions.get(fileID, revisionID);
    const date = new Date(doc.modifiedDate).getTime();
    const url = doc.exportLinks["text/plain"];
    const res = UrlFetchApp.fetch(url, { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() } });
    const len = res.getContentText().length;
    updatedData.set(revisionID, [date, len]);
  });

  return new Map([...updatedData].sort((a, b) => a[0] - b[0]));
}

function updateSheet(column, updatedData, currentData) {
  // 更新されていなかったら即座にfalseを返す
  if (updatedData.size <= currentData.size) {
    return false;
    // そうじゃない場合、更新されているので、更新してtrueを返す
  } else {
    sheet.getRange(START_ROW, column, updatedData.size, COLUMNS_PER_FILE)
      .setValues(Array.from(updatedData).map(([key, value]) => [key, new Date(value[0]), value[1]]));
    return true;
  }
}

function mergeRevisions() {

  const allData = sheet.getRange(START_ROW, START_COLUMN, sheet.getLastRow(), COLUMNS_PER_FILE * fileList.length).getValues();
  const data = new Map();

  fileList.forEach(fileID => {
    data.set(fileID, new Map());
  });

  allData.forEach(row => {
    for (let i = 0; i < fileList.length; i++) {
      const fileID = fileList[i];
      const startColumn = i * COLUMNS_PER_FILE;
      const [revisionID, date, len] = row.slice(startColumn, startColumn + COLUMNS_PER_FILE);
      if (!revisionID) continue;
      data.get(fileID).set(new Date(date).getTime(), len);
    }
  });


  let newData = new Map();

  // ファイルについてイテレート
  for (let [fileID, value] of data) {

    // あるファイルのリビジョンについてイテレート
    for (let [date, len] of value) {

      // このリビジョンの文字数
      let current = len;

      // 他のファイルを見に行く
      for (let [_fileID, _value] of data) {

        // 自分自身とは比較しない
        if (fileID == _fileID) continue;

        let lenResult = 0;

        // 他のファイルのリビジョンを一個ずつ確認
        for (let [_date, _len] of _value) {

          // 他のファイルの日付の方が古かったら採用
          if (_date < date) {
            lenResult = _len;

            // 古い方から見ていくので、超えたら即終わり。次のファイルへ
          } else {
            break;
          }
        }
        current += lenResult;
      }
      newData.set(new Date(date), current);
    }
  }

  const resultSheet = spreadsheet.getSheetByName("合計");
  const sortedData = Array.from(newData).sort((a, b) => a[0].getTime() - b[0].getTime());
  resultSheet.getRange(START_ROW, START_COLUMN, sortedData.length, 2).setValues(sortedData);
}



function countDialogues() {
  const SHEET_NAME = 'ワード数';

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  const docUrls = sheet.getRange("A:A").getValues().flat().filter(Boolean);  // "A:A"はURLが格納されている列を指定
  // const charNames = sheet.getRange("B1:AZ1").getValues().flat().filter(Boolean);  // "B1:Z1"はキャラ名が格納されている行を指定

  const counts = {};

  docUrls.forEach((docId, rowIndex) => {
    if (!docId) return;
    
    const doc = DocumentApp.openById(docId);
    const text = doc.getBody().getText();
    
    const regex = /(.+?)「.+」$/gm;
    let match;
    
    while ((match = regex.exec(text)) !== null) {
      const matchedNames = match[1].split(/[・？（）]/).filter(Boolean);
      
      matchedNames.forEach((matchedName) => {
        // if (!charNames.includes(matchedName)) return;

        if (!counts[matchedName]) {
          counts[matchedName] = {};
        }

        if (!counts[matchedName][docId]) {
          counts[matchedName][docId] = 0;
        }

        counts[matchedName][docId]++;

      });
    }
    
  });

  const charNames = Object.keys(counts);
  const excludedNames = sheet.getRange(20, 2).getValue().split(',');

  sheet.getRange(1, 2, 1, charNames.length).setValues([charNames]);
  
  charNames.forEach((charName, colIndex) => {
    if (excludedNames.includes(charName)) return;

    docUrls.forEach((docId, rowIndex) => {
      if (!docId) return;

      const count = counts[charName] && counts[charName][docId] || 0;
      sheet.getRange(rowIndex + 2, colIndex + 2).setValue(count);
    });
  });
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // メニュー項目を追加
  ui.createMenu('カスタムメニュー')
      .addItem('ワード数をカウント', 'countDialogues')
      .addToUi();
}
