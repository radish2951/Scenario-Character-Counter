let data = new Map();
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const sheet = spreadsheet.getSheetByName("文字数");
const range = spreadsheet.getSheetByName("対象ファイルID").getDataRange();
const fileList = range.getValues().map(row => { return row[0]; });

const START_ROW = 3;
  


function execute() {
  
  for (let i = 0; i < fileList.length; i++) {

    const column = 1 + i * 3;
    
    const fileID = fileList[i];
    // const name = DocumentApp.openById(fileID).getName();
    // sheet.getRange(1, column).setValue(name);

    fetchCurrent(column);
    
    listFileRevisions(fileID, column);
    
    updateSheet(column);
    

  }

  mergeRevisions();

}



function fetchCurrent(column) {

  data.clear();

    const values = sheet.getRange(START_ROW, column, 1000, 3).getValues();

    for (const row of values) {

      const revisionID = parseInt(row[0]);
      const date = new Date(row[1]).getTime();
      const len = row[2];

      if (!revisionID) continue;

      data.set(revisionID, {
        date: date,
        len: len
      });
    }

}



function listFileRevisions(fileID, column) {

  
  const list = Drive.Revisions.list(fileID);
  const revisions = list.items;

  let row = START_ROW;

  for (const revision of revisions) {

    const revisionID = parseInt(revision.id);

    // すでに記録している場合はスキップ
    if (data.has(revisionID)) {
      Logger.log("リビジョン %s をスキップします", revisionID);
      continue;
    }

    Logger.log("新しいリビジョン %s を取得するよ", revisionID);

    try {

    const doc = Drive.Revisions.get(fileID, revisionID);
    const date = new Date(doc.modifiedDate).getTime();
    const url = doc.exportLinks["text/plain"];
    const res = UrlFetchApp.fetch(url, {headers: {Authorization: "Bearer " + ScriptApp.getOAuthToken()}});
    const len = res.getContentText().length;
    // Logger.log('%s: %s 文字', date, length);

    data.set(revisionID, {
      date: date,
      len: len
    });

    Logger.log("新しいリビジョン %s を取得しました", revisionID);

    } catch {
      Logger.log("リビジョンなさそう");

    }
  
  }

  data = new Map([...data].sort((a, b) => a - b));

}



function updateSheet(column) {

  let row = START_ROW;

  let existsNewRevision = false;

  for (let [revisionID, value] of data) {

    const revisionFinder = sheet.getRange(START_ROW, column, 1000).createTextFinder(revisionID);

    const date = new Date(value.date);
    const len = value.len;

    if (!revisionFinder.findNext()) {

      sheet.getRange(row, column).setValue(revisionID);
      sheet.getRange(row, column + 1).setValue(date);
      sheet.getRange(row, column + 2).setValue(len);

      existsNewRevision = true;

      Logger.log("新しいリビジョン。ほんまか？");

    }

    row += 1;

  }

  if (existsNewRevision) {

    sheet.getRange(START_ROW, column, 1000, 3).sort(column);
  
  }

}



function mergeRevisions() {

  data.clear();
  
  for (let i = 0; i < fileList.length; i++) {

    const fileID = fileList[i];

    data.set(fileID, new Map());

    const column = 1 + i * 3

    const values = sheet.getRange(START_ROW, column, 1000, 3).getValues();

    for (const row of values) {

      const revisionID = row[0];
      const date = row[1];
      const len = row[2];

      if (!revisionID) continue;

      data.get(fileID).set(new Date(date).getTime(), len);
    }

  }

  let newData = new Map();

  for (let [fileID, value] of data) {

    const _data = value;
    
    // あるファイルのリビジョンを見ていく
    for (let [date, len] of _data) {

      // このファイルの文字数
      let current = len;

      // 他のファイルを見に行く
      for (let [_fileID, _value] of data) {

        // 自分自身とは比較しない
        if (fileID == _fileID) continue;

        let lenResult = 0;

        // 他のファイルのログを一個ずつ確認
        for (let [_date, _len] of _value) {

          // 比較対象より古かったら採用
          if (_date < date) {

            lenResult = _len;

          // 古い方から見ていくので、超えたら即終わり
          } else {

            break;

          }

        }

        current += lenResult;

      }

      newData.set(date, current);

    }

  }

  newData = new Map([...newData].sort());

  let row = START_ROW;

  const column = 1;

  const resultSheet = spreadsheet.getSheetByName("合計");

  for (let [date, len] of newData) {

    const currentDate = new Date(resultSheet.getRange(row, column).getValue()).getTime();

    if (currentDate != date) {

      resultSheet.getRange(row, column).setValue(new Date(date));
      resultSheet.getRange(row, column + 1).setValue(len);

    }

    row += 1;

  }
}
