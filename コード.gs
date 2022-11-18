let data = new Map();
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const sheet = spreadsheet.getSheetByName("文字数");
const range = spreadsheet.getSheetByName("対象ファイルID").getDataRange();
const fileList = range.getValues().map(value => { return value[0]; });
  


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

    const values = sheet.getRange(2, column, 1000, 3).getValues();

    for (const row of values) {

      const revisionID = row[0];
      const date = row[1];
      const len = row[2];

      if (!revisionID) continue;

      data.set(revisionID, {
        date: new Date(date).getTime(),
        len: len
      });
    }

}



function listFileRevisions(fileID, column) {

  
  const list = Drive.Revisions.list(fileID);
  const revisions = list.items;

  let row = 2;

  for (const revision of revisions) {

    const revisionID = revision.id;

    // すでに記録している場合はスキップ
    if (data.has(parseInt(revisionID))) {
      Logger.log("リビジョン %s をスキップします", revisionID);
      continue;
    }

    Logger.log("新しいリビジョン %s を取得するよ", revisionID);

    try {

    const doc = Drive.Revisions.get(fileID, revisionID);
    const date = new Date(doc.modifiedDate);
    const url = doc.exportLinks["text/plain"];
    const res = UrlFetchApp.fetch(url, {headers: {Authorization: "Bearer " + ScriptApp.getOAuthToken()}});
    const length = res.getContentText().length;
    // Logger.log('%s: %s 文字', date, length);

    data.set(revisionID, {
      date: date.getTime(),
      len: length
    });

    Logger.log("新しいリビジョン %s を取得しました", revisionID);

    } catch {
      Logger.log("リビジョンなさそう");

    }
  
  }

  data = new Map([...data].sort((a, b) => a - b));

}



function updateSheet(column) {

  let row = 2;

  for (let [revisionID, value] of data) {

    // const revisionFinder = sheet.getRange(2, column, 1000).createTextFinder(revisionID);

    const date = value.date;
    const len = value.len;

    // if (!revisionFinder.findNext()) {
      sheet.getRange(row, column).setValue(revisionID);
      sheet.getRange(row, column + 1).setValue(new Date(date));
      sheet.getRange(row, column + 2).setValue(len);
    // }

    row += 1;

  }

  sheet.getRange(2, column, 500, 3).sort(column);
  
}



function mergeRevisions() {

  data.clear();
  
  for (let i = 0; i < fileList.length; i++) {

    const fileID = fileList[i];

    data.set(fileID, new Map());

    const column = 1 + i * 3

    const values = sheet.getRange(2, column, 1000, 3).getValues();

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

  let row = 2;

  for (let [date, len] of newData) {

    sheet.getRange(row, 20).setValue(new Date(date));
    sheet.getRange(row, 20 + 1).setValue(len);

    row += 1;

  }
}
