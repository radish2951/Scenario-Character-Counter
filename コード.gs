let data = new Map();
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const sheet = spreadsheet.getSheetByName("文字数");
const range = spreadsheet.getSheetByName("対象ファイルID").getDataRange();
const IDList = range.getValues().map(value => { return value[0]; });
  

function execute() {
  
  for (let i = 0; i < IDList.length; i++) {

    //if (i < 3) continue;

    const column = 1 + i * 2;

    data.clear();

    const values = sheet.getRange(2, column, 500, 2).getValues();

    for (const row of values) {

      if (!row[0]) continue;

      data.set((new Date(row[0])).getTime(), row[1]);
    }

    const fileID = IDList[i];
    // const name = DocumentApp.openById(fileID).getName();
    // sheet.getRange(1, column).setValue(name);

    listFileRevisions(fileID, column);
    updateSheet(column);
    

  }

  mergeRevisions();
}

function listFileRevisions(fileID, column) {

  
  const list = Drive.Revisions.list(fileID);
  const revisions = list.items;

  Logger.log(list.nextPageToken);

  let row = 2;

  for (const revision of revisions) {

    try {

    const revisionID = revision.id;
    const doc = Drive.Revisions.get(fileID, revisionID);
    const date = new Date(doc.modifiedDate);
    const url = doc.exportLinks["text/plain"];
    const res = UrlFetchApp.fetch(url, {headers: {Authorization: "Bearer " + ScriptApp.getOAuthToken()}});
    const length = res.getContentText().length;
    Logger.log('%s: %s 文字', date, length);

    data.set(date.getTime(), length);

    } catch {
      Logger.log("リビジョンエラー出たね");
    }

    //sheet.getRange(row, column).setValue(date);
    //sheet.getRange(row, column + 1).setValue(length);

    // row += 1;
  
  }

}

function updateSheet(column) {

  let row = 2;

  for (let [key, value] of data) {
    sheet.getRange(row, column).setValue(new Date(key));
    sheet.getRange(row, column + 1).setValue(value);

    row += 1;

  }

  sheet.getRange(2, column, 500, 2).sort(column)
}

function mergeRevisions() {

  data.clear();
  
  for (let i = 0; i < IDList.length; i++) {

    const fileID = IDList[i];

    data.set(fileID, new Map());

    const column = 1 + i * 2;

    const values = sheet.getRange(2, column, 500, 2).getValues();

    for (const row of values) {

      if (!row[0]) continue;

      data.get(fileID).set((new Date(row[0])).getTime(), row[1]);
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

  let row = 2;
  let column = 15;

  for (let [key, value] of newData) {
    sheet.getRange(row, column).setValue(new Date(key));
    sheet.getRange(row, column + 1).setValue(value);

    row += 1;

  }

  sheet.getRange(2, column, 500, 2).sort(column)



}


