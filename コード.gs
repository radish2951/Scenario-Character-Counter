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

  // if (updated) mergeRevisions();
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
      newData.set(new Date(date), current);
    }
  }

  const resultSheet = spreadsheet.getSheetByName("合計");
  const sortedData = Array.from(newData).sort((a, b) => new Date(a[0]) - new Date(b[0]));
  resultSheet.getRange(START_ROW, START_COLUMN, sortedData.length, 2).setValues(sortedData);
}