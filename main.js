function setup() { // 初期設定
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName('Config');
  if (configSheet == null) {
    configSheet = ss.insertSheet();
    configSheet.setName('Config');
    configSheet = ss.getSheetByName('Config');
  }
  let baseSheet = ss.getSheetByName('Base');
  if (baseSheet == null) {
    baseSheet = ss.insertSheet();
    baseSheet.setName('Base');
    baseSheet = ss.getSheetByName('Base');
  }

  setupConfig(configSheet);
}

function setupConfig(sheet) { // Configシートの自動作成
  const data1 = [[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]];
  const data2 = [['実施回'], ['時間帯'], ['場所'], ['班数'], ['集計区別']];
  sheet.getRange(2, 3, 1, 15).setValues(data1);
  sheet.getRange(3, 2, 5, 1).setValues(data2);
}

function createBase() { // Baseシートの自動作成
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName('Config');
  if (configSheet == null) {
    configSheet = ss.insertSheet();
    configSheet.setName('Config');
    configSheet = ss.getSheetByName('Config');
  }
  let baseSheet = ss.getSheetByName('Base');
  if (baseSheet == null) {
    baseSheet = ss.insertSheet();
    baseSheet.setName('Base');
    baseSheet = ss.getSheetByName('Base');
  }
  let rowNum = 3;
  let data1 = configSheet.getRange(rowNum++, 3, 1, 15).getValues();
  let data2 = configSheet.getRange(rowNum++, 3, 1, 15).getValues();
  let data3 = configSheet.getRange(rowNum++, 3, 1, 15).getValues();
  let data4 = configSheet.getRange(rowNum++, 3, 1, 15).getValues();
  let data5 = configSheet.getRange(rowNum++, 3, 1, 15).getValues();
  console.log(data1);
}

function delNullValue(data) { // 配列の末尾要素がNULLなら除去
  return data;
}