var max_width = 20; // Configシートの横項目読み取り最大数

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
  let data1 = [];
  let in_data1 = [];
  for (let i = 0; i < max_width; i++)
    in_data1.push(i + 1);
  data1.push(in_data1); // 与えるデータは二次元配列
  const data2 = [['実施回'], ['時間帯'], ['場所'], ['班数'], ['統計区別']];
  sheet.getRange(2, 3, 1, max_width).setValues(data1);
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

  // Configの解析
  let rowNum = 3;
  let data1 = configSheet.getRange(rowNum++, 3, 1, max_width).getValues();
  let data2 = configSheet.getRange(rowNum++, 3, 1, max_width).getValues();
  let data3 = configSheet.getRange(rowNum++, 3, 1, max_width).getValues();
  let data4 = configSheet.getRange(rowNum++, 3, 1, max_width).getValues();
  let data5 = configSheet.getRange(rowNum++, 3, 1, max_width).getValues();
  let part = data1[0].filter(word => word != '');
  let section = data2[0].filter(word => word != '');
  let place = data3[0].filter(word => word != '');
  let group = data4[0].filter(word => word != '');
  let groupCount = [];
  for (let i = 0; i < group.length; i++)
    groupCount.push(parseInt(group[i], 10));
  let statisticOption = data5[0].filter(word => word != '');

  console.log(groupCount);

  // Baseの作成
  var totalStartColNum = 5;
}