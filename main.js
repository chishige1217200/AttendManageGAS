var max_width = 20; // Configシートの横項目読み取り最大数

function setup() { // 初期設定
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName('Config');
  if (configSheet === null) {
    configSheet = ss.insertSheet();
    configSheet.setName('Config');
    //configSheet = ss.getSheetByName('Config');
  }
  let baseSheet = ss.getSheetByName('Base');
  if (baseSheet === null) {
    baseSheet = ss.insertSheet();
    baseSheet.setName('Base');
    //baseSheet = ss.getSheetByName('Base');
  }

  setupConfig(configSheet);
}

function setupConfig(sheet) { // Configシートの自動作成
  let data1 = [];
  let in_data1 = [];
  for (let i = 0; i < max_width; i++)
    in_data1.push(i + 1);
  data1.push(in_data1); // 与えるデータは二次元配列
  const data2 = [['実施回'], ['時間帯'], ['場所'], ['班数'], ['統計区別'], ['出席要素'], ['未処理要素']];
  sheet.getRange(1, 2, 1, 1).setValue('シートを生成すると既存のシートは失われます．').setFontColor('red');
  sheet.getRange(2, 3, 1, max_width).setValues(data1);
  sheet.getRange(3, 2, data2.length, 1).setValues(data2);

  console.log('Configを記入してください．');
}

function backSum(num, flag) {
  let sum = 0;
  for (let i = num.length - 1; i >= flag; i--)
    sum += num[i];
  return sum;
}

function createBase() { // Baseシートの自動作成
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName('Config');
  if (configSheet === null) {
    configSheet = ss.insertSheet();
    configSheet.setName('Config');
    //configSheet = ss.getSheetByName('Config');
    setupConfig(configSheet);
    return;
  }
  let baseSheet = ss.getSheetByName('Base');
  if (baseSheet !== null) ss.deleteSheet(baseSheet);
  baseSheet = ss.insertSheet();
  baseSheet.setName('Base');
  //baseSheet = ss.getSheetByName('Base');

  // Configの解析
  let rowNum = 3; // 解析行番
  let data1 = configSheet.getRange(rowNum++, 3, 1, max_width).getValues(); // データ取得
  let data2 = configSheet.getRange(rowNum++, 3, 1, max_width).getValues();
  let data3 = configSheet.getRange(rowNum++, 3, 1, max_width).getValues();
  let data4 = configSheet.getRange(rowNum++, 3, 1, max_width).getValues();
  let data5 = configSheet.getRange(rowNum++, 3, 1, max_width).getValues();
  let data6 = configSheet.getRange(rowNum++, 3, 1, max_width).getValues();
  let data7 = configSheet.getRange(rowNum++, 3, 1, max_width).getValues();
  let part = data1[0].filter(word => word != ''); // 実施回（ex: 1st）
  if (part.length <= 0) console.error('実施回は1回以上ある必要があります．');
  let section = data2[0].filter(word => word != ''); // 曜日時間帯（ex: 月曜前半）
  if (section.length <= 0) console.error('時間帯は1つ以上ある必要があります．');
  let place = data3[0].filter(word => word != ''); // 実施場所（ex: A教室）
  if (place.length <= 0) console.error('実施場所は1箇所以上ある必要があります．');
  let group = data4[0].filter(word => word != ''); // 班数（ex: 3）
  let groupCount = []; // 教室ごとの班数をカウント
  for (let i = 0; i < group.length; i++) {
    groupCount.push(parseInt(group[i], 10)); // 10進数でパース
    if (groupCount[i] <= 0) console.error('班数には自然数を指定してください．');
  }

  // 出席周りは例外処理がないため，入力に注意
  let statisticOption = data5[0].filter(word => word != ''); // 集計区分要素
  let attends = data6[0].filter(word => word != ''); // 出席と記録する要素
  let unattends = data7[0].filter(word => word != ''); // 未処理とする要素

  // Baseの作成
  baseSheet.getRange(1, 1, 1, 1).setValue('これは基準シートです．このシートが複製されます．');
  let statisticLines = []; //各実施時間帯の集計行をマーク
  let halfSectionCount = Math.ceil(section.length / 2); // 開始行の推定用（切り上げ）
  console.log(halfSectionCount);

  let totalStartRowNum = 4 + halfSectionCount; // 1番目の表の開始行
  let tableRowCount = 0; // 表の行数カウント

  let placeStatisticFormula = '=';

  let firstLineArray = [section[0]];
  for (let l = 0; l < statisticOption.length; l++)
    firstLineArray.push(statisticOption[l]);
  firstLineArray.push('総数');
  firstLineArray.push('出席率');
  firstLineArray.push('未処理　計');
  firstLineArray = [firstLineArray];
  baseSheet.getRange(totalStartRowNum, 2, 1, firstLineArray[0].length).setValues(firstLineArray).setHorizontalAlignment('center');
  tableRowCount++;

  for (let j = 0; j < place.length; j++) { // 実施場所毎のループ（1つの表）
    baseSheet.getRange(totalStartRowNum + tableRowCount, 3, groupCount[j], statisticOption.length).setBackground('aqua'); // 色をつける
    for (let k = 0; k < groupCount[j]; k++) { // 実施場所入力行生成部
      baseSheet.getRange(totalStartRowNum + tableRowCount, 2, 1, 1).setValue(place[j] + (k + 1) + '班').setHorizontalAlignment('center');
      baseSheet.getRange(totalStartRowNum + tableRowCount, 3 + statisticOption.length, 1, 1).setFormulaR1C1('=SUM(RC[' + (-statisticOption.length) + '],RC[-1])');
      tableRowCount++;
    }
    baseSheet.getRange(totalStartRowNum + tableRowCount, 2, 1, 1).setValue(place[j] + '合計').setHorizontalAlignment('center'); // 実施場所毎合計部
    baseSheet.getRange(totalStartRowNum + tableRowCount, 3, 1, statisticOption.length).setFormulaR1C1('=SUM(R[' + (-groupCount[j]) + ']C,R[-1]C)');
    baseSheet.getRange(totalStartRowNum + tableRowCount, 3 + statisticOption.length, 1, 1).setFormulaR1C1('=SUM(RC[' + (-statisticOption.length) + '],RC[-1])');
    tableRowCount++;
  }

  // 合計計算（相対仕様に変更）
  baseSheet.getRange(totalStartRowNum + tableRowCount, 2, 1, 1).setValue('合計').setFontColor('red').setHorizontalAlignment('center');
  for (let j = place.length - 1; j >= 0; j--) {
    let back = 0;
    if (j === place.length - 1) back = -1;
    else
      back = -backSum(groupCount, j + 1) - (place.length - j);
    placeStatisticFormula += 'R[' + back + ']C';
    if (j !== 0) placeStatisticFormula += '+';
  }
  baseSheet.getRange(totalStartRowNum + tableRowCount, 3, 1, statisticOption.length).setFormulaR1C1(placeStatisticFormula); // 要素ごとの最終合計
  baseSheet.getRange(totalStartRowNum + tableRowCount, 3 + statisticOption.length, 1, 1).setFormulaR1C1('=SUM(RC[' + (-statisticOption.length) + '],RC[-1])');

  // ここに出席率と未処理 計を計算
  let attendsIndex = [];
  let unattendsIndex = [];

  for (let j = 0; j < attends.length; j++) {
    let index = statisticOption.findIndex(element => element === attends[j]);
    if (index === -1) console.error('出席要素が見つかりません');
    else attendsIndex.push(index + 1);
  }
  for (let j = 0; j < unattends.length; j++) {
    let index = statisticOption.findIndex(element => element === unattends[j]);
    if (index === -1) console.error('未処理要素が見つかりません');
    else unattendsIndex.push(index + 1);
  }

  let attendsFormula = '=(';
  for (let j = 0; j < attendsIndex.length; j++) {
    attendsFormula += 'RC[' + (-2 - statisticOption.length + attendsIndex[j]) + ']';
    if (j + 1 !== attendsIndex.length) attendsFormula += '+';
  }
  attendsFormula += ')/RC[-1]';
  baseSheet.getRange(totalStartRowNum + tableRowCount, statisticOption.length + 4, 1, 1).setFormulaR1C1(attendsFormula);

  let unattendsFormula = '=';
  for (let j = 0; j < unattendsIndex.length; j++) {
    unattendsFormula += 'RC[' + (-3 - statisticOption.length + unattendsIndex[j]) + ']';
    if (j + 1 !== unattendsIndex.length) unattendsFormula += '+';
  }
  baseSheet.getRange(totalStartRowNum + tableRowCount, statisticOption.length + 5, 1, 1).setFormulaR1C1(unattendsFormula);

  statisticLines.push(tableRowCount + totalStartRowNum);
  tableRowCount++;

  baseSheet.getRange(totalStartRowNum, 2, tableRowCount, statisticOption.length + 2).setBorder(true, true, true, true, true, true); // 枠線を引く

  // 表の複製処理
  for (let i = 1; i < section.length; i++) {
    baseSheet.getRange(totalStartRowNum, 2, tableRowCount, statisticOption.length + 5).copyTo(baseSheet.getRange(totalStartRowNum + (tableRowCount + 2) * i, 2, tableRowCount, statisticOption.length + 5));
    baseSheet.getRange(totalStartRowNum + (tableRowCount + 2) * i, 2, 1, 1).setValue(section[i]);
    statisticLines.push(totalStartRowNum + (tableRowCount + 2) * i + tableRowCount - 1);
  }
  console.log(tableRowCount);
  console.log(statisticLines);

  // 上部の集計部を生成
  let baseColumn = 1;
  let rowCount = 0;
  baseSheet.getRange(2, baseColumn + 1, 1, 2).setValues([['出席率', '未処理']]);
  for (let i = 0; i < section.length; i++) {
    baseSheet.getRange(rowCount + 3, baseColumn, 1, 1).setValue(section[i]);
    baseSheet.getRange(rowCount + 3, baseColumn + 1, 1, 1).setFormulaR1C1('=R' + statisticLines[i] + 'C' + (statisticOption.length + 4));
    baseSheet.getRange(rowCount + 3, baseColumn + 2, 1, 1).setFormulaR1C1('=R' + statisticLines[i] + 'C' + (statisticOption.length + 5));
    // 代入処理
    rowCount++;
    if (i + 1 === halfSectionCount) {
      baseColumn += 4;
      rowCount = 0;
      baseSheet.getRange(2, baseColumn + 1, 1, 2).setValues([['出席率', '未処理']]);
    }
  }

  baseColumn += 4;
  baseSheet.getRange(1, baseColumn, 1, 1).setValue('計');
  baseSheet.getRange(2, baseColumn, 1, statisticOption.length).setValues([statisticOption]).setHorizontalAlignment('center');
  baseSheet.getRange(2, baseColumn + statisticOption.length, 1, 1).setValue('総数').setHorizontalAlignment('center');

  for (let i = 0; i < statisticOption.length; i++) {
    let allStatisticFormula = '=';
    for (let j = 0; j < statisticLines.length; j++) {
      allStatisticFormula += 'R' + statisticLines[j] + 'C' + (i + 3);
      if (j + 1 !== statisticLines.length) allStatisticFormula += '+';
    }
    baseSheet.getRange(3, baseColumn + i, 1, 1).setFormulaR1C1(allStatisticFormula);
  }
  baseSheet.getRange(3, baseColumn + statisticOption.length, 1, 1).setFormulaR1C1('=SUM(RC[' + (-statisticOption.length) + ']:RC[-1])');

  // シート複製
  let completedsheet = [];
  let finalSheet;
  for (let i = 0; i < part.length; i++) {
    finalSheet = ss.getSheetByName(part[i]);
    if (finalSheet !== null) ss.deleteSheet(finalSheet);
    finalSheet = baseSheet.copyTo(ss);
    finalSheet.setName(part[i]);
    //finalSheet = ss.getSheetByName('Base');
    finalSheet.getRange(1, 1, 1, 1).setValue(part[i]);
    completedsheet.push(finalSheet);
  }
}