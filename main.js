var max_width = 20; // Configシートの横項目読み取り最大数
var makeStatisticsSheets = true; // 出席率記入用シートを作成するか true/false
var makeAggregateSheet = true; // 全体集計シート（グラフ）を作成するか true/false

function setup() { // 初期設定
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let scriptSheet = ss.getSheetByName('Script');
  let configSheet = ss.getSheetByName('Config');
  if (scriptSheet === null) {
    scriptSheet = ss.insertSheet();
    scriptSheet.setName('Script');
  }
  if (configSheet === null) {
    configSheet = ss.insertSheet();
    configSheet.setName('Config');
  }

  // Configシート作成
  let data1 = [];
  let in_data1 = [];
  for (let i = 0; i < max_width; i++)
    in_data1.push(i + 1);
  data1.push(in_data1); // 与えるデータは二次元配列
  const data2 = [['実施回'], ['時間帯'], ['場所'], ['班数'], ['統計区別'], ['出席要素'], ['未処理要素']];
  configSheet.getRange(1, 2, 1, 1).setValue('シートを生成すると既存のシートは失われます．').setFontColor('red');
  configSheet.getRange(2, 3, 1, max_width).setValues(data1);
  configSheet.getRange(3, 2, data2.length, 1).setValues(data2);
  configSheet.getRange(3, 2, 2, 1).setFontColor('blue'); // この項目のみの入力で全体集計シートを作成可能
  configSheet.getRange(1, 6, 1, 1).setValue('Config，Script，Base，出席率集計は予約語です．シート名及びConfigの入力値として使用できません．').setFontColor('red');

  // Scriptシート作成
  scriptSheet.getRange(2, 2, 1, 1).setValue('Coded by chishige1217200');
  scriptSheet.getRange(3, 2, 1, 1).setValue('https://github.com/chishige1217200/AttendManageGAS');

  console.log('Configを記入してください．');
}

function createStatisticSheet() { // 集計シートの自動作成
  function backSum(num, flag) { // 全体合計処理に使用
    let sum = 0;
    for (let i = num.length - 1; i >= flag; i--)
      sum += num[i];
    return sum;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName('Config');
  if (configSheet === null) {
    configSheet = ss.insertSheet();
    configSheet.setName('Config');
    setup();
    return;
  }
  let baseSheet = ss.getSheetByName('Base');
  if (baseSheet !== null) ss.deleteSheet(baseSheet);
  baseSheet = ss.insertSheet();
  baseSheet.setName('Base');

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
  if (part.length <= 0) {
    console.error('実施回は1回以上ある必要があります．');
    return;
  }
  let section = data2[0].filter(word => word != ''); // 曜日時間帯（ex: 月曜前半）
  if (section.length <= 0) {
    console.error('時間帯は1つ以上ある必要があります．');
    return;
  }
  let place = data3[0].filter(word => word != ''); // 実施場所（ex: A教室）
  if (makeStatisticsSheets & place.length <= 0) {
    console.error('実施場所は1箇所以上ある必要があります．');
    return;
  }
  let group = data4[0].filter(word => word != ''); // 班数（ex: 3）
  let groupCount = []; // 教室ごとの班数をカウント
  for (let i = 0; i < group.length; i++) {
    groupCount.push(parseInt(group[i], 10)); // 10進数でパース
    if (makeStatisticsSheets & groupCount[i] <= 0) {
      console.error('班数には自然数を指定してください．');
      return;
    }
  }

  if (place.length !== group.length) { // 実施場所を入力したにも関わらず，班数を入力していない場合の例外
    console.error('実施場所と班数の組み合わせが1対1で対応していません．');
    return;
  }

  // 出席周りは例外処理がないため，入力に注意
  let statisticOption = data5[0].filter(word => word != ''); // 集計区分要素
  if (makeStatisticsSheets & statisticOption.length === 0) {
    console.error('統計区別に1つ以上の統計要素を指定してください．');
    return;
  }
  let attends = data6[0].filter(word => word != ''); // 出席と記録する要素
  if (makeStatisticsSheets & attends.length === 0) {
    console.error('出席要素に1つ以上の統計要素を指定してください．');
    return;
  }
  let unattends = data7[0].filter(word => word != ''); // 未処理とする要素
  if (makeStatisticsSheets & unattends.length === 0) {
    console.error('未処理要素に1つ以上の統計要素を指定してください．');
    return;
  }

  console.log('入力されたConfigの検証OK．');

  let statisticLines = []; //各実施時間帯の集計行をマーク
  let halfSectionCount = Math.ceil(section.length / 2); // 開始行の推定用（切り上げ）
  //console.log(halfSectionCount);

  if (makeStatisticsSheets) {
    // Baseの作成
    console.log('Baseシートの作成中...');
    baseSheet.getRange(1, 1, 1, 1).setValue('これは基準シートです．このシートが複製されます．');
    baseSheet.setFrozenRows(halfSectionCount + 2); // 行の表示範囲を固定

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
        baseSheet.getRange(totalStartRowNum + tableRowCount, 3 + statisticOption.length, 1, 1).setFormulaR1C1('=SUM(RC[' + (-statisticOption.length) + ']:RC[-1])');
        tableRowCount++;
      }
      baseSheet.getRange(totalStartRowNum + tableRowCount, 2, 1, 1).setValue(place[j] + '合計').setHorizontalAlignment('center'); // 実施場所毎合計部
      baseSheet.getRange(totalStartRowNum + tableRowCount, 3, 1, statisticOption.length).setFormulaR1C1('=SUM(R[' + (-groupCount[j]) + ']C:R[-1]C)');
      baseSheet.getRange(totalStartRowNum + tableRowCount, 3 + statisticOption.length, 1, 1).setFormulaR1C1('=SUM(RC[' + (-statisticOption.length) + ']:RC[-1])');
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
    baseSheet.getRange(totalStartRowNum + tableRowCount, 3 + statisticOption.length, 1, 1).setFormulaR1C1('=SUM(RC[' + (-statisticOption.length) + ']:RC[-1])');

    // ここに出席率と未処理 計を計算
    let attendsIndex = [];
    let unattendsIndex = [];

    for (let j = 0; j < attends.length; j++) {
      let index = statisticOption.findIndex(element => element === attends[j]);
      if (index === -1) {
        console.error('出席要素が見つかりません．統計区別と同じ要素が出席要素に記入されているか確認してください．');
        return;
      }
      else attendsIndex.push(index + 1);
    }
    for (let j = 0; j < unattends.length; j++) {
      let index = statisticOption.findIndex(element => element === unattends[j]);
      if (index === -1) {
        console.error('未処理要素が見つかりません．統計区別と同じ要素が未処理要素に記入されているか確認してください．');
        return;
      }
      else unattendsIndex.push(index + 1);
    }

    let attendsFormula = '=(';
    for (let j = 0; j < attendsIndex.length; j++) {
      attendsFormula += 'RC[' + (-2 - statisticOption.length + attendsIndex[j]) + ']';
      if (j + 1 !== attendsIndex.length) attendsFormula += '+';
    }
    attendsFormula += ')/RC[-1]';
    baseSheet.getRange(totalStartRowNum + tableRowCount, statisticOption.length + 4, 1, 1).setFormulaR1C1(attendsFormula).setNumberFormat("0%");

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
    //console.log(tableRowCount);
    //console.log(statisticLines);

    // 上部の集計部を生成
    let baseColumn = 1;
    let rowCount = 0;
    baseSheet.getRange(2, baseColumn + 1, 1, 2).setValues([['出席率', '未処理']]);
    for (let i = 0; i < section.length; i++) {
      baseSheet.getRange(rowCount + 3, baseColumn, 1, 1).setValue(section[i]);
      baseSheet.getRange(rowCount + 3, baseColumn + 1, 1, 1).setFormulaR1C1('=R' + statisticLines[i] + 'C' + (statisticOption.length + 4)).setNumberFormat("0%");
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

    console.log('Baseシートの作成完了．');
    console.log('シートの複製中...');

    // シート複製
    let completedsheet = [];
    let finalSheet;
    for (let i = 0; i < part.length; i++) {
      finalSheet = ss.getSheetByName(part[i]);
      if (finalSheet !== null) ss.deleteSheet(finalSheet);
      finalSheet = baseSheet.copyTo(ss);
      finalSheet.setName(part[i]);
      finalSheet.getRange(1, 1, 1, 1).setValue(part[i]);
      completedsheet.push(finalSheet);
    }
    console.log('シートの複製完了．');
  }

  if (makeAggregateSheet) { // 全体集計シートの作成
    console.log('全体集計シートの作成中...');
    let statisticSheet = ss.getSheetByName('出席率集計');
    if (statisticSheet !== null) ss.deleteSheet(statisticSheet);
    statisticSheet = ss.insertSheet('出席率集計');

    let sectionCol = [];
    for (let i = 0; i < section.length; i++) {
      sectionCol.push([section[i]]);
    }

    // 見出し行生成
    statisticSheet.getRange(2, 1, sectionCol.length, 1).setValues(sectionCol);
    statisticSheet.getRange(1, 2, 1, part.length).setValues([part]);

    // 1列ずつ処理
    for (let i = 0; i < part.length; i++) { // 実施回
      rowCount = 0;
      baseColumn = 2;
      // 1行ずつ処理
      for (let j = 0; j < section.length; j++) { // 実施時間帯
        statisticSheet.getRange(j + 2, i + 2, 1, 1).setFormulaR1C1('=\'' + part[i] + '\'!R' + (rowCount + 3) + 'C' + baseColumn).setNumberFormat("0%");
        rowCount++;
        if (j + 1 === halfSectionCount) {
          baseColumn += 4;
          rowCount = 0;
        }
      }
    }

    // グラフ描画
    let graphRange = statisticSheet.getRange(1, 1, section.length + 1, part.length + 1);
    let graph = statisticSheet.newChart().addRange(graphRange).setChartType(Charts.ChartType.LINE).setPosition(section.length + 3, 1, 0, 0).setNumHeaders(1).setOption('title', '出席率集計');
    statisticSheet.insertChart(graph.build());
    console.log('全体集計シートの作成完了．');
  }
}