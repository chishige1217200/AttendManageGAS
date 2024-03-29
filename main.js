var maxWidth = 20; // Configシートの横項目読み取り最大数
var makeStatisticsSheets = true; // 出席率記入用シートを作成するか true/false
var makeAggregateSheet = false; // 全体集計シート（グラフ）を作成するか true/false

// 集計区分のプルダウンリスト項目設定
var statisticClass = ['出席', '欠席', '未処理', '集計除外'];

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
  for (let i = 0; i < maxWidth; i++)
    in_data1.push(i + 1);
  data1.push(in_data1); // 与えるデータは二次元配列
  const data2 = [['実施回'], ['時間帯'], ['場所'], ['集計分類'], ['集計区分']];
  configSheet.getRange(1, 2, 1, 1).setValue('シートを生成すると既存のシートは失われます．').setFontColor('red');
  configSheet.getRange(2, 3, 1, maxWidth).setValues(data1);
  configSheet.getRange(3, 2, data2.length, 1).setValues(data2);
  configSheet.getRange(3, 2, 2, 1).setFontColor('blue'); // この項目のみの入力で全体集計シートを作成可能
  configSheet.getRange(1, 6, 1, 1).setValue('Config，Script，Base，出席率集計は予約語です．シート名及びConfigの入力値として使用できません．').setFontColor('red');

  // 集計区分のプルダウンリスト項目設定
  let statisticRule = SpreadsheetApp.newDataValidation().requireValueInList(statisticClass).build();
  configSheet.getRange(7, 3, 1, maxWidth).setDataValidation(statisticRule);

  // Scriptシート作成
  scriptSheet.getRange(2, 2, 1, 1).setValue('Coded by chishige1217200');
  scriptSheet.getRange(3, 2, 1, 1).setValue('https://github.com/chishige1217200/AttendManageGAS');

  console.log('Configを記入してください．');
}

function createStatisticSheet() { // 集計シートの自動作成
  // Configシートの情報を取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName('Config');
  if (configSheet === null) {
    configSheet = ss.insertSheet();
    configSheet.setName('Config');
    setup();
    return;
  }

  // Configの解析
  let rowNum = 3; // 解析行番
  let data1 = configSheet.getRange(rowNum++, 3, 1, maxWidth).getValues(); // データ取得
  let data2 = configSheet.getRange(rowNum++, 3, 1, maxWidth).getValues();
  let data3 = configSheet.getRange(rowNum++, 3, 1, maxWidth).getValues();
  let data4 = configSheet.getRange(rowNum++, 3, 1, maxWidth).getValues();
  let data5 = configSheet.getRange(rowNum++, 3, 1, maxWidth).getValues();

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

  let statisticOption = data4[0].filter(word => word != ''); // 集計分類要素
  if (makeStatisticsSheets & statisticOption.length === 0) {
    console.error('集計分類に1つ以上の分類要素を指定してください．');
    return;
  }

  let statisticRule = data5[0].filter(word => word != ''); // 集計区分要素
  if (makeStatisticsSheets & statisticRule.length === 0) {
    console.error('集計区分に1つ以上の集計区分を指定してください．');
    return;
  }

  if (makeStatisticsSheets & statisticOption.length != statisticRule.length) {
    console.error('集計分類と集計区分の個数が一致しません．個数を合わせて実行してください．');
    return;
  }

  // ここに出席者，欠席者，未処理者，集計除外者のセルの場所を取得
  let attendsIndex = [];
  let absentsIndex = [];
  let unattendsIndex = [];
  let ignoreIndex = [];

  for (let j = 0; j < statisticRule.length; j++)
    if (statisticRule[j] === statisticClass[0])
      attendsIndex.push(j + 1);

  for (let j = 0; j < statisticRule.length; j++)
    if (statisticRule[j] === statisticClass[1])
      absentsIndex.push(j + 1);

  for (let j = 0; j < statisticRule.length; j++)
    if (statisticRule[j] === statisticClass[2])
      unattendsIndex.push(j + 1);

  for (let j = 0; j < statisticRule.length; j++)
    if (statisticRule[j] === statisticClass[3])
      ignoreIndex.push(j + 1);

  if (attendsIndex.length === 0) {
    console.error('出席区分の集計分類が存在しません．1つ以上の出席区分の要素が必要です．');
    return;
  }

  if (unattendsIndex.length === 0) {
    console.error('未処理区分の集計分類が存在しません．1つ以上の未処理区分の要素が必要です．');
    return;
  }

  console.log('入力されたConfigの検証OK．');

  let linkLines = []; //各実施時間帯の初めの行をマーク
  let statisticLines = []; //各実施時間帯の集計行をマーク
  let halfSectionCount = Math.ceil(section.length / 2); // 上部集計行の折り返し推定用（切り上げ）
  //console.log(halfSectionCount);

  if (makeStatisticsSheets) {
    // Baseの作成

    // 衝突するシートが存在するか確認
    let sheetExist = 0;
    for (let i = 0; i < part.length; i++) {
      let checkSheet = ss.getSheetByName(part[i]);
      if (checkSheet !== null) sheetExist++;
    }
    if (sheetExist > 0) {
      console.log('作成済みのシートを検出しました．処理を実行するには，\"スプレッドシート\"のメッセージウィンドウから許可する必要があります．');
      let wantcontinue = Browser.msgBox('作成済みのシートが存在します．実行すると作成済みのシートが上書きされます．それでも実行しますか？', Browser.Buttons.YES_NO);

      if (wantcontinue === null) console.log('メッセージウィンドウが表示されない場合は，Google Chromeを使用してみてください．');
      if (wantcontinue === 'no' || wantcontinue === 'cancel' || wantcontinue === null) return;
    }

    let baseSheet = ss.getSheetByName('Base');
    if (baseSheet !== null) ss.deleteSheet(baseSheet);
    baseSheet = ss.insertSheet();
    baseSheet.setName('Base');
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
    firstLineArray.push('総出席率');
    firstLineArray.push('純出席率');
    firstLineArray.push('未処理　計');
    firstLineArray = [firstLineArray];
    baseSheet.getRange(totalStartRowNum, 2, 1, firstLineArray[0].length).setValues(firstLineArray).setHorizontalAlignment('center');
    linkLines.push(totalStartRowNum + tableRowCount);
    tableRowCount++;

    // 総出席率計算式
    let attendsFormula = '=(';
    for (let j = 0; j < attendsIndex.length; j++) {
      attendsFormula += 'RC[' + (-2 - statisticOption.length + attendsIndex[j]) + ']';
      if (j + 1 !== attendsIndex.length) attendsFormula += '+';
    }
    attendsFormula += ')/RC[-1]';

    // 純出席率計算式
    let ignoreFormula = '=(';
    for (let j = 0; j < attendsIndex.length; j++) {
      ignoreFormula += 'RC[' + (-3 - statisticOption.length + attendsIndex[j]) + ']';
      if (j + 1 !== attendsIndex.length) ignoreFormula += '+';
    }
    ignoreFormula += ')/(RC[-2]';
    for (let j = 0; j < ignoreIndex.length; j++) {
      ignoreFormula += '-';
      ignoreFormula += 'RC[' + (-3 - statisticOption.length + ignoreIndex[j]) + ']';
    }
    ignoreFormula += ')';

    // 未処理者 計 計算式
    let unattendsFormula = '=';
    for (let j = 0; j < unattendsIndex.length; j++) {
      unattendsFormula += 'RC[' + (-4 - statisticOption.length + unattendsIndex[j]) + ']';
      if (j + 1 !== unattendsIndex.length) unattendsFormula += '+';
    }

    for (let j = 0; j < place.length; j++) { // 実施場所毎のループ（1つの表）
      baseSheet.getRange(totalStartRowNum + tableRowCount, 3, 1, statisticOption.length).setBackground('aqua'); // 色をつける
      baseSheet.getRange(totalStartRowNum + tableRowCount, 2, 1, 1).setValue(place[j]).setHorizontalAlignment('center'); // 実施場所毎合計部
      baseSheet.getRange(totalStartRowNum + tableRowCount, statisticOption.length + 3, 1, 1).setFormulaR1C1('=SUM(RC[' + (-statisticOption.length) + ']:RC[-1])'); // 時間帯の1場所の合計
      baseSheet.getRange(totalStartRowNum + tableRowCount, statisticOption.length + 4, 1, 1).setFormulaR1C1(attendsFormula).setNumberFormat("0%"); // 時間帯の1場所の総出席率
      baseSheet.getRange(totalStartRowNum + tableRowCount, statisticOption.length + 5, 1, 1).setFormulaR1C1(ignoreFormula).setNumberFormat("0%"); // 時間帯の1場所の純出席率
      baseSheet.getRange(totalStartRowNum + tableRowCount, statisticOption.length + 6, 1, 1).setFormulaR1C1(unattendsFormula); // 時間帯の1場所の未処理 計
      tableRowCount++;
    }

    // 合計計算（相対仕様に変更）
    baseSheet.getRange(totalStartRowNum + tableRowCount, 2, 1, 1).setValue('合計').setFontColor('red').setHorizontalAlignment('center');
    placeStatisticFormula += 'SUM(R[' + (-place.length) + ']C:R[-1]C)'
    baseSheet.getRange(totalStartRowNum + tableRowCount, 3, 1, statisticOption.length).setFormulaR1C1(placeStatisticFormula); // 要素ごとの最終合計
    baseSheet.getRange(totalStartRowNum + tableRowCount, statisticOption.length + 3, 1, 1).setFormulaR1C1('=SUM(RC[' + (-statisticOption.length) + ']:RC[-1])'); // 時間帯の合計
    baseSheet.getRange(totalStartRowNum + tableRowCount, statisticOption.length + 4, 1, 1).setFormulaR1C1(attendsFormula).setNumberFormat("0%"); // 時間帯の総出席率
    baseSheet.getRange(totalStartRowNum + tableRowCount, statisticOption.length + 5, 1, 1).setFormulaR1C1(ignoreFormula).setNumberFormat("0%"); // 時間帯の純出席率
    baseSheet.getRange(totalStartRowNum + tableRowCount, statisticOption.length + 6, 1, 1).setFormulaR1C1(unattendsFormula); // 時間帯の未処理 計

    statisticLines.push(tableRowCount + totalStartRowNum);
    tableRowCount++;

    baseSheet.getRange(totalStartRowNum, 2, tableRowCount, statisticOption.length + 2).setBorder(true, true, true, true, true, true); // 枠線を引く

    // 表の複製処理（1つだけ作成したテンプレートをコピーする）
    for (let i = 1; i < section.length; i++) {
      baseSheet.getRange(totalStartRowNum, 2, tableRowCount, statisticOption.length + 5).copyTo(baseSheet.getRange(totalStartRowNum + (tableRowCount + 2) * i, 2, tableRowCount, statisticOption.length + 5));
      baseSheet.getRange(totalStartRowNum + (tableRowCount + 2) * i, 2, 1, 1).setValue(section[i]);
      linkLines.push(totalStartRowNum + (tableRowCount + 2) * i); // ジャンプリンクを貼る行番号を保存
      statisticLines.push(totalStartRowNum + (tableRowCount + 2) * i + tableRowCount - 1); // 集計値のリンクを貼る行番号を保存
    }
    //console.log(tableRowCount);
    //console.log(statisticLines);

    // 上部の集計部を生成
    let baseColumn = 1;
    let rowCount = 0;
    baseSheet.getRange(2, baseColumn + 1, 1, 3).setValues([['総出席率', '純出席率', '未処理　計']]);

    let basegid = baseSheet.getSheetId();

    // 各時間帯の総出席率，純出席率，未処理　計を読み出す
    for (let i = 0; i < section.length; i++) {
      baseSheet.getRange(rowCount + 3, baseColumn, 1, 1).setFormula('=HYPERLINK(\"#gid=' + basegid + '&range=B' + linkLines[i] + '\", \"' + section[i] + '\")'); // リンク作成
      baseSheet.getRange(rowCount + 3, baseColumn + 1, 1, 1).setFormulaR1C1('=R' + statisticLines[i] + 'C' + (statisticOption.length + 4)).setNumberFormat("0%");
      baseSheet.getRange(rowCount + 3, baseColumn + 2, 1, 1).setFormulaR1C1('=R' + statisticLines[i] + 'C' + (statisticOption.length + 5)).setNumberFormat("0%");
      baseSheet.getRange(rowCount + 3, baseColumn + 3, 1, 1).setFormulaR1C1('=R' + statisticLines[i] + 'C' + (statisticOption.length + 6));
      rowCount++;
      if (i + 1 === halfSectionCount) {
        baseColumn += 4;
        rowCount = 0;
        baseSheet.getRange(2, baseColumn + 1, 1, 3).setValues([['総出席率', '純出席率', '未処理　計']]);
      }
    }

    baseColumn += 5;
    baseSheet.getRange(1, baseColumn, 1, 1).setValue('計');
    baseSheet.getRange(2, baseColumn, 1, statisticOption.length).setValues([statisticOption]).setHorizontalAlignment('center');
    baseSheet.getRange(2, baseColumn + statisticOption.length, 1, 4).setValues([['総数', '総出席率', '純出席率', '未処理　計']]).setHorizontalAlignment('center');

    // シートの全体集計式の作成と挿入
    for (let i = 0; i < statisticOption.length; i++) {
      let allStatisticFormula = '=';
      for (let j = 0; j < statisticLines.length; j++) {
        allStatisticFormula += 'R' + statisticLines[j] + 'C' + (i + 3);
        if (j + 1 !== statisticLines.length) allStatisticFormula += '+';
      }
      baseSheet.getRange(3, baseColumn + i, 1, 1).setFormulaR1C1(allStatisticFormula);
    }
    baseSheet.getRange(3, baseColumn + statisticOption.length, 1, 1).setFormulaR1C1('=SUM(RC[' + (-statisticOption.length) + ']:RC[-1])');

    // 出席率の項目を追加
    baseSheet.getRange(3, baseColumn + statisticOption.length + 1, 1, 1).setFormulaR1C1(attendsFormula).setNumberFormat("0%"); // 時間帯の1場所の総出席率
    baseSheet.getRange(3, baseColumn + statisticOption.length + 2, 1, 1).setFormulaR1C1(ignoreFormula).setNumberFormat("0%"); // 時間帯の1場所の純出席率
    baseSheet.getRange(3, baseColumn + statisticOption.length + 3, 1, 1).setFormulaR1C1(unattendsFormula); // 時間帯の1場所の未処理 計

    baseSheet.getRange(2, baseColumn, 2, statisticOption.length + 1).setBorder(true, true, true, true, true, true); // 枠線を引く

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

      let baseColumn = 1;
      let rowCount = 0;
      let finalgid = finalSheet.getSheetId();
      for (let i = 0; i < section.length; i++) {
        finalSheet.getRange(rowCount + 3, baseColumn, 1, 1).setFormula('=HYPERLINK(\"#gid=' + finalgid + '&range=B' + linkLines[i] + '\", \"' + section[i] + '\")'); // リンク作成
        rowCount++;
        if (i + 1 === halfSectionCount) {
          baseColumn += 4;
          rowCount = 0;
        }
      }

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