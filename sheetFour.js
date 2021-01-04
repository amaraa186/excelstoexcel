function sheetFour(workbook, inputDataOne, inputDataTwo) {
  var sheet = workbook.worksheets().add("Extra Sheet");
  var tNTwo = 4;

  var mergedCellA1J2 = sheet.mergedCellsRegions().add(0, 0, 1, 9);
  var mergedCellA4B4 = sheet.mergedCellsRegions().add(3, 0, 3, 1);
  var mergedCellC4D4 = sheet.mergedCellsRegions().add(3, 2, 3, 3);
  var mergedCellE4F4 = sheet.mergedCellsRegions().add(3, 4, 3, 5);
  var mergedCellH4J4 = sheet.mergedCellsRegions().add(3, 7, 3, 9);

  var mergedCellE3F3 = sheet.mergedCellsRegions().add(2, 4, 2, 5);
  var mergedCellH3I3 = sheet.mergedCellsRegions().add(2, 7, 2, 8);

  mergedCellA1J2.value("諸経費 • 交通費集計表");
  mergedCellA1J2
    .cellFormat()
    .alignment($.ig.excel.HorizontalCellAlignment.center);
  mergedCellA1J2
    .cellFormat()
    .font()
    .height(16 * 20);
  mergedCellE3F3.value(kanjidate.format("2016-06-12"));
  mergedCellH3I3.value(kanjidate.format("2016-06-12"));

  mergedCellA4B4.value("氏名");
  mergedCellC4D4.value("諸経費");
  mergedCellE4F4.value("交通費");
  sheet.getCell("G4").value("計");
  sheet
    .getCell("G4")
    .cellFormat()
    .alignment($.ig.excel.HorizontalCellAlignment.center);
  mergedCellH4J4.value("振込口座");

  mergedCellA4B4
    .cellFormat()
    .alignment($.ig.excel.HorizontalCellAlignment.center);
  mergedCellC4D4
    .cellFormat()
    .alignment($.ig.excel.HorizontalCellAlignment.center);
  mergedCellE4F4
    .cellFormat()
    .alignment($.ig.excel.HorizontalCellAlignment.center);
  mergedCellH4J4
    .cellFormat()
    .alignment($.ig.excel.HorizontalCellAlignment.center);

  sheet.rows(3).cellFormat().font().bold(true);

  var count = inputDataOne.length;

  for (let i = 0; i < inputDataOne.length; i++) {
    var mergedOne = sheet.mergedCellsRegions().add(4 + i, 0, 4 + i, 1);
    mergedOne.value(inputDataOne[i].name);
    var sum = 0;
    for (let j = 0; j < inputDataTwo.length; j++) {
      if (inputDataOne[i].name === inputDataTwo[j].請求者氏名) {
        sum += inputDataTwo[j].金額;
      }
    }
    var mergedTwo = sheet.mergedCellsRegions().add(4 + i, 2, 4 + i, 3);
    mergedTwo.value(sum);
    var mergedThree = sheet.mergedCellsRegions().add(4 + i, 4, 4 + i, 5);
    mergedThree.value(inputDataOne[i].total);
    sheet
      .getCell("G" + (5 + i))
      .applyFormula("=SUM(C" + (5 + i) + ":E" + (5 + i) + ")");
    var mergedFour = sheet.mergedCellsRegions().add(4 + i, 7, 4 + i, 9);
    mergedFour.value(inputDataOne[i].code);
  }

  var mergedOne = sheet
    .mergedCellsRegions()
    .add(tNTwo + count, 0, tNTwo + count, 1);
  mergedOne.value("計");

  var mergedTwo = sheet
    .mergedCellsRegions()
    .add(tNTwo + count, 2, tNTwo + count, 3);
  mergedTwo.applyFormula("=SUM(C5:C" + (tNTwo + count) + ")");

  var mergedThree = sheet
    .mergedCellsRegions()
    .add(tNTwo + count, 4, tNTwo + count, 5);
  mergedThree.applyFormula("=SUM(E5:E" + (tNTwo + count) + ")");

  var mergedFour = sheet
    .mergedCellsRegions()
    .add(tNTwo + count, 7, tNTwo + count, 9);
  sheet
    .getCell("G" + (5 + count))
    .applyFormula("=SUM(C" + (5 + count) + ":E" + (5 + count) + ")");

  sheet.getCell("D" + (7 + count)).value("決裁");
  sheet.getCell("E" + (7 + count)).value("起票");
  sheet.getCell("F" + (7 + count)).value("支払");
  sheet.getCell("G" + (7 + count)).value("EB");
  sheet.getCell("H" + (7 + count)).value("確認");
  sheet.getCell("I" + (7 + count)).value("確認");
  sheet.getCell("J" + (7 + count)).value("担当");

  sheet.mergedCellsRegions().add(7 + count, 3, 8 + count, 3);
  sheet.mergedCellsRegions().add(7 + count, 4, 8 + count, 4);
  sheet.mergedCellsRegions().add(7 + count, 5, 8 + count, 5);
  sheet.mergedCellsRegions().add(7 + count, 6, 8 + count, 6);
  sheet.mergedCellsRegions().add(7 + count, 7, 8 + count, 7);
  sheet.mergedCellsRegions().add(7 + count, 8, 8 + count, 8);
  sheet.mergedCellsRegions().add(7 + count, 9, 8 + count, 9);

  //table 1
  for (var index = 3; index < 5 + count; index++) {
    for (var i = 0; i < 10; i++) {
      sheet
        .rows(index)
        .cells(i)
        .cellFormat()
        .bottomBorderStyle($.ig.excel.CellBorderLineStyle.thin);
      sheet
        .rows(index)
        .cells(i)
        .cellFormat()
        .leftBorderStyle($.ig.excel.CellBorderLineStyle.thin);
      sheet
        .rows(index)
        .cells(i)
        .cellFormat()
        .rightBorderStyle($.ig.excel.CellBorderLineStyle.thin);
      sheet
        .rows(index)
        .cells(i)
        .cellFormat()
        .topBorderStyle($.ig.excel.CellBorderLineStyle.thin);
    }
  }

  //table 2
  for (var j = 6 + count; j < 9 + count; j++) {
    for (var l = 3; l < 10; l++) {
      sheet
        .rows(j)
        .cells(l)
        .cellFormat()
        .bottomBorderStyle($.ig.excel.CellBorderLineStyle.thin);
      sheet
        .rows(j)
        .cells(l)
        .cellFormat()
        .leftBorderStyle($.ig.excel.CellBorderLineStyle.thin);
      sheet
        .rows(j)
        .cells(l)
        .cellFormat()
        .rightBorderStyle($.ig.excel.CellBorderLineStyle.thin);
      sheet
        .rows(j)
        .cells(l)
        .cellFormat()
        .topBorderStyle($.ig.excel.CellBorderLineStyle.thin);
    }
  }
}
