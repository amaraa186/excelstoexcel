function sheetTwo(workbook, inputDataOne, inputDataTwo) {
  var table;
  var long = inputDataOne.length;
  //sheet2
  var sheet = workbook.worksheets().add("Sheet2");

  for (var i = 0; i < 3; i++) {
    sheet.columns(i).setWidth(90, $.ig.excel.WorksheetColumnWidthUnit.pixel);
  }
  for (var i = 3; i < 20; i++) {
    sheet.columns(i).setWidth(60, $.ig.excel.WorksheetColumnWidthUnit.pixel);
  }
  // Create a to-do list table with columns for tasks and their priorities.
  sheet.getCell("A1").value("列1");
  sheet.getCell("B1").value("物件コード");
  sheet.getCell("C1").value("姓名");
  sheet.getCell("D1").value("仕入");
  sheet.getCell("E1").value("支払手数料");
  sheet.getCell("F1").value("外注費");
  sheet.getCell("G1").value("消耗品費");
  sheet.getCell("H1").value("通信費");
  sheet.getCell("I1").value("アフター費");
  sheet.getCell("J1").value("接待交際費");
  sheet.getCell("K1").value("会議費");
  sheet.getCell("L1").value("諸経費");
  sheet.getCell("M1").value("交通費");
  sheet.getCell("N1").value("合計");

  // Populate the table with data
  sheet.getCell("E2").value(5215);
  sheet.getCell("F2").value(3016);
  sheet.getCell("G2").value(5126);
  sheet.getCell("H2").value(5130);
  sheet.getCell("I2").value(5170);
  sheet.getCell("J2").value(5160);
  sheet.getCell("K2").value(5504);
  sheet.getCell("M2").value(5140);

  sheet.getCell("D3").value("課税");
  sheet.getCell("E3").value("非課税");
  sheet.getCell("F3").value("課税");
  sheet.getCell("G3").value("課税");
  sheet.getCell("H3").value("課税");
  sheet.getCell("I3").value("課税");
  sheet.getCell("J3").value("課税");
  sheet.getCell("K3").value("課税");
  sheet.getCell("M3").value("課税");
  var p = 0;
  for (var i = 0; i < long; i++) {
    var count = 0;
    sheet.getCell("A" + (4 + p)).value(inputDataOne[i].name);
    sheet.getCell("B" + (4 + p)).value(20010000);
    sheet.getCell("C" + (4 + p)).value("共通経費");
    sheet
      .getCell("L" + (4 + p))
      .applyFormula("=SUM(D" + (4 + p) + ":K" + (4 + p) + ")");
    sheet.getCell("M" + (4 + p)).value(inputDataOne[i].total);
    sheet
      .getCell("N" + (4 + p))
      .applyFormula("=SUM(L" + (4 + p) + ":M" + (4 + p) + ")");
    p++;
    count++;
    function input(p, j, column) {
      sheet.getCell("A" + (4 + p)).value(inputDataTwo[j].請求者氏名);
      sheet.getCell("B" + (4 + p)).value(inputDataTwo[j].物件コード);
      sheet.getCell("C" + (4 + p)).value(inputDataTwo[j].姓名);
      sheet.getCell(column + (4 + p)).value(inputDataTwo[j].金額);
      sheet
        .getCell("L" + (4 + p))
        .applyFormula("=SUM(D" + (4 + p) + ":K" + (4 + p) + ")");
      sheet
        .getCell("N" + (4 + p))
        .applyFormula("=SUM(L" + (4 + p) + ":M" + (4 + p) + ")");
    }
    for (var j = 0; j < inputDataTwo.length; j++) {
      if (inputDataOne[i].name === inputDataTwo[j].請求者氏名) {
        switch (inputDataTwo[j].費) {
          case "支払手数料":
            input(p, j, "E");
            p++;
            count++;
            break;
          case "外注費":
            input(p, j, "F");
            p++;
            count++;
            break;
          case "消耗品費":
            input(p, j, "G");
            p++;
            count++;
            break;
          case "会議費":
            input(p, j, "K");
            p++;
            count++;
        }
      }
    }
    sheet.getCell("A" + (4 + p)).value(inputDataOne[i].name);
    sheet.getCell("B" + (4 + p)).value("データの個数");

    for (var u = 0; u < alphabet2.length; u++) {
      sheet
        .getCell(alphabet2[u] + (4 + p))
        .applyFormula(
          "=SUM(" +
            alphabet2[u] +
            (4 + p - count) +
            ":" +
            alphabet2[u] +
            (4 + p - 1) +
            ")"
        );
    }
    p++;
  }
  var plus = 4;
  var plus2 = 3;
  sheet.getCell("A" + (plus + p)).value("総合計");
  for (var u = 0; u < alphabet2.length - 3; u++) {
    sheet
      .getCell(alphabet2[u] + (plus + p))
      .applyFormula(
        "=SUM(" + alphabet2[u] + "4:" + alphabet2[u] + (plus2 + p) + ")" + "/2"
      );
  }

  sheet
    .getCell("L" + (plus + p))
    .applyFormula("=SUM(D" + (plus + p) + ":K" + (plus + p) + ")");
  sheet
    .getCell("M" + (plus + p))
    .applyFormula("=SUM(M4:" + "M" + (plus2 + p) + ")" + "/2");
  sheet
    .getCell("N" + (plus + p))
    .applyFormula("=SUM(L" + (plus + p) + ":M" + (plus + p) + ")");

  sheet.getCell("A" + (6 + p)).value("資産項目");
  sheet.getCell("B" + (6 + p)).value("立替金");
  sheet.getCell("C" + (6 + p)).value("安全協力会");
  sheet.getCell("D" + (6 + p)).value(" ");

  sheet.getCell("A" + (7 + p)).value("負債項目");
  sheet.getCell("B" + (7 + p)).value("工事未払金");

  sheet.getCell("A" + (8 + p)).value("販管費項目");
  sheet.getCell("B" + (8 + p)).value("研修費");
  sheet.getCell("C" + (8 + p)).value("課税対象外");

  sheet.getCell("B" + (9 + p)).value("研修費");
  sheet.getCell("C" + (9 + p)).value("課税");

  sheet.getCell("B" + (10 + p)).value("アフター");

  sheet.getCell("B" + (11 + p)).value("旅費交通費");
  sheet.getCell("C" + (11 + p)).value("出張費-宿泊");

  sheet.getCell("C" + (12 + p)).value("出張費-交通費");
  sheet.getCell("D" + (12 + p)).value("");

  sheet.getCell("C" + (13 + p)).value("出張費-交通費非課税");
  sheet.getCell("D" + (13 + p)).value("");

  sheet.getCell("B" + (14 + p)).value("旅費交通費");
  sheet.getCell("C" + (14 + p)).value("出張費-食事");

  sheet.getCell("C" + (15 + p)).value("出張費-食事課税");
  sheet.getCell("D" + (15 + p)).value("");

  sheet.getCell("B" + (16 + p)).value("消耗品費");
  sheet.getCell("C" + (16 + p)).value("出張時消耗品費");
  sheet.getCell("D" + (16 + p)).value("");

  sheet.getCell("B" + (17 + p)).value("消耗品費");
  sheet.getCell("C" + (17 + p)).value("出張時消耗品費　課税対象外");
  sheet.getCell("D" + (17 + p)).value("");

  sheet.getCell("B" + (18 + p)).value("会議費");
  sheet.getCell("B" + (19 + p)).value("消耗品費");
  sheet.getCell("B" + (20 + p)).value("福利厚生費");
  sheet.getCell("B" + (21 + p)).value("接待交際費");
  sheet.getCell("B" + (22 + p)).value("事務用品費");

  sheet.getCell("C" + (23 + p)).value("資産負債項目・販管費支払合計");
  sheet.getCell("D" + (23 + p)).value(0);

  sheet.getCell("C" + (24 + p)).value("支払合計");
  sheet
    .getCell("D" + (24 + p))
    .applyFormula("=SUM(L" + (plus + p) + ":M" + (plus + p) + ")");

  //table styles
  table = sheet.tables().add("A1:N" + (4 + p), true);
  table.style(workbook.standardTableStyles("TableStyleMedium2"));
  table = sheet.tables().add("B" + (6 + p) + ":" + "D" + (24 + p), true);
  table.style(workbook.standardTableStyles("TableStyleLight15"));
}
