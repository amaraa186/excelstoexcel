$(document).ready(function () {
  $("#input1").change(inputData1);
  $("button").click(buttonClick);
  $("#input2").change(inputData2);
});
//vars
var rABS = true;
let inputDataOne = [];
let inputDataTwo = [];
let diffCatergory;
let diff = [];
let alphabet = "HIJKLMNOPQRSTUVWXYZ".split("");
let alphabet2 = "DEFGHIJKLMN".split("");
Array.prototype.sum = function (prop) {
  var total = 0;
  for (var i = 0, _len = this.length; i < _len; i++) {
    total += this[i][prop];
  }
  return total;
};

function call_cell(col, row, sheet) {
  /* Find desired cell */
  this.desired_cell = sheet[col + row];
  /* Get the value */
  this.value = this.desired_cell ? this.desired_cell.v : null;
  return this.value;
}

function inputData1() {
  var files = $("#input1").prop("files");
  for (var j = 0; j < files.length; j++) {
    f = files[j];
    var reader = new FileReader();
    reader.onload = function (e) {
      var data = e.target.result;
      if (!rABS) data = new Uint8Array(data);
      var workbook = XLSX.read(data, {
        type: rABS ? "binary" : "array",
      });

      // ==============
      var full = [];
      for (var r = 0; r < workbook.SheetNames.length; r++) {
        var first_sheet_name = workbook.SheetNames[r];
        var worksheet = workbook.Sheets[first_sheet_name];
        for (var i = 0; i < 30; i = i + 2) {
          var obj = {
            code: call_cell("A", 3, worksheet),
            employee: call_cell("O", 2, worksheet),
            請求者氏名: call_cell("O", 3, worksheet),
            物件コード: call_cell("I", 14 + i, worksheet),
            姓名: call_cell("I", 15 + i, worksheet),
            駐車料金: call_cell("L", 14 + i, worksheet),
            交通機関利用料: call_cell("O", 14 + i, worksheet),
            その他: call_cell("R", 14 + i, worksheet),
          };
          full.push(obj);
        }
      }
      //get total function
      var insert = {
        name: obj.請求者氏名,
        employee: obj.employee,
        code: obj.code,
        total:
          full.sum("駐車料金") +
          full.sum("交通機関利用料") +
          full.sum("その他"),
      };
      inputDataOne.push(insert);
      // ===========================================================
    };
    if (rABS) reader.readAsBinaryString(f);
    else reader.readAsArrayBuffer(f);
  }
}

function inputData2() {
  var files = $("#input2").prop("files");
  for (var j = 0; j < files.length; j++) {
    f = files[j];
    var reader = new FileReader();
    reader.onload = function (e) {
      var data = e.target.result;
      if (!rABS) data = new Uint8Array(data);
      var workbook = XLSX.read(data, {
        type: rABS ? "binary" : "array",
      });

      // ==============
      var full = [];
      for (var r = 0; r < workbook.SheetNames.length; r++) {
        var first_sheet_name = workbook.SheetNames[r];
        var worksheet = workbook.Sheets[first_sheet_name];
        for (var i = 0; i < 30; i = i + 2) {
          var obj = {
            請求者氏名: call_cell("N", 3, worksheet),
            費: call_cell("J", 9 + i, worksheet),
            物件コード: call_cell("K", 9 + i, worksheet),
            姓名: call_cell("K", 10 + i, worksheet),
            金額: call_cell("P", 9 + i, worksheet),
          };
          if (null !== obj.物件コード) {
            full.push(obj);
          }
        }
      }
      //Sum similar keys in an array of objects
      var obj2 = {};
      full.forEach(function (a) {
        if (!(obj2[a.物件コード] && obj2[a.費])) {
          obj2[a.物件コード] = {
            請求者氏名: a.請求者氏名,
            物件コード: a.物件コード,
            姓名: a.姓名,
            費: a.費,
            金額: 0,
          };
          inputDataTwo.push(obj2[a.物件コード]);
        }
        obj2[a.物件コード].金額 += a.金額;
      });
      for (var q = 0; q < inputDataTwo.length; q++) {
        switch (inputDataTwo[q].費) {
          case "支払手数料":
            break;
          case "外注費":
            break;
          case "消耗品費":
            break;
          case "会議費":
            break;
          default:
            diff.push(inputDataTwo[q].費);
        }
      }
      diffCatergory = diff.filter((c, index) => {
        return diff.indexOf(c) === index;
      });
    };
    if (rABS) reader.readAsBinaryString(f);
    else reader.readAsArrayBuffer(f);
  }
}

function buttonClick() {
  var table;
  var longOne = inputDataOne.length;
  var longTwo = inputDataTwo.length;
  if ($("#input1").prop("files").length === $("#input2").prop("files").length) {
    //sheet1
    var workbook = new $.ig.excel.Workbook($.ig.excel.WorkbookFormat.excel2007);
    var sheet = workbook.worksheets().add("Sheet1");

    for (var i = 0; i < 20; i++) {
      sheet.columns(i).setWidth(70, $.ig.excel.WorksheetColumnWidthUnit.pixel);
    }
    // Create a to-do list table with columns for tasks and their priorities.
    sheet.getCell("A1").value("列1");
    sheet.getCell("B1").value("物件コード");
    sheet.getCell("C1").value("姓名");
    sheet.getCell("D1").value("支払手数料");
    sheet.getCell("E1").value("外注費");
    sheet.getCell("F1").value("消耗品費");
    sheet.getCell("G1").value("会議費");
    if (diffCatergory.length === 0) {
      sheet.getCell("H1").value("諸経費");
      sheet.getCell("I1").value("交通費");
      sheet.getCell("J1").value("合計");
      table = sheet.tables().add("A1:J" + (4 + longOne + longTwo), true);
    } else {
      for (var j = 0; j < diffCatergory.length; j++) {
        sheet.getCell(alphabet[j] + "1").value(diffCatergory[j]);
      }
      sheet.getCell(alphabet[diffCatergory.length] + "1").value("諸経費");
      sheet.getCell(alphabet[diffCatergory.length + 1] + "1").value("交通費");
      sheet.getCell(alphabet[diffCatergory.length + 2] + "1").value("合計");
      table = sheet
        .tables()
        .add(
          "A1:" + alphabet[diffCatergory.length + 2] + (4 + longOne + longTwo),
          true
        );
    }

    // Specify the style to use in the table (this can also be specified as an optional 3rd argument to the 'add' call above).
    table.style(workbook.standardTableStyles("TableStyleMedium2"));

    // Populate the table with data
    sheet.getCell("D2").value(5215);
    sheet.getCell("E2").value(3016);
    sheet.getCell("F2").value(5126);
    sheet.getCell("G2").value(5504);
    sheet.getCell(alphabet[diffCatergory.length + 1] + "2").value(5140);

    sheet.getCell("D3").value("非課税");
    sheet.getCell("E3").value("課税");
    sheet.getCell("F3").value("課税");
    sheet.getCell("G3").value("課税");
    sheet.getCell(alphabet[diffCatergory.length + 1] + "3").value("課税");

    $.each(inputDataOne, function (index, val) {
      if (val.total !== 0) {
        sheet.getCell("A" + (4 + index)).value(val.name);
        sheet.getCell("B" + (4 + index)).value(20010000);
        sheet.getCell("C" + (4 + index)).value("共通経費");

        if (diffCatergory.length !== 0) {
          sheet
            .getCell(alphabet[diffCatergory.length] + (4 + index))
            .applyFormula(
              "=SUM(D" +
                (4 + index) +
                ":" +
                alphabet[diffCatergory.length - 1] +
                (4 + index) +
                ")"
            );
          sheet
            .getCell(alphabet[diffCatergory.length + 1] + (4 + index))
            .value(val.total);
          sheet
            .getCell(alphabet[diffCatergory.length + 2] + (4 + index))
            .applyFormula(
              "=SUM(" +
                alphabet[diffCatergory.length] +
                (4 + index) +
                ":" +
                alphabet[diffCatergory.length + 1] +
                (4 + index) +
                ")"
            );
        } else {
          sheet
            .getCell("H" + (4 + index))
            .applyFormula("=SUM(D" + (4 + index) + ":G" + (4 + index) + ")");
          sheet.getCell("I" + (4 + index)).value(val.total);
          sheet
            .getCell("J" + (4 + index))
            .applyFormula("=SUM(H" + (4 + index) + ":I" + (4 + index) + ")");
        }
      }
    });

    $.each(inputDataTwo, function (index, val) {
      sheet.getCell("A" + (4 + longOne + index)).value(val.請求者氏名);
      sheet.getCell("B" + (4 + longOne + index)).value(val.物件コード);
      sheet.getCell("C" + (4 + longOne + index)).value(val.姓名);
      switch (val.費) {
        case "支払手数料":
          sheet.getCell("D" + (4 + longOne + index)).value(val.金額);
          break;
        case "外注費":
          sheet.getCell("E" + (4 + longOne + index)).value(val.金額);
          break;
        case "消耗品費":
          sheet.getCell("F" + (4 + longOne + index)).value(val.金額);
          break;
        case "会議費":
          sheet.getCell("G" + (4 + longOne + index)).value(val.金額);
          break;
        default:
          for (let indec = 0; indec < diffCatergory.length; indec++) {
            if (diffCatergory[indec] === val.費) {
              sheet
                .getCell(alphabet[indec] + (4 + longOne + index))
                .value(val.金額);
              break;
            }
          }
      }

      if (diffCatergory.length !== 0) {
        sheet
          .getCell(alphabet[diffCatergory.length] + (4 + longOne + index))
          .applyFormula(
            "=SUM(D" +
              (4 + longOne + index) +
              ":" +
              alphabet[diffCatergory.length - 1] +
              (4 + longOne + index) +
              ")"
          );
        sheet
          .getCell(alphabet[diffCatergory.length + 2] + (4 + longOne + index))
          .applyFormula(
            "=SUM(" +
              alphabet[diffCatergory.length] +
              (4 + longOne + index) +
              ":" +
              alphabet[diffCatergory.length + 1] +
              (4 + longOne + index) +
              ")"
          );
      } else {
        sheet
          .getCell("H" + (4 + longOne + index))
          .applyFormula(
            "=SUM(D" +
              (4 + longOne + index) +
              ":G" +
              (4 + longOne + index) +
              ")"
          );
        sheet
          .getCell("J" + (4 + longOne + index))
          .applyFormula(
            "=SUM(" +
              alphabet[diffCatergory.length] +
              (4 + longOne + index) +
              ":" +
              alphabet[diffCatergory.length + 1] +
              (4 + longOne + index) +
              ")"
          );
      }
    });

    sheet.getCell("A" + (4 + longOne + longTwo)).value("総合計");
    sheet
      .getCell("D" + (4 + longOne + longTwo))
      .applyFormula("=SUM(D4:D" + (3 + longOne + longTwo) + ")");
    sheet
      .getCell("E" + (4 + longOne + longTwo))
      .applyFormula("=SUM(E4:E" + (3 + longOne + longTwo) + ")");
    sheet
      .getCell("F" + (4 + longOne + longTwo))
      .applyFormula("=SUM(F4:F" + (3 + longOne + longTwo) + ")");
    sheet
      .getCell("G" + (4 + longOne + longTwo))
      .applyFormula("=SUM(G4:G" + (3 + longOne + longTwo) + ")");

    if (diffCatergory.length !== 0) {
      for (var indes = 0; indes < diffCatergory.length; indes++) {
        sheet
          .getCell(alphabet[indes] + (4 + longOne + longTwo))
          .applyFormula(
            "=SUM(" +
              alphabet[indes] +
              4 +
              ":" +
              alphabet[indes] +
              (3 + longOne + longTwo) +
              ")"
          );
      }
      sheet
        .getCell(alphabet[diffCatergory.length] + (4 + longOne + longTwo))
        .applyFormula(
          "=SUM(" +
            alphabet[diffCatergory.length] +
            4 +
            ":" +
            alphabet[diffCatergory.length] +
            (3 + longOne + longTwo) +
            ")"
        );
      sheet
        .getCell(alphabet[diffCatergory.length + 1] + (4 + longOne + longTwo))
        .applyFormula(
          "=SUM(" +
            alphabet[diffCatergory.length + 1] +
            4 +
            ":" +
            alphabet[diffCatergory.length + 1] +
            (3 + longOne + longTwo) +
            ")"
        );
      sheet
        .getCell(alphabet[diffCatergory.length + 2] + (4 + longOne + longTwo))
        .applyFormula(
          "=SUM(" +
            alphabet[diffCatergory.length] +
            (4 + longOne + longTwo) +
            ":" +
            alphabet[diffCatergory.length + 1] +
            (4 + longOne + longTwo) +
            ")"
        );
    } else {
      sheet
        .getCell("H" + (4 + longOne + longTwo))
        .applyFormula(
          "=SUM(D" +
            (4 + longOne + longTwo) +
            ":G" +
            (4 + longOne + longTwo) +
            ")"
        );
      sheet
        .getCell("I" + (4 + longOne + longTwo))
        .applyFormula("=SUM(I4:" + "I" + (3 + longOne + longTwo) + ")");
      sheet
        .getCell("J" + (4 + longOne + longTwo))
        .applyFormula(
          "=SUM(H" +
            (4 + longOne + longTwo) +
            ":I" +
            (4 + longOne + longTwo) +
            ")"
        );
    }

    sheet.getCell("A" + (6 + longOne + longTwo)).value("業者ｺｰﾄﾞ：34");
    sheet.getCell("C" + (6 + longOne + longTwo)).value("どっと：");
    sheet.getCell("D" + (6 + longOne + longTwo)).value("□発注依頼");
    sheet.getCell("E" + (6 + longOne + longTwo)).value("□発注依頼");
    sheet.getCell("F" + (6 + longOne + longTwo)).value("□発注依頼");
    sheet.getCell("G" + (6 + longOne + longTwo)).value("□発注依頼");
    sheet.getCell("I" + (6 + longOne + longTwo)).value("□発注依頼");

    sheet.getCell("D" + (7 + longOne + longTwo)).value("□発注");
    sheet.getCell("E" + (7 + longOne + longTwo)).value("□発注");
    sheet.getCell("F" + (7 + longOne + longTwo)).value("□発注");
    sheet.getCell("G" + (7 + longOne + longTwo)).value("□発注");
    sheet.getCell("I" + (7 + longOne + longTwo)).value("□発注");

    sheet.getCell("D" + (8 + longOne + longTwo)).value("□支払伝票");
    sheet.getCell("E" + (8 + longOne + longTwo)).value("□支払伝票");
    sheet.getCell("F" + (8 + longOne + longTwo)).value("□支払伝票");
    sheet.getCell("G" + (8 + longOne + longTwo)).value("□支払伝票");
    sheet.getCell("I" + (8 + longOne + longTwo)).value("□支払伝票");

    sheet.getCell("J" + (9 + longOne + longTwo)).value("□仕入");
    sheet.getCell("J" + (10 + longOne + longTwo)).value("□仕入転送");
    sheet.getCell("A" + (11 + longOne + longTwo)).value("支払伝票");
    sheet.getCell("B" + (11 + longOne + longTwo)).value("22/7013");
    sheet.getCell("J" + (11 + longOne + longTwo)).value("□支払転送");

    sheet
      .getCell("B" + (12 + longOne + longTwo))
      .value("交通費は合計で1行で入力していい");
    sheet.getCell("A" + (14 + longOne + longTwo)).value("支払日：");
    sheet.getCell("B" + (14 + longOne + longTwo)).value("9/10/2020");
    sheet.getCell("C" + (14 + longOne + longTwo)).value("（8/31〆）");

    //sheet2
    sheetTwo(workbook, inputDataOne, inputDataTwo);

    //sheet3
    sheetThree(workbook, inputDataOne, inputDataTwo);

    // Save the workbook
    saveWorkbook(workbook, "Result.xlsx");
    start();
  } else {
    alert("アップロードされたファイルの数が異なります。");
    start();
  }
}

function saveWorkbook(workbook, name) {
  workbook.save(
    { type: "blob" },
    function (data) {
      saveAs(data, name);
    },
    function (error) {
      alert("Error exporting: : " + error);
    }
  );
}

function start() {
  rABS = true;
  inputDataOne = [];
  inputDataTwo = [];
  diffCatergory = [];
}
