const SHEET_ID = "13fBGGhmlv44AdMMTtobF6CUopJxC2ZXPGdUTLwLH-ro";
const SHEET_NAME = "Stamping";

// โหลดหน้าเว็บ
function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Stamping Production System");
}

// บันทึกข้อมูล
function saveData(data) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);

  const total = Number(data.good) + Number(data.ng);

  sheet.appendRow([
    data.machine,
    data.partNo,
    data.partName,
    data.shift,
    data.good,
    data.ng,
    total,
    data.hours,
    data.operator,
    new Date()
  ]);

  createChartByMachinePart();

  return "บันทึกสำเร็จ";
}

// 🔥 สร้างกราฟใน Google Sheet (มีตัวเลขบนแท่ง)
function createChartByMachinePart() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  let summary = {};

  for (let i = 1; i < data.length; i++) {
    const machine = data[i][0];
    const part = data[i][2];
    const good = Number(data[i][4]);
    const ng = Number(data[i][5]);

    const key = machine + " | " + part;

    if (!summary[key]) {
      summary[key] = { good: 0, ng: 0 };
    }

    summary[key].good += good;
    summary[key].ng += ng;
  }

  // ล้างข้อมูลเก่า
  sheet.getRange("M:O").clearContent();

  // header
  sheet.getRange("M1").setValue("Machine-Part");
  sheet.getRange("N1").setValue("Good");
  sheet.getRange("O1").setValue("NG");

  let output = [];
  for (let key in summary) {
    output.push([key, summary[key].good, summary[key].ng]);
  }

  // เรียง NG มาก → น้อย
  output.sort((a, b) => b[2] - a[2]);

  if (output.length > 0) {
    sheet.getRange(2, 13, output.length, 3).setValues(output);
  }

  // ลบกราฟเก่า
  sheet.getCharts().forEach(c => sheet.removeChart(c));

  // 🔥 สร้างกราฟ + แสดงตัวเลข
  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sheet.getRange("M1:O" + (output.length + 1)))
    .setPosition(2, 12, 0, 0)
    .setOption("title", "NG vs Good by Machine & Part")

    // 🔥 แสดงตัวเลขบนกราฟ
    .setOption("series", {
      0: { dataLabel: "value" },
      1: { dataLabel: "value" }
    })

    .build();

  sheet.insertChart(chart);
}

// 🔥 ส่งข้อมูลไปหน้าเว็บ
function getChartData() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  let summary = {};

  for (let i = 1; i < data.length; i++) {
    const machine = data[i][0];
    const part = data[i][2];
    const good = Number(data[i][4]);
    const ng = Number(data[i][5]);

    const key = machine + " | " + part;

    if (!summary[key]) {
      summary[key] = { good: 0, ng: 0 };
    }

    summary[key].good += good;
    summary[key].ng += ng;
  }

  let labels = [];
  let goodData = [];
  let ngData = [];

  for (let key in summary) {
    labels.push(key);
    goodData.push(summary[key].good);
    ngData.push(summary[key].ng);
  }

  return {
    labels: labels,
    good: goodData,
    ng: ngData
  };
}
