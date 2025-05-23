function createPivotTable() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getDataRange();

  const pivotSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("PivotSheet");
  const pivotTableRange = pivotSheet.getRange('A1');

  const pivotTable = pivotTableRange.createPivotTable(range);
  pivotTable.addRowGroup(1);  // e.g., Category
  pivotTable.addPivotValue(2, SpreadsheetApp.PivotTableSummarizeFunction.SUM); // e.g., Value
}

function createChart() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getRange("A1:B" + sheet.getLastRow());

  const chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(range)
    .setPosition(5, 5, 0, 0)
    .build();

  sheet.insertChart(chart);
}

function callAI(instruction) {
  const payload = {
    instruction: instruction
  };
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };
  const response = UrlFetchApp.fetch("http://localhost:8000/ai-command", options);
  const result = JSON.parse(response.getContentText());

  // 👉 Auto-detect logic from instruction
  const text = instruction.toLowerCase();
  if (text.includes("pivot")) {
    createPivotTable();
    result.output += "\n📊 Pivot table created.";
  } else if (text.includes("chart")) {
    createChart();
    result.output += "\n📈 Chart inserted.";
  }

  return result;
}

const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
const numRows = sheet.getLastRow();
const numCols = sheet.getLastColumn();
const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];

let valueCol = 2; // default
for (let i = 0; i < headers.length; i++) {
  if (headers[i].toLowerCase().includes("sales") || headers[i].toLowerCase().includes("revenue")) {
    valueCol = i + 1;
    break;
  }
}
const range = sheet.getRange(1, 1, numRows, valueCol);
