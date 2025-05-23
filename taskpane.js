Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("cmd-btn").onclick = runCommand;
  }
});

async function runCommand() {
  const cmd = document.getElementById("cmd").value.trim();
  const resultDiv = document.getElementById("result");

  if (!cmd) {
    resultDiv.innerText = "❗ Please enter a command.";
    return;
  }

  resultDiv.innerText = "⏳ Processing...";

  const response = await fetch("http://localhost:8000/ai-command", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ instruction: cmd })
  });

  const result = await response.json();
  resultDiv.innerText = "✅ AI Result: " + result.output;

  const lower = cmd.toLowerCase();
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    if (lower.includes("chart")) {
      const dataRange = sheet.getRange("A1:B10");
      const chart = sheet.charts.add("ColumnClustered", dataRange, "Auto");
      chart.setPosition("D2", "G20");
      chart.title.text = "Auto Chart";
    }

    await context.sync();
  });
}

if (lower.includes("pivot")) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const dataRange = sheet.getRange("A1").getSurroundingRegion();

    const pivotSheet = context.workbook.worksheets.add("PivotTableSheet");
    const pivotTable = context.workbook.pivotTables.add("SalesPivot", dataRange, pivotSheet.getRange("A1"));

    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItemAt(0)); // e.g., Category
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItemAt(1)); // e.g., Sales

    pivotTable.dataHierarchies.getItemAt(0).summarizeBy = "sum";
    pivotTable.layoutType = Excel.PivotLayoutType.compact;

    pivotSheet.activate();
    await context.sync();
  });
}
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  const usedRange = sheet.getUsedRange();
  usedRange.load("address");
  await context.sync();

  const chart = sheet.charts.add("ColumnClustered", usedRange, "Auto");
  chart.setPosition("E5", "H25");
  chart.title.text = "📊 Auto Chart";
  await context.sync();
});

const usedRange = sheet.getUsedRange();
const chart = sheet.charts.add("ColumnClustered", usedRange, "Auto");
chart.setPosition("E5", "H25");
chart.title.text = "📊 Auto Chart";
