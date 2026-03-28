console.log("🔥 VERSION 29 - CENTER ALIGNMENT ADDED 🔥");

import express from "express";
import ExcelJS from "exceljs";
import path from "path";
import { fileURLToPath } from "url";

const app = express();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const TEMPLATE_FILES = {
styled: "template.xlsx",
editableAp: "template-editable-ap.xlsx",
};

const DEFAULT_START_ROW = Number(process.env.FMEA_START_ROW || 16);

app.use((req, res, next) => {
res.header("Access-Control-Allow-Origin", "*");
res.header("Access-Control-Allow-Headers", "Content-Type");
res.header("Access-Control-Allow-Methods", "POST, OPTIONS, GET");
if (req.method === "OPTIONS") return res.sendStatus(200);
next();
});

app.use(express.json({ limit: "10mb" }));

function quoteSheetName(sheetName) {
return `'${String(sheetName).replace(/'/g, "''")}'`;
}

// ✅ Safe AP formula
function buildApFormula({ sheetName, rowNumber, sCol, oCol, dCol }) {
const fmeaSheet = quoteSheetName(sheetName);

return (
"IFERROR(INDEX('AP Table'!$E$3:$E$1002, MATCH(1, " +
"('AP Table'!$B$3:$B$1002=" + fmeaSheet + "!" + sCol + rowNumber + ")*" +
"('AP Table'!$C$3:$C$1002=" + fmeaSheet + "!" + oCol + rowNumber + ")*" +
"('AP Table'!$D$3:$D$1002=" + fmeaSheet + "!" + dCol + rowNumber + "), 0)), "")"
);
}

function copyTemplateRowStyle(templateRow, targetRow) {
templateRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
targetRow.getCell(colNumber).style = JSON.parse(JSON.stringify(cell.style));
if (cell.numFmt) targetRow.getCell(colNumber).numFmt = cell.numFmt;
if (cell.alignment) targetRow.getCell(colNumber).alignment = JSON.parse(JSON.stringify(cell.alignment));
if (cell.font) targetRow.getCell(colNumber).font = JSON.parse(JSON.stringify(cell.font));
if (cell.fill) targetRow.getCell(colNumber).fill = JSON.parse(JSON.stringify(cell.fill));
if (cell.border) targetRow.getCell(colNumber).border = JSON.parse(JSON.stringify(cell.border));
});

if (templateRow.height) {
targetRow.height = templateRow.height;
}
}

function setFormulaCell(cell, formula, result) {
cell.value = {
formula,
result: result ?? "",
};
}

// ✅ Excel date handling
function toExcelDate(value) {
if (!value) return null;
const d = new Date(value);
return isNaN(d) ? null : d;
}

// ✅ NEW: Center all cells
function centerRowCells(row) {
row.eachCell({ includeEmpty: true }, (cell) => {
cell.alignment = {
horizontal: "center",
vertical: "middle",
wrapText: true,
};
});
}

app.get("/health", (req, res) => {
res.status(200).json({ ok: true });
});

app.post("/export", async (req, res) => {
try {
const {
rows = [],
exportFormat = "styled",
editableAp = false,
} = req.body || {};

```
const useEditableAp =
  editableAp === true ||
  exportFormat === "editable-ap" ||
  exportFormat === "excel-formatted-editable-ap" ||
  exportFormat === "advanced";

const templateFile = useEditableAp
  ? TEMPLATE_FILES.editableAp
  : TEMPLATE_FILES.styled;

const templatePath = path.join(__dirname, templateFile);

const workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile(templatePath);

const sheet = workbook.worksheets[0];
const sheetName = sheet.name || "FMEA";

const START_ROW = DEFAULT_START_ROW;
const templateRow = sheet.getRow(START_ROW);

rows.forEach((r, i) => {
  const rowIndex = START_ROW + i;
  const row = sheet.getRow(rowIndex);

  copyTemplateRowStyle(templateRow, row);

  // ===== DATA =====
  row.getCell(3).value = r.process_step || "";
  row.getCell(4).value = r.function || "";
  row.getCell(5).value = r.failure_mode || "";
  row.getCell(6).value = r.effect || "";

  row.getCell(7).value = r.severity != null ? Number(r.severity) : null;
  row.getCell(8).value = r.cause || "";
  row.getCell(9).value = r.occurrence != null ? Number(r.occurrence) : null;

  row.getCell(10).value =
    (r.current_prevention_controls || "") +
    "\n" +
    (r.current_detection_controls || "");

  row.getCell(11).value = r.detection != null ? Number(r.detection) : null;

  // ===== AP =====
  if (useEditableAp) {
    const apFormula = buildApFormula({
      sheetName,
      rowNumber: rowIndex,
      sCol: "G",
      oCol: "I",
      dCol: "K",
    });

    setFormulaCell(row.getCell(12), apFormula, r.action_priority || "");
  } else {
    row.getCell(12).value = r.action_priority || "";
  }

  row.getCell(15).value = r.recommended_action || "";
  row.getCell(16).value = r.responsibility || r.assigned_to || "";

  // ===== DATES =====
  const targetDate = toExcelDate(r.target_completion_date || r.action_due_date);
  const completionDate = toExcelDate(r.completion_date);

  const cellQ = row.getCell(17);
  cellQ.value = targetDate;
  cellQ.numFmt = "dd.mm.yyyy";

  row.getCell(18).value = r.action_status || "";

  const cellS = row.getCell(19);
  cellS.value = completionDate;
  cellS.numFmt = "dd.mm.yyyy";

  // ===== RESIDUAL =====
  row.getCell(20).value = r.severity_override != null ? Number(r.severity_override) : null;
  row.getCell(21).value = r.occurrence_override != null ? Number(r.occurrence_override) : null;
  row.getCell(22).value = r.detection_override != null ? Number(r.detection_override) : null;

  if (useEditableAp) {
    const residualApFormula = buildApFormula({
      sheetName,
      rowNumber: rowIndex,
      sCol: "T",
      oCol: "U",
      dCol: "V",
    });

    setFormulaCell(row.getCell(23), residualApFormula, r.action_priority_override || "");
  } else {
    row.getCell(23).value = r.action_priority_override || "";
  }

  row.getCell(24).value = r.moderation_notes || "";

  // ✅ APPLY CENTER ALIGNMENT
  centerRowCells(row);

  row.commit();
});

const buffer = await workbook.xlsx.writeBuffer();

res.setHeader(
  "Content-Type",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
);

res.setHeader(
  "Content-Disposition",
  `attachment; filename="${
    useEditableAp ? "P-FMEA-Editable-AP.xlsx" : "P-FMEA-Export.xlsx"
  }"`
);

res.send(buffer);
```

} catch (e) {
console.error("EXPORT ERROR:", e);
res.status(500).send(e?.message || "Export failed");
}
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
console.log(`Running on ${PORT}`);
});
