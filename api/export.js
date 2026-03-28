import ExcelJS from "exceljs";
import path from "path";
import fs from "fs";

// ✅ SAFE formula builder (same as Railway)
function quoteSheetName(sheetName) {
  return "'" + String(sheetName).replace(/'/g, "''") + "'";
}

function buildApFormula({ sheetName, rowNumber, sCol, oCol, dCol }) {
  const fmeaSheet = quoteSheetName(sheetName);

  return (
    "IFERROR(INDEX('AP Table'!$E$3:$E$1002, MATCH(1, " +
    "('AP Table'!$B$3:$B$1002=" + fmeaSheet + "!" + sCol + rowNumber + ")*" +
    "('AP Table'!$C$3:$C$1002=" + fmeaSheet + "!" + oCol + rowNumber + ")*" +
    "('AP Table'!$D$3:$D$1002=" + fmeaSheet + "!" + dCol + rowNumber + "), 0)), \"\")"
  );
}

// ✅ Copy styles + enforce center alignment
function copyTemplateRowStyle(templateRow, targetRow) {
  templateRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const targetCell = targetRow.getCell(colNumber);

    targetCell.style = JSON.parse(JSON.stringify(cell.style || {}));

    if (cell.numFmt) targetCell.numFmt = cell.numFmt;
    if (cell.font) targetCell.font = JSON.parse(JSON.stringify(cell.font));
    if (cell.fill) targetCell.fill = JSON.parse(JSON.stringify(cell.fill));
    if (cell.border) targetCell.border = JSON.parse(JSON.stringify(cell.border));

    targetCell.alignment = {
      ...(cell.alignment || {}),
      horizontal: "center",
      vertical: "middle",
      wrapText: true,
    };
  });

  if (templateRow.height) {
    targetRow.height = templateRow.height;
  }
}

function setFormulaCell(cell, formula, result) {
  cell.value = {
    formula: formula,
    result: result || "",
  };

  cell.alignment = {
    horizontal: "center",
    vertical: "middle",
  };
}

function toExcelDate(value) {
  if (!value) return null;
  const d = new Date(value);
  return isNaN(d) ? null : d;
}

export default async function handler(req, res) {
  // ✅ CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const { rows = [], exportFormat = "styled", editableAp = false } = req.body;

    const useEditableAp =
      editableAp === true ||
      exportFormat === "editable-ap" ||
      exportFormat === "excel-formatted-editable-ap" ||
      exportFormat === "advanced";

    const templateFile = useEditableAp
      ? "template-editable-ap.xlsx"
      : "template.xlsx";

    // ✅ VERCEL PATH FIX
    const templatePath = path.join(process.cwd(), "templates", templateFile);

    if (!fs.existsSync(templatePath)) {
      console.error("❌ Template not found:", templatePath);
      return res.status(500).json({ error: "Template missing on server" });
    }

    const fileBuffer = fs.readFileSync(templatePath);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(fileBuffer);

    const sheet = workbook.worksheets[0];
    const sheetName = sheet.name || "FMEA";

    const START_ROW = 16;
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
      row.getCell(20).value =
        r.severity_override != null ? Number(r.severity_override) : null;

      row.getCell(21).value =
        r.occurrence_override != null ? Number(r.occurrence_override) : null;

      row.getCell(22).value =
        r.detection_override != null ? Number(r.detection_override) : null;

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

      row.commit();
    });

    const buffer = await workbook.xlsx.writeBuffer();

    const filename = useEditableAp
      ? "P-FMEA-Editable-AP.xlsx"
      : "P-FMEA-Export.xlsx";

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    res.setHeader(
      "Content-Disposition",
      "attachment; filename=\"" + filename + "\""
    );

    return res.status(200).send(buffer);
  } catch (e) {
    console.error("EXPORT ERROR:", e);
    return res.status(500).json({ error: "Export failed" });
  }
}
