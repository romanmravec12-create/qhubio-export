import ExcelJS from "exceljs";
import path from "path";

export default async function handler(req, res) {
  // ✅ CORS (CRITICAL FOR VERCEL)
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  // ✅ Handle preflight request
  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  // ❌ Block other methods
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

    // ✅ Vercel-safe path
    const templatePath = path.join(process.cwd(), "templates", templateFile);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templatePath);

    const sheet = workbook.worksheets[0];

    const START_ROW = 16;
    const templateRow = sheet.getRow(START_ROW);

    function toExcelDate(value) {
      if (!value) return null;
      const d = new Date(value);
      return isNaN(d) ? null : d;
    }

    rows.forEach((r, i) => {
      const rowIndex = START_ROW + i;
      const row = sheet.getRow(rowIndex);

      // ✅ Copy style + enforce CENTER alignment
      templateRow.eachCell({ includeEmpty: true }, (cell, col) => {
        const target = row.getCell(col);

        target.style = JSON.parse(JSON.stringify(cell.style || {}));

        target.alignment = {
          horizontal: "center",
          vertical: "middle",
          wrapText: true,
        };
      });

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

      // ===== AP (STATIC HERE — OK FOR YOUR CURRENT SETUP) =====
      row.getCell(12).value = r.action_priority || "";

      row.getCell(15).value = r.recommended_action || "";
      row.getCell(16).value = r.responsibility || "";

      // ===== DATES =====
      const d1 = toExcelDate(r.target_completion_date);
      const d2 = toExcelDate(r.completion_date);

      const c1 = row.getCell(17);
      c1.value = d1;
      c1.numFmt = "dd.mm.yyyy";

      row.getCell(18).value = r.action_status || "";

      const c2 = row.getCell(19);
      c2.value = d2;
      c2.numFmt = "dd.mm.yyyy";

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
