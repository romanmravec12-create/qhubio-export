console.log("🔥 VERSION 24 - FINAL INDEX MATCH 🔥");

import express from "express";
import ExcelJS from "exceljs";

const app = express();

app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.sendStatus(200);
  next();
});

app.use(express.json({ limit: "10mb" }));

app.post("/export", async (req, res) => {
  try {
    const { rows } = req.body;

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("template.xlsx");

    workbook.calcProperties.fullCalcOnLoad = true;

    const sheet = workbook.worksheets[0];

    const START_ROW = 16;
    const templateRow = sheet.getRow(START_ROW);

    rows.forEach((r, i) => {
      const rowIndex = START_ROW + i;
      const row = sheet.getRow(rowIndex);

      // styles
      templateRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        row.getCell(colNumber).style = JSON.parse(JSON.stringify(cell.style));
      });

      // data
      row.getCell(3).value = r.process_step || "";
      row.getCell(4).value = r.function || "";
      row.getCell(5).value = r.failure_mode || "";
      row.getCell(6).value = r.effect || "";

      row.getCell(7).value = r.severity != null ? Number(r.severity) : null;
      row.getCell(8).value = r.cause || "";
      row.getCell(9).value = r.occurrence != null ? Number(r.occurrence) : null;

      row.getCell(10).value =
        `${r.current_prevention_controls || ""}\n${r.current_detection_controls || ""}`.trim();

      row.getCell(11).value = r.detection != null ? Number(r.detection) : null;

      // 🔴 FINAL FORMULA (INDEX MATCH, NO ARRAYS)
      row.getCell(12).value = {
        formula: `INDEX('AP Table'!E4:E1003, MATCH(VALUE(G${rowIndex}&I${rowIndex}&K${rowIndex}), 'AP Table'!A4:A1003, 0))`
      };

      row.getCell(15).value = r.recommended_action || "";
      row.getCell(16).value = r.responsibility || r.assigned_to || "";
      row.getCell(17).value = r.target_completion_date || r.action_due_date || "";
      row.getCell(18).value = r.action_status || "";
      row.getCell(19).value = r.completion_date || "";

      row.getCell(20).value = r.severity_override != null ? Number(r.severity_override) : null;
      row.getCell(21).value = r.occurrence_override != null ? Number(r.occurrence_override) : null;
      row.getCell(22).value = r.detection_override != null ? Number(r.detection_override) : null;

      row.getCell(23).value = {
        formula: `INDEX('AP Table'!E4:E1003, MATCH(VALUE(T${rowIndex}&U${rowIndex}&V${rowIndex}), 'AP Table'!A4:A1003, 0))`
      };

      row.getCell(24).value = r.moderation_notes || "";

      row.commit();
    });

    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    res.send(buffer);
  } catch (e) {
    console.error("EXPORT ERROR:", e);
    res.status(500).send(e.message);
  }
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`Running on ${PORT}`);
});
