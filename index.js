console.log("🔥 VERSION 7 - NEW BUILD 🔥");
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
    const { rows, processName, userName } = req.body;

    // 🔴 THIS IS THE KEY
    console.log("=== FIRST ROW DATA ===");
    console.log(JSON.stringify(rows[0], null, 2));

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("template.xlsx");

    const sheet = workbook.worksheets[0];

    const START_ROW = 16;

    rows.forEach((r, i) => {
      const row = sheet.getRow(START_ROW + i);

      row.getCell(3).value = r.process_step || "";
      row.getCell(4).value = r.task || "";
      row.getCell(5).value = r.failure_mode || "";
      row.getCell(6).value = r.failure_effect || "";
      row.getCell(7).value = r.severity ?? "";
      row.getCell(8).value = r.failure_cause || "";
      row.getCell(9).value = r.occurrence ?? "";
      row.getCell(10).value = r.current_controls || "";
      row.getCell(11).value = r.detection ?? "";
      row.getCell(12).value = r.action_priority || "";
      row.getCell(13).value = r.recommended_action || "";
      row.getCell(14).value = r.responsibility || "";
      row.getCell(15).value = r.target_completion_date || "";
      row.getCell(16).value = r.failure_status || "";
      row.getCell(17).value = r.action_status || "";
      row.getCell(18).value = r.completion_date || "";

      row.getCell(19).value = r.severity_override ?? "";
      row.getCell(20).value = r.occurrence_override ?? "";
      row.getCell(21).value = r.detection_override ?? "";
      row.getCell(22).value = r.action_priority_override ?? "";

      if (r.moderation_notes) {
        row.getCell(24).value = r.moderation_notes;
      }

      row.commit();
    });

    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    res.send(buffer);
  } catch (e) {
    console.error(e);
    res.status(500).send(e.message);
  }
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`Running on ${PORT}`);
});
