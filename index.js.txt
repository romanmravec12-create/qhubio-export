import express from "express";
import ExcelJS from "exceljs";

const app = express();
app.use(express.json({ limit: "10mb" }));

app.post("/export", async (req, res) => {
  try {
    const { rows, processName, userName } = req.body;

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("template.xlsx");

    const sheet = workbook.worksheets[0];

    rows.forEach((r, i) => {
      const row = sheet.getRow(16 + i);

      row.getCell(3).value = r.process_step || "";
      row.getCell(4).value = r.function || "";
      row.getCell(5).value = r.failure_mode || "";
      row.getCell(6).value = r.effect || "";
      row.getCell(7).value = r.severity ?? "";
      row.getCell(8).value = r.cause || "";
      row.getCell(9).value = r.occurrence ?? "";
      row.getCell(10).value = r.current_controls || "";
      row.getCell(11).value = r.detection ?? "";
      row.getCell(12).value = r.action_priority || "";
      row.getCell(13).value = r.recommended_action || "";
      row.getCell(14).value = r.responsibility || "";
      row.getCell(15).value = r.target_completion_date || "";
      row.getCell(16).value = r.action_status || "";
      row.getCell(17).value = r.completion_date || "";
      row.getCell(18).value = r.severity_override ?? "";
      row.getCell(19).value = r.occurrence_override ?? "";
      row.getCell(20).value = r.detection_override ?? "";
      row.getCell(21).value = r.action_priority_override ?? "";

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

app.listen(3000, () => console.log("Running on 3000"));