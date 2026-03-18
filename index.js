console.log("🔥 VERSION 13 - FINAL ALIGNMENT FIX 🔥");

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

    const sheet = workbook.worksheets[0];

    const START_ROW = 16;

    rows.forEach((r, i) => {
      const row = sheet.getRow(START_ROW + i);

      // LEFT SIDE (already correct)
      row.getCell(3).value = r.process_step || "";   // C
      row.getCell(4).value = r.function || "";       // D
      row.getCell(5).value = r.failure_mode || "";   // E
      row.getCell(6).value = r.effect || "";         // F
      row.getCell(7).value = r.severity ?? "";       // G
      row.getCell(8).value = r.cause || "";          // H
      row.getCell(9).value = r.occurrence ?? "";     // I

      row.getCell(10).value =
        `${r.current_prevention_controls || ""}\n${r.current_detection_controls || ""}`.trim(); // J

      row.getCell(11).value = r.detection ?? "";     // K
      row.getCell(12).value = r.action_priority || ""; // L

      // 🔴 FIXED RIGHT SIDE (shifted to correct columns)

      row.getCell(15).value = r.recommended_action || "";           // O
      row.getCell(16).value = r.responsibility || r.assigned_to || ""; // P
      row.getCell(17).value = r.target_completion_date || r.action_due_date || ""; // Q
      row.getCell(18).value = r.action_status || "";                // R
      row.getCell(19).value = r.completion_date || "";              // S

      row.getCell(20).value = r.severity_override ?? "";            // T
      row.getCell(21).value = r.occurrence_override ?? "";          // U
      row.getCell(22).value = r.detection_override ?? "";           // V
      row.getCell(23).value = r.action_priority_override ?? "";     // W

      row.getCell(24).value = r.moderation_notes || "";             // X

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
