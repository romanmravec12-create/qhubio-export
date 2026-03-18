import express from "express";
import ExcelJS from "exceljs";

const app = express();

// CORS (needed for browser → Railway)
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

    if (!rows || !Array.isArray(rows)) {
      return res.status(400).send("Invalid rows");
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("template.xlsx");

    const sheet = workbook.worksheets[0];

    rows.forEach((r, i) => {
      const excelRow = sheet.getRow(16 + i);

      // MAIN DATA (C → R)
      excelRow.getCell(3).value = r.process_step || "";                 // C
      excelRow.getCell(4).value = r.task || "";                         // D
      excelRow.getCell(5).value = r.failure_mode || "";                 // E
      excelRow.getCell(6).value = r.failure_effect || "";               // F
      excelRow.getCell(7).value = r.severity ?? "";                     // G
      excelRow.getCell(8).value = r.failure_cause || "";                // H
      excelRow.getCell(9).value = r.occurrence ?? "";                   // I
      excelRow.getCell(10).value = r.current_controls || "";            // J
      excelRow.getCell(11).value = r.detection ?? "";                   // K
      excelRow.getCell(12).value = r.action_priority || "";             // L (initial AP)
      excelRow.getCell(13).value = r.recommended_action || "";          // M
      excelRow.getCell(14).value = r.responsibility || "";              // N
      excelRow.getCell(15).value = r.target_completion_date || "";      // O
      excelRow.getCell(16).value = r.failure_status || "";              // P
      excelRow.getCell(17).value = r.action_status || "";               // Q
      excelRow.getCell(18).value = r.completion_date || "";             // R

      // RESIDUAL / MODERATION (S → V)
      excelRow.getCell(19).value = r.severity_override ?? "";           // S
      excelRow.getCell(20).value = r.occurrence_override ?? "";         // T
      excelRow.getCell(21).value = r.detection_override ?? "";          // U
      excelRow.getCell(22).value = r.action_priority_override ?? "";    // V

      // NOTES (X only if exists)
      if (r.moderation_notes && r.moderation_notes.trim() !== "") {
        excelRow.getCell(24).value = r.moderation_notes;                // X
      }

      excelRow.commit();
    });

    const buffer = await workbook.xlsx.writeBuffer();

    const safeProcess = (processName || "Export").replace(/[^a-z0-9]/gi, "_");
    const safeUser = (userName || "User").replace(/[^a-z0-9]/gi, "_");

    const today = new Date();
    const dateStr = `${String(today.getDate()).padStart(2, "0")}.${String(
      today.getMonth() + 1
    ).padStart(2, "0")}.${today.getFullYear()}`;

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="P-FMEA_${safeProcess}_${safeUser}_${dateStr}.xlsx"`
    );

    res.send(buffer);
  } catch (e) {
    console.error("EXPORT ERROR:", e);
    res.status(500).send(e.message || "Export failed");
  }
});

// PORT (Railway handles it automatically)
const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
