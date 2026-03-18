console.log("🔥 VERSION 10 - STRICT MAPPING 🔥");

import express from "express";
import ExcelJS from "exceljs";

const app = express();

// CORS
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

    console.log("=== DATA SAMPLE ===");
    console.log(JSON.stringify(rows[0], null, 2));

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("template.xlsx");

    const sheet = workbook.worksheets[0];

    const START_ROW = 16;

    rows.forEach((r, i) => {
      const row = sheet.getRow(START_ROW + i);

      // 🔴 EXACT COLUMN MAPPING (NO FALLBACKS)

      row.getCell(3).value = r.process_step || "";                 // C
      row.getCell(4).value = r.task || "";                         // D
      row.getCell(5).value = r.failure_mode || "";                 // E
      row.getCell(6).value = r.failure_effect || "";               // F
      row.getCell(7).value = r.severity ?? "";                     // G
      row.getCell(8).value = r.failure_cause || "";                // H
      row.getCell(9).value = r.occurrence ?? "";                   // I
      row.getCell(10).value = r.current_controls || "";            // J
      row.getCell(11).value = r.detection ?? "";                   // K
      row.getCell(12).value = r.action_priority || "";             // L
      row.getCell(13).value = r.recommended_action || "";          // M
      row.getCell(14).value = r.responsibility || "";              // N
      row.getCell(15).value = r.target_completion_date || "";      // O
      row.getCell(16).value = r.action_status || "";               // P
      row.getCell(17).value = r.completion_date || "";             // Q

      row.getCell(18).value = r.severity_override ?? "";           // R
      row.getCell(19).value = r.occurrence_override ?? "";         // S
      row.getCell(20).value = r.detection_override ?? "";          // T
      row.getCell(21).value = r.action_priority_override ?? "";    // U

      // V (22) and W (23) intentionally untouched

      row.getCell(24).value = r.moderation_notes || "";            // X

      row.commit();
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

// Railway port (do not change)
const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`Running on ${PORT}`);
});
