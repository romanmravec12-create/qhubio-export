console.log("🔥 VERSION 8 - FINAL BUILD 🔥");

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

// helper to resolve unknown key names (THIS is what you were missing)
const pick = (obj, ...keys) => {
  for (const k of keys) {
    if (obj[k] !== undefined && obj[k] !== null) return obj[k];
  }
  return "";
};

app.post("/export", async (req, res) => {
  try {
    const { rows, processName, userName } = req.body;

    if (!rows || !Array.isArray(rows)) {
      return res.status(400).send("Invalid rows");
    }

    // 🔴 DEBUG — THIS WILL SHOW REAL STRUCTURE
    console.log("=== REAL ROW DATA ===");
    console.log(JSON.stringify(rows[0], null, 2));

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("template.xlsx");

    const sheet = workbook.worksheets[0];

    const START_ROW = 16;

    rows.forEach((r, i) => {
      const rowIndex = START_ROW + i;

      // 🔴 CRITICAL: insert row instead of overwriting → preserves formatting
      sheet.spliceRows(rowIndex, 0, []);

      const row = sheet.getRow(rowIndex);

      // EXACT COLUMN ORDER (C → V) — aligned with your spec
      row.getCell(3).value = pick(r, "process_step", "processStep");
      row.getCell(4).value = pick(r, "task", "function");
      row.getCell(5).value = pick(r, "failure_mode", "failureMode");
      row.getCell(6).value = pick(r, "failure_effect", "effect");
      row.getCell(7).value = pick(r, "severity");
      row.getCell(8).value = pick(r, "failure_cause", "cause");
      row.getCell(9).value = pick(r, "occurrence");
      row.getCell(10).value = pick(r, "current_controls");
      row.getCell(11).value = pick(r, "detection");
      row.getCell(12).value = pick(r, "action_priority");

      row.getCell(13).value = pick(r, "recommended_action");
      row.getCell(14).value = pick(r, "responsibility");
      row.getCell(15).value = pick(r, "target_completion_date", "targetDate");

      row.getCell(16).value = pick(r, "failure_status");
      row.getCell(17).value = pick(r, "action_status");
      row.getCell(18).value = pick(r, "completion_date");

      row.getCell(19).value = pick(r, "severity_override");
      row.getCell(20).value = pick(r, "occurrence_override");
      row.getCell(21).value = pick(r, "detection_override");
      row.getCell(22).value = pick(r, "action_priority_override");

      const notes = pick(r, "moderation_notes");
      if (notes) {
        row.getCell(24).value = notes;
      }

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

// Railway port (DO NOT TOUCH)
const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`Running on ${PORT}`);
});
