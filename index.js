console.log("🔥 VERSION 11 - REAL FIX 🔥");

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

// 🔴 universal key resolver (handles snake_case + camelCase)
const get = (obj, ...keys) => {
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

    // 🔴 THIS SHOWS TRUTH
    console.log("=== REAL ROW ===");
    console.log(JSON.stringify(rows[0], null, 2));

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("template.xlsx");

    const sheet = workbook.worksheets[0];

    const START_ROW = 16;

    rows.forEach((r, i) => {
      const row = sheet.getRow(START_ROW + i);

      // 🟩 STRICT COLUMN POSITIONS + FLEXIBLE KEYS

      row.getCell(3).value = get(r, "process_step", "processStep");                // C
      row.getCell(4).value = get(r, "task", "function");                           // D
      row.getCell(5).value = get(r, "failure_mode", "failureMode");                // E
      row.getCell(6).value = get(r, "failure_effect", "failureEffect");            // F
      row.getCell(7).value = get(r, "severity");                                   // G
      row.getCell(8).value = get(r, "failure_cause", "failureCause");              // H
      row.getCell(9).value = get(r, "occurrence");                                 // I
      row.getCell(10).value = get(r, "current_controls", "currentControls");       // J
      row.getCell(11).value = get(r, "detection");                                 // K
      row.getCell(12).value = get(r, "action_priority", "actionPriority");         // L

      row.getCell(13).value = get(r, "recommended_action", "recommendedAction");   // M
      row.getCell(14).value = get(r, "responsibility");                            // N
      row.getCell(15).value = get(r, "target_completion_date", "targetDate");      // O
      row.getCell(16).value = get(r, "action_status", "actionStatus");             // P
      row.getCell(17).value = get(r, "completion_date", "completionDate");         // Q

      row.getCell(18).value = get(r, "severity_override");                         // R
      row.getCell(19).value = get(r, "occurrence_override");                       // S
      row.getCell(20).value = get(r, "detection_override");                        // T
      row.getCell(21).value = get(r, "action_priority_override");                  // U

      row.getCell(24).value = get(r, "moderation_notes", "notes");                 // X

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

// Railway port
const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`Running on ${PORT}`);
});
