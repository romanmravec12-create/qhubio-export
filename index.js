import express from "express";
import ExcelJS from "exceljs";

const app = express();

// CORS (required for browser calls)
app.use((req, res, next) => {
  res.header("Access-Control-Allow-Origin", "*");
  res.header("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.sendStatus(200);
  next();
});

app.use(express.json({ limit: "10mb" }));

// helper: safe getter with fallbacks
const get = (obj, keys, fallback = "") => {
  for (const k of keys) {
    if (obj[k] !== undefined && obj[k] !== null) return obj[k];
  }
  return fallback;
};

app.post("/export", async (req, res) => {
  try {
    const { rows, processName, userName } = req.body;

    if (!rows || !Array.isArray(rows)) {
      return res.status(400).send("Invalid rows");
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("template.xlsx");

    const sheet = workbook.worksheets[0];

    const TEMPLATE_ROW_INDEX = 16;

    rows.forEach((r, i) => {
      const targetRowIndex = TEMPLATE_ROW_INDEX + i;

      // 🔴 CLONE TEMPLATE ROW (preserves styles, merges, formatting)
      const templateRow = sheet.getRow(TEMPLATE_ROW_INDEX);
      const newRow = sheet.getRow(targetRowIndex);

      newRow.model = JSON.parse(JSON.stringify(templateRow.model));

      // 🟩 FLEXIBLE FIELD MAPPING (handles naming differences)
      newRow.getCell(3).value = get(r, ["process_step", "processStep"]);
      newRow.getCell(4).value = get(r, ["task", "function"]);
      newRow.getCell(5).value = get(r, ["failure_mode", "failureMode"]);
      newRow.getCell(6).value = get(r, ["failure_effect", "effect"]);
      newRow.getCell(7).value = get(r, ["severity"]);
      newRow.getCell(8).value = get(r, ["failure_cause", "cause"]);
      newRow.getCell(9).value = get(r, ["occurrence"]);
      newRow.getCell(10).value = get(r, ["current_controls"]);
      newRow.getCell(11).value = get(r, ["detection"]);
      newRow.getCell(12).value = get(r, ["action_priority"]);

      newRow.getCell(13).value = get(r, ["recommended_action"]);
      newRow.getCell(14).value = get(r, ["responsibility"]);
      newRow.getCell(15).value = get(r, ["target_completion_date", "targetDate"]);

      newRow.getCell(16).value = get(r, ["failure_status"]);
      newRow.getCell(17).value = get(r, ["action_status"]);
      newRow.getCell(18).value = get(r, ["completion_date"]);

      // residual
      newRow.getCell(19).value = get(r, ["severity_override"]);
      newRow.getCell(20).value = get(r, ["occurrence_override"]);
      newRow.getCell(21).value = get(r, ["detection_override"]);
      newRow.getCell(22).value = get(r, ["action_priority_override"]);

      if (r.moderation_notes) {
        newRow.getCell(24).value = r.moderation_notes;
      }

      newRow.commit();
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

// Railway port handling (DO NOT CHANGE)
const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
