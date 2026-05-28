import pptxgen from "pptxgenjs";

function applyCors(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader(
    "Access-Control-Allow-Headers",
    "Content-Type, Authorization"
  );
}

function normalizeAp(ap) {
  const v = String(ap || "").toLowerCase();

  if (v.includes("high")) return "High";
  if (v.includes("medium")) return "Medium";
  return "Low";
}

function barPercent(value, total) {
  if (!total || total <= 0) return 0;
  return Math.round((value / total) * 100);
}

function groupThemes(rows) {
  const map = {};

  rows.forEach((r) => {
    const txt = `
      ${r.process_step || ""}
      ${r.failure_mode || ""}
      ${r.cause || ""}
      ${r.effect || ""}
    `.toLowerCase();

    let theme = "Other Operational Areas";

    if (
      txt.includes("software") ||
      txt.includes("system") ||
      txt.includes("data")
    ) {
      theme = "Software & Data";
    } else if (
      txt.includes("document") ||
      txt.includes("trace") ||
      txt.includes("record")
    ) {
      theme = "Documentation & Traceability";
    } else if (
      txt.includes("label") ||
      txt.includes("identification") ||
      txt.includes("barcode")
    ) {
      theme = "Identification & Traceability";
    }

    if (!map[theme]) {
      map[theme] = [];
    }

    map[theme].push(r);
  });

  return Object.entries(map)
    .map(([theme, items]) => ({
      theme,
      count: items.length,
      sample: items[0],
    }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 4);
}

function addFooter(slide, processName) {
  slide.addText(
    `Qhubio Executive Summary • ${processName}`,
    {
      x: 0.3,
      y: 7.0,
      w: 6,
      h: 0.2,
      fontSize: 9,
      color: "777777",
    }
  );

  slide.addText(
    `Confidential`,
    {
      x: 11,
      y: 7.0,
      w: 1.5,
      h: 0.2,
      align: "right",
      fontSize: 9,
      color: "777777",
    }
  );
}

export default async function handler(req, res) {
  applyCors(req, res);

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({
      success: false,
      error: "Method not allowed",
    });
  }

  try {
    const {
      rows = [],
      processName = "Process",
      companyName = "Company",
      generatedBy = "Qhubio",
    } = req.body;

    if (!Array.isArray(rows) || rows.length === 0) {
      return res.status(400).json({
        success: false,
        error: "No rows provided",
      });
    }

    const pptx = new pptxgen();

    pptx.layout = "LAYOUT_WIDE";
    pptx.author = "Qhubio";
    pptx.company = companyName;
    pptx.subject = "P-FMEA Executive Summary";
    pptx.title = `P-FMEA Executive Summary - ${processName}`;
    pptx.lang = "en-US";

    const total = rows.length;

    const high = rows.filter(
      (r) => normalizeAp(r.action_priority) === "High"
    ).length;

    const medium = rows.filter(
      (r) => normalizeAp(r.action_priority) === "Medium"
    ).length;

    const low = rows.filter(
      (r) => normalizeAp(r.action_priority) === "Low"
    ).length;

    const openActions = rows.filter(
      (r) =>
        !r.action_status ||
        String(r.action_status).toLowerCase() !== "closed"
    ).length;

    const themes = groupThemes(rows);

    const criticalItems = [...rows]
      .filter((r) => normalizeAp(r.action_priority) === "High")
      .slice(0, 6);

    // =========================================================
    // SLIDE 1
    // =========================================================

    {
      const slide = pptx.addSlide();

      slide.background = { color: "F8F6FC" };

      slide.addText("P-FMEA Executive Summary", {
        x: 0.6,
        y: 0.4,
        w: 6,
        h: 0.5,
        fontFace: "Aptos",
        bold: true,
        fontSize: 28,
        color: "4B1F6F",
      });

      slide.addText(processName, {
        x: 0.6,
        y: 1.0,
        w: 9,
        h: 0.3,
        fontSize: 18,
        color: "555555",
      });

      slide.addShape(pptx.ShapeType.rect, {
        x: 0.6,
        y: 1.6,
        w: 2.2,
        h: 1.2,
        fill: { color: "7C3AED" },
        line: { color: "7C3AED" },
        radius: 0.1,
      });

      slide.addText(String(high), {
        x: 0.6,
        y: 1.8,
        w: 2.2,
        h: 0.4,
        align: "center",
        fontSize: 28,
        bold: true,
        color: "FFFFFF",
      });

      slide.addText("High AP Risks", {
        x: 0.6,
        y: 2.3,
        w: 2.2,
        h: 0.2,
        align: "center",
        fontSize: 12,
        color: "FFFFFF",
      });

      slide.addShape(pptx.ShapeType.rect, {
        x: 3.0,
        y: 1.6,
        w: 2.2,
        h: 1.2,
        fill: { color: "F59E0B" },
        line: { color: "F59E0B" },
        radius: 0.1,
      });

      slide.addText(String(openActions), {
        x: 3.0,
        y: 1.8,
        w: 2.2,
        h: 0.4,
        align: "center",
        fontSize: 28,
        bold: true,
        color: "FFFFFF",
      });

      slide.addText("Open Actions", {
        x: 3.0,
        y: 2.3,
        w: 2.2,
        h: 0.2,
        align: "center",
        fontSize: 12,
        color: "FFFFFF",
      });

      slide.addText(
        `The analysis identified ${high} high-priority risks out of ${total} reviewed failure modes.`,
        {
          x: 0.6,
          y: 3.4,
          w: 10,
          h: 0.5,
          fontSize: 18,
          color: "333333",
          bold: true,
        }
      );

      slide.addText(
        "This presentation provides a management-level overview focused on risk exposure, operational concentration, and recommended execution priorities.",
        {
          x: 0.6,
          y: 4.1,
          w: 11,
          h: 1,
          fontSize: 15,
          color: "555555",
        }
      );

      slide.addText(
        `Generated by ${generatedBy} • ${new Date().toLocaleDateString()}`,
        {
          x: 0.6,
          y: 6.4,
          w: 5,
          h: 0.3,
          fontSize: 11,
          color: "777777",
        }
      );

      addFooter(slide, processName);
    }

    // =========================================================
    // SLIDE 2 — PIE CHART
    // =========================================================

    {
      const slide = pptx.addSlide();

      slide.addText("Risk Distribution", {
        x: 0.5,
        y: 0.4,
        w: 5,
        h: 0.4,
        fontSize: 24,
        bold: true,
        color: "4B1F6F",
      });

      slide.addChart(
        pptx.ChartType.pie,
        [
          {
            name: "Risk Split",
            labels: ["High", "Medium", "Low"],
            values: [high, medium, low],
          },
        ],
        {
          x: 0.8,
          y: 1.2,
          w: 4.5,
          h: 4.5,
          showLegend: true,
          showTitle: false,
          dataLabelPosition: "bestFit",
        }
      );

      slide.addText(
        `High AP items represent ${barPercent(high, total)}% of all identified risks.`,
        {
          x: 6,
          y: 2,
          w: 5,
          h: 0.5,
          fontSize: 20,
          bold: true,
          color: "333333",
        }
      );

      slide.addText(
        "Management focus should remain on ownership assignment, mitigation prioritization, and closure tracking of the high-priority exposure areas.",
        {
          x: 6,
          y: 3,
          w: 5.5,
          h: 1.2,
          fontSize: 15,
          color: "555555",
        }
      );

      addFooter(slide, processName);
    }

    // =========================================================
    // SLIDE 3 — THEMES BAR CHART
    // =========================================================

    {
      const slide = pptx.addSlide();

      slide.addText("Top Risk Themes", {
        x: 0.5,
        y: 0.4,
        w: 5,
        h: 0.4,
        fontSize: 24,
        bold: true,
        color: "4B1F6F",
      });

      slide.addChart(
        pptx.ChartType.bar,
        [
          {
            name: "Themes",
            labels: themes.map((t) => t.theme),
            values: themes.map((t) => t.count),
          },
        ],
        {
          x: 0.7,
          y: 1.2,
          w: 5.5,
          h: 4.5,
          catAxisLabelFontSize: 12,
          valAxisLabelFontSize: 10,
          showLegend: false,
        }
      );

      let y = 1.3;

      themes.forEach((t) => {
        slide.addText(
          `${t.theme}: ${t.count} risks`,
          {
            x: 6.6,
            y,
            w: 5,
            h: 0.3,
            fontSize: 16,
            bold: true,
            color: "333333",
          }
        );

        y += 0.45;
      });

      addFooter(slide, processName);
    }

    // =========================================================
    // SLIDE 4 — CRITICAL ITEMS
    // =========================================================

    {
      const slide = pptx.addSlide();

      slide.addText("Top Critical Items", {
        x: 0.5,
        y: 0.4,
        w: 5,
        h: 0.4,
        fontSize: 24,
        bold: true,
        color: "4B1F6F",
      });

      const tableRows = [
        [
          { text: "Process Step", options: { bold: true } },
          { text: "Failure Mode", options: { bold: true } },
          { text: "AP", options: { bold: true } },
          { text: "Recommended Action", options: { bold: true } },
        ],
      ];

      criticalItems.forEach((r) => {
        tableRows.push([
          r.process_step || "",
          r.failure_mode || "",
          normalizeAp(r.action_priority),
          r.recommended_action || "",
        ]);
      });

      slide.addTable(tableRows, {
        x: 0.4,
        y: 1.1,
        w: 12.2,
        border: {
          type: "solid",
          color: "CCCCCC",
          pt: 1,
        },
        fontSize: 11,
        color: "333333",
        fill: "FFFFFF",
      });

      addFooter(slide, processName);
    }

    // =========================================================
    // SLIDE 5 — NEXT STEPS
    // =========================================================

    {
      const slide = pptx.addSlide();

      slide.addText("Recommended Next Steps", {
        x: 0.5,
        y: 0.4,
        w: 6,
        h: 0.4,
        fontSize: 24,
        bold: true,
        color: "4B1F6F",
      });

      const steps = [
        "Assign ownership for all High AP items.",
        "Validate mitigation feasibility with process owners.",
        "Track closure progress in the Excel execution register.",
        "Review recurring operational themes quarterly.",
      ];

      let y = 1.5;

      steps.forEach((s, i) => {
        slide.addShape(pptx.ShapeType.ellipse, {
          x: 0.7,
          y,
          w: 0.35,
          h: 0.35,
          fill: { color: "7C3AED" },
          line: { color: "7C3AED" },
        });

        slide.addText(String(i + 1), {
          x: 0.77,
          y: y + 0.02,
          w: 0.2,
          h: 0.2,
          fontSize: 10,
          bold: true,
          color: "FFFFFF",
          align: "center",
        });

        slide.addText(s, {
          x: 1.2,
          y: y - 0.02,
          w: 9,
          h: 0.3,
          fontSize: 18,
          color: "333333",
        });

        y += 0.9;
      });

      addFooter(slide, processName);
    }

    // =========================================================
    // SLIDE 6 — CLOSING
    // =========================================================

    {
      const slide = pptx.addSlide();

      slide.background = { color: "4B1F6F" };

      slide.addText("Thank you", {
        x: 0.8,
        y: 2,
        w: 4,
        h: 0.5,
        fontSize: 30,
        bold: true,
        color: "FFFFFF",
      });

      slide.addText(
        "Use this executive summary to guide the management review discussion and action prioritization.",
        {
          x: 0.8,
          y: 3,
          w: 8,
          h: 1,
          fontSize: 18,
          color: "FFFFFF",
        }
      );

      slide.addText(
        `${processName} • ${total} failure modes reviewed`,
        {
          x: 0.8,
          y: 5.8,
          w: 6,
          h: 0.3,
          fontSize: 13,
          color: "DDDDDD",
        }
      );
    }

    const buffer = await pptx.write({
      outputType: "nodebuffer",
    });

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="P-FMEA-Executive-Summary.pptx"`
    );

    return res.status(200).send(buffer);

  } catch (e) {
    console.error("PPTX EXPORT ERROR:", e);

    return res.status(500).json({
      success: false,
      error: e.message || "PPTX export failed",
      stage: "write-buffer",
    });
  }
}
