import PptxGenJS from "pptxgenjs";

// =========================
// CORS
// =========================
function applyCors(res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
}

// =========================
// HELPERS
// =========================
function safe(value) {
  return value == null ? "" : String(value);
}

function percent(value, total) {
  if (!total) return 0;
  return Math.round((value / total) * 100);
}

function normalizeText(text) {
  return safe(text).toLowerCase();
}

function detectTheme(row) {
  const combined = [
    row.process_step,
    row.failure_mode,
    row.cause,
    row.effect,
    row.recommended_action,
  ]
    .map(normalizeText)
    .join(" ");

  // ===== IMPORTANT =====
  // Restored deterministic grouping logic
  // Keep this stable and predictable

  if (
    combined.includes("software") ||
    combined.includes("system") ||
    combined.includes("interface") ||
    combined.includes("data") ||
    combined.includes("erp") ||
    combined.includes("sap") ||
    combined.includes("api")
  ) {
    return "Software & Data";
  }

  if (
    combined.includes("traceability") ||
    combined.includes("identification") ||
    combined.includes("barcode") ||
    combined.includes("serial") ||
    combined.includes("label")
  ) {
    return "Identification & Traceability";
  }

  if (
    combined.includes("documentation") ||
    combined.includes("record") ||
    combined.includes("signature") ||
    combined.includes("report") ||
    combined.includes("inspection")
  ) {
    return "Documentation & Traceability";
  }

  return "Other Operational Areas";
}

function buildThemeStats(rows) {
  const map = {};

  rows.forEach((row) => {
    const theme = detectTheme(row);

    if (!map[theme]) {
      map[theme] = {
        name: theme,
        count: 0,
        sample: row,
      };
    }

    map[theme].count += 1;
  });

  return Object.values(map).sort((a, b) => b.count - a.count);
}

function topCriticalItems(rows) {
  return rows
    .filter((r) => safe(r.action_priority).toLowerCase().includes("high"))
    .slice(0, 6);
}

function addTitle(slide, title, subtitle, slideNumber) {
  slide.addShape("rect", {
    x: 0,
    y: 0,
    w: 13.33,
    h: 0.6,
    fill: { color: "5B2D91" },
    line: { color: "5B2D91" },
  });

  slide.addText(title, {
    x: 0.4,
    y: 0.12,
    w: 5,
    h: 0.3,
    fontSize: 24,
    bold: true,
    color: "FFFFFF",
  });

  slide.addText(subtitle, {
    x: 0.4,
    y: 0.38,
    w: 5,
    h: 0.2,
    fontSize: 9,
    color: "F3EFFF",
  });

  slide.addText(String(slideNumber), {
    x: 12.9,
    y: 0.12,
    w: 0.2,
    h: 0.2,
    fontSize: 10,
    color: "FFFFFF",
    bold: true,
  });
}

function addFooter(slide) {
  slide.addText("Confidential", {
    x: 11.7,
    y: 7.15,
    w: 1.2,
    h: 0.2,
    fontSize: 8,
    color: "777777",
    italic: true,
  });
}

// =========================
// MAIN
// =========================
export default async function handler(req, res) {
  applyCors(res);

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
      processName = "P-FMEA Process",
      userName = "User",
    } = req.body || {};

    if (!Array.isArray(rows) || rows.length === 0) {
      return res.status(400).json({
        success: false,
        error: "No rows provided",
      });
    }

    const pptx = new PptxGenJS();

    pptx.layout = "LAYOUT_WIDE";
    pptx.author = "Qhubio";
    pptx.company = "Qhubio";
    pptx.subject = "P-FMEA Executive Summary";
    pptx.title = processName;
    pptx.lang = "en-US";
    pptx.theme = {
      headFontFace: "Aptos",
      bodyFontFace: "Aptos",
      lang: "en-US",
    };

    // =========================
    // STATS
    // =========================
    const total = rows.length;

    const high = rows.filter((r) =>
      safe(r.action_priority).toLowerCase().includes("high")
    ).length;

    const medium = rows.filter((r) =>
      safe(r.action_priority).toLowerCase().includes("medium")
    ).length;

    const low = rows.filter((r) =>
      safe(r.action_priority).toLowerCase().includes("low")
    ).length;

    const highPct = percent(high, total);
    const mediumPct = percent(medium, total);
    const lowPct = percent(low, total);

    const themes = buildThemeStats(rows);

    // =========================
    // SLIDE 1
    // =========================
    {
      const slide = pptx.addSlide();

      addTitle(
        slide,
        "P-FMEA Executive Summary",
        processName,
        1
      );

      slide.addText(
        `Generated for ${userName}`,
        {
          x: 0.5,
          y: 1.0,
          w: 4,
          h: 0.3,
          fontSize: 16,
          color: "5B2D91",
          bold: true,
        }
      );

      slide.addText(
        `Total reviewed failure modes: ${total}`,
        {
          x: 0.5,
          y: 1.4,
          w: 5,
          h: 0.3,
          fontSize: 20,
          bold: true,
        }
      );

      slide.addText(
        `High AP risks identified: ${high}`,
        {
          x: 0.5,
          y: 1.8,
          w: 5,
          h: 0.3,
          fontSize: 20,
          bold: true,
          color: "D62828",
        }
      );

      slide.addText(
        "This presentation provides a management-level overview focused on operational exposure, dominant risk concentration, and execution priorities.",
        {
          x: 0.5,
          y: 2.5,
          w: 7,
          h: 1,
          fontSize: 16,
          color: "444444",
        }
      );

      addFooter(slide);
    }

    // =========================
    // SLIDE 2
    // =========================
    {
      const slide = pptx.addSlide();

      addTitle(
        slide,
        "Executive Summary",
        "High-level operational overview",
        2
      );

      slide.addText(
        `• ${highPct}% of all risks are classified as High AP\n\n• ${total} actions currently require tracking and execution\n\n• Main operational focus should remain on ownership, mitigation prioritization, and execution follow-up\n\n• Dominant recurring themes indicate concentrated operational exposure`,
        {
          x: 0.8,
          y: 1.2,
          w: 10,
          h: 3,
          fontSize: 18,
          color: "333333",
          breakLine: false,
        }
      );

      addFooter(slide);
    }

    // =========================
    // SLIDE 3 - FIXED
    // =========================
    {
      const slide = pptx.addSlide();

      addTitle(
        slide,
        "Risk Distribution",
        "Quick view of the Action Priority breakdown",
        3
      );

      slide.addText("Action Priority split", {
        x: 0.5,
        y: 0.9,
        w: 3,
        h: 0.3,
        fontSize: 18,
        bold: true,
        color: "5B2D91",
      });

      const totalWidth = 9.5;
      const highW = (high / total) * totalWidth;
      const medW = (medium / total) * totalWidth;
      const lowW = (low / total) * totalWidth;

      slide.addShape("rect", {
        x: 0.7,
        y: 1.4,
        w: highW,
        h: 0.45,
        fill: { color: "E52521" },
        line: { color: "E52521" },
      });

      slide.addShape("rect", {
        x: 0.7 + highW,
        y: 1.4,
        w: medW,
        h: 0.45,
        fill: { color: "F28C28" },
        line: { color: "F28C28" },
      });

      slide.addShape("rect", {
        x: 0.7 + highW + medW,
        y: 1.4,
        w: lowW,
        h: 0.45,
        fill: { color: "2CA02C" },
        line: { color: "2CA02C" },
      });

      slide.addText(`High ${high}`, {
        x: 0.7,
        y: 1.47,
        w: highW,
        h: 0.2,
        fontSize: 11,
        bold: true,
        align: "center",
        color: "FFFFFF",
      });

      // ===== KPI CARDS =====
      const cards = [
        {
          title: "High AP",
          count: high,
          pct: highPct,
          color: "E52521",
          x: 0.6,
        },
        {
          title: "Medium AP",
          count: medium,
          pct: mediumPct,
          color: "F28C28",
          x: 4.2,
        },
        {
          title: "Low AP",
          count: low,
          pct: lowPct,
          color: "2CA02C",
          x: 7.8,
        },
      ];

      cards.forEach((c) => {
        slide.addShape("rect", {
          x: c.x,
          y: 2.2,
          w: 3.0,
          h: 1.2,
          fill: { color: "FFFFFF" },
          line: { color: "DADADA" },
        });

        slide.addShape("rect", {
          x: c.x,
          y: 2.2,
          w: 0.18,
          h: 1.2,
          fill: { color: c.color },
          line: { color: c.color },
        });

        slide.addText(c.title, {
          x: c.x + 0.25,
          y: 2.35,
          w: 1.5,
          h: 0.2,
          fontSize: 13,
          bold: true,
          color: "666666",
        });

        slide.addText(`${c.count} (${c.pct}%)`, {
          x: c.x + 0.25,
          y: 2.58,
          w: 2,
          h: 0.3,
          fontSize: 22,
          bold: true,
          color: "222222",
        });
      });

      slide.addShape("rect", {
        x: 0.6,
        y: 3.7,
        w: 10.2,
        h: 1,
        fill: { color: "F7F7F7" },
        line: { color: "DADADA" },
      });

      slide.addText("Action focus", {
        x: 0.8,
        y: 3.9,
        w: 2,
        h: 0.2,
        fontSize: 16,
        bold: true,
        color: "5B2D91",
      });

      slide.addText(
        `Open actions: ${total} open actions (${percent(total, total)}%) • Completed/closed: 0`,
        {
          x: 0.8,
          y: 4.15,
          w: 6,
          h: 0.2,
          fontSize: 14,
          color: "333333",
        }
      );

      addFooter(slide);
    }

    // =========================
    // SLIDE 4 - FIXED THEMES
    // =========================
    {
      const slide = pptx.addSlide();

      addTitle(
        slide,
        "Top Risk Themes",
        "Consolidated risk clusters based on deterministic grouping",
        4
      );

      const colors = [
        "6B7A90",
        "2F6BFF",
        "1DA1F2",
        "E84393",
      ];

      themes.slice(0, 4).forEach((theme, idx) => {
        const x = idx % 2 === 0 ? 0.6 : 6.1;
        const y = idx < 2 ? 1.0 : 3.3;

        slide.addShape("rect", {
          x,
          y,
          w: 4.8,
          h: 1.8,
          fill: { color: "FFFFFF" },
          line: { color: "DADADA" },
        });

        slide.addShape("rect", {
          x,
          y,
          w: 0.18,
          h: 1.8,
          fill: { color: colors[idx] },
          line: { color: colors[idx] },
        });

        slide.addText(theme.name, {
          x: x + 0.3,
          y: y + 0.18,
          w: 3.5,
          h: 0.2,
          fontSize: 18,
          bold: true,
          color: "5B2D91",
        });

        slide.addText(`Count: ${theme.count}`, {
          x: x + 0.3,
          y: y + 0.5,
          w: 2,
          h: 0.2,
          fontSize: 14,
          bold: true,
          color: colors[idx],
        });

        slide.addText(
          `Representative item: ${safe(theme.sample.process_step)} — ${safe(theme.sample.failure_mode)}`,
          {
            x: x + 0.3,
            y: y + 0.9,
            w: 4,
            h: 0.4,
            fontSize: 11,
            color: "333333",
          }
        );
      });

      addFooter(slide);
    }

    // =========================
    // SLIDE 5
    // =========================
    {
      const slide = pptx.addSlide();

      addTitle(
        slide,
        "Top Critical Items",
        "Highest-priority execution focus areas",
        5
      );

      const critical = topCriticalItems(rows);

      slide.addText(
        "Process Step",
        {
          x: 0.5,
          y: 1.0,
          w: 2,
          h: 0.2,
          bold: true,
          fontSize: 12,
          color: "FFFFFF",
        }
      );

      slide.addText(
        "Failure Mode",
        {
          x: 2.6,
          y: 1.0,
          w: 3.2,
          h: 0.2,
          bold: true,
          fontSize: 12,
          color: "FFFFFF",
        }
      );

      slide.addText(
        "AP",
        {
          x: 6.0,
          y: 1.0,
          w: 0.5,
          h: 0.2,
          bold: true,
          fontSize: 12,
          color: "FFFFFF",
        }
      );

      slide.addText(
        "Recommended Action",
        {
          x: 6.8,
          y: 1.0,
          w: 4.5,
          h: 0.2,
          bold: true,
          fontSize: 12,
          color: "FFFFFF",
        }
      );

      slide.addShape("rect", {
        x: 0.4,
        y: 0.9,
        w: 11,
        h: 0.35,
        fill: { color: "5B2D91" },
        line: { color: "5B2D91" },
      });

      critical.forEach((item, idx) => {
        const y = 1.4 + idx * 0.7;

        slide.addShape("line", {
          x: 0.4,
          y,
          w: 11,
          h: 0,
          line: { color: "DDDDDD", pt: 1 },
        });

        slide.addText(safe(item.process_step), {
          x: 0.5,
          y: y + 0.08,
          w: 2,
          h: 0.3,
          fontSize: 10,
        });

        slide.addText(safe(item.failure_mode), {
          x: 2.6,
          y: y + 0.08,
          w: 3.2,
          h: 0.3,
          fontSize: 10,
        });

        slide.addShape("roundRect", {
          x: 6.0,
          y: y + 0.04,
          w: 0.6,
          h: 0.25,
          rectRadius: 0.04,
          fill: { color: "E52521" },
          line: { color: "E52521" },
        });

        slide.addText("HIGH", {
          x: 6.02,
          y: y + 0.07,
          w: 0.55,
          h: 0.1,
          fontSize: 7,
          bold: true,
          align: "center",
          color: "FFFFFF",
        });

        slide.addText(safe(item.recommended_action), {
          x: 6.8,
          y: y + 0.08,
          w: 4.3,
          h: 0.3,
          fontSize: 10,
        });
      });

      addFooter(slide);
    }

    // =========================
    // SLIDE 6
    // =========================
    {
      const slide = pptx.addSlide();

      addTitle(
        slide,
        "Recommended Next Steps",
        "Management execution priorities",
        6
      );

      const steps = [
        "Assign ownership for all High AP items",
        "Validate mitigation feasibility with process owners",
        "Track closure progress in the Excel execution register",
        "Review recurring operational themes quarterly",
      ];

      const colors = [
        "5B2D91",
        "F28C28",
        "2CA02C",
        "1DA1F2",
      ];

      steps.forEach((step, idx) => {
        const x = 0.7 + idx * 2.8;

        slide.addShape("rect", {
          x,
          y: 2.0,
          w: 2.3,
          h: 1.6,
          fill: { color: "FFFFFF" },
          line: { color: "DADADA" },
        });

        slide.addShape("roundRect", {
          x: x + 0.15,
          y: 2.15,
          w: 0.35,
          h: 0.35,
          rectRadius: 0.05,
          fill: { color: colors[idx] },
          line: { color: colors[idx] },
        });

        slide.addText(String(idx + 1), {
          x: x + 0.15,
          y: 2.2,
          w: 0.35,
          h: 0.1,
          fontSize: 10,
          bold: true,
          align: "center",
          color: "FFFFFF",
        });

        slide.addText(step, {
          x: x + 0.2,
          y: 2.7,
          w: 1.9,
          h: 0.6,
          fontSize: 11,
          color: "333333",
          align: "center",
        });
      });

      addFooter(slide);
    }

    // =========================
    // SLIDE 7
    // =========================
    {
      const slide = pptx.addSlide();

      addTitle(
        slide,
        "Thank you",
        "Management review ready",
        7
      );

      slide.addText(
        "Use this executive summary to guide management review discussion and mitigation prioritization.",
        {
          x: 1.0,
          y: 2.0,
          w: 9,
          h: 0.6,
          fontSize: 24,
          align: "center",
          bold: true,
          color: "5B2D91",
        }
      );

      slide.addText(
        `${processName} • ${total} failure modes reviewed`,
        {
          x: 1.0,
          y: 3.0,
          w: 9,
          h: 0.3,
          fontSize: 16,
          align: "center",
          color: "666666",
        }
      );

      addFooter(slide);
    }

    // =========================
    // EXPORT
    // =========================
    const buffer = await pptx.write({
      outputType: "nodebuffer",
    });

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="P-FMEA_Executive_${processName}.pptx"`
    );

    return res.status(200).send(buffer);
  } catch (err) {
    console.error("PPTX EXPORT ERROR:", err);

    return res.status(500).json({
      success: false,
      error: err?.message || "PPTX export failed",
    });
  }
}
