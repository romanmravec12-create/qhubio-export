import pptxgen from "pptxgenjs";

const ALLOWED_ORIGINS = (process.env.ALLOWED_ORIGINS ?? "*")
  .split(",")
  .map((s) => s.trim())
  .filter(Boolean);

const BRAND = {
  violet: "6B3FD4",
  violetDark: "4E2A84",
  orange: "F28C28",
  green: "16A34A",
  red: "DC2626",
  gray: "64748B",
  light: "F8FAFC",
  border: "D8D5E4",
  text: "1F2937",
  muted: "6B7280",
};

function applyCors(req, res) {
  const origin = req.headers.origin ?? "";
  const allow =
    ALLOWED_ORIGINS.includes("*") || ALLOWED_ORIGINS.includes(origin)
      ? origin || "*"
      : ALLOWED_ORIGINS[0] ?? "*";

  res.setHeader("Access-Control-Allow-Origin", allow);
  res.setHeader("Vary", "Origin");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization, apikey");
}

function jsonError(res, status, error, stage = "input") {
  res.status(status).json({ success: false, error, stage });
}

function safeText(value) {
  return String(value ?? "").trim();
}

function normalizeText(value) {
  return safeText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

function toNumber(value, fallback = 0) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function formatDate(date = new Date()) {
  return new Intl.DateTimeFormat("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  }).format(date);
}

function sanitizeFilenamePart(value) {
  return safeText(value)
    .replace(/[^\w\s-]+/g, "")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 60) || "FMEA";
}

function isOpenAction(row) {
  const s = normalizeText(row.action_status);
  if (!s) return true;
  return !(
    s.includes("complete") ||
    s.includes("closed") ||
    s.includes("done")
  );
}

function actionPriorityWeight(value) {
  const v = normalizeText(value);
  if (v.includes("high")) return 3;
  if (v.includes("medium")) return 2;
  if (v.includes("low")) return 1;
  return 0;
}

function riskScore(row) {
  return (
    toNumber(row.severity, 0) *
    toNumber(row.occurrence, 0) *
    toNumber(row.detection, 0)
  );
}

const THEME_RULES = [
  {
    label: "Scheduling & Dispatching",
    color: BRAND.violet,
    keywords: ["schedule", "dispatch", "arrival", "route", "timing", "priority", "sequence", "delay"],
  },
  {
    label: "Inventory & Spare Parts",
    color: BRAND.orange,
    keywords: ["inventory", "spare part", "spare parts", "stock", "warehouse", "supplier", "availability", "packaging", "part number"],
  },
  {
    label: "Technician Competency",
    color: BRAND.violetDark,
    keywords: ["skill", "training", "competency", "experience", "qualification", "operator", "knowledge", "oversight"],
  },
  {
    label: "Documentation & Traceability",
    color: "2563EB",
    keywords: ["document", "record", "report", "traceability", "log", "signature", "checklist", "documentation"],
  },
  {
    label: "Communication & Handover",
    color: "0F766E",
    keywords: ["communication", "handover", "customer", "signoff", "feedback", "coordination", "stakeholder"],
  },
  {
    label: "Equipment, Tooling & Maintenance",
    color: "B45309",
    keywords: ["equipment", "tool", "maintenance", "calibration", "diagnostic", "wear", "lubrication", "fixture", "torque", "machine"],
  },
  {
    label: "Inspection & Verification",
    color: "7C3AED",
    keywords: ["inspect", "verification", "test", "confirm", "check", "validation", "control", "audit"],
  },
  {
    label: "Specification & Quality",
    color: BRAND.red,
    keywords: ["specification", "tolerance", "out of tolerance", "dimension", "quality", "nonconforming", "deviation", "defect"],
  },
  {
    label: "Identification & Traceability",
    color: "EC4899",
    keywords: ["identify", "identification", "barcode", "label", "mix-up", "wrong part", "trace", "traceable"],
  },
  {
    label: "Software & Data",
    color: "0EA5E9",
    keywords: ["software", "parameter", "database", "system", "backup", "digital", "sensor", "data", "signal"],
  },
  {
    label: "Safety & Interlock",
    color: BRAND.green,
    keywords: ["safety", "interlock", "lockout", "tagout", "energy source", "hazard", "unsafe", "protection"],
  },
];

function deriveTheme(row) {
  const text = [
    row.process_step,
    row.function,
    row.failure_mode,
    row.effect,
    row.cause,
    row.current_prevention_controls,
    row.current_detection_controls,
    row.recommended_action,
    row.responsibility,
    row.assigned_to,
  ]
    .map(normalizeText)
    .join(" ");

  for (const theme of THEME_RULES) {
    if (theme.keywords.some((kw) => text.includes(normalizeText(kw)))) {
      return theme.label;
    }
  }
  return "Other Operational Areas";
}

function themeColor(themeLabel) {
  return THEME_RULES.find((t) => t.label === themeLabel)?.color ?? BRAND.gray;
}

function buildAnalytics(rows) {
  const total = rows.length;
  const counts = {
    high: 0,
    medium: 0,
    low: 0,
    openActions: 0,
    completedActions: 0,
  };

  const themeMap = new Map();
  const items = [];

  for (const row of rows) {
    const ap = normalizeText(row.action_priority);
    if (ap.includes("high")) counts.high += 1;
    else if (ap.includes("medium")) counts.medium += 1;
    else if (ap.includes("low")) counts.low += 1;

    if (isOpenAction(row)) counts.openActions += 1;
    else counts.completedActions += 1;

    const theme = deriveTheme(row);
    const existing = themeMap.get(theme) ?? {
      label: theme,
      color: themeColor(theme),
      count: 0,
      topScore: -1,
      sample: null,
    };

    existing.count += 1;
    const score = riskScore(row);
    if (score > existing.topScore) {
      existing.topScore = score;
      existing.sample = row;
    }
    themeMap.set(theme, existing);

    items.push({
      ...row,
      theme,
      score,
      apWeight: actionPriorityWeight(row.action_priority),
    });
  }

  const themes = Array.from(themeMap.values()).sort((a, b) => {
    if (b.count !== a.count) return b.count - a.count;
    return a.label.localeCompare(b.label);
  });

  const criticalItems = items
    .sort((a, b) => {
      if (b.apWeight !== a.apWeight) return b.apWeight - a.apWeight;
      if (b.score !== a.score) return b.score - a.score;
      return toNumber(b.severity, 0) - toNumber(a.severity, 0);
    })
    .slice(0, 6);

  return { total, counts, themes, criticalItems };
}

function actionLine(row) {
  return safeText(row.recommended_action) || "Recommended action not specified";
}

function sampleLine(row) {
  if (!row) return "No representative item available";
  const step = safeText(row.process_step) || "Step not specified";
  const fm = safeText(row.failure_mode) || "Failure mode not specified";
  return `${step} — ${fm}`;
}

function openActionsText(count, total) {
  const pct = total ? Math.round((count / total) * 100) : 0;
  return `${count} open actions (${pct}%)`;
}

function buildFilename(processName, userName) {
  const p = sanitizeFilenamePart(processName);
  const u = sanitizeFilenamePart(userName);
  const d = new Date().toISOString().slice(0, 10);
  return `P-FMEA-Executive-Summary-${p}-${u}-${d}.pptx`;
}

function addPageHeader(slide, ST, title, subtitle, pageNo) {
  slide.addShape(ST.rect, {
    x: 0,
    y: 0,
    w: 13.333,
    h: 0.72,
    fill: { color: BRAND.violetDark },
    line: { color: BRAND.violetDark },
  });

  slide.addText(title, {
    x: 0.5,
    y: 0.14,
    w: 11.2,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 22,
    bold: true,
    color: "FFFFFF",
    margin: 0,
  });

  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.5,
      y: 0.44,
      w: 11.7,
      h: 0.18,
      fontFace: "Arial",
      fontSize: 9.5,
      color: "E9D5FF",
      margin: 0,
    });
  }

  slide.addText(String(pageNo), {
    x: 12.55,
    y: 0.16,
    w: 0.5,
    h: 0.24,
    fontFace: "Arial",
    fontSize: 9,
    color: "FFFFFF",
    align: "right",
    margin: 0,
  });
}

function addCard(slide, ST, x, y, w, h, title, value, detail, accentColor) {
  slide.addShape(ST.rect, {
    x,
    y,
    w,
    h,
    fill: { color: "FFFFFF" },
    line: { color: BRAND.border, pt: 1 },
  });

  slide.addShape(ST.rect, {
    x,
    y,
    w: 0.16,
    h,
    fill: { color: accentColor },
    line: { color: accentColor },
  });

  slide.addText(title, {
    x: x + 0.26,
    y: y + 0.1,
    w: w - 0.36,
    h: 0.18,
    fontFace: "Arial",
    fontSize: 10,
    color: BRAND.muted,
    bold: true,
    margin: 0,
  });

  slide.addText(value, {
    x: x + 0.26,
    y: y + 0.28,
    w: w - 0.36,
    h: 0.28,
    fontFace: "Arial",
    fontSize: 18,
    color: BRAND.text,
    bold: true,
    margin: 0,
  });

  if (detail) {
    slide.addText(detail, {
      x: x + 0.26,
      y: y + 0.56,
      w: w - 0.36,
      h: h - 0.62,
      fontFace: "Arial",
      fontSize: 8.5,
      color: BRAND.muted,
      margin: 0,
    });
  }
}

function buildDeck(rows, processName, userName) {
  const { total, counts, themes, criticalItems } = buildAnalytics(rows);

  const pptx = new pptxgen();
  const ST = pptx.ShapeType;

  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Qhubio";
  pptx.company = "Qhubio";
  pptx.subject = "P-FMEA Executive Summary";
  pptx.title = `P-FMEA Executive Summary — ${processName || "Process"}`;
  pptx.lang = "en-US";

  // Slide 1 - Title
  {
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    slide.addShape(ST.rect, {
      x: 0,
      y: 0,
      w: 13.333,
      h: 7.5,
      fill: { color: "FFFFFF" },
      line: { color: "FFFFFF" },
    });
    slide.addShape(ST.rect, {
      x: 0,
      y: 0,
      w: 13.333,
      h: 1.0,
      fill: { color: BRAND.violetDark },
      line: { color: BRAND.violetDark },
    });
    slide.addText("P-FMEA Executive Summary", {
      x: 0.7,
      y: 1.05,
      w: 7.2,
      h: 0.45,
      fontFace: "Arial",
      fontSize: 28,
      bold: true,
      color: BRAND.violetDark,
      margin: 0,
    });
    slide.addText(processName || "Unnamed process", {
      x: 0.7,
      y: 1.55,
      w: 7.2,
      h: 0.35,
      fontFace: "Arial",
      fontSize: 18,
      color: BRAND.text,
      bold: true,
      margin: 0,
    });
    slide.addText(
      `Prepared for ${userName || "the customer"} • Generated ${formatDate()} • ${total} failure modes reviewed`,
      {
        x: 0.7,
        y: 1.95,
        w: 7.4,
        h: 0.3,
        fontFace: "Arial",
        fontSize: 10.5,
        color: BRAND.muted,
        margin: 0,
      }
    );

    addCard(slide, ST, 0.7, 2.55, 3.0, 1.0, "High AP", String(counts.high), "Risks requiring direct management attention", BRAND.red);
    addCard(slide, ST, 4.0, 2.55, 3.0, 1.0, "Medium AP", String(counts.medium), "Important items for review and ownership", BRAND.orange);
    addCard(slide, ST, 7.3, 2.55, 3.0, 1.0, "Low AP", String(counts.low), "Lower priority items still tracked in Excel", BRAND.green);
    addCard(slide, ST, 10.6, 2.55, 2.0, 1.0, "Open actions", String(counts.openActions), "Items not yet closed", BRAND.violet);

    slide.addText("Executive takeaway", {
      x: 0.7,
      y: 4.0,
      w: 3.5,
      h: 0.24,
      fontFace: "Arial",
      fontSize: 11,
      bold: true,
      color: BRAND.violetDark,
      margin: 0,
    });

    const takeaway = [
      `The analysis identified ${counts.high} high-priority risks out of ${total} failure modes.`,
      `The strongest concentration sits in ${themes.slice(0, 3).map((t) => `${t.label}`).join(", ") || "the dominant operational themes"}.`,
      `This deck summarizes the management view only; the full actionable register remains in the Excel export.`,
    ].join("\n");

    slide.addText(takeaway, {
      x: 0.7,
      y: 4.28,
      w: 7.6,
      h: 1.35,
      fontFace: "Arial",
      fontSize: 14,
      color: BRAND.text,
      valign: "top",
      margin: 0,
    });

    slide.addShape(ST.rect, {
      x: 9.0,
      y: 4.0,
      w: 3.5,
      h: 1.75,
      fill: { color: BRAND.light },
      line: { color: BRAND.border, pt: 1 },
    });
    slide.addText("What management gets", {
      x: 9.25,
      y: 4.18,
      w: 3.0,
      h: 0.22,
      fontFace: "Arial",
      fontSize: 11,
      bold: true,
      color: BRAND.violetDark,
      margin: 0,
    });
    slide.addText(
      [
        "• Risk distribution at a glance",
        "• Top operational themes",
        "• Critical items needing ownership",
        "• Recommended next steps",
      ].join("\n"),
      {
        x: 9.25,
        y: 4.48,
        w: 3.0,
        h: 1.05,
        fontFace: "Arial",
        fontSize: 10,
        color: BRAND.text,
        margin: 0,
      }
    );

    slide.addText("PPTX generated deterministically from FMEA data", {
      x: 0.7,
      y: 6.95,
      w: 5.5,
      h: 0.2,
      fontFace: "Arial",
      fontSize: 8.5,
      color: BRAND.muted,
      italic: true,
      margin: 0,
    });
  }

  // Slide 2 - Executive Summary
  {
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    addPageHeader(slide, ST, "Executive Summary", "Management-ready view of risk exposure and operational focus areas", 2);

    slide.addShape(ST.rect, {
      x: 0.55,
      y: 1.0,
      w: 7.3,
      h: 5.85,
      fill: { color: "FFFFFF" },
      line: { color: BRAND.border, pt: 1 },
    });
    slide.addText("Summary narrative", {
      x: 0.75,
      y: 1.2,
      w: 2.6,
      h: 0.2,
      fontFace: "Arial",
      fontSize: 12,
      bold: true,
      color: BRAND.violetDark,
      margin: 0,
    });

    const narrative = [
      `The process FMEA reviewed ${total} failure modes.`,
      `High Action Priority items: ${counts.high}.`,
      `Open actions: ${counts.openActions}.`,
      `Main exposure themes: ${themes.slice(0, 4).map((t) => t.label).join(", ") || "see operational themes slide"}.`,
      `This summary is intended for management discussion, prioritization, and ownership assignment.`,
    ].join("\n");

    slide.addText(narrative, {
      x: 0.75,
      y: 1.5,
      w: 6.7,
      h: 1.7,
      fontFace: "Arial",
      fontSize: 14,
      color: BRAND.text,
      margin: 0,
    });

    slide.addShape(ST.rect, {
      x: 0.75,
      y: 3.45,
      w: 6.65,
      h: 1.0,
      fill: { color: BRAND.light },
      line: { color: BRAND.border, pt: 1 },
    });
    slide.addText("Management interpretation", {
      x: 0.95,
      y: 3.62,
      w: 2.6,
      h: 0.2,
      fontFace: "Arial",
      fontSize: 10.5,
      bold: true,
      color: BRAND.violetDark,
      margin: 0,
    });
    slide.addText(
      "The deck highlights where the process is most exposed, what is driving the risk concentration, and what should be reviewed first in the management meeting.",
      {
        x: 0.95,
        y: 3.85,
        w: 6.1,
        h: 0.42,
        fontFace: "Arial",
        fontSize: 10,
        color: BRAND.text,
        margin: 0,
      }
    );

    addCard(slide, ST, 8.15, 1.1, 4.6, 1.1, "Total failure modes", String(total), "All rows included in the analysis", BRAND.violet);
    addCard(slide, ST, 8.15, 2.4, 4.6, 1.1, "Risk concentration", `${counts.high} high / ${counts.medium} medium / ${counts.low} low`, "Action Priority split", BRAND.orange);
    addCard(slide, ST, 8.15, 3.7, 4.6, 1.1, "Open actions", String(counts.openActions), "Items not yet closed", BRAND.green);
    addCard(slide, ST, 8.15, 5.0, 4.6, 1.1, "Top themes", `${Math.min(4, themes.length)} highlighted`, themes.slice(0, 4).map((t) => t.label).join(", "), BRAND.violetDark);
  }

  // Slide 3 - Risk Distribution
  {
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    addPageHeader(slide, ST, "Risk Distribution", "Quick view of the Action Priority breakdown", 3);

    slide.addText("Action Priority split", {
      x: 0.6,
      y: 1.0,
      w: 3,
      h: 0.22,
      fontFace: "Arial",
      fontSize: 12,
      bold: true,
      color: BRAND.violetDark,
      margin: 0,
    });

    const totalCount = total || 1;
    const barX = 0.8;
    const barY = 1.45;
    const barW = 11.8;
    const barH = 0.55;

    slide.addShape(ST.rect, {
      x: barX,
      y: barY,
      w: barW,
      h: barH,
      fill: { color: "F1F5F9" },
      line: { color: BRAND.border, pt: 1 },
    });

    const segments = [
      { label: "High", count: counts.high, color: BRAND.red },
      { label: "Medium", count: counts.medium, color: BRAND.orange },
      { label: "Low", count: counts.low, color: BRAND.green },
    ];

    let offset = barX;
    for (const seg of segments) {
      const w = (barW * seg.count) / totalCount;
      if (w > 0) {
        slide.addShape(ST.rect, {
          x: offset,
          y: barY,
          w,
          h: barH,
          fill: { color: seg.color },
          line: { color: seg.color },
        });
        if (w > 1.1) {
          slide.addText(`${seg.label} ${seg.count}`, {
            x: offset,
            y: barY + 0.12,
            w,
            h: 0.2,
            fontFace: "Arial",
            fontSize: 10,
            bold: true,
            color: "FFFFFF",
            align: "center",
            margin: 0,
          });
        }
      }
      offset += w;
    }

    addCard(slide, ST, 0.75, 2.35, 3.7, 1.05, "High AP", `${counts.high} (${barPercent(counts.high, total)}%)`, "Most urgent items", BRAND.red);
    addCard(slide, ST, 4.75, 2.35, 3.7, 1.05, "Medium AP", `${counts.medium} (${barPercent(counts.medium, total)}%)`, "Prioritize with owners", BRAND.orange);
    addCard(slide, ST, 8.75, 2.35, 3.7, 1.05, "Low AP", `${counts.low} (${barPercent(counts.low, total)}%)`, "Track in background", BRAND.green);

    slide.addShape(ST.rect, {
      x: 0.75,
      y: 3.7,
      w: 12.0,
      h: 1.25,
      fill: { color: BRAND.light },
      line: { color: BRAND.border, pt: 1 },
    });
    slide.addText("Action focus", {
      x: 0.95,
      y: 3.86,
      w: 2.2,
      h: 0.2,
      fontFace: "Arial",
      fontSize: 11,
      bold: true,
      color: BRAND.violetDark,
      margin: 0,
    });
    slide.addText(`Open actions: ${openActionsText(counts.openActions, total)} • Completed/closed: ${counts.completedActions}`, {
      x: 0.95,
      y: 4.08,
      w: 11.4,
      h: 0.2,
      fontFace: "Arial",
      fontSize: 11,
      color: BRAND.text,
      margin: 0,
    });
    slide.addText(
      "Management should focus first on the High AP items, then review the dominant themes and confirm ownership for the open actions.",
      {
        x: 0.95,
        y: 4.3,
        w: 11.2,
        h: 0.34,
        fontFace: "Arial",
        fontSize: 10,
        color: BRAND.muted,
        margin: 0,
      }
    );
  }

  // Slide 4 - Top Themes
  {
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    addPageHeader(slide, ST, "Top Risk Themes", "Consolidated risk clusters based on the FMEA row content", 4);

    const topThemes = themes.slice(0, 4);
    const cardPositions = [
      [0.7, 1.15],
      [6.8, 1.15],
      [0.7, 3.95],
      [6.8, 3.95],
    ];

    for (let i = 0; i < 4; i++) {
      const t = topThemes[i];
      const [x, y] = cardPositions[i];
      slide.addShape(ST.rect, {
        x,
        y,
        w: 5.8,
        h: 2.2,
        fill: { color: "FFFFFF" },
        line: { color: BRAND.border, pt: 1 },
      });

      if (t) {
        slide.addShape(ST.rect, {
          x,
          y,
          w: 0.18,
          h: 2.2,
          fill: { color: t.color },
          line: { color: t.color },
        });
        slide.addText(t.label, {
          x: x + 0.28,
          y: y + 0.16,
          w: 5.2,
          h: 0.22,
          fontFace: "Arial",
          fontSize: 13,
          bold: true,
          color: BRAND.violetDark,
          margin: 0,
        });
        slide.addText(`Count: ${t.count}`, {
          x: x + 0.28,
          y: y + 0.46,
          w: 2.2,
          h: 0.2,
          fontFace: "Arial",
          fontSize: 11,
          bold: true,
          color: t.color,
          margin: 0,
        });
        slide.addText(`Representative item: ${sampleLine(t.sample)}`, {
          x: x + 0.28,
          y: y + 0.76,
          w: 5.2,
          h: 0.62,
          fontFace: "Arial",
          fontSize: 10,
          color: BRAND.text,
          margin: 0,
        });
        slide.addText(`Why it matters: ${actionLine(t.sample)}`, {
          x: x + 0.28,
          y: y + 1.42,
          w: 5.2,
          h: 0.55,
          fontFace: "Arial",
          fontSize: 9.5,
          italic: true,
          color: BRAND.muted,
          margin: 0,
        });
      } else {
        slide.addText("No additional theme data", {
          x: x + 0.28,
          y: y + 0.45,
          w: 5,
          h: 0.25,
          fontFace: "Arial",
          fontSize: 12,
          color: BRAND.muted,
          margin: 0,
        });
      }
    }
  }

  // Slide 5 - Top Critical Items
  {
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    addPageHeader(slide, ST, "Top Critical Items", "Highest priority findings to review with management", 5);

    slide.addText("Showing the highest priority items by Action Priority and calculated risk score", {
      x: 0.65,
      y: 1.0,
      w: 12,
      h: 0.2,
      fontFace: "Arial",
      fontSize: 10.5,
      color: BRAND.muted,
      margin: 0,
    });

    const headers = [
      { text: "Process step", w: 2.1 },
      { text: "Failure mode", w: 3.8 },
      { text: "AP", w: 0.7 },
      { text: "Recommended action", w: 6.0 },
    ];
    const rowStartY = 1.45;
    const rowH = 0.78;

    slide.addShape(ST.rect, {
      x: 0.65,
      y: rowStartY - 0.32,
      w: 12.0,
      h: 0.34,
      fill: { color: BRAND.violetDark },
      line: { color: BRAND.violetDark },
    });

    let headerX = 0.75;
    for (const h of headers) {
      slide.addText(h.text, {
        x: headerX,
        y: rowStartY - 0.28,
        w: h.w - 0.1,
        h: 0.18,
        fontFace: "Arial",
        fontSize: 9,
        bold: true,
        color: "FFFFFF",
        margin: 0,
      });
      headerX += h.w;
    }

    const items = criticalItems.slice(0, 6);
    items.forEach((row, idx) => {
      const y = rowStartY + idx * rowH;
      slide.addShape(ST.rect, {
        x: 0.65,
        y,
        w: 12.0,
        h: rowH - 0.06,
        fill: { color: idx % 2 === 0 ? "FFFFFF" : "FAFAFF" },
        line: { color: BRAND.border, pt: 1 },
      });

      slide.addText(safeText(row.process_step) || "—", {
        x: 0.75,
        y: y + 0.08,
        w: 2.0,
        h: 0.42,
        fontFace: "Arial",
        fontSize: 10,
        bold: true,
        color: BRAND.text,
        margin: 0,
      });
      slide.addText(safeText(row.failure_mode) || "—", {
        x: 2.95,
        y: y + 0.08,
        w: 3.7,
        h: 0.42,
        fontFace: "Arial",
        fontSize: 10,
        color: BRAND.text,
        margin: 0,
      });

      const ap = safeText(row.action_priority) || "—";
      const apColor = normalizeText(ap).includes("high")
        ? BRAND.red
        : normalizeText(ap).includes("medium")
        ? BRAND.orange
        : BRAND.green;

      slide.addShape(ST.rect, {
        x: 6.9,
        y: y + 0.13,
        w: 0.58,
        h: 0.32,
        fill: { color: apColor },
        line: { color: apColor },
      });
      slide.addText(ap, {
        x: 6.9,
        y: y + 0.17,
        w: 0.58,
        h: 0.12,
        fontFace: "Arial",
        fontSize: 8,
        bold: true,
        color: "FFFFFF",
        align: "center",
        margin: 0,
      });

      slide.addText(actionLine(row), {
        x: 7.7,
        y: y + 0.08,
        w: 4.75,
        h: 0.42,
        fontFace: "Arial",
        fontSize: 9.5,
        color: BRAND.text,
        margin: 0,
      });
    });
  }

  // Slide 6 - Next Steps
  {
    const slide = pptx.addSlide();
    slide.background = { color: "FFFFFF" };
    addPageHeader(slide, ST, "Recommended Next Steps", "Simple management actions to turn the analysis into execution", 6);

    const steps = [
      {
        n: "1",
        title: "Assign ownership",
        text: "Confirm a responsible owner for each top critical item and move the action out of backlog.",
        color: BRAND.violet,
      },
      {
        n: "2",
        title: "Review high-priority risks",
        text: "Validate the High AP findings with the process owner and agree whether additional containment is needed.",
        color: BRAND.orange,
      },
      {
        n: "3",
        title: "Track closure",
        text: "Use the full Excel export for detailed action tracking, while the PowerPoint stays focused on management.",
        color: BRAND.green,
      },
    ];

    steps.forEach((s, idx) => {
      const x = 0.8 + idx * 4.1;
      slide.addShape(ST.rect, {
        x,
        y: 1.55,
        w: 3.65,
        h: 3.5,
        fill: { color: "FFFFFF" },
        line: { color: BRAND.border, pt: 1 },
      });
      slide.addShape(ST.rect, {
        x: x + 0.18,
        y: 1.75,
        w: 0.62,
        h: 0.62,
        fill: { color: s.color },
        line: { color: s.color },
      });
      slide.addText(s.n, {
        x: x + 0.18,
        y: 1.88,
        w: 0.62,
        h: 0.14,
        fontFace: "Arial",
        fontSize: 16,
        bold: true,
        color: "FFFFFF",
        align: "center",
        margin: 0,
      });
      slide.addText(s.title, {
        x: x + 0.92,
        y: 1.78,
        w: 2.3,
        h: 0.22,
        fontFace: "Arial",
        fontSize: 13,
        bold: true,
        color: BRAND.violetDark,
        margin: 0,
      });
      slide.addText(s.text, {
        x: x + 0.18,
        y: 2.52,
        w: 3.2,
        h: 1.1,
        fontFace: "Arial",
        fontSize: 11,
        color: BRAND.text,
        margin: 0,
      });
    });

    slide.addShape(ST.rect, {
      x: 0.8,
      y: 5.4,
      w: 11.75,
      h: 0.8,
      fill: { color: BRAND.light },
      line: { color: BRAND.border, pt: 1 },
    });
    slide.addText("This deck is intentionally concise. The Excel export remains the source of truth for full row-level execution details.", {
      x: 1.0,
      y: 5.64,
      w: 11.3,
      h: 0.2,
      fontFace: "Arial",
      fontSize: 10.5,
      color: BRAND.text,
      align: "center",
      margin: 0,
    });
  }

  // Slide 7 - Closing
  {
    const slide = pptx.addSlide();
    slide.background = { color: BRAND.violetDark };
    slide.addText("Thank you", {
      x: 0.9,
      y: 1.45,
      w: 5.0,
      h: 0.5,
      fontFace: "Arial",
      fontSize: 30,
      bold: true,
      color: "FFFFFF",
      margin: 0,
    });
    slide.addText("Use this summary to guide the management review discussion.", {
      x: 0.9,
      y: 2.0,
      w: 6.5,
      h: 0.3,
      fontFace: "Arial",
      fontSize: 14,
      color: "E9D5FF",
      margin: 0,
    });
    slide.addShape(ST.rect, {
      x: 0.9,
      y: 2.7,
      w: 11.4,
      h: 1.6,
      fill: { color: "5B2FB3" },
      line: { color: "5B2FB3" },
    });
    slide.addText(
      [
        `• Process reviewed: ${processName || "FMEA process"}`,
        `• Failure modes included: ${total}`,
        `• High AP: ${counts.high} • Medium AP: ${counts.medium} • Low AP: ${counts.low}`,
        `• Generated on: ${formatDate()}`,
      ].join("\n"),
      {
        x: 1.2,
        y: 3.0,
        w: 10.8,
        h: 0.95,
        fontFace: "Arial",
        fontSize: 13,
        color: "FFFFFF",
        margin: 0,
      }
    );
    slide.addText("Full row-level details remain available in Excel export.", {
      x: 0.9,
      y: 5.15,
      w: 6.5,
      h: 0.25,
      fontFace: "Arial",
      fontSize: 11,
      color: "E9D5FF",
      italic: true,
      margin: 0,
    });
  }

  return pptx;
}

export default async function handler(req, res) {
  applyCors(req, res);

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const t0 = Date.now();
  const reqId = Math.random().toString(36).slice(2, 10);
  const log = (msg) => console.log(`[export-pptx][${reqId}] ${msg}`);

  try {
    const body = req.body ?? {};
    const rows = Array.isArray(body.rows) ? body.rows : [];
    const processName = safeText(body.processName) || "FMEA Process";
    const userName = safeText(body.userName) || "User";

    if (rows.length === 0) {
      return jsonError(res, 400, "No rows provided", "input");
    }

    const payloadSize = Buffer.byteLength(JSON.stringify(body));
    log(`request received`);
    log(`payload: ${rows.length} rows, ${payloadSize} bytes`);

    const pptx = buildDeck(rows, processName, userName);
    const buffer = await pptx.write({ outputType: "nodebuffer" });

    const filename = buildFilename(processName, userName);
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.setHeader("Content-Length", String(buffer.length));

    log(`pptx built (${buffer.length} bytes)`);
    res.status(200).send(buffer);
    log(`response sent in ${Date.now() - t0}ms total`);
  } catch (err) {
    const msg = err?.message ?? String(err);
    const stack = err?.stack ?? "";
    console.error(`[export-pptx][${reqId}] FAILED after ${Date.now() - t0}ms: ${msg}\n${stack}`);
    return jsonError(res, 500, msg || "Export failed", "write-buffer");
  }
}
