import { jsPDF } from "jspdf";

type Audit = Record<string, any>;
const safe = (value: unknown) => String(value ?? "").replace(/[\u0000-\u001f]/g, " ").trim();
const title = (value: unknown) => safe(value).replace(/_/g, " ").replace(/\b\w/g, (character) => character.toUpperCase());

function score(document: Audit) {
  const values = (document.template?.chapters ?? []).flatMap((chapter: Audit) => chapter.questions ?? []).map((question: Audit) => {
    const grade = document.responses?.[question.id]?.grade;
    return typeof grade === "number" ? grade : grade === "green" ? 10 : grade === "yellow" ? 6 : grade === "red" ? 0 : null;
  }).filter((item: number | null): item is number => item !== null);
  return values.length ? Math.round((values.reduce((sum: number, item: number) => sum + item, 0) / values.length) * 10) : null;
}

export function buildAuditPortfolioPdf({ audits, generatedAt, includeDrafts, reportType = "portfolio" }: { audits: Audit[]; generatedAt: string; includeDrafts: boolean; reportType?: "portfolio" | "supplier_scorecard" }) {
  const pdf = new jsPDF({ unit: "mm", format: "a4", compress: true });
  let y = 20;
  const page = () => { pdf.addPage(); y = 20; };
  const text = (value: unknown, size = 9, bold = false) => {
    const lines = pdf.splitTextToSize(safe(value) || "—", 174);
    pdf.setFont("helvetica", bold ? "bold" : "normal"); pdf.setFontSize(size);
    for (const line of lines) { if (y > 278) page(); pdf.text(line, 18, y); y += size * .48 + 1.8; }
  };
  const heading = (value: string) => { if (y > 264) page(); y += 3; pdf.setDrawColor(32, 66, 104); pdf.line(18, y, 192, y); y += 6; text(value, 14, true); };
  const supplierScorecard = reportType === "supplier_scorecard";
  pdf.setFillColor(21, 55, 91); pdf.rect(0, 0, 210, 13, "F"); pdf.setTextColor(255, 255, 255); pdf.setFont("helvetica", "bold"); pdf.setFontSize(10); pdf.text(supplierScorecard ? "QHUBIO · SUPPLIER PERFORMANCE" : "QHUBIO · AUDIT PORTFOLIO", 18, 8); pdf.setTextColor(20, 25, 30);
  text(supplierScorecard ? "Supplier Scorecard & Multi-site Report" : "Audit Portfolio Report", 19, true); text(`Generated ${generatedAt} · ${audits.length} selected audit${audits.length === 1 ? "" : "s"}`, 9);
  if (includeDrafts) { pdf.setTextColor(160, 35, 35); text("DRAFT RECORDS INCLUDED — NOT ISSUED", 10, true); pdf.setTextColor(20, 25, 30); }
  const issued = audits.filter((audit) => !!audit.issued_snapshot); const openFindings = audits.reduce((sum, audit) => sum + ((audit.document?.findings ?? []).filter((finding: Audit) => finding.status !== "closed").length), 0);
  heading(supplierScorecard ? "Supplier performance overview" : "Management overview"); text(`Issued audits: ${issued.length}`); text(`Open findings: ${openFindings}`); text(`Average supplier / audit score: ${(() => { const scores = audits.map((audit) => score(audit.issued_snapshot ?? audit.document)).filter((value): value is number => value !== null); return scores.length ? `${Math.round(scores.reduce((sum, value) => sum + value, 0) / scores.length)}%` : "Not calculated"; })()}`);
  heading("Included audits");
  for (const audit of audits) { const document = audit.issued_snapshot ?? audit.document; const open = (audit.document?.findings ?? document.findings ?? []).filter((finding: Audit) => finding.status !== "closed"); text(`${document.name} · ${document.organization}`, 10, true); text(`${document.template?.name ?? "Audit template"} · planned ${document.plannedOn ?? "—"} · ${audit.issued_snapshot ? "Issued snapshot" : "Draft"}`); text(`Result: ${score(document) == null ? "Not calculated" : `${score(document)}%`} · Release: ${title(document.releaseDecision)} · Open findings: ${open.length}`); if (document.conclusion) text(`Conclusion: ${document.conclusion}`, 8); y += 2; }
  heading("Report integrity manifest"); for (const audit of audits) text(`${audit.id} · revision ${audit.revision} · ${audit.issued_snapshot ? "issued snapshot" : "draft"}`, 8);
  const pages = pdf.getNumberOfPages(); for (let index = 1; index <= pages; index++) { pdf.setPage(index); if (includeDrafts) { pdf.saveGraphicsState(); pdf.setTextColor(190, 45, 45); pdf.setFontSize(25); pdf.setFont("helvetica", "bold"); pdf.text("DRAFT — NOT ISSUED", 36, 156, { angle: 35 }); pdf.restoreGraphicsState(); } pdf.setDrawColor(190); pdf.line(18, 286, 192, 286); pdf.setFontSize(8); pdf.setTextColor(80); pdf.text(`Qhubio Audit Portfolio · ${index} / ${pages}`, 18, 291); }
  return pdf.output("arraybuffer");
}


