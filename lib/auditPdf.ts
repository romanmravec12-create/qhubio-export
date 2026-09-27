import { jsPDF } from "jspdf";

export type AuditPdfMode = "executive" | "detailed";
type AnyRecord = Record<string, any>;

const safe = (value: unknown) => String(value ?? "").replace(/[\u0000-\u001f]/g, " ").trim();
const titleCase = (value: unknown) => safe(value).replace(/_/g, " ").replace(/\b\w/g, (c) => c.toUpperCase());
const gradeLabel = (grade: unknown) => grade === "na" ? "Not applicable" : grade === "not_assessed" || grade == null ? "Not assessed" : typeof grade === "number" ? `${grade}/10` : titleCase(grade);
const score = (grade: unknown) => typeof grade === "number" ? grade : grade === "green" ? 10 : grade === "yellow" ? 6 : grade === "red" ? 0 : null;

export function buildAuditPdf({ audit, evidence, links, history, approvals = [], branding, mode }: { audit: AnyRecord; evidence: AnyRecord[]; links: AnyRecord[]; history: AnyRecord[]; approvals?: AnyRecord[]; branding?: { bytes: Uint8Array; format: "PNG" | "JPEG" } | null; mode: AuditPdfMode }) {
  const pdf = new jsPDF({ unit: "mm", format: "a4", compress: true });
  const document = audit.issued_snapshot ?? audit.document;
  const draft = !audit.issued_snapshot;
  const result = evaluate(document);
  let y = 20;
  const page = () => { pdf.addPage(); y = 20; };
  const write = (value: unknown, size = 9, bold = false, indent = 0) => {
    const text = safe(value) || "—";
    pdf.setFont("helvetica", bold ? "bold" : "normal"); pdf.setFontSize(size);
    for (const line of pdf.splitTextToSize(text, 174 - indent)) { if (y > 276) page(); pdf.text(line, 18 + indent, y); y += size * 0.48 + 1.8; }
  };
  const heading = (value: string, level = 14) => { if (y > 260) page(); y += 3; pdf.setDrawColor(32, 66, 104); pdf.setLineWidth(.45); pdf.line(18, y, 192, y); y += 6; write(value, level, true); };
  const key = (label: string, value: unknown) => write(`${label}: ${safe(value) || "—"}`, 9, false);
  const header = () => { pdf.setFillColor(21, 55, 91); pdf.rect(0, 0, 210, 13, "F"); pdf.setTextColor(255, 255, 255); pdf.setFont("helvetica", "bold"); pdf.setFontSize(10); pdf.text("QHUBIO · AUDIT", 18, 8); pdf.setTextColor(20, 25, 30); };
  header();
  if (branding) { try { pdf.addImage(branding.bytes, branding.format, 146, 18, 44, 18, undefined, "FAST"); } catch { /* branding is optional, never fail an audit record because of it */ } }
  write(mode === "executive" ? "Audit Executive Summary" : "Detailed Audit Report", 19, true);
  if (draft) { pdf.setTextColor(160, 35, 35); write("DRAFT — NOT ISSUED", 11, true); pdf.setTextColor(20, 25, 30); }
  write(document.name, 15, true);
  key("Audit reference", audit.id); key("Template", `${document.template?.name ?? "Audit template"} · v${document.template?.version ?? "—"}`); key("Audited organisation", document.organization); key("Site", document.site); key("Audit date", document.plannedOn); key("Lead auditor", document.auditor); key("Scope", document.scope); key("Sample / project", document.sample);
  heading("Assessment result");
  key("Completion", `${result.coverage ?? 0}%`); key("Result", result.result); key("Score", result.score == null ? "Not calculated" : `${Number(result.score).toFixed(1)}% · band ${result.band ?? "—"}`); key("Conclusion", document.conclusion); if (document.template?.family === "supplier") key("Supplier / release decision", titleCase(document.releaseDecision));
  if (mode === "executive") {
    heading("Chapter performance"); drawRadar(pdf, document, 105, y + 35, 30); y += 70;
    heading("Key findings");
    const open = (document.findings ?? []).filter((f: AnyRecord) => f.status !== "closed");
    if (!open.length) write("No open findings recorded.");
    for (const finding of open.slice(0, 8)) { write(`${titleCase(finding.severity)} ${titleCase(finding.type)} · ${finding.title}`, 10, true); key("Owner / due", `${finding.owner || "—"} / ${finding.due || "—"}`); if (finding.severity === "major") key("Major finding deadlines", `Containment ${finding.containmentDue || "—"}; action plan ${finding.planDue || "—"}`); }
    heading("Management attention"); key("Open findings", open.length); key("Overdue findings", open.filter((f: AnyRecord) => f.due && f.due < new Date().toISOString().slice(0, 10)).length); key("CAPA links", links.length); key("Issued status", audit.issued_snapshot ? "Issued assessment snapshot" : "Working audit record");
  } else {
    heading("Scoring policy"); write("Scores and responses are reproduced from the active audit template. Not applicable responses are excluded from the calculated score. The audit conclusion remains an auditor judgement, not a certificate or automatic release authorization.");
    for (const chapter of document.template?.chapters ?? []) {
      heading(chapter.name);
      for (const question of chapter.questions ?? []) {
        const response = document.responses?.[question.id] ?? {}; write(`${question.id} · ${question.text}`, 10, true); key("Reference", question.reference); key("Assessment", gradeLabel(response.grade)); if (response.note) key("Observation / evidence examined", response.note); if (response.justification) key("Non-applicability rationale", response.justification); if (response.evidenceRefs) key("Document / sample references", response.evidenceRefs);
        const docs = evidence.filter((item) => item.question_id === question.id); if (docs.length) write(`Attached documents: ${docs.map((item) => `${item.name}${item.source_media_type ? ` (${item.source_media_type})` : ""}`).join("; ")}`, 8, false);
      }
    }
    heading(audit.issued_snapshot ? "Findings at issue" : "Draft findings");
    for (const finding of document.findings ?? []) { write(`${titleCase(finding.severity)} ${titleCase(finding.type)} · ${finding.title}`, 10, true); key("Evidence", finding.evidence); key("Owner / due", `${finding.owner || "—"} / ${finding.due || "—"}`); if (finding.severity === "major") key("Major-finding deadlines", `Containment ${finding.containmentDue || "—"}; action plan ${finding.planDue || "—"}; implementation ${finding.due || "—"}`); key("Action status", titleCase(finding.status)); key("Corrective action", finding.action); key("Implementation / effectiveness", `${finding.implementation || "—"} / ${finding.effectiveness || "—"}`); }
    if (audit.issued_snapshot) { heading("Current follow-up after issue"); for (const finding of audit.document?.findings ?? []) { write(`${finding.title} · ${titleCase(finding.status)}`, 10, true); key("Current implementation", finding.implementation); key("Verification", `${finding.effectiveness || "Pending"} · ${finding.verifiedBy || "—"} ${finding.verifiedOn || ""}`); } }
    heading("CAPA and audit history"); if (links.length) for (const link of links) write(`${link.capa_code || "CAPA"} · ${link.capa_title || "Linked corrective action"} · ${titleCase(link.capa_status)}`); else write("No CAPA links recorded."); if (history.length) { y += 2; for (const event of history.slice(0, 30)) write(`${event.created_at || ""} · ${titleCase(event.event)} · revision ${event.revision ?? "—"}`, 8); }
    heading("Controlled electronic sign-off"); if (approvals.length) for (const approval of approvals) write(`${approval.created_at || ""} · ${titleCase(approval.event)} · ${approval.meaning || "Controlled sign-off event"} · document SHA-256 ${String(approval.document_hash || "—").slice(0, 16)}…`, 8); else write("No controlled electronic sign-off has been recorded for this audit.");
  }
  const pages = pdf.getNumberOfPages(); for (let i = 1; i <= pages; i++) { pdf.setPage(i); if (draft) { pdf.setTextColor(232, 232, 232); pdf.setFont("helvetica", "bold"); pdf.setFontSize(28); pdf.text("DRAFT — NOT ISSUED", 105, 155, { align: "center", angle: 35 }); } pdf.setDrawColor(190); pdf.line(18, 286, 192, 286); pdf.setFontSize(8); pdf.setTextColor(80); pdf.text(`Qhubio audit record · ${audit.id} · ${i} / ${pages}`, 18, 291); }
  return pdf.output("arraybuffer");
}

function evaluate(document: AnyRecord) {
  const values = (document.template?.chapters ?? []).flatMap((chapter: AnyRecord) => chapter.questions ?? []).map((q: AnyRecord) => score(document.responses?.[q.id]?.grade)).filter((value: number | null): value is number => value != null);
  const assessed = values.length, total = (document.template?.chapters ?? []).reduce((sum: number, c: AnyRecord) => sum + (c.questions?.length ?? 0), 0);
  const percent = total ? Math.round((assessed / total) * 100) : 0; const average = assessed ? values.reduce((sum: number, value: number) => sum + value, 0) / assessed : null; const result = average == null ? "Not evaluated" : average >= 9 ? "A" : average >= 7 ? "B" : "C";
  return { coverage: percent, score: average == null ? null : average * 10, band: result, result };
}

function drawRadar(pdf: jsPDF, document: AnyRecord, cx: number, cy: number, radius: number) {
  const chapters = document.template?.chapters ?? []; const count = Math.max(chapters.length, 3); const value = (chapter: AnyRecord) => { const scores = (chapter.questions ?? []).map((q: AnyRecord) => score(document.responses?.[q.id]?.grade)).filter((x: number | null): x is number => x != null); return scores.length ? scores.reduce((a: number, b: number) => a + b, 0) / (scores.length * 10) : 0; };
  for (let ring = 1; ring <= 4; ring++) { const r = radius * ring / 4; pdf.setDrawColor(210); for (let i = 0; i < count; i++) { const angle = -Math.PI / 2 + i * 2 * Math.PI / count; const next = -Math.PI / 2 + (i + 1) * 2 * Math.PI / count; pdf.line(cx + Math.cos(angle) * r, cy + Math.sin(angle) * r, cx + Math.cos(next) * r, cy + Math.sin(next) * r); } }
  pdf.setDrawColor(130); for (let i = 0; i < count; i++) { const a = -Math.PI / 2 + i * 2 * Math.PI / count; pdf.line(cx, cy, cx + Math.cos(a) * radius, cy + Math.sin(a) * radius); const label = safe(chapters[i]?.name).replace(/^\d+\.\s*/, "").slice(0, 16); pdf.setFontSize(6); pdf.text(label, cx + Math.cos(a) * (radius + 7), cy + Math.sin(a) * (radius + 7), { align: "center" }); }
  pdf.setDrawColor(23, 91, 151); pdf.setFillColor(82, 147, 204); const points = chapters.map((chapter: AnyRecord, i: number) => { const a = -Math.PI / 2 + i * 2 * Math.PI / count, r = Math.max(2, radius * value(chapter)); return [cx + Math.cos(a) * r, cy + Math.sin(a) * r] as [number, number]; }); for (let i = 0; i < points.length; i++) { const a = points[i], b = points[(i + 1) % points.length]; pdf.line(a[0], a[1], b[0], b[1]); pdf.circle(a[0], a[1], 1.2, "F"); }
}


