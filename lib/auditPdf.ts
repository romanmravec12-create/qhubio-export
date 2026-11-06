import { jsPDF } from "jspdf";

export type AuditPdfMode = "executive" | "detailed";
type AnyRecord = Record<string, any>;

const safe = (value: unknown) => String(value ?? "").replace(/[\u0000-\u001f]/g, " ").trim();
const titleCase = (value: unknown) => safe(value).replace(/_/g, " ").replace(/\b\w/g, (c) => c.toUpperCase());
export const auditPdfGradeLabel = (grade: unknown) => ({ green: "Fully met", yellow: "Partially met", red: "Not met", na: "Not applicable", not_assessed: "Not assessed" } as Record<string, string>)[String(grade)] ?? (grade == null ? "Unanswered" : typeof grade === "number" ? `${grade} / 10` : titleCase(grade));
const score = (grade: unknown) => typeof grade === "number" ? grade : grade === "green" ? 10 : grade === "yellow" ? 6 : grade === "red" ? 0 : null;

export function buildAuditPdf({ audit, evidence, links, history, approvals = [], branding, mode }: { audit: AnyRecord; evidence: AnyRecord[]; links: AnyRecord[]; history: AnyRecord[]; approvals?: AnyRecord[]; branding?: { bytes: Uint8Array; format: "PNG" | "JPEG" } | null; mode: AuditPdfMode }) {
  const pdf = new jsPDF({ unit: "mm", format: "a4", compress: true });
  const document = audit.issued_snapshot ?? audit.document;
  const draft = !audit.issued_snapshot;
  const result = evaluateAuditForPdf(document);
  let y = 20;
  const watermark = () => { if (draft) { pdf.setTextColor(232, 232, 232); pdf.setFont("helvetica", "bold"); pdf.setFontSize(28); pdf.text("DRAFT — NOT ISSUED", 105, 155, { align: "center", angle: 35 }); pdf.setTextColor(20, 25, 30); } };
  const page = () => { pdf.addPage(); y = 20; watermark(); };
  const write = (value: unknown, size = 9, bold = false, indent = 0) => {
    const text = safe(value) || "—";
    pdf.setFont("helvetica", bold ? "bold" : "normal"); pdf.setFontSize(size);
    for (const line of pdf.splitTextToSize(text, 174 - indent)) { if (y > 276) { page(); pdf.setFont("helvetica", bold ? "bold" : "normal"); pdf.setFontSize(size); } pdf.text(line, 18 + indent, y); y += size * 0.48 + 1.8; }
  };
  const heading = (value: string, level = 14) => { if (y > 260) page(); y += 3; pdf.setDrawColor(32, 66, 104); pdf.setLineWidth(.45); pdf.line(18, y, 192, y); y += 6; write(value, level, true); };
  const key = (label: string, value: unknown) => write(`${label}: ${safe(value) || "—"}`, 9, false);
  const findingDetails = (finding: AnyRecord) => {
    write(`${titleCase(finding.severity)} ${titleCase(finding.type)} · ${finding.title}`, 10, true);
    key("Finding ID", finding.id); key("Related question ID", finding.questionId); key("Finding date", finding.createdOn);
    key("Evidence", finding.evidence); key("Owner", finding.owner); key("Containment due", finding.containmentDue);
    key("Action plan due", finding.planDue); key("Implementation due", finding.due); key("Containment", finding.containment);
    key("Root cause", finding.rootCause); key("Action status", titleCase(finding.status)); key("Corrective action", finding.action);
    key("Implementation", finding.implementation); key("Effectiveness", finding.effectiveness); key("Verified by", finding.verifiedBy); key("Verified on", finding.verifiedOn);
  };
  const header = () => { pdf.setFillColor(21, 55, 91); pdf.rect(0, 0, 210, 13, "F"); pdf.setTextColor(255, 255, 255); pdf.setFont("helvetica", "bold"); pdf.setFontSize(10); pdf.text("QHUBIO · AUDIT", 18, 8); pdf.setTextColor(20, 25, 30); };
  header();
  watermark();
  if (branding) { try { pdf.addImage(branding.bytes, branding.format, 146, 18, 44, 18, undefined, "FAST"); } catch { /* branding is optional, never fail an audit record because of it */ } }
  write(mode === "executive" ? "Audit Executive Summary" : "Detailed Audit Report", 19, true);
  if (draft) { pdf.setTextColor(160, 35, 35); write("DRAFT — NOT ISSUED", 11, true); pdf.setTextColor(20, 25, 30); }
  write(document.name, 15, true);
  key("Audit reference", audit.id); key("Template", `${document.template?.name ?? "Audit template"} · v${document.template?.version ?? "—"}`); key("Audited organisation", document.organization); key("Site", document.site); key("Audit date", document.plannedOn); key("Lead auditor", document.auditor); key("Co-auditor", document.coAuditor); key("Scope", document.scope); key("Risk / reason for audit", document.trigger); key("Sample / project", document.sample);
  if (document.template?.sourceExtractionWarning === "word_elements_omitted") key("Source review note", "Some Word document elements were not interpreted during questionnaire preparation. Review the original source for coverage.");
  heading("Assessment result");
  key("Completion", `${result.coverage ?? 0}%`); key("Result", result.result); key("Score", result.score == null ? "Not calculated" : `${Number(result.score).toFixed(1)}% · band ${result.band ?? "—"}`); key("Conclusion", document.conclusion); if (document.template?.family === "supplier") key("Supplier / release decision", titleCase(document.releaseDecision));
  if (mode === "executive") {
    if (y > 190) page();
    heading("Chapter performance"); drawRadar(pdf, document, 105, y + 35, 30); y += 70;
    heading("Key findings");
    const open = (document.findings ?? []).filter((f: AnyRecord) => f.status !== "closed");
    if (!open.length) write("No open findings recorded.");
    for (const finding of open.slice(0, 8)) { write(`${titleCase(finding.severity)} ${titleCase(finding.type)} · ${finding.title}`, 10, true); key("Owner / due", `${finding.owner || "—"} / ${finding.due || "—"}`); if (finding.severity === "major") key("Major finding deadlines", `Containment ${finding.containmentDue || "—"}; action plan ${finding.planDue || "—"}`); }
    heading("Management attention"); key("Open findings", open.length); key("Overdue findings", open.filter((f: AnyRecord) => f.due && f.due < new Date().toISOString().slice(0, 10)).length); key("CAPA links", links.length); key("Issued status", audit.issued_snapshot ? "Issued assessment snapshot" : "Working audit record");
  } else {
    heading("Scoring policy"); write("Scores and responses are reproduced from the active audit template. Not applicable responses are excluded from the calculated score. The audit conclusion remains an auditor judgement, not a certificate or automatic release authorization.");
    let questionNumber = 0;
    for (const chapter of document.template?.chapters ?? []) {
      heading(chapter.name);
      for (const question of chapter.questions ?? []) {
        const response = document.responses?.[question.id] ?? {};
        if (y > 245) page();
        write(`Question ${++questionNumber} · ${question.text}`, 10, true); key("Question ID", question.id); key("Reference", question.reference);
        if (typeof question.required === "boolean") key("Required", question.required ? "Yes" : "No");
        if (typeof question.critical === "boolean") key("Critical control", question.critical ? "Yes" : "No");
        if (typeof question.weight === "number") key("Weight", question.weight);
        if (question.guidance) key("Auditor guidance", question.guidance); if (question.provenance?.label) key("Source", question.provenance.label); if (question.provenance?.rationale) key("Source rationale", question.provenance.rationale); if (question.provenance?.edited) key("Source note", "The reference relates to the original suggestion; verify it against the edited question."); key("Assessment", auditPdfGradeLabel(response.grade)); if (response.note) key("Observation / evidence examined", response.note); if (response.justification) key("Non-applicability rationale", response.justification); if (response.evidenceRefs) key("Document / sample references", response.evidenceRefs);
        const docs = evidence.filter((item) => item.question_id === question.id); if (docs.length) write(`Attached documents: ${docs.map((item) => `${item.name}${item.source_media_type ? ` (${item.source_media_type})` : ""}`).join("; ")}`, 8, false);
      }
    }
    heading(audit.issued_snapshot ? "Findings at issue" : "Draft findings");
    for (const finding of document.findings ?? []) findingDetails(finding);
    if (audit.issued_snapshot) { heading("Current follow-up after issue"); for (const finding of audit.document?.findings ?? []) findingDetails(finding); }
    heading("CAPA and audit history"); if (links.length) for (const link of links) { write(`CAPA ID: ${safe(link.capa_id) || "—"}`, 9, true); key("Related finding ID", link.finding_id); key("Linked at", link.linked_at); } else write("No CAPA links recorded."); if (history.length) { y += 2; for (const event of history) write(`${event.created_at || ""} · ${titleCase(event.event)} · revision ${event.revision ?? "—"} · actor ID ${event.actor_id || "—"}`, 8); }
    heading("Controlled electronic sign-off"); if (approvals.length) for (const approval of approvals) { write(`${approval.created_at || ""} · ${titleCase(approval.event)} · ${approval.meaning || "Controlled sign-off event"}`, 8); key("Actor ID", approval.actor_id); key("Document SHA-256", approval.document_hash); } else write("No controlled electronic sign-off has been recorded for this audit.");
  }
  const pages = pdf.getNumberOfPages(); for (let i = 1; i <= pages; i++) { pdf.setPage(i); pdf.setDrawColor(190); pdf.line(18, 286, 192, 286); pdf.setFontSize(8); pdf.setTextColor(80); pdf.text(`Qhubio audit record · ${audit.id} · ${i} / ${pages}`, 18, 291); }
  return pdf.output("arraybuffer");
}

/** Mirrors the audit workspace evaluation contract; keep in sync with model.ts. */
export function evaluateAuditForPdf(document: AnyRecord) {
  const questions = (document.template?.chapters ?? []).flatMap((chapter: AnyRecord) => chapter.questions ?? []);
  const responses = document.responses ?? {};
  const assessed = questions.filter((q: AnyRecord) => responses[q.id]?.grade != null && !["na", "not_assessed"].includes(responses[q.id].grade));
  const complete = questions.filter((q: AnyRecord) => {
    const r = responses[q.id];
    return r?.grade != null && r.grade !== "not_assessed" && (r.grade !== "na" || Boolean(r.justification?.trim()));
  }).length;
  let value: number | null = null;
  if (document.template?.profile === "supplier_score" && assessed.length) {
    const numerator = assessed.reduce((sum: number, q: AnyRecord) => sum + Number(responses[q.id].grade) * q.weight, 0);
    const denominator = assessed.reduce((sum: number, q: AnyRecord) => sum + 10 * q.weight, 0);
    value = numerator / denominator * 100;
  }
  const red = assessed.some((q: AnyRecord) => responses[q.id].grade === "red" || (typeof responses[q.id].grade === "number" && responses[q.id].grade <= 4));
  const yellow = assessed.some((q: AnyRecord) => responses[q.id].grade === "yellow");
  const blocked = questions.some((q: AnyRecord) => q.critical && !["green", 10, "na"].includes(responses[q.id]?.grade)) || (document.findings ?? []).some((f: AnyRecord) => f.type === "nonconformity" && f.severity === "major" && f.status !== "closed");
  const band = value === null ? null : value >= 90 ? "A" : value >= 70 ? "B" : "C";
  const decision = blocked ? "Decision blocked" : document.releaseDecision === "released" ? "Approved / released" : document.releaseDecision === "conditional" ? "Conditionally accepted" : document.releaseDecision === "not_released" ? "Not approved / not released" : "Decision required";
  const result = !assessed.length ? "Not evaluated" : complete < questions.length ? "Assessment incomplete" : document.template?.profile === "supplier_score" ? `${band} — ${decision}` : red ? "Not met" : yellow ? "Partially met" : "Fully met";
  return { coverage: questions.length ? Math.round(complete / questions.length * 100) : 0, score: value, band, result };
}

function drawRadar(pdf: jsPDF, document: AnyRecord, cx: number, cy: number, radius: number) {
  const chapters = document.template?.chapters ?? []; const count = Math.max(chapters.length, 3); const value = (chapter: AnyRecord) => { const scores = (chapter.questions ?? []).map((q: AnyRecord) => score(document.responses?.[q.id]?.grade)).filter((x: number | null): x is number => x != null); return scores.length ? scores.reduce((a: number, b: number) => a + b, 0) / (scores.length * 10) : 0; };
  for (let ring = 1; ring <= 4; ring++) { const r = radius * ring / 4; pdf.setDrawColor(210); for (let i = 0; i < count; i++) { const angle = -Math.PI / 2 + i * 2 * Math.PI / count; const next = -Math.PI / 2 + (i + 1) * 2 * Math.PI / count; pdf.line(cx + Math.cos(angle) * r, cy + Math.sin(angle) * r, cx + Math.cos(next) * r, cy + Math.sin(next) * r); } }
  pdf.setDrawColor(130); for (let i = 0; i < count; i++) { const a = -Math.PI / 2 + i * 2 * Math.PI / count; pdf.line(cx, cy, cx + Math.cos(a) * radius, cy + Math.sin(a) * radius); const label = safe(chapters[i]?.name).replace(/^\d+\.\s*/, "").slice(0, 16); pdf.setFontSize(6); pdf.text(label, cx + Math.cos(a) * (radius + 7), cy + Math.sin(a) * (radius + 7), { align: "center" }); }
  pdf.setDrawColor(23, 91, 151); pdf.setFillColor(82, 147, 204); const points = chapters.map((chapter: AnyRecord, i: number) => { const a = -Math.PI / 2 + i * 2 * Math.PI / count, r = Math.max(2, radius * value(chapter)); return [cx + Math.cos(a) * r, cy + Math.sin(a) * r] as [number, number]; }); for (let i = 0; i < points.length; i++) { const a = points[i], b = points[(i + 1) % points.length]; pdf.line(a[0], a[1], b[0], b[1]); pdf.circle(a[0], a[1], 1.2, "F"); }
}
