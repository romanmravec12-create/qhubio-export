import { describe, expect, it, vi } from "vitest";

const captured = vi.hoisted(() => [] as string[]);
const capturedSizes = vi.hoisted(() => [] as Array<{ text: string; size: number }>);
const capturedPositions = vi.hoisted(() => [] as Array<{ text: string; y: number }>);
const capturedPoints = vi.hoisted(() => [] as Array<[number, number]>);
const imageFailure = vi.hoisted(() => ({ enabled: false }));
vi.mock("jspdf", () => ({ jsPDF: class {
  size = 0;
  setFont() {} setFontSize(size: number) { this.size = size; } setDrawColor() {} setLineWidth() {} line() {} setFillColor() {} rect() {} setTextColor() {} addPage() {} setPage() {} getImageProperties() { return { width: 100, height: 50 }; } addImage() { if (imageFailure.enabled) throw new Error("bad image"); } circle(x: number, y: number) { capturedPoints.push([x, y]); }
  splitTextToSize(text: string) { return [text]; }
  text(value: string, _x: number, y: number) { captured.push(value); capturedSizes.push({ text: value, size: this.size }); capturedPositions.push({ text: value, y }); }
  getNumberOfPages() { return 1; }
  output() { return new ArrayBuffer(1); }
} }));
import { auditRadarChapterValue, AuditReportBrandingError, buildAuditPdf } from "../lib/auditPdf";

describe("detailed audit report", () => {
  it("leaves space between the bottom radar label and its explanatory caption", () => {
    capturedPositions.length = 0;
    const chapters = ["Planning", "Execution", "Monitoring", "Improvement"].map((name, i) => ({ name, questions: [{ id: `q${i}` }] }));
    const document = { name: "Audit", template: { chapters }, responses: {}, findings: [] };
    buildAuditPdf({ audit: { id: "audit", document }, evidence: [], links: [], history: [], mode: "executive" });
    const label = capturedPositions.find((item) => item.text === "Monitoring");
    const caption = capturedPositions.find((item) => item.text.startsWith("Chart uses the unweighted"));
    expect(label).toBeDefined(); expect(caption).toBeDefined();
    expect(caption!.y - label!.y).toBeGreaterThanOrEqual(8);
  });
  it("distinguishes unassessed chapters from zero and explains the executive chart", () => {
    const chapter = { name: "Controls", questions: [{ id: "a" }, { id: "b" }, { id: "c" }] };
    expect(auditRadarChapterValue(chapter, { a: { grade: "na" }, b: { grade: "not_assessed" } })).toBeNull();
    expect(auditRadarChapterValue(chapter, { a: { grade: "red" }, b: { grade: "na" } })).toBe(0);
    expect(auditRadarChapterValue(chapter, { a: { grade: "green" }, b: { grade: "yellow" }, c: { grade: "na" } })).toBe(0.8);
    captured.length = 0; capturedPoints.length = 0;
    const document = { name: "Audit", template: { chapters: [chapter] }, responses: {}, findings: [] };
    buildAuditPdf({ audit: { id: "audit", document, issued_snapshot: null }, evidence: [], links: [], history: [], mode: "executive" });
    expect(capturedPoints).toHaveLength(0);
    expect(captured.join(" ")).toContain("Chapters without assessed answers have no plotted point.");
    document.responses = { a: { grade: "red" } };
    buildAuditPdf({ audit: { id: "audit", document, issued_snapshot: null }, evidence: [], links: [], history: [], mode: "executive" });
    expect(capturedPoints).toHaveLength(1);
    expect(capturedPoints[0][0]).toBe(105);
  });
  it("includes audit context, source context and all submitted finding details", () => {
    captured.length = 0;
    const document = {
      name: "Audit", organization: "Example", coAuditor: "Second auditor", trigger: "Annual review",
      template: { name: "Process", chapters: [{ name: "Controls", questions: [{ id: "q1", text: "Controlled?", guidance: "Check log", reference: "Procedure", provenance: { label: "process.txt", rationale: "Source excerpt" } }] }] },
      responses: { q1: { grade: "red", note: "Observation", evidenceRefs: "Record 42" } },
      findings: [{ id: "f1", questionId: "q1", title: "Missing check", type: "nonconformity", severity: "major", evidence: "Log missing", owner: "Owner", createdOn: "2026-11-06", containmentDue: "2026-11-09", planDue: "2026-11-27", due: "2027-02-04", containment: "Hold material", rootCause: "No owner", action: "Assign owner", implementation: "Assigned", effectiveness: "Verified", status: "verification", verifiedBy: "Verifier", verifiedOn: "2026-12-01" }],
    };
    buildAuditPdf({ audit: { id: "audit-1", document, issued_snapshot: null }, evidence: [], links: [], history: [], mode: "detailed" });
    const text = captured.join("\n");
    expect(captured.indexOf("DRAFT — NOT ISSUED")).toBeLessThan(captured.indexOf("Detailed Audit Report"));
    for (const value of ["Second auditor", "Annual review", "Check log", "process.txt", "Source excerpt", "Observation", "Record 42", "Finding ID: f1", "Related question ID: q1", "2026-11-06", "Log missing", "Owner", "2026-11-09", "2026-11-27", "2027-02-04", "Hold material", "No owner", "Assign owner", "Assigned", "Verified", "Verifier", "2026-12-01"]) expect(text).toContain(value);
  });

  it("keeps issued findings distinct from complete current follow-up and includes all supplied history", () => {
    captured.length = 0;
    const snapshot = { name: "Audit", template: { chapters: [] }, responses: {}, findings: [{ id: "finding", title: "Finding", rootCause: "Original root cause", action: "Original action" }] };
    const current = { ...snapshot, findings: [{ ...snapshot.findings[0], rootCause: "Updated root cause", action: "Updated action", containment: "Updated containment" }] };
    const history = Array.from({ length: 35 }, (_, index) => ({ event: `event_${index}`, revision: index, actor_id: "history-actor" }));
    const links = [{ capa_id: "capa-123", finding_id: "finding", linked_at: "2026-11-06" }];
    const approvals = [{ event: "ISSUED", meaning: "Approved", actor_id: "approval-actor", document_hash: "a".repeat(64) }];
    buildAuditPdf({ audit: { id: "audit", issued_snapshot: snapshot, document: current }, evidence: [], links, history, approvals, mode: "detailed" });
    const text = captured.join("\n");
    expect(text).not.toContain("DRAFT — NOT ISSUED");
    expect(text.indexOf("Findings at issue")).toBeLessThan(text.indexOf("Original root cause"));
    expect(text.indexOf("Original root cause")).toBeLessThan(text.indexOf("Current follow-up after issue"));
    expect(text.indexOf("Current follow-up after issue")).toBeLessThan(text.indexOf("Updated root cause"));
    expect(text).toContain("Updated action");
    expect(text).toContain("Updated containment");
    expect(text).toContain("Event 34");
    expect(text).toContain("history-actor");
    expect(text).toContain("CAPA ID: capa-123");
    expect(text).toContain("Related finding ID: finding");
    expect(text).toContain("approval-actor");
    expect(text).toContain("a".repeat(64));
  });

  it("restores body font size when a draft history entry flows to a new page", () => {
    captured.length = 0; capturedSizes.length = 0;
    const document = { name: "Audit", template: { chapters: [] }, responses: {}, findings: [] };
    const history = Array.from({ length: 90 }, (_, index) => ({ event: `event_${index}`, revision: index }));
    buildAuditPdf({ audit: { id: "audit", issued_snapshot: null, document }, evidence: [], links: [], history, mode: "detailed" });
    const entries = capturedSizes.filter((entry) => entry.text.includes("Event "));
    expect(entries).toHaveLength(90);
    expect(entries.every((entry) => entry.size === 8)).toBe(true);
  });

  it("does not silently omit a logo that cannot be decoded", () => {
    imageFailure.enabled = true;
    try {
      expect(() => buildAuditPdf({ audit: { id: "audit", issued_snapshot: null, document: { template: { chapters: [] }, responses: {}, findings: [] } }, evidence: [], links: [], history: [], branding: { bytes: new Uint8Array([1]), format: "PNG" }, mode: "detailed" })).toThrow(AuditReportBrandingError);
    } finally { imageFailure.enabled = false; }
  });
});

