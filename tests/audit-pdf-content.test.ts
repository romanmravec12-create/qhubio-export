import { describe, expect, it, vi } from "vitest";

const captured = vi.hoisted(() => [] as string[]);
const capturedSizes = vi.hoisted(() => [] as Array<{ text: string; size: number }>);
vi.mock("jspdf", () => ({ jsPDF: class {
  size = 0;
  setFont() {} setFontSize(size: number) { this.size = size; } setDrawColor() {} setLineWidth() {} line() {} setFillColor() {} rect() {} setTextColor() {} addPage() {} setPage() {} addImage() {} circle() {}
  splitTextToSize(text: string) { return [text]; }
  text(value: string) { captured.push(value); capturedSizes.push({ text: value, size: this.size }); }
  getNumberOfPages() { return 1; }
  output() { return new ArrayBuffer(1); }
} }));
import { buildAuditPdf } from "../lib/auditPdf";

describe("detailed audit report", () => {
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
});
