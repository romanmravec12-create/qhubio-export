import { describe, expect, it, vi } from "vitest";

const captured = vi.hoisted(() => [] as string[]);
vi.mock("jspdf", () => ({ jsPDF: class {
  setFont() {} setFontSize() {} setDrawColor() {} setLineWidth() {} line() {} setFillColor() {} rect() {} setTextColor() {} addPage() {} setPage() {} addImage() {} circle() {}
  splitTextToSize(text: string) { return [text]; }
  text(value: string) { captured.push(value); }
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
    for (const value of ["Second auditor", "Annual review", "Check log", "process.txt", "Source excerpt", "Observation", "Record 42", "Finding ID: f1", "Related question ID: q1", "2026-11-06", "Log missing", "Owner", "2026-11-09", "2026-11-27", "2027-02-04", "Hold material", "No owner", "Assign owner", "Assigned", "Verified", "Verifier", "2026-12-01"]) expect(text).toContain(value);
  });
});
