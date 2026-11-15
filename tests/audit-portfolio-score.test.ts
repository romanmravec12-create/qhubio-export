import { describe, expect, it } from "vitest";
import { buildAuditPortfolioPdf } from "../lib/auditPortfolioPdf";

async function reportText(document: Record<string, unknown>) {
  const bytes = buildAuditPortfolioPdf({ audits: [{ id: "fictional", revision: 1, document, issued_snapshot: null }], generatedAt: "2026-01-01T00:00:00Z", includeDrafts: true });
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const task = pdfjs.getDocument({ data: new Uint8Array(bytes), useSystemFonts: false });
  const pdf = await task.promise;
  const pages = await Promise.all(Array.from({ length: pdf.numPages }, async (_, index) => (await (await pdf.getPage(index + 1)).getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" ")));
  await task.destroy();
  return pages.join(" ");
}
const template = (profile: string) => ({ name: "Test", profile, chapters: [{ questions: [{ id: "one", weight: 1 }, { id: "two", weight: 3 }] }] });
describe("portfolio scores", () => {
  it("does not invent a percentage for a compliance audit", async () => {
    const text = await reportText({ name: "Compliance", organization: "QA", template: template("compliance"), responses: { one: { grade: "green" }, two: { grade: "green" } }, findings: [] });
    expect(text).toContain("Score: Not calculated");
    expect(text).toContain("Average completed supplier score: Not calculated");
  }, 15000);
  it("does not present an incomplete supplier audit as a completed score", async () => {
    const text = await reportText({ name: "Supplier", organization: "QA", template: template("supplier_score"), responses: { one: { grade: 10 } }, findings: [] });
    expect(text).toContain("Score: Not calculated");
    expect(text).not.toContain("Score: 100%");
  });
  it("uses question weights for a completed supplier score", async () => {
    const text = await reportText({ name: "Supplier", organization: "QA", template: template("supplier_score"), responses: { one: { grade: 10 }, two: { grade: 0 } }, findings: [] });
    expect(text).toContain("Score: 25%");
    expect(text).toContain("Average completed supplier score: 25%");
  });
  it("requires a justification before excluding a question as not applicable", async () => {
    const base = { name: "Supplier", organization: "QA", template: template("supplier_score"), findings: [] };
    expect(await reportText({ ...base, responses: { one: { grade: 10 }, two: { grade: "na", justification: "" } } })).toContain("Score: Not calculated");
    expect(await reportText({ ...base, responses: { one: { grade: 10 }, two: { grade: "na", justification: "Outside scope" } } })).toContain("Score: 100%");
  });
});
