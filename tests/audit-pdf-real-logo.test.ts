import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { AuditReportBrandingError, buildAuditPdf } from "../lib/auditPdf";
import { AuditReportTextError } from "../lib/auditReportText";

describe("real report logo rendering", () => {
  const document = { name: "Audit", template: { chapters: [] }, responses: {}, findings: [] };
  const request = { audit: { id: "audit", document, issued_snapshot: null }, evidence: [], links: [], history: [], mode: "detailed" as const };

  it("embeds a valid PNG logo", () => {
    const bytes = new Uint8Array(readFileSync("assets/qhubio-logo.png"));
    const pdf = buildAuditPdf({ ...request, branding: { bytes, format: "PNG" } });
    expect(Buffer.from(pdf).includes(Buffer.from("/Subtype /Image"))).toBe(true);
  });

  it("embeds a valid JPEG logo", () => {
    const bytes = new Uint8Array(readFileSync("assets/qhubio-logo.jpg"));
    const pdf = buildAuditPdf({ ...request, branding: { bytes, format: "JPEG" } });
    expect(Buffer.from(pdf).includes(Buffer.from("/Subtype /Image"))).toBe(true);
  });

  it("rejects a corrupt logo rather than silently omitting it", () => {
    expect(() => buildAuditPdf({ ...request, branding: { bytes: new Uint8Array([1, 2, 3]), format: "PNG" } })).toThrow(AuditReportBrandingError);
  });

  it("preserves Slovak and Chinese text in an actual PDF", async () => {
    const chapter = { name: "Čistenie a údržba 中文", questions: [{ id: "q1", text: "Ako sa riešia sťažnosti?" }] };
    const input = { name: "Výrobný proces", organization: "Žilinská spoločnosť 中文", template: { chapters: [chapter] }, responses: {}, findings: [] };
    const bytes = buildAuditPdf({ ...request, audit: { id: "audit", document: input, issued_snapshot: null } });
    const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
    const task = pdfjs.getDocument({ data: new Uint8Array(bytes), useSystemFonts: false });
    const pdf = await task.promise; const result: string[] = [];
    for (let i = 1; i <= pdf.numPages; i++) result.push((await (await pdf.getPage(i)).getTextContent()).items.map((item) => "str" in item ? item.str : "").join(""));
    const text = result.join("");
    for (const value of ["Výrobný proces", "Žilinská spoločnosť", "中文", "Čistenie a údržba", "Ako sa riešia sťažnosti?"]) expect(text).toContain(value);
    await task.destroy();
  });

  it("preserves combined extraction warnings in actual executive and detailed PDFs", async () => {
    const input = { ...document, template: { chapters: [], sourceExtractionWarning: "pdf_word_and_spreadsheet_incomplete" } };
    const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
    for (const mode of ["executive", "detailed"] as const) {
      const bytes = buildAuditPdf({ ...request, mode, audit: { id: "audit", document: input, issued_snapshot: null } });
      const task = pdfjs.getDocument({ data: new Uint8Array(bytes), useSystemFonts: false });
      const pdf = await task.promise; const pages: string[] = [];
      for (let i = 1; i <= pdf.numPages; i++) pages.push((await (await pdf.getPage(i)).getTextContent()).items.map((item) => "str" in item ? item.str : "").join(" "));
      const text = pages.join(" ");
      for (const warning of ["Some Word document elements", "Some spreadsheet formulas", "Some PDF pages had no extractable text"]) expect(text).toContain(warning);
      await task.destroy();
    }
  });

  it("rejects unsupported characters instead of silently corrupting the report", () => {
    expect(() => buildAuditPdf({ ...request, audit: { ...request.audit, document: { ...document, name: "Audit 😀" } } })).toThrow(AuditReportTextError);
  });
});

