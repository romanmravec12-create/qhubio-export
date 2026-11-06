import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { AuditReportBrandingError, buildAuditPdf } from "../lib/auditPdf";

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
});
