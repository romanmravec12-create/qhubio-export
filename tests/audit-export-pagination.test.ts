import { afterEach, describe, expect, it, vi } from "vitest";

const build = vi.hoisted(() => vi.fn((_input: unknown) => new Uint8Array([37, 80, 68, 70]).buffer));
vi.mock("../lib/auditPdf", () => ({ buildAuditPdf: build, AuditReportBrandingError: class AuditReportBrandingError extends Error {} }));
import handler from "../api/audit-export";
import { AuditReportTextError } from "../lib/auditReportText";

function response(status = 200) {
  return { statusCode: status, headers: {} as Record<string, string>, body: undefined as unknown, setHeader(name: string, value: string) { this.headers[name] = value; }, status(value: number) { this.statusCode = value; return this; }, json(value: unknown) { this.body = value; return this; }, send(value: unknown) { this.body = value; return this; }, end() { return this; } };
}
const json = (body: unknown, range?: string, status = 200) => new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json", ...(range ? { "content-range": range } : {}) } });
const request = { method: "POST", headers: { authorization: "Bearer synthetic-user-token" }, body: { auditId: "audit-1", format: "detailed" } };

describe("audit report endpoint pagination", () => {
  afterEach(() => { vi.unstubAllGlobals(); vi.clearAllMocks(); });
  it("does not fetch related data or render a foreign audit", async () => {
    const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
      expect((init?.headers as Record<string, string>).Authorization).toBe("Bearer synthetic-user-token");
      if (url.endsWith("/auth/v1/user")) return json({ id: "user-1" });
      if (url.includes("/rpc/get_user_organization_id")) return json("org-1");
      if (url.includes("audit_v2_runs?")) {
        expect(url).toContain("id=eq.foreign-audit");
        expect(url).toContain("organization_id=eq.org-1");
        return json([]); // The user's organization has no access to this audit.
      }
      throw new Error(`Unexpected request: ${url}`);
    });
    vi.stubGlobal("fetch", fetchMock);
    const res = response();
    await handler({ ...request, body: { ...request.body, auditId: "foreign-audit" } } as never, res as never);
    expect(res.statusCode).toBe(404);
    expect(build).not.toHaveBeenCalled();
    expect(fetchMock).toHaveBeenCalledTimes(3);
  });

  it("does not deliver a report if audit access disappears during export", async () => {
    let auditReads = 0;
    vi.stubGlobal("fetch", vi.fn(async (url: string) => {
      if (url.endsWith("/auth/v1/user")) return json({ id: "user-1" });
      if (url.includes("/rpc/get_user_organization_id")) return json("org-1");
      if (url.includes("audit_v2_runs?")) return json(++auditReads === 1 ? [{ id: "audit-1", revision: 1, status: "review", document: { name: "Test" } }] : []);
      if (url.includes("audit_v2_report_branding?")) return json([]);
      return json([], "*/0");
    }));
    const res = response(); await handler(request as never, res as never);
    expect(res.statusCode).toBe(409);
    expect(res.body).toMatchObject({ success: false });
  });
  it("passes source warnings from the issued audit snapshot through to the PDF renderer", async () => {
    const snapshot = { name: "Issued QA audit", template: { sourceExtractionWarning: "pdf_word_and_spreadsheet_incomplete", chapters: [] } };
    vi.stubGlobal("fetch", vi.fn(async (url: string) => {
      if (url.endsWith("/auth/v1/user")) return json({ id: "user-1" });
      if (url.includes("/rpc/get_user_organization_id")) return json("org-1");
      if (url.includes("audit_v2_runs?")) return json([{ id: "audit-1", revision: 1, status: "issued", document: { name: "Current" }, issued_snapshot: snapshot }]);
      if (url.includes("audit_v2_report_branding?")) return json([]);
      return json([], "*/0");
    }));
    const res = response(); await handler(request as never, res as never);
    expect(res.statusCode).toBe(200);
    expect(res.headers['X-Qhubio-Audit-Contract']).toBe('internal-audit-2026-11-14-v1');
    expect(res.headers['Access-Control-Expose-Headers']).toContain('X-Qhubio-Audit-Contract');
    expect(build).toHaveBeenCalledWith(expect.objectContaining({ audit: expect.objectContaining({ issued_snapshot: snapshot }) }));
  });
  it("returns a clear error without a partial PDF for unsupported text", async () => {
    vi.stubGlobal("fetch", vi.fn(async (url: string) => {
      if (url.endsWith("/auth/v1/user")) return json({ id: "user-1" });
      if (url.includes("/rpc/get_user_organization_id")) return json("org-1");
      if (url.includes("audit_v2_runs?")) return json([{ id: "audit-1", revision: 1, status: "review", document: { name: "Test" } }]);
      if (url.includes("audit_v2_report_branding?")) return json([]);
      return json([], "*/0");
    }));
    build.mockImplementationOnce(() => { throw new AuditReportTextError("This report contains characters its PDF fonts cannot render. No partial report was exported."); });
    const res = response(); await handler(request as never, res as never);
    expect(res.statusCode).toBe(422);
    expect(res.body).toMatchObject({ success: false, error: expect.stringContaining("fonts cannot render") });
  });
  it("passes every evidence page to the report with the user authorization", async () => {
    const fetchMock = vi.fn(async (url: string, init?: RequestInit) => {
      if (url.endsWith("/auth/v1/user")) return json({ id: "user-1" });
      if (url.includes("/rpc/get_user_organization_id")) return json("org-1");
      if (url.includes("audit_v2_runs?")) return json([{ id: "audit-1", revision: 1, status: "review", document: { name: "Test" } }]);
      if (url.includes("audit_v2_report_branding?")) return json([]);
      if (url.includes("audit_v2_events?")) expect(url).toContain("select=id,created_at,actor_id,event,revision&");
      if (url.includes("audit_v2_capa_links?")) expect(url).toContain("select=id,capa_id,finding_id,linked_at&");
      const headers = init?.headers as Record<string, string>;
      expect(headers.Authorization).toBe("Bearer synthetic-user-token");
      expect(headers.Prefer).toBe("count=exact");
      if (url.includes("audit_v2_evidence?")) {
        const start = Number(headers.Range.split("-")[0]);
        return start === 0 ? json(Array.from({ length: 500 }, (_, i) => ({ id: `e${i}` })), "0-499/501") : json([{ id: "e500" }], "500-500/501");
      }
      return json([], "*/0");
    });
    vi.stubGlobal("fetch", fetchMock);
    const res = response();
    await handler(request as never, res as never);
    expect(res.statusCode).toBe(200);
    expect((build.mock.calls[0][0] as { evidence: unknown[] }).evidence).toHaveLength(501);
  });

  it("does not generate a PDF when a page count cannot be verified", async () => {
    vi.stubGlobal("fetch", vi.fn(async (url: string) => {
      if (url.endsWith("/auth/v1/user")) return json({ id: "user-1" });
      if (url.includes("/rpc/get_user_organization_id")) return json("org-1");
      if (url.includes("audit_v2_runs?")) return json([{ id: "audit-1", document: { name: "Test" } }]);
      return json([]);
    }));
    const res = response();
    await handler(request as never, res as never);
    expect(res.statusCode).toBe(500);
    expect(build).not.toHaveBeenCalled();
    expect(res.body).toMatchObject({ error: expect.stringContaining("completeness") });
  });

  it("does not deliver a report if the audit changes during export", async () => {
    let auditReads = 0;
    vi.stubGlobal("fetch", vi.fn(async (url: string) => {
      if (url.endsWith("/auth/v1/user")) return json({ id: "user-1" });
      if (url.includes("/rpc/get_user_organization_id")) return json("org-1");
      if (url.includes("audit_v2_runs?")) return json([{ id: "audit-1", revision: ++auditReads, status: "review", document: { name: "Test" } }]);
      if (url.includes("audit_v2_report_branding?")) return json([]);
      return json([], "*/0");
    }));
    const res = response();
    await handler(request as never, res as never);
    expect(res.statusCode).toBe(409);
    expect(res.body).toMatchObject({ error: expect.stringContaining("changed during export") });
  });

  it.each(["added", "changed"])("does not deliver a report if evidence is %s during export", async (change) => {
    let reads = 0;
    vi.stubGlobal("fetch", vi.fn(async (url: string) => {
      if (url.endsWith("/auth/v1/user")) return json({ id: "user-1" });
      if (url.includes("/rpc/get_user_organization_id")) return json("org-1");
      if (url.includes("audit_v2_runs?")) return json([{ id: "audit-1", revision: 1, status: "review", document: { name: "Test" } }]);
      if (url.includes("audit_v2_report_branding?")) return json([]);
      if (url.includes("audit_v2_evidence?")) {
        const first = [{ id: "e1", question_id: "q1", name: "file.pdf", source_media_type: "application/pdf" }];
        if (++reads === 1) return json(first, "0-0/1");
        return change === "added" ? json([...first, { id: "e2", question_id: "q1", name: "new.pdf", source_media_type: "application/pdf" }], "0-1/2")
          : json([{ ...first[0], name: "changed.pdf" }], "0-0/1");
      }
      return json([], "*/0");
    }));
    const res = response();
    await handler(request as never, res as never);
    expect(res.statusCode).toBe(409);
    expect(res.body).toMatchObject({ error: expect.stringContaining("evidence changed during export") });
  });

  it("does not silently omit an unavailable saved company logo", async () => {
    vi.stubGlobal("fetch", vi.fn(async (url: string) => {
      if (url.endsWith("/auth/v1/user")) return json({ id: "user-1" });
      if (url.includes("/rpc/get_user_organization_id")) return json("org-1");
      if (url.includes("audit_v2_runs?")) return json([{ id: "audit-1", revision: 1, status: "review", document: { name: "Test" } }]);
      if (url.includes("audit_v2_report_branding?")) return json([{ logo_path: "org/logo", logo_media_type: "image/png" }]);
      if (url.includes("/storage/v1/object/")) return json({}, undefined, 403);
      return json([], "*/0");
    }));
    const res = response();
    await handler(request as never, res as never);
    expect(res.statusCode).toBe(422);
    expect(res.body).toMatchObject({ error: expect.stringContaining("logo could not be loaded") });
    expect(build).not.toHaveBeenCalled();
  });

  it("allows an explicit Qhubio-branding export without loading the saved logo", async () => {
    const fetchMock = vi.fn(async (url: string) => {
      if (url.endsWith("/auth/v1/user")) return json({ id: "user-1" });
      if (url.includes("/rpc/get_user_organization_id")) return json("org-1");
      if (url.includes("audit_v2_runs?")) return json([{ id: "audit-1", revision: 1, status: "review", document: { name: "Test" } }]);
      if (url.includes("audit_v2_report_branding?")) return json([{ logo_path: "org/logo", logo_media_type: "image/png" }]);
      if (url.includes("/storage/v1/object/")) throw new Error("Unexpected logo download");
      return json([], "*/0");
    });
    vi.stubGlobal("fetch", fetchMock);
    const res = response();
    await handler({ ...request, body: { ...request.body, useQhubioBranding: true } } as never, res as never);
    expect(res.statusCode).toBe(200);
    expect(fetchMock.mock.calls.some(([url]) => String(url).includes("/storage/v1/object/"))).toBe(false);
  });
});

