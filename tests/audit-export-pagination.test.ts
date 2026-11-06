import { afterEach, describe, expect, it, vi } from "vitest";

const build = vi.hoisted(() => vi.fn((_input: unknown) => new Uint8Array([37, 80, 68, 70]).buffer));
vi.mock("../lib/auditPdf", () => ({ buildAuditPdf: build }));
import handler from "../api/audit-export";

function response(status = 200) {
  return { statusCode: status, headers: {} as Record<string, string>, body: undefined as unknown, setHeader(name: string, value: string) { this.headers[name] = value; }, status(value: number) { this.statusCode = value; return this; }, json(value: unknown) { this.body = value; return this; }, send(value: unknown) { this.body = value; return this; }, end() { return this; } };
}
const json = (body: unknown, range?: string, status = 200) => new Response(JSON.stringify(body), { status, headers: { "content-type": "application/json", ...(range ? { "content-range": range } : {}) } });
const request = { method: "POST", headers: { authorization: "Bearer synthetic-user-token" }, body: { auditId: "audit-1", format: "detailed" } };

describe("audit report endpoint pagination", () => {
  afterEach(() => { vi.unstubAllGlobals(); vi.clearAllMocks(); });
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
});
