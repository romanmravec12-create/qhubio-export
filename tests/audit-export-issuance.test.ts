import { afterEach, describe, expect, it, vi } from "vitest";

const render = vi.hoisted(() => vi.fn(() => new Uint8Array([37, 80, 68, 70]).buffer));
vi.mock("../lib/auditPdf", () => ({ buildAuditPdf: render, AuditReportBrandingError: class AuditReportBrandingError extends Error {} }));
import handler from "../api/audit-export";

const json = (data: unknown, range?: string) => new Response(JSON.stringify(data), { headers: { "content-type": "application/json", ...(range ? { "content-range": range } : {}) } });
const response = () => ({ statusCode: 200, headers: {} as Record<string, string>, body: undefined as unknown, setHeader(k: string, v: string) { this.headers[k] = v; }, status(n: number) { this.statusCode = n; return this; }, json(v: unknown) { this.body = v; return this; }, send(v: unknown) { this.body = v; return this; }, end() { return this; } });
const request = { method: "POST", headers: { authorization: "Bearer synthetic-user-token" }, body: { auditId: "audit-1", format: "executive" } };

describe("internal issued report snapshot", () => {
  afterEach(() => { vi.unstubAllGlobals(); vi.clearAllMocks(); });
  it.each([false, true])("rechecks approval requests and fails on changes: %s", async (changed) => {
    let requestReads = 0;
    vi.stubGlobal("fetch", vi.fn(async (url: string, init?: RequestInit) => {
      expect((init?.headers as Record<string, string>).Authorization).toBe("Bearer synthetic-user-token");
      if (url.endsWith("/auth/v1/user")) return json({ id: "user-1" });
      if (url.includes("/rpc/get_user_organization_id")) return json("org-1");
      if (url.includes("audit_v2_runs?")) return json([{ id: "audit-1", revision: 6, status: "issued", document: { name: "QA", template: { family: "internal" } }, issued_snapshot: { template: { family: "internal" } } }]);
      if (url.includes("audit_v2_report_branding?")) return json([]);
      if (url.includes("audit_v2_approval_requests?")) {
        expect(url).toContain("audit_id=eq.audit-1");
        requestReads++;
        return json([{ id: changed && requestReads > 1 ? "changed" : "approval", requested_revision: 5 }], "0-0/1");
      }
      if (url.includes("audit_v2_approval_events?")) return json([{ id: "event", approval_request_id: "approval" }], "0-0/1");
      return json([], "*/0");
    }));
    const res = response();
    await handler(request as never, res as never);
    expect(requestReads).toBe(2);
    expect(res.statusCode).toBe(changed ? 409 : 200);
    expect(render).toHaveBeenCalledWith(expect.objectContaining({ approvalRequests: [expect.objectContaining({ requested_revision: 5 })], approvals: [expect.objectContaining({ approval_request_id: "approval" })] }));
  });
});

