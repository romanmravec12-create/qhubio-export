import { describe, expect, it, vi } from "vitest";
import { auditReportRows } from "../lib/auditReportRows";

const page = (values: unknown[], range: string, status = 200) => new Response(JSON.stringify(values), { status, headers: { "content-range": range, "content-type": "application/json" } });
const items = (from: number, count: number) => Array.from({ length: count }, (_, index) => ({ id: String(from + index) }));

describe("audit report pagination", () => {
  it("handles empty results and multiple pages without dropping rows", async () => {
    await expect(auditReportRows(async () => page([], "*/0"))).resolves.toEqual([]);
    const fetchPage = vi.fn().mockResolvedValueOnce(page(items(0, 500), "0-499/501")).mockResolvedValueOnce(page(items(500, 1), "500-500/501"));
    const rows = await auditReportRows(fetchPage);
    expect(rows).toHaveLength(501);
    expect(fetchPage).toHaveBeenNthCalledWith(2, 500, 999);
    expect(rows[500].id).toBe("500");
  });

  it("fails closed for an unavailable or unverifiable response", async () => {
    await expect(auditReportRows(async () => page([], "*/0", 503))).rejects.toThrow("could not be retrieved");
    await expect(auditReportRows(async () => new Response("[]"))).rejects.toThrow("completeness");
    await expect(auditReportRows(async () => page(items(0, 10), "0-9/100"))).rejects.toThrow("completeness");
  });

  it("fails closed when counts change, rows repeat, or an export exceeds its safe limit", async () => {
    const changed = vi.fn().mockResolvedValueOnce(page(items(0, 500), "0-499/501")).mockResolvedValueOnce(page(items(500, 2), "500-501/502"));
    await expect(auditReportRows(changed)).rejects.toThrow("changed");
    const duplicates = vi.fn().mockResolvedValueOnce(page(items(0, 500), "0-499/501")).mockResolvedValueOnce(page([{ id: "0" }], "500-500/501"));
    await expect(auditReportRows(duplicates)).rejects.toThrow("changed");
    await expect(auditReportRows(async () => page([], "*/10001"))).rejects.toThrow("10,000");
  });
});
