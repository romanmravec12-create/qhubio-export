// @vitest-environment node
import { describe, expect, it, vi, afterEach } from "vitest";
import { readFileSync } from "node:fs";
import { inflateSync } from "node:zlib";
import { Writable } from "node:stream";
import { buildEightDReportPdf, D2_PICTURE_MAX_MM, reportRcaMethod, ReportInputError } from "../lib/eightDPdf";
import handler from "../api/export-8d-report";
import { reportFixture } from "./eightDFixture";

afterEach(() => { vi.unstubAllGlobals(); vi.unstubAllEnvs(); });
const response = () => {
  const chunks: Buffer[] = [];
  const r = Object.assign(new Writable({ write(chunk, _encoding, done) { chunks.push(Buffer.from(chunk)); done(); } }), { status: vi.fn(), setHeader: vi.fn(), json: vi.fn(), send: vi.fn(), pdf: () => Buffer.concat(chunks) });
  r.status.mockReturnValue(r); return r;
};
const request = (body: unknown = { caseId: reportFixture().caseRow.id, branding: { type: "qhubio" } }, authorization = "Bearer test-token") => ({ method: "POST", headers: { authorization, origin: "https://qhubio.com" }, body });
const configure = () => { vi.stubEnv("SUPABASE_URL", "https://example.supabase.co"); vi.stubEnv("SUPABASE_ANON_KEY", "test-anon"); };
const backend = (fixture = reportFixture()) => vi.fn()
  .mockResolvedValueOnce({ ok: true, json: async () => [fixture.caseRow] })
  .mockResolvedValueOnce({ ok: true, json: async () => fixture.sections });
const webp = Buffer.from("UklGRj4AAABXRUJQVlA4IDIAAADQAgCdASogABQAPmEuk0akIqGhKAgAgAwJZQAAPaOgAP7uQx//+Ic/ly/iCw1CXgAAAA==", "base64");
const pictureFiles = (count = 5) => Array.from({ length: count }, (_, n) => ({ file_name: `Evidence-${n + 1}.webp`, file_type: "image/webp", storage_path: `org-id/${reportFixture().caseRow.id}/picture-${n + 1}.webp` }));
const pictureRequest = () => request({ caseId: reportFixture().caseRow.id, branding: { type: "qhubio" }, includeD2Pictures: true });
const contentText = (pdf: Buffer) => [...pdf.toString("latin1").matchAll(/stream\r?\n([\s\S]*?)\r?\nendstream/g)].map(match => {
  try { return inflateSync(Buffer.from(match[1], "latin1")).toString("latin1"); } catch { return match[1]; }
}).join("\n");

describe("8D portrait report", () => {
  it.each(["ishikawa", "five_why", "fault_tree"] as const)("renders %s as an A3 portrait PDF", method => {
    const pdf = buildEightDReportPdf(reportFixture(method));
    expect(pdf.subarray(0, 5).toString()).toBe("%PDF-");
    const box = pdf.toString("latin1").match(/MediaBox \[0 0 ([\d.]+) ([\d.]+)\]/)!;
    expect(Number(box[1])).toBeCloseTo(841.89, 1); expect(Number(box[2])).toBeCloseTo(1190.55, 1);
  });
  it("follows the same legacy single-method precedence as D4", () => {
    expect(reportRcaMethod({})).toBe("ishikawa");
    expect(reportRcaMethod({ d4_rca_methods: { ishikawa: true, five_why: true } })).toBe("ishikawa");
    expect(reportRcaMethod({ d4_rca_methods: { ishikawa: false, five_why: true } })).toBe("five_why");
    expect(reportRcaMethod({ d4_rca_methods: { ishikawa: true, fault_tree: true } })).toBe("fault_tree");
  });
  it("adds continuation pages for unbounded action lists and long text", () => {
    const fixture = reportFixture();
    fixture.sections.find(s => s.step_code === "D5")!.content = { actions: Array.from({ length: 30 }, (_, n) => ({ permanent_action: `Action ${n} ` + "Full retained action text. ".repeat(100), root_cause: "Cause", verification: "Verify", effectiveness_measure: "Measure" })) };
    const pdf = buildEightDReportPdf(fixture);
    expect((pdf.toString("latin1").match(/\/Type \/Page\b/g) || []).length).toBeGreaterThan(3);
    expect(contentText(pdf)).toContain("Action 29");
  });
  it("retains the actual D6 validation, D7 documents and D8 approvals", () => {
    const text = contentText(buildEightDReportPdf(reportFixture()));
    for (const expected of ["Validation record VR-030", "WI-WH-07 and process FMEA", "Example service approver", "Stock reconciliation record SR-014", "Bin count BC-011"]) expect(text).toContain(expected);
    expect(text).toContain("CLOSED"); expect(text).not.toContain("APPROVED");
  });
  it("excludes inactive RCA workspaces and never promotes an unverified candidate", () => {
    const fixture = reportFixture("five_why");
    const d4 = fixture.sections.find(s => s.step_code === "D4")!.content as Record<string, unknown>;
    d4.root_cause_verification = { status: "unverified", evidence: [] };
    const text = contentText(buildEightDReportPdf(fixture));
    expect(text).toContain("NOT VERIFIED"); expect(text).not.toContain("Scanner connectivity"); expect(text).not.toContain("Stale top event");
  });
  it("retains causes beyond the compact diagram and FTA branches beyond seven", () => {
    const fixture = reportFixture(); const d4 = fixture.sections.find(s => s.step_code === "D4")!.content as Record<string, unknown>;
    d4.ishikawa = { material: Array.from({ length: 20 }, (_, i) => `Material-cause-${i + 1}`) };
    expect(contentText(buildEightDReportPdf(fixture))).toContain("Material-cause-20");
    const fta = reportFixture("fault_tree"); (fta.sections.find(s => s.step_code === "D4")!.content as Record<string, unknown>).fault_tree = { root: { text: "Old problem", type: "top", gate: "or", children: Array.from({ length: 20 }, (_, i) => ({ text: `FTA-event-${i + 1}`, type: "basic", children: [] })) } };
    const text = contentText(buildEightDReportPdf(fta)); expect(text).toContain("FTA-event-20"); expect(text).not.toContain("Old problem");
  });
  it("accepts custom PNG and JPEG branding and rejects corrupt logos", () => {
    for (const format of ["PNG", "JPEG"] as const) {
      const fixture = reportFixture();
      fixture.branding = { bytes: readFileSync(`assets/qhubio-logo.${format === "PNG" ? "png" : "jpg"}`), format };
      expect(buildEightDReportPdf(fixture).subarray(0, 5).toString()).toBe("%PDF-");
    }
    const fixture = reportFixture(); fixture.branding = { bytes: Buffer.from("broken"), format: "PNG" };
    expect(() => buildEightDReportPdf(fixture)).toThrow(ReportInputError);
  });
  it("fails explicitly rather than producing missing glyphs", () => {
    const fixture = reportFixture(); fixture.sections.find(s => s.step_code === "D8")!.content = { closure_summary: "Unsupported emoji 🚀" };
    expect(() => buildEightDReportPdf(fixture)).toThrow(/No partial report/);
  });
  it("includes all five WEBP reference pictures inside D2, before D3, with no appendix", () => {
    const fixture = reportFixture();
    fixture.pictures = pictureFiles().map(file => ({ fileName: file.file_name, format: "WEBP", bytes: webp }));
    const text = contentText(buildEightDReportPdf(fixture));
    expect(text).toContain("VISUAL REFERENCE / 5 PICTURES");
    for (let n = 1; n <= 5; n++) { expect(text).toContain(`PICTURE ${n}`); expect(text).toContain(`Evidence-${n}.webp`); }
    expect(text).toContain("not AI inputs or proof of root cause");
    expect(text).not.toContain("appendix");
    expect(text.indexOf("PICTURE 5")).toBeLessThan(text.indexOf("Interim containment / customer protection"));
    expect(contentText(buildEightDReportPdf(reportFixture()))).not.toContain("VISUAL REFERENCE");
  });
  it.each([1, 5])("keeps %i reference pictures within the same fixed thumbnail bounds", count => {
    const fixture = reportFixture();
    fixture.pictures = pictureFiles(count).map(file => ({ fileName: file.file_name, format: "WEBP", bytes: webp }));
    const text = contentText(buildEightDReportPdf(fixture));
    const images = [...text.matchAll(/([\d.]+) 0 0 ([\d.]+) ([\d.]+) ([\d.]+) cm\s*\/I(\d+) Do/g)]
      .filter(match => Number(match[5]) > 0); // I0 is the repeated branding logo.
    expect(images).toHaveLength(count);
    const boundPoints = D2_PICTURE_MAX_MM * 72 / 25.4;
    for (const image of images) {
      expect(Number(image[1])).toBeLessThanOrEqual(boundPoints + .01);
      expect(Number(image[2])).toBeLessThanOrEqual(boundPoints + .01);
    }
    expect(new Set(images.map(image => image[4])).size).toBe(1);
  });
  it("keeps the gallery inside continued D2 when the problem description is long", () => {
    const fixture = reportFixture();
    const d2 = fixture.sections.find(section => section.step_code === "D2")!.content as Record<string, unknown>;
    d2.summary = "Documented problem detail. ".repeat(1000) + "END-OF-D2-PROBLEM";
    fixture.pictures = pictureFiles().map(file => ({ fileName: file.file_name, format: "WEBP", bytes: webp }));
    const text = contentText(buildEightDReportPdf(fixture));
    expect(text).toContain("END-OF-D2-PROBLEM");
    expect(text.includes("Problem definition \\(continued\\)")).toBe(true);
    expect(text.indexOf("END-OF-D2-PROBLEM")).toBeLessThan(text.indexOf("PICTURE 1"));
    expect(text.indexOf("PICTURE 5")).toBeLessThan(text.indexOf("Interim containment / customer protection"));
    expect(text).not.toContain("appendix");
  });
  it("does not silently omit a sixth or unreadable picture", () => {
    const fixture = reportFixture();
    fixture.pictures = pictureFiles(6).map(file => ({ fileName: file.file_name, format: "WEBP", bytes: webp }));
    expect(() => buildEightDReportPdf(fixture)).toThrow(/maximum of 5/);
    fixture.pictures = [{ fileName: "bad.webp", format: "WEBP", bytes: Buffer.from("broken") }];
    expect(() => buildEightDReportPdf(fixture)).toThrow(/No partial report/);
  });
});

describe("authenticated 8D export API", () => {
  it("reads saved records using caller authorization and the anon key (RLS), never a service role", async () => {
    configure(); const fetcher = backend(); vi.stubGlobal("fetch", fetcher); const res = response();
    await handler(request() as never, res as never);
    expect(res.status).toHaveBeenCalledWith(200);
    expect(res.setHeader).toHaveBeenCalledWith("Cache-Control", "no-store");
    expect(fetcher).toHaveBeenCalledTimes(2);
    expect(fetcher.mock.calls[0][1].headers).toEqual({ apikey: "test-anon", Authorization: "Bearer test-token" });
    expect(res.pdf().subarray(0, 5).toString()).toBe("%PDF-");
    expect(res.send).not.toHaveBeenCalled(); // Streamed, not a Vercel buffered response.
    expect(res.writableFinished).toBe(true);
  });
  it("denies a case invisible to the caller", async () => {
    configure(); vi.stubGlobal("fetch", vi.fn().mockResolvedValue({ ok: true, json: async () => [] })); const res = response();
    await handler(request() as never, res as never); expect(res.status).toHaveBeenCalledWith(404); expect(res.send).not.toHaveBeenCalled();
  });
  it("requires all eight saved approvals", async () => {
    configure(); const fixture = reportFixture(); fixture.sections[5].approval_status = "draft"; vi.stubGlobal("fetch", backend(fixture)); const res = response();
    await handler(request() as never, res as never); expect(res.status).toHaveBeenCalledWith(409); expect(res.send).not.toHaveBeenCalled();
  });
  it("requires a bearer token without accessing the database", async () => {
    configure(); const fetcher = vi.fn(); vi.stubGlobal("fetch", fetcher); const res = response();
    await handler(request(undefined, "") as never, res as never); expect(res.status).toHaveBeenCalledWith(401); expect(fetcher).not.toHaveBeenCalled();
  });
  it("rejects malformed request and RCA overrides", async () => {
    const res = response(); await handler(request({ caseId: reportFixture().caseRow.id, branding: { type: "qhubio" }, method: "fault_tree" }) as never, res as never);
    expect(res.status).toHaveBeenCalledWith(400);
  });
  it("returns an actionable failure for a corrupt logo instead of dropping it", async () => {
    configure(); vi.stubGlobal("fetch", backend()); const res = response();
    await handler(request({ caseId: reportFixture().caseRow.id, branding: { type: "custom", imageData: "data:image/png;base64,YnJva2Vu", format: "PNG", width: 10, height: 10 } }) as never, res as never);
    expect(res.status).toHaveBeenCalledWith(422); expect(res.send).not.toHaveBeenCalled();
  });
  it("loads all five pictures from private case-scoped Storage using the caller JWT", async () => {
    configure(); const fetcher = backend().mockResolvedValueOnce({ ok: true, json: async () => pictureFiles() }).mockImplementation(async () => new Response(webp));
    vi.stubGlobal("fetch", fetcher); const res = response();
    await handler(pictureRequest() as never, res as never);
    expect(res.status).toHaveBeenCalledWith(200); expect(fetcher).toHaveBeenCalledTimes(8);
    expect(fetcher.mock.calls[2][0]).toContain(`case_id=eq.${reportFixture().caseRow.id}&file_category=eq.image`);
    for (const [url, options] of fetcher.mock.calls.slice(3)) {
      expect(url).toContain(`/storage/v1/object/authenticated/eight_d_files/org-id/${reportFixture().caseRow.id}/`);
      expect(options.headers).toEqual({ apikey: "test-anon", Authorization: "Bearer test-token" });
      expect(options.redirect).toBe("error");
    }
    expect(contentText(res.pdf())).toContain("PICTURE 5");
  });
  it("streams a picture-heavy PDF above Vercel's 4.5 MB buffered limit", async () => {
    configure(); const files = pictureFiles(2).map(file => ({ ...file, file_type: "image/jpeg" }));
    const jpeg = readFileSync("assets/qhubio-logo.jpg");
    let n = 0;
    const fetcher = backend().mockResolvedValueOnce({ ok: true, json: async () => files }).mockImplementation(async () => new Response(Buffer.concat([jpeg, Buffer.alloc(2_600_000, ++n)])));
    vi.stubGlobal("fetch", fetcher); const res = response();
    await handler(pictureRequest() as never, res as never);
    expect(res.status).toHaveBeenCalledWith(200);
    expect(res.pdf().length).toBeGreaterThan(4_500_000);
    expect(res.pdf().subarray(0, 5).toString()).toBe("%PDF-");
    expect(res.send).not.toHaveBeenCalled(); expect(res.writableFinished).toBe(true);
    expect(res.setHeader.mock.calls.map(call => call[0])).not.toContain("Content-Length");
  });
  it.each([0, 6])("fails on %i pictures without downloading or omitting anything", async count => {
    configure(); const fetcher = backend().mockResolvedValueOnce({ ok: true, json: async () => pictureFiles(count) });
    vi.stubGlobal("fetch", fetcher); const res = response();
    await handler(pictureRequest() as never, res as never);
    expect(res.status).toHaveBeenCalledWith(422); expect(fetcher).toHaveBeenCalledTimes(3); expect(res.pdf().length).toBe(0);
  });
  it("rejects a picture reference belonging to another case", async () => {
    configure(); const files = pictureFiles(1); files[0].storage_path = "org-id/another-case/picture.webp";
    const fetcher = backend().mockResolvedValueOnce({ ok: true, json: async () => files });
    vi.stubGlobal("fetch", fetcher); const res = response();
    await handler(pictureRequest() as never, res as never);
    expect(res.status).toHaveBeenCalledWith(422); expect(fetcher).toHaveBeenCalledTimes(3);
  });
  it.each(["denied", "corrupt", "oversized"])("does not send a partial report for a %s picture", async failure => {
    configure(); const fetcher = backend().mockResolvedValueOnce({ ok: true, json: async () => pictureFiles(1) }).mockImplementation(async () => failure === "denied" ? new Response("Denied", { status: 403 }) : failure === "oversized" ? new Response(webp, { headers: { "content-length": String(5 * 1024 * 1024 + 1) } }) : new Response("broken"));
    vi.stubGlobal("fetch", fetcher); const res = response();
    await handler(pictureRequest() as never, res as never);
    expect(res.status).toHaveBeenCalledWith(422); expect(res.pdf().length).toBe(0);
  });
});
