import type { VercelRequest, VercelResponse } from "@vercel/node";
import { Readable } from "node:stream";
import { pipeline } from "node:stream/promises";
import { z } from "zod";
import { buildEightDReportPdf, ReportInputError, type EightDReportBranding, type EightDReportPicture } from "../lib/eightDPdf";

export const config = { runtime: "nodejs" } as const;
const origin = "https://qhubio.com";
const allowedOrigins = new Set([origin, "https://www.qhubio.com"]);
const requestSchema = z.object({
  caseId: z.string().uuid(),
  includeD2Pictures: z.boolean().optional().default(false),
  branding: z.union([
    z.object({ type: z.literal("qhubio") }),
    z.object({ type: z.literal("custom"), imageData: z.string().max(2_800_000), format: z.enum(["PNG", "JPEG"]), width: z.number().int().min(1).max(4096), height: z.number().int().min(1).max(4096) }),
  ]),
}).strict();

function cors(req: VercelRequest, res: VercelResponse) {
  const value = req.headers.origin;
  res.setHeader("Access-Control-Allow-Origin", value && allowedOrigins.has(value) ? value : origin);
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Vary", "Origin");
}
function env(name: string) { return process.env[name]?.trim().replace(/^["']|["']$/g, ""); }
function bodyOf(value: unknown) {
  if (typeof value === "string") { try { return JSON.parse(value); } catch { return value; } }
  if (Buffer.isBuffer(value)) { try { return JSON.parse(value.toString("utf8")); } catch { return value; } }
  if (value && typeof value === "object" && "type" in value && (value as { type?: unknown }).type === "Buffer" && Array.isArray((value as unknown as { data?: unknown }).data)) { try { return JSON.parse(Buffer.from((value as unknown as { data: number[] }).data).toString("utf8")); } catch { return value; } }
  return value;
}
function customBranding(value: z.infer<typeof requestSchema>["branding"]): EightDReportBranding {
  if (value.type !== "custom") return null;
  const match = value.imageData.match(/^data:image\/(png|jpeg);base64,([A-Za-z0-9+/=]+)$/i);
  if (!match || (match[1].toLowerCase() === "png" ? "PNG" : "JPEG") !== value.format) throw new ReportInputError("The company logo is invalid.");
  return { bytes: Buffer.from(match[2], "base64"), format: match[1].toLowerCase() === "png" ? "PNG" : "JPEG" };
}

async function readPicture(response: Response): Promise<Uint8Array> {
  const limit = 5 * 1024 * 1024;
  if (!response.ok || !response.body || Number(response.headers.get("content-length") || 0) > limit) throw new ReportInputError("A D2 picture is unavailable or too large. No partial report was exported.");
  const reader = response.body.getReader(), chunks: Uint8Array[] = []; let total = 0;
  try {
    for (;;) {
      const { done, value } = await reader.read(); if (done) break;
      total += value.length;
      if (total > limit) { await reader.cancel(); throw new ReportInputError("A D2 picture exceeds the existing 5 MB limit. No partial report was exported."); }
      chunks.push(value);
    }
  } finally { reader.releaseLock(); }
  return Buffer.concat(chunks, total);
}

async function loadPictures(base: string, caseId: string, headers: Record<string, string>): Promise<EightDReportPicture[]> {
  const filesResponse = await fetch(`${base}/rest/v1/eight_d_files?select=file_name,file_type,storage_path&case_id=eq.${encodeURIComponent(caseId)}&file_category=eq.image&order=created_at.asc,id.asc`, { headers, signal: AbortSignal.timeout(15_000) });
  if (!filesResponse.ok) throw new ReportInputError("The D2 pictures could not be read. Refresh the case and retry.");
  const files = await filesResponse.json() as Array<{ file_name?: string; file_type?: string; storage_path?: string }>;
  if (!Array.isArray(files) || !files.length) throw new ReportInputError("No D2 pictures are available. Refresh the case or turn off picture inclusion.");
  if (files.length > 5) throw new ReportInputError("D2 supports a maximum of 5 pictures. No pictures were omitted.");
  return Promise.all(files.map(async file => {
    const parts = file.storage_path?.split("/");
    if (!parts || parts.length !== 3 || parts[1] !== caseId || parts.some(part => !part || part === "." || part === ".." || part.includes("\\"))) throw new ReportInputError("A D2 picture has an invalid storage reference. Review it in D2 and retry.");
    const format = file.file_type === "image/webp" ? "WEBP" : file.file_type === "image/png" ? "PNG" : file.file_type === "image/jpeg" ? "JPEG" : null;
    if (!format) throw new ReportInputError("A D2 picture has an unsupported image format. Review it in D2 and retry.");
    // Private Storage download under the caller JWT; never a public URL or service role.
    const response = await fetch(`${base}/storage/v1/object/authenticated/eight_d_files/${parts.map(encodeURIComponent).join("/")}`, { headers, signal: AbortSignal.timeout(15_000), redirect: "error" });
    return { fileName: file.file_name || parts[2], format, bytes: await readPicture(response) };
  }));
}

export default async function handler(req: VercelRequest, res: VercelResponse) {
  cors(req, res);
  if (req.method === "OPTIONS") return res.status(204).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });
  if (Number(req.headers["content-length"] ?? 0) > 3_000_000) return res.status(413).json({ error: "Export request is too large" });
  const parsed = requestSchema.safeParse(bodyOf(req.body));
  if (!parsed.success) return res.status(400).json({ error: "Invalid export request" });
  const authorization = req.headers.authorization;
  const base = env("SUPABASE_URL");
  const anon = env("SUPABASE_ANON_KEY");
  if (!authorization?.startsWith("Bearer ")) return res.status(401).json({ error: "Sign in to export this report." });
  if (!base || !anon) return res.status(503).json({ error: "The report service is not configured." });
  const headers = { apikey: anon, Authorization: authorization };
  try {
    const caseUrl = `${base.replace(/\/$/, "")}/rest/v1/eight_d_cases?select=*&id=eq.${encodeURIComponent(parsed.data.caseId)}&limit=1`;
    const caseResponse = await fetch(caseUrl, { headers });
    if (!caseResponse.ok) return res.status(403).json({ error: "This account cannot export the requested 8D report." });
    const cases = await caseResponse.json() as Record<string, unknown>[];
    const caseRow = cases[0];
    if (!caseRow) return res.status(404).json({ error: "8D report not found." });
    const sectionsResponse = await fetch(`${base.replace(/\/$/, "")}/rest/v1/eight_d_sections?select=step_code,content,approval_status&case_id=eq.${encodeURIComponent(parsed.data.caseId)}`, { headers });
    if (!sectionsResponse.ok) return res.status(403).json({ error: "This account cannot read this 8D report." });
    const sections = await sectionsResponse.json() as Array<Record<string, unknown>>;
    const requiredSteps = ["D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8"];
    if (!requiredSteps.every((step) => sections.some((item) => item.step_code === step && item.approval_status === "approved"))) {
      return res.status(409).json({ error: "Approve all eight disciplines before exporting the final report." });
    }
    const pictures = parsed.data.includeD2Pictures ? await loadPictures(base.replace(/\/$/, ""), parsed.data.caseId, headers) : [];
    const pdf = buildEightDReportPdf({ caseRow, sections, branding: customBranding(parsed.data.branding), pictures });
    const safeName = String(caseRow.title || "8D_Case").replace(/[^a-z0-9_-]+/gi, "_").slice(0, 60) || "8D_Case";
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${safeName}_8D_Report.pdf"`);
    res.setHeader("Cache-Control", "no-store");
    // Streaming avoids Vercel's buffered-response limit for five large pictures.
    // Generate/validate the complete PDF before sending any bytes: never a partial report.
    res.status(200);
    await pipeline(Readable.from((function* () {
      for (let offset = 0; offset < pdf.length; offset += 64 * 1024) yield pdf.subarray(offset, offset + 64 * 1024);
    })()), res);
  } catch (error) {
    if (res.headersSent) { res.destroy(error instanceof Error ? error : new Error("Report download interrupted")); return; }
    if (error instanceof ReportInputError) return res.status(422).json({ error: (error as ReportInputError).message });
    console.warn("[export-8d-report] generation_failed", error instanceof Error ? error.message : "unknown");
    return res.status(502).json({ error: "The report service could not create this PDF." });
  }
}
