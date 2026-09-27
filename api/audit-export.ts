import type { VercelRequest, VercelResponse } from "@vercel/node";
import { buildAuditPdf, type AuditPdfMode } from "../lib/auditPdf";

const origin = "https://qhubio.com";
// This is the browser-safe publishable key, not a service-role secret. Every
// data and storage request below carries the caller's JWT, so Supabase RLS is
// the authorization boundary for a report export.
const supabaseUrl = "https://faggswvpcctwlyuyllnc.supabase.co";
const supabasePublishableKey = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZhZ2dzd3ZwY2N0d2x5dXlsbG5jIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzE0MTAzNDYsImV4cCI6MjA4Njk4NjM0Nn0.ughHxVaf-7YEFzi9Y6AqeAlXhHg1HDQmeeuzY4PMIUs";
const json = (res: VercelResponse, status: number, error: string) => res.status(status).json({ success: false, error });
const headers = (res: VercelResponse) => { res.setHeader("Access-Control-Allow-Origin", origin); res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS"); res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization"); };
type TemporaryLogo = { mediaType?: string; byteSize?: number; width?: number; height?: number; base64?: string };

function decodeTemporaryLogo(value: TemporaryLogo | undefined): { bytes: Uint8Array; format: "PNG" | "JPEG" } | null {
  if (!value) return null;
  if (!['image/png', 'image/jpeg'].includes(value.mediaType ?? '') || !Number.isInteger(value.byteSize) || !Number.isInteger(value.width) || !Number.isInteger(value.height) || value.byteSize! <= 0 || value.byteSize! > 2 * 1024 * 1024 || value.width! <= 0 || value.height! <= 0 || value.width! > 4096 || value.height! > 4096 || typeof value.base64 !== 'string' || value.base64.length > 2_800_000 || !/^[A-Za-z0-9+/]+={0,2}$/.test(value.base64)) throw new Error('Temporary logo does not meet the PNG/JPEG, 2 MB and 4096 px limits.');
  const bytes = new Uint8Array(Buffer.from(value.base64, 'base64'));
  if (bytes.byteLength !== value.byteSize) throw new Error('Temporary logo data is invalid.');
  const png = value.mediaType === 'image/png';
  if (png) {
    if (bytes.byteLength < 24 || ![137, 80, 78, 71, 13, 10, 26, 10].every((byte, index) => bytes[index] === byte)) throw new Error('Temporary logo is not a valid PNG.');
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength); if (view.getUint32(16) !== value.width || view.getUint32(20) !== value.height) throw new Error('Temporary logo dimensions are invalid.');
  } else {
    if (bytes.byteLength < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) throw new Error('Temporary logo is not a valid JPEG.');
    let offset = 2, matched = false;
    while (offset + 9 < bytes.length) { if (bytes[offset] !== 0xff) { offset++; continue; } const marker = bytes[offset + 1]; const length = (bytes[offset + 2] << 8) + bytes[offset + 3]; if (length < 2 || offset + 2 + length > bytes.length) break; if ([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(marker)) { const height = (bytes[offset + 5] << 8) + bytes[offset + 6], width = (bytes[offset + 7] << 8) + bytes[offset + 8]; matched = width === value.width && height === value.height; break; } offset += 2 + length; }
    if (!matched) throw new Error('Temporary logo dimensions are invalid.');
  }
  return { bytes, format: png ? 'PNG' : 'JPEG' };
}

export default async function handler(req: VercelRequest, res: VercelResponse) {
  headers(res); if (req.method === "OPTIONS") return res.status(204).end(); if (req.method !== "POST") return json(res, 405, "Method not allowed");
  try {
    const authorization = req.headers.authorization;
    if (!authorization) return json(res, 401, "Sign in is required.");
    const { auditId, format: requestedFormat, temporaryLogo, useQhubioBranding } = (req.body ?? {}) as { auditId?: string; format?: AuditPdfMode; temporaryLogo?: TemporaryLogo; useQhubioBranding?: boolean };
    if (!auditId || !["executive", "detailed"].includes(requestedFormat ?? "")) return json(res, 400, "A valid audit and report format are required.");
    const format = requestedFormat as AuditPdfMode;
    const userResponse = await fetch(`${supabaseUrl}/auth/v1/user`, { headers: { apikey: supabasePublishableKey, Authorization: authorization } }); if (!userResponse.ok) return json(res, 401, "Your session is invalid.");
    const user = await userResponse.json() as { id: string };
    const orgResponse = await fetch(`${supabaseUrl}/rest/v1/rpc/get_user_organization_id`, { method: "POST", headers: { apikey: supabasePublishableKey, Authorization: authorization, "Content-Type": "application/json" }, body: JSON.stringify({ _user_id: user.id }) });
    if (!orgResponse.ok) return json(res, 403, "No organisation access.");
    const organizationId = String(await orgResponse.json() ?? ""); if (!organizationId) return json(res, 403, "No organisation access.");
    const rest = (path: string) => fetch(`${supabaseUrl}/rest/v1/${path}`, { headers: { apikey: supabasePublishableKey, Authorization: authorization } });
    const auditResponse = await rest(`audit_v2_runs?select=*&id=eq.${encodeURIComponent(auditId)}&organization_id=eq.${encodeURIComponent(organizationId)}`);
    if (!auditResponse.ok) return json(res, 502, "The audit record could not be retrieved.");
    const auditRows = await auditResponse.json() as Array<Record<string, any>>; const audit = auditRows?.[0]; if (!audit) return json(res, 404, "Audit not found.");
    const [evidence, links, history, approvals, brandingRows] = await Promise.all([rest(`audit_v2_evidence?select=*&audit_id=eq.${encodeURIComponent(auditId)}`), rest(`audit_v2_capa_links?select=*&audit_id=eq.${encodeURIComponent(auditId)}`), rest(`audit_v2_events?select=*&audit_id=eq.${encodeURIComponent(auditId)}&order=created_at.desc`), rest(`audit_v2_approval_events?select=*&audit_id=eq.${encodeURIComponent(auditId)}&order=created_at.asc`), rest(`audit_v2_report_branding?select=*&organization_id=eq.${encodeURIComponent(organizationId)}`)]);
    const branding = (await brandingRows.json() as Array<{ logo_path?: string; logo_media_type?: string }>)?.[0]; let logo: { bytes: Uint8Array; format: "PNG" | "JPEG" } | null = decodeTemporaryLogo(temporaryLogo);
    if (!logo && !useQhubioBranding && branding?.logo_path) { const file = await fetch(`${supabaseUrl}/storage/v1/object/authenticated/audit-v2-report-branding/${branding.logo_path}`, { headers: { apikey: supabasePublishableKey, Authorization: authorization } }); if (file.ok) logo = { bytes: new Uint8Array(await file.arrayBuffer()), format: branding.logo_media_type === "image/png" ? "PNG" : "JPEG" }; }
    const buffer = buildAuditPdf({ audit, evidence: await evidence.json() as Array<Record<string, unknown>>, links: await links.json() as Array<Record<string, unknown>>, history: await history.json() as Array<Record<string, unknown>>, approvals: await approvals.json() as Array<Record<string, unknown>>, branding: logo, mode: format }); if (buffer.byteLength > 4 * 1024 * 1024) return json(res, 413, "This report is too large to export safely. Reduce the report scope and try again.");
    const name = String(audit.document?.name ?? "audit").replace(/[^a-z0-9]+/gi, "-").replace(/^-|-$/g, "").toLowerCase(); res.setHeader("Content-Type", "application/pdf"); res.setHeader("Content-Disposition", `attachment; filename=\"${name}-${format}.pdf\"`); res.setHeader("Content-Length", String(buffer.byteLength)); return res.status(200).send(Buffer.from(buffer));
  } catch (error) { console.error("[audit-export] failed", error); return json(res, 500, "Audit report generation failed. Please try again."); }
}


