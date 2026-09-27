import type { VercelRequest, VercelResponse } from "@vercel/node";
import { buildAuditPortfolioPdf } from "../lib/auditPortfolioPdf";

const origin = "https://qhubio.com";
const supabaseUrl = "https://faggswvpcctwlyuyllnc.supabase.co";
const supabasePublishableKey = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZhZ2dzd3ZwY2N0d2x5dXlsbG5jIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzE0MTAzNDYsImV4cCI6MjA4Njk4NjM0Nn0.ughHxVaf-7YEFzi9Y6AqeAlXhHg1HDQmeeuzY4PMIUs";
const headers = (res: VercelResponse) => { res.setHeader("Access-Control-Allow-Origin", origin); res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS"); res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization"); };
const fail = (res: VercelResponse, status: number, error: string) => res.status(status).json({ success: false, error });
const csv = (value: unknown) => `"${String(value ?? "").replace(/"/g, '""').replace(/[\r\n]+/g, " ")}"`;

export default async function handler(req: VercelRequest, res: VercelResponse) {
  headers(res); if (req.method === "OPTIONS") return res.status(204).end(); if (req.method !== "POST") return fail(res, 405, "Method not allowed");
  try {
    const authorization = req.headers.authorization;
    const { auditIds, format, includeDrafts = false, reportType = "portfolio" } = (req.body ?? {}) as { auditIds?: unknown; format?: "pdf" | "csv"; includeDrafts?: boolean; reportType?: "portfolio" | "supplier_scorecard" };
    if (!authorization) return fail(res, 401, "Sign in is required.");
    if (!Array.isArray(auditIds) || auditIds.length < 1 || auditIds.length > 50 || !auditIds.every((id) => typeof id === "string" && /^[0-9a-f-]{36}$/i.test(id)) || !["pdf", "csv"].includes(format ?? "") || !["portfolio", "supplier_scorecard"].includes(reportType)) return fail(res, 400, "Choose between one and fifty valid audits and an export format.");
    const userResponse = await fetch(`${supabaseUrl}/auth/v1/user`, { headers: { apikey: supabasePublishableKey, Authorization: authorization } });
    if (!userResponse.ok) return fail(res, 401, "Your session is invalid.");
    const user = await userResponse.json() as { id: string };
    const orgResponse = await fetch(`${supabaseUrl}/rest/v1/rpc/get_user_organization_id`, { method: "POST", headers: { apikey: supabasePublishableKey, Authorization: authorization, "Content-Type": "application/json" }, body: JSON.stringify({ _user_id: user.id }) });
    if (!orgResponse.ok) return fail(res, 403, "No organisation access.");
    const organizationId = String(await orgResponse.json() ?? ""); if (!organizationId) return fail(res, 403, "No organisation access.");
    const ids = auditIds.join(",");
    const rowsResponse = await fetch(`${supabaseUrl}/rest/v1/audit_v2_runs?select=*&organization_id=eq.${encodeURIComponent(organizationId)}&id=in.(${encodeURIComponent(ids)})`, { headers: { apikey: supabasePublishableKey, Authorization: authorization } });
    if (!rowsResponse.ok) return fail(res, 502, "The selected audits could not be retrieved.");
    const rows = await rowsResponse.json() as Array<Record<string, any>>;
    if (rows.length !== auditIds.length) return fail(res, 404, "One or more selected audits are unavailable.");
    if (!includeDrafts && rows.some((row) => !row.issued_snapshot)) return fail(res, 400, "Portfolio exports include issued audits only unless draft inclusion is explicitly selected.");
    const generatedAt = new Date().toISOString();
    if (format === "csv") {
      const lines = [["audit_id","revision","record_state","supplier_id","supplier_site_id","audit_name","audited_organization","site","planned_date","template","finding_id","question_id","finding_title","finding_type","severity","finding_status","owner","containment_due","action_plan_due","due_date","containment","root_cause","action","implementation","effectiveness","verified_by","verified_on","generated_at"]];
      for (const row of rows) {
        const document = row.issued_snapshot ?? row.document;
        const findings = document.findings ?? [];
        const base = [row.id,row.revision,row.issued_snapshot ? "ISSUED" : "DRAFT",row.supplier_id,row.supplier_site_id,document.name,document.organization,document.site,document.plannedOn,document.template?.name];
        if (!findings.length) lines.push([...base,"","","","","","","","","","","","","","","","","",generatedAt]);
        for (const finding of findings) lines.push([...base,finding.id,finding.questionId,finding.title,finding.type,finding.severity,finding.status,finding.owner,finding.containmentDue,finding.planDue,finding.due,finding.containment,finding.rootCause,finding.action,finding.implementation,finding.effectiveness,finding.verifiedBy,finding.verifiedOn,generatedAt]);
      }
      res.setHeader("Content-Type", "text/csv; charset=utf-8"); res.setHeader("Content-Disposition", 'attachment; filename="qhubio-audit-portfolio.csv"'); return res.status(200).send(lines.map((line) => line.map(csv).join(",")).join("\n"));
    }
    if (reportType === "supplier_scorecard" && rows.some((row) => row.document?.template?.family !== "supplier")) return fail(res, 400, "Supplier scorecard reports can include supplier audits only.");
    const pdf = buildAuditPortfolioPdf({ audits: rows, generatedAt, includeDrafts, reportType });
    if (pdf.byteLength > 4 * 1024 * 1024) return fail(res, 413, "This portfolio is too large to export safely. Narrow the selection and try again.");
    res.setHeader("Content-Type", "application/pdf"); res.setHeader("Content-Disposition", `attachment; filename="qhubio-${reportType === "supplier_scorecard" ? "supplier-scorecards" : "audit-portfolio"}.pdf"`); res.setHeader("Content-Length", String(pdf.byteLength)); return res.status(200).send(Buffer.from(pdf));
  } catch (error) { console.error("[audit-portfolio-export] failed", error); return fail(res, 500, "Portfolio export failed. Please try again."); }
}


