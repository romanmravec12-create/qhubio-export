export class AuditReportDataError extends Error {}

/** Read a complete ordered PostgREST result, or fail instead of exporting a partial report. */
export async function auditReportRows(fetchPage: (from: number, to: number) => Promise<Response>): Promise<Array<Record<string, unknown>>> {
  const pageSize = 500;
  const maxRows = 10_000;
  const rows: Array<Record<string, unknown>> = [];
  const seen = new Set<string>();
  let expectedTotal: number | undefined;
  for (let from = 0; from <= maxRows; from += pageSize) {
    const response = await fetchPage(from, from + pageSize - 1);
    if (!response.ok) throw new AuditReportDataError("Report data could not be retrieved. Please try again.");
    const range = response.headers.get("content-range");
    const match = range?.match(/^(?:(\d+)-(\d+)|\*)\/(\d+)$/);
    if (!match) throw new AuditReportDataError("Report data completeness could not be verified. Please try again.");
    const total = Number(match[3]);
    if (!Number.isSafeInteger(total)) throw new AuditReportDataError("Report data completeness could not be verified. Please try again.");
    if (total > maxRows) throw new AuditReportDataError("This report contains more than 10,000 records in one category and cannot be exported safely.");
    if (expectedTotal !== undefined && expectedTotal !== total) throw new AuditReportDataError("Report data changed during export. Please try again.");
    expectedTotal = total;
    const page = await response.json();
    if (!Array.isArray(page) || page.length > pageSize || rows.length + page.length > total || (page.length === 0 && rows.length < total)) throw new AuditReportDataError("Report data completeness could not be verified. Please try again.");
    if (page.length ? Number(match[1]) !== from || Number(match[2]) !== from + page.length - 1 : match[1] !== undefined) throw new AuditReportDataError("Report data completeness could not be verified. Please try again.");
    for (const row of page) {
      if (!row || typeof row !== "object" || typeof row.id !== "string" || seen.has(row.id)) throw new AuditReportDataError("Report data changed during export. Please try again.");
      seen.add(row.id);
      rows.push(row);
    }
    if (rows.length === total) return rows;
    // A server-side page cap smaller than requested would make the next range skip rows.
    if (page.length !== pageSize) throw new AuditReportDataError("Report data completeness could not be verified. Please try again.");
  }
  throw new AuditReportDataError("This report contains more than 10,000 records in one category and cannot be exported safely.");
}
