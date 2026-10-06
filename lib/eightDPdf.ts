import { jsPDF } from "jspdf";
import * as fontkit from "fontkit";
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";

type RecordValue = Record<string, unknown>;
type Colour = [number, number, number];
type Column = { key: string; label: string; weight?: number };
export type EightDReportBranding = { bytes: Uint8Array; format: "PNG" | "JPEG" } | null;
export type EightDReportPicture = { fileName: string; bytes: Uint8Array; format: "WEBP" | "PNG" | "JPEG" };
export type EightDReportInput = { caseRow: RecordValue; sections: RecordValue[]; branding: EightDReportBranding; pictures?: EightDReportPicture[] };
export class ReportInputError extends Error {}

const W = 297, H = 420, M = 12, BOTTOM = H - 16, BODY = 9, LINE = 4.2;
// Reference thumbnails, not full-page exhibits: matches D2's five-column gallery.
export const D2_PICTURE_MAX_MM = 42;
const C = { ink: [38, 41, 54] as Colour, muted: [108, 111, 125] as Colour, purple: [92, 58, 152] as Colour,
  border: [224, 225, 232] as Colour, pale: [248, 246, 252] as Colour, green: [29, 111, 79] as Colour };
const CATEGORIES = ["man", "machine", "method", "material", "measurement", "environment"];
const LABELS = ["People (Man)", "Machine", "Method", "Material", "Measurement", "Environment"];
const obj = (v: unknown): RecordValue => v && typeof v === "object" && !Array.isArray(v) ? v as RecordValue : {};
const str = (v: unknown) => typeof v === "string" ? v.normalize("NFC").replace(/\r\n?/g, "\n").trim() : "";
const array = (v: unknown): unknown[] => Array.isArray(v) ? v : [];
const date = (v: unknown) => { const d = new Date(String(v || "")); return Number.isNaN(d.getTime()) ? "Not recorded" : d.toISOString().slice(0, 10); };
const asset = (name: string) => {
  const local = join(process.cwd(), "assets", name);
  return existsSync(local) ? local : join(process.cwd(), "vercel-export-service", "assets", name);
};
let fontBytes: Buffer | undefined;
let fontFace: ReturnType<typeof fontkit.openSync> | undefined;
function trustedFont() {
  fontBytes ??= readFileSync(asset("NotoSansSC-Variable.ttf"));
  fontFace ??= fontkit.openSync(asset("NotoSansSC-Variable.ttf"));
  return { bytes: fontBytes, face: fontFace! };
}

/** Match D4's legacy normalization without changing or repairing saved work. */
export function reportRcaMethod(d4: RecordValue): "ishikawa" | "five_why" | "fault_tree" {
  const flags = obj(d4.d4_rca_methods);
  return flags.fault_tree === true ? "fault_tree" : flags.ishikawa === false && flags.five_why === true ? "five_why" : "ishikawa";
}
const METHOD_LABEL = { ishikawa: "Ishikawa / cause-and-effect", five_why: "5-Why analysis", fault_tree: "Fault tree analysis" };

/** Content grows onto new pages; font size and saved content are never reduced to fit. */
class Report {
  readonly doc = new jsPDF({ unit: "mm", format: "a3", orientation: "portrait", compress: true, putOnlyUsedFonts: true });
  y = M;
  x = M;
  width = W - M * 2;
  private active: { step: string; title: string; start: number; x: number; width: number } | null = null;
  private pairFrames: Array<NonNullable<Report["active"]>> | null = null;
  private readonly logo: NonNullable<EightDReportBranding>;
  private readonly logoSize: { width: number; height: number };
  constructor(private readonly input: EightDReportInput) {
    const font = trustedFont();
    this.doc.addFileToVFS("NotoSansSC.ttf", font.bytes.toString("base64"));
    this.doc.addFont("NotoSansSC.ttf", "NotoSans", "normal");
    this.logo = input.branding ?? { bytes: readFileSync(asset("qhubio-8d-logo.png")), format: "PNG" };
    try { this.logoSize = this.doc.getImageProperties(this.logo.bytes); }
    catch { throw new ReportInputError("The company logo cannot be decoded. Upload a valid PNG or JPEG."); }
    if (!this.logoSize.width || !this.logoSize.height || this.logoSize.width > 4096 || this.logoSize.height > 4096 || this.logo.bytes.length > 2 * 1024 * 1024) {
      throw new ReportInputError("The company logo exceeds the 2 MB or 4096 px limit.");
    }
    this.doc.setProperties({ title: `8D Corrective Action Report - ${str(input.caseRow.title)}`, author: "Qhubio", subject: "Saved 8D investigation" });
    this.header(false);
  }
  private font(value: string, bold = false, size = BODY) {
    const latin = /^[\x20-\x7E\t\n]*$/.test(value);
    if (!latin) {
      const face = trustedFont().face;
      for (const character of value) {
        const point = character.codePointAt(0)!;
        if (!["\t", "\n"].includes(character) && (point > 0xFFFF || !face.hasGlyphForCodePoint(point))) {
          throw new ReportInputError("This report contains characters the PDF font cannot render. No partial report was exported.");
        }
      }
    }
    this.doc.setFont(latin ? "helvetica" : "NotoSans", latin && bold ? "bold" : "normal");
    this.doc.setFontSize(size);
  }
  lines(value: unknown, width = this.width - 10, bold = false, size = BODY): string[] {
    const content = str(value) || "Not recorded";
    if ([...content].some(c => { const n = c.codePointAt(0)!; return (n < 32 && c !== "\t" && c !== "\n") || (n >= 127 && n <= 159); })) throw new ReportInputError("This report contains unsupported control characters.");
    this.font(content, bold, size);
    return this.doc.splitTextToSize(content.replace(/\t/g, "    "), width) as string[];
  }
  text(value: string, x: number, y: number, options: { bold?: boolean; size?: number; colour?: Colour; align?: "left" | "right" | "center" } = {}) {
    this.font(value, options.bold, options.size ?? BODY);
    this.doc.setTextColor(...(options.colour ?? C.ink));
    this.doc.text(value, x, y, { align: options.align });
  }
  box(x: number, y: number, w: number, h: number, fill?: Colour, border = C.border) {
    this.doc.setDrawColor(...border); this.doc.setLineWidth(0.23);
    if (fill) this.doc.setFillColor(...fill);
    this.doc.roundedRect(x, y, w, h, 1.5, 1.5, fill ? "FD" : "S");
  }
  private reportStatus() {
    const state = str(this.input.caseRow.status);
    // A final report is exportable only after D1-D8 are approved. Some legacy
    // rows retain `draft` after D8, so the governed discipline state is the
    // authoritative report status rather than the stale case label.
    const allDisciplinesApproved = ["D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8"].every(step =>
      this.input.sections.some(section => str(section.step_code) === step && str(section.approval_status) === "approved"),
    );
    const closed = state === "closed" || allDisciplinesApproved;
    return { closed, status: closed ? "CLOSED" : state ? state.replace(/_/g, " ").toUpperCase() : "STATUS NOT RECORDED" };
  }
  private drawBrand() {
    const slot = { x: W - M - 69, y: M, w: 69, h: 18 };
    const scale = Math.min(slot.w / this.logoSize.width, slot.h / this.logoSize.height);
    const lw = this.logoSize.width * scale, lh = this.logoSize.height * scale;
    this.doc.addImage(this.logo.bytes, this.logo.format, slot.x + (slot.w - lw) / 2, slot.y + (slot.h - lh) / 2, lw, lh, undefined, "FAST");
    const { closed, status } = this.reportStatus();
    this.text(status, slot.x + slot.w / 2, M + 25, { size: 8, bold: true, colour: closed ? C.green : C.muted, align: "center" });
  }
  private header(continued: boolean) {
    const doc = this.doc;
    doc.setFillColor(...C.purple); doc.rect(M, M, 1.4, continued ? 19 : 27, "F");
    this.text("QUALITY INVESTIGATION / CUSTOMER COMPLAINT", M + 5, M + 4, { size: 7, bold: true, colour: C.purple });
    this.text(continued ? "8D Report / continued" : "8D Corrective Action Report", M + 5, M + 12, { size: continued ? 17 : 20, bold: true });
    const titleLines = this.lines(this.input.caseRow.title || "8D investigation", 184, false, 10);
    let titleY = M + 20;
    if (titleLines.length <= 3) { for (const line of titleLines) { this.text(line, M + 5, titleY, { size: 10, colour: C.muted }); titleY += 4.5; } }
    this.drawBrand();
    this.y = Math.max(M + 33, titleY + 1);
    if (!continued) {
      this.begin("", "Report identification");
      const c = this.input.caseRow;
      this.table([
        { key: "reference", label: "Complaint / source reference", weight: 1.1 }, { key: "customer", label: "Customer" },
        { key: "project", label: "Project / process" }, { key: "product", label: "Product / part number" },
      ], [{ reference: c.source_reference, customer: c.customer, project: c.project, product: c.product }]);
      this.table([{ key: "id", label: "Case ID", weight: 2 }, { key: "created", label: "Created" }, { key: "updated", label: "Updated" }], [{ id: c.id, created: date(c.created_at), updated: date(c.updated_at) }]);
      if (str(c.team_name) || str(c.severity)) this.table([{ key: "team_name", label: "Team" }, { key: "severity", label: "Complaint / investigation priority" }], [c]);
      if (titleLines.length > 3) this.field("Investigation title", c.title);
      this.end();
    }
  }
  private closeSegment() {
    if (!this.active) return;
    const a = this.active;
    if (this.pairFrames) { this.pairFrames.push(a); return; }
    this.box(a.x, a.start, a.width, Math.max(12, this.y - a.start + 3));
  }
  private newPage() {
    const a = this.active;
    this.closeSegment(); this.active = null;
    this.doc.addPage(); this.x = M; this.width = W - M * 2;
    this.header(true);
    if (a) this.begin(a.step, `${a.title.replace(/ \(continued\)$/, "")} (continued)`);
  }
  ensure(height: number) { if (this.y + height > BOTTOM - 3) this.newPage(); }
  begin(step: string, title: string) {
    this.ensure(35);
    const start = this.y;
    this.active = { step, title, start, x: this.x, width: this.width };
    this.doc.setFillColor(...C.pale); this.doc.roundedRect(this.x, start, this.width, 10, 1.5, 1.5, "F");
    if (step) {
      this.doc.setFillColor(...C.purple); this.doc.roundedRect(this.x + 3, start + 2, 10, 6, 1, 1, "F");
      this.text(step, this.x + 8, start + 6.3, { size: 8, bold: true, colour: [255, 255, 255], align: "center" });
    }
    this.text(title, this.x + (step ? 16 : 5), start + 6.5, { size: 10, bold: true, colour: C.purple });
    this.y = start + 14;
  }
  end() { this.closeSegment(); this.active = null; this.y += 6; }
  estimate(fields: unknown[], tables: Array<{ columns: Column[]; rows: unknown }>, width: number) {
    let h = 20;
    fields.forEach(value => { h += 4.5 + this.lines(value, width - 10).length * LINE; });
    tables.forEach(({ columns, rows }) => {
      const usable = width - 10, total = columns.reduce((n, c) => n + (c.weight ?? 1), 0);
      const widths = columns.map(c => usable * (c.weight ?? 1) / total);
      h += Math.max(...columns.map((c, i) => this.lines(c.label, widths[i] - 4, true, 7).length)) * 3.2 + 6;
      array(rows).forEach(raw => { const row = obj(raw); h += Math.max(...columns.map((c, i) => this.lines(row[c.key], widths[i] - 4).length)) * LINE + 3.5; });
    });
    return h;
  }
  pair(left: () => void, right: () => void, height: number) {
    if (height > 125) { left(); right(); return; }
    this.ensure(height + 5);
    const start = this.y, width = (W - M * 2 - 4) / 2;
    this.pairFrames = [];
    this.width = width; left(); const bottom = this.y;
    this.x = M + width + 4; this.y = start; right();
    this.y = Math.max(bottom, this.y); this.x = M; this.width = W - M * 2;
    for (const frame of this.pairFrames) this.box(frame.x, frame.start, frame.width, this.y - 6 - frame.start + 3);
    this.pairFrames = null;
  }
  field(label: string, value: unknown, options: { emphasize?: boolean } = {}) {
    const lines = this.lines(value);
    this.ensure(11);
    this.text(label.toUpperCase(), this.x + 5, this.y, { size: 7, bold: true, colour: options.emphasize ? C.purple : C.muted });
    this.y += 3.5;
    for (const line of lines) {
      if (this.y + LINE > BOTTOM - 3) {
        this.newPage();
        this.text(`${label.toUpperCase()} (continued)`, this.x + 5, this.y, { size: 7, bold: true, colour: C.muted });
        this.y += 3.5;
      }
      this.text(line, this.x + 5, this.y, { bold: options.emphasize }); this.y += LINE;
    }
    this.y += 1;
  }
  list(label: string, values: unknown) {
    const items = array(values);
    if (!items.length) return this.field(label, "None recorded");
    items.forEach((item, i) => this.field(`${label} ${i + 1}`, typeof item === "string" ? item : JSON.stringify(item)));
  }
  table(columns: Column[], values: unknown) {
    const rows = array(values);
    if (!rows.length) { this.field(columns.map(c => c.label).join(" / "), "None recorded"); return; }
    const usable = this.width - 10;
    const total = columns.reduce((n, c) => n + (c.weight ?? 1), 0);
    const widths = columns.map(c => usable * (c.weight ?? 1) / total);
    const heading = () => {
      this.ensure(15);
      const labelLines = columns.map((c, i) => this.lines(c.label, widths[i] - 4, true, 7));
      const h = Math.max(...labelLines.map(l => l.length)) * 3.2 + 3;
      this.doc.setFillColor(243, 242, 247); this.doc.rect(this.x + 5, this.y - 2, usable, h, "F");
      let x = this.x + 7;
      labelLines.forEach((lines, i) => { lines.forEach((l, j) => this.text(l, x, this.y + 1 + j * 3.2, { size: 7, bold: true, colour: C.muted })); x += widths[i]; });
      this.y += h + 1;
    };
    heading();
    rows.forEach((raw, index) => {
      const row = typeof raw === "string" ? { [columns[0].key]: raw } : obj(raw);
      const cells = columns.map((c, i) => this.lines(row[c.key], widths[i] - 4));
      let offset = 0;
      const count = Math.max(...cells.map(l => l.length));
      if (this.y + Math.min(count * LINE + 4, 80) > BOTTOM - 3) { this.newPage(); heading(); }
      while (offset < count) {
        let capacity = Math.floor((BOTTOM - 6 - this.y) / LINE);
        if (capacity < 1) { this.newPage(); heading(); capacity = Math.floor((BOTTOM - 6 - this.y) / LINE); }
        const size = Math.min(capacity, count - offset);
        const rowTop = this.y - 2;
        let x = this.x + 7;
        cells.forEach((lines, col) => {
          lines.slice(offset, offset + size).forEach((line, j) => this.text(line, x, this.y + 1.5 + j * LINE));
          x += widths[col];
        });
        const rowHeight = size * LINE + 3;
        this.doc.setDrawColor(...C.border); this.doc.setLineWidth(0.18); this.doc.rect(this.x + 5, rowTop, usable, rowHeight);
        this.y += rowHeight + 0.5;
        offset += size;
        if (offset < count) {
          this.newPage(); this.text(`Record ${index + 1} (continued)`, this.x + 5, this.y, { size: 7, colour: C.muted }); this.y += 5; heading();
        }
      }
    });
    this.y += 2;
  }
  /** Dense maps use counts in the visual plus a complete cause register, never shortened labels. */
  ishikawa(d4: RecordValue) {
    const branches = obj(d4.ishikawa);
    const values = CATEGORIES.map(key => array(branches[key]).map(str));
    const full = values.every(items => items.length <= 2 && items.every(v => this.lines(v, 54).length <= 2));
    const height = full ? 91 : 68;
    this.ensure(height + 8);
    const y = this.y, left = this.x + 7, spineY = y + (full ? 43 : 31), doc = this.doc;
    const conclusion = str(d4.selected_root_cause);
    const verified = obj(d4.root_cause_verification).status === "verified";
    const headX = this.x + this.width - 44;
    doc.setDrawColor(...C.purple); doc.setLineWidth(0.65); doc.line(left, spineY, headX - 2, spineY);
    doc.setFillColor(...C.purple); doc.triangle(headX - 4, spineY - 1.6, headX - 4, spineY + 1.6, headX - 1, spineY, "F");
    this.box(headX, spineY - 8, 36, 16, C.pale, C.purple);
    this.text("D2 PROBLEM", headX + 18, spineY - 1.4, { size: 8, bold: true, colour: C.purple, align: "center" });
    this.text("See problem definition", headX + 18, spineY + 4, { size: 7, align: "center" });
    values.forEach((items, i) => {
      const upper = i < 3, column = i % 3, nodeX = left + column * 65, junction = nodeX + 50;
      const endY = upper ? y + 9 : y + height - 11;
      doc.setDrawColor(157, 150, 174); doc.setLineWidth(0.4); doc.line(junction, spineY, nodeX + 18, endY);
      this.box(nodeX, upper ? y : y + height - 8, 58, 6, [243, 242, 247]);
      this.text(LABELS[i].toUpperCase(), nodeX + 29, upper ? y + 4.1 : y + height - 3.9, { size: 7, bold: true, colour: C.muted, align: "center" });
      const nodeValues = full ? (items.length ? items : ["No recorded cause"]) : [items.length ? `${items.length} causes / register below` : "No recorded cause"];
      let nodeY = upper ? y + 11 : spineY + 8;
      nodeValues.forEach(value => {
        const lines = this.lines(value, 54), h = lines.length * LINE + 4;
        const isConclusion = verified && !!conclusion && value === conclusion;
        this.box(nodeX, nodeY, 58, h, isConclusion ? C.pale : [255, 255, 255], isConclusion ? C.purple : C.border);
        lines.forEach((line, j) => this.text(line, nodeX + 2, nodeY + 5 + j * LINE, { colour: isConclusion ? C.purple : C.ink, bold: isConclusion }));
        nodeY += h + 2;
      });
    });
    this.y += height + 3;
    this.field("Cause-map interpretation", full ? "Candidate causes are not proof. Purple marks the verified selected cause only." : "The diagram shows category counts. Every cause and its category is recorded in full below; candidates are not proof.");
    if (!full) values.forEach((items, i) => items.forEach((value, j) => this.field(`${i + 1}.${j + 1} / ${LABELS[i]}`, value)));
  }
  fiveWhy(problem: string, value: unknown) {
    this.field("Problem / D2", problem);
    const entries = array(value);
    if (!entries.length) return this.field("5-Why chain", "None recorded");
    entries.forEach((entry, i) => { this.ensure(15); if (i) { this.text("|", this.x + 7, this.y, { size: 10, colour: C.purple }); this.y += 4; } this.field(`Why ${i + 1}`, entry, { emphasize: true }); });
  }
  faultTree(problem: string, value: unknown) {
    this.field("Fault-tree interpretation", "Qualitative analysis. AND requires all child events; OR requires any child event. Event paths retain parent / child relationships across continuation pages. No probability is inferred.");
    const tree = obj(value), root = obj(tree.root || value);
    if (!Object.keys(root).length) return this.field("Fault tree", "None recorded");
    const nodes: Array<{ node: RecordValue; path: string; depth: number; centre: number; children: number[]; lines: string[] }> = [];
    const treeSeen = new Set<object>(); let leaf = 0;
    const layout = (node: RecordValue, path: string, depth: number): number => {
      if (treeSeen.has(node)) throw new ReportInputError("The fault tree contains an invalid cyclic relationship.");
      treeSeen.add(node);
      const index = nodes.length;
      const lines = this.lines(depth ? node.text : problem, depth ? 55 : 114);
      nodes.push({ node, path, depth, centre: 0, children: [], lines });
      const children = array(node.children).map((child, i) => layout(obj(child), `${path}.${i + 1}`, depth + 1));
      nodes[index].children = children;
      nodes[index].centre = children.length ? children.reduce((sum, i) => sum + nodes[i].centre, 0) / children.length : leaf++;
      return index;
    };
    layout(root, "1", 0);
    const compact = leaf <= 4 && nodes.length <= 12 && nodes.every(n => n.depth <= 2 && n.lines.length <= 5);
    if (compact) {
      const levels = Array.from({ length: Math.max(...nodes.map(n => n.depth)) + 1 }, (_, depth) => Math.max(...nodes.filter(n => n.depth === depth).map(n => n.lines.length * LINE + 11)) + 14);
      const height = levels.reduce((sum, h) => sum + h, 0);
      this.ensure(height + 4);
      const start = this.y, usable = this.width - 16, slot = usable / leaf;
      const pos = (n: typeof nodes[number]) => ({ x: this.x + 8 + (n.centre + 0.5) * slot, y: start + levels.slice(0, n.depth).reduce((sum, h) => sum + h, 0), h: n.lines.length * LINE + 11, w: n.depth ? 59 : 118 });
      nodes.forEach(n => {
        const p = pos(n);
        this.doc.setDrawColor(159, 150, 178); this.doc.setLineWidth(0.45);
        n.children.forEach(i => { const c = pos(nodes[i]); const linkY = p.y + p.h + 9; this.doc.line(p.x, p.y + p.h, p.x, linkY); this.doc.line(p.x, linkY, c.x, linkY); this.doc.line(c.x, linkY, c.x, c.y); });
      });
      nodes.forEach(n => {
        const p = pos(n);
        this.box(p.x - p.w / 2, p.y, p.w, p.h, n.depth ? [255, 255, 255] : C.pale, n.depth ? C.border : C.purple);
        this.text(`EVENT ${n.path} / ${str(n.node.type).toUpperCase() || "EVENT"}`, p.x - p.w / 2 + 2, p.y + 4.5, { size: 6.5, bold: true, colour: C.purple });
        n.lines.forEach((line, i) => this.text(line, p.x - p.w / 2 + 2, p.y + 10 + i * LINE));
        if (str(n.node.gate)) {
          this.box(p.x - 7, p.y + p.h + 2, 14, 5, C.pale, C.purple);
          this.text(str(n.node.gate).toUpperCase(), p.x, p.y + p.h + 5.5, { size: 7, bold: true, colour: C.purple, align: "center" });
        }
      });
      this.y += height + 2;
      return;
    }
    this.field("Expanded fault-tree event register", "This tree exceeds the compact diagram area. Every event follows below with its hierarchical path, parent, event type and gate.");
    const visited = new Set<object>();
    const visit = (node: RecordValue, path: string, parent: string | null) => {
      if (visited.has(node)) throw new ReportInputError("The fault tree contains an invalid cyclic relationship.");
      visited.add(node);
      const gate = str(node.gate).toUpperCase(), kind = str(node.type) || "event";
      const title = `Event ${path} / ${parent ? `parent ${parent}` : "TOP EVENT"} / ${kind}${gate ? ` / ${gate}` : ""}`;
      this.field(title, parent ? node.text : problem, { emphasize: !parent });
      array(node.children).forEach((child, i) => visit(obj(child), `${path}.${i + 1}`, path));
    };
    visit(root, "1", null);
  }
  visualEvidence(pictures: EightDReportPicture[]) {
    if (!pictures.length) return;
    if (pictures.length > 5) throw new ReportInputError("D2 supports a maximum of 5 pictures. No pictures were omitted.");
    const decoded = pictures.map(picture => {
      try {
        const image = this.doc.getImageProperties(picture.bytes);
        if (!image.width || !image.height || image.width > 4096 || image.height > 4096 || picture.bytes.length > 5 * 1024 * 1024) throw new Error("Invalid size");
        // jsPDF's WEBP decoder already returns a PDF-compatible JPEG payload.
        // Reuse it, rather than decoding/encoding each large picture twice.
        return { ...picture, bytes: image.fileType === "WEBP" ? Buffer.from(image.data, "binary") : picture.bytes, format: image.fileType === "WEBP" ? "JPEG" as const : picture.format, width: image.width, height: image.height };
      } catch { throw new ReportInputError("A D2 picture could not be decoded. No partial report was exported; review the pictures in D2 and retry."); }
    });
    this.field(`Visual reference / ${pictures.length} ${pictures.length === 1 ? "picture" : "pictures"}`, "Uploaded D2 pictures are for visual reference only. They are not AI inputs or proof of root cause.");
    const gap = 3, columns = 5, imageH = D2_PICTURE_MAX_MM, captionLine = 3.5;
    const cardW = Math.min(D2_PICTURE_MAX_MM + 8, (this.width - 10 - gap * (columns - 1)) / columns);
    for (let offset = 0; offset < decoded.length; offset += columns) {
      const row = decoded.slice(offset, offset + columns).map(picture => ({ picture, caption: this.lines(picture.fileName, cardW - 8, false, 7.5) }));
      const rowH = imageH + 18 + Math.max(...row.map(item => item.caption.length)) * captionLine;
      this.ensure(rowH + 6);
      const top = this.y - 2;
      row.forEach(({ picture, caption }, index) => {
        const x = this.x + 5 + index * (cardW + gap);
        this.box(x, top, cardW, rowH);
        this.text(`PICTURE ${offset + index + 1}`, x + 4, top + 6, { size: 7, bold: true, colour: C.purple });
        this.doc.setFillColor(246, 247, 250); this.doc.rect(x + 4, top + 9, cardW - 8, imageH, "F");
        const scale = Math.min(Math.min(D2_PICTURE_MAX_MM, cardW - 8) / picture.width, imageH / picture.height);
        const w = picture.width * scale, h = picture.height * scale;
        this.doc.addImage(picture.bytes, picture.format, x + 4 + (cardW - 8 - w) / 2, top + 9 + (imageH - h) / 2, w, h, undefined, "FAST");
        caption.forEach((line, i) => this.text(line, x + 4, top + imageH + 15 + i * captionLine, { size: 7.5 }));
      });
      this.y = top + rowH + 6;
    }
  }
  finish() {
    const total = this.doc.getNumberOfPages();
    for (let page = 1; page <= total; page++) {
      this.doc.setPage(page); this.doc.setDrawColor(...C.border); this.doc.setLineWidth(0.2); this.doc.line(M, H - 11, W - M, H - 11);
      this.text(`Qhubio / Case ${str(this.input.caseRow.id)}`, M, H - 6, { size: 7, colour: C.muted });
      this.text(`Generated ${date(new Date())} / Confidential / Page ${page} of ${total}`, W - M, H - 6, { size: 7, colour: C.muted, align: "right" });
      // Headers are painted again as the final layer. This prevents page-one
      // branding from being visually covered by later PDF drawing operations.
      this.drawBrand();
    }
    return Buffer.from(this.doc.output("arraybuffer"));
  }
}

/** Saved human-facing fields, not AI metadata, suggested signals or inactive RCA workspaces. */
export function buildEightDReportPdf(input: EightDReportInput): Buffer {
  const r = new Report(input);
  const section = (step: string) => obj(input.sections.find(s => s.step_code === step)?.content);
  const d1 = section("D1"), d2 = section("D2"), d3 = section("D3"), d4 = section("D4"), d5 = section("D5"), d6 = section("D6"), d7 = section("D7"), d8 = section("D8");
  const problem = str(d2.summary) || "Problem description not recorded";
  const teamColumns = [{ key: "role", label: "Team member / role" }, { key: "responsibility", label: "Responsibility", weight: 2 }];
  const problemColumns = [{ key: "what", label: "What", weight: 2 }, { key: "where", label: "Where" }, { key: "when", label: "When" }, { key: "who", label: "Who" }];
  const contextColumns = [{ key: "why_significant", label: "Why significant", weight: 1.5 }, { key: "how_detected", label: "How detected", weight: 1.5 }, { key: "how_many", label: "How many" }];
  const half = (W - M * 2 - 4) / 2;
  // D2 needs full width for its compact five-picture row. Text-only exports keep
  // the existing paired D1/D2 layout; picture exports stay in discipline order.
  const pairHeight = input.pictures?.length ? Infinity : Math.max(r.estimate([d1.champion, d1.scope], [{ columns: teamColumns, rows: d1.team_members }], half), r.estimate([problem], [{ columns: problemColumns, rows: [d2] }, { columns: contextColumns, rows: [d2] }], half));
  r.pair(() => {
    r.begin("D1", "Team and scope"); r.field("Champion", d1.champion); r.field("Scope", d1.scope); r.table(teamColumns, d1.team_members); r.end();
  }, () => {
    r.begin("D2", "Problem definition"); r.field("Problem description", problem, { emphasize: true }); r.table(problemColumns, [d2]); r.table(contextColumns, [d2]); r.visualEvidence(input.pictures ?? []); r.end();
  }, pairHeight);
  r.begin("D3", "Interim containment / customer protection"); r.field("Customer protection", d3.customer_protection_summary);
  r.table([{ key: "action", label: "Containment action", weight: 3 }, { key: "owner_role", label: "Owner" }, { key: "due", label: "Due" }, { key: "verification", label: "Verification", weight: 2 }], d3.actions); r.end();
  const selected = reportRcaMethod(d4);
  r.ensure(selected === "ishikawa" ? 130 : 50);
  r.begin("D4", `Root cause analysis / ${METHOD_LABEL[selected]}`);
  if (selected === "ishikawa") r.ishikawa(d4);
  else if (selected === "five_why") r.fiveWhy(problem, d4.five_why);
  else r.faultTree(problem, d4.fault_tree);
  const verification = obj(d4.root_cause_verification);
  r.field(verification.status === "verified" ? "Verified selected root cause" : "Selected root-cause candidate / not verified", d4.selected_root_cause, { emphasize: true });
  r.field("Occurrence root cause", d4.occurrence_root_cause);
  r.field("Detection / escape root cause", d4.detection_escape_not_applicable === true ? "Not applicable (recorded in D4)" : d4.detection_escape_root_cause);
  r.field("Verification method", d4.verification_method); r.list("Root-cause verification evidence", verification.evidence);
  array(d4.investigation_findings).forEach((raw, i) => { const finding = obj(raw); r.field(`Recorded finding ${i + 1}`, finding.finding); r.list(`Finding ${i + 1} supporting evidence`, finding.evidence_support); });
  if (selected === "ishikawa" && array(d4.hypotheses).length) r.table([
    { key: "hypothesis", label: "Working hypothesis", weight: 1.5 }, { key: "reasoning", label: "Reasoning", weight: 2 }, { key: "validation_required", label: "Validation required", weight: 1.5 },
  ], d4.hypotheses);
  if (selected === "ishikawa") array(d4.hypotheses).forEach((h, i) => { if (array(obj(h).evidence_support).length) r.list(`Hypothesis ${i + 1} evidence`, obj(h).evidence_support); });
  if (array(d4.rejected_hypotheses).length) r.list("Ruled-out hypothesis / reason", d4.rejected_hypotheses);
  r.end();
  r.begin("D5", "Permanent corrective actions");
  r.table([{ key: "root_cause", label: "Addresses root cause", weight: 1.2 }, { key: "permanent_action", label: "Permanent action", weight: 2 }, { key: "verification", label: "Verification", weight: 1.5 }, { key: "effectiveness_measure", label: "Effectiveness measure", weight: 1.5 }], d5.actions); r.end();
  r.begin("D6", "Implementation and effectiveness validation");
  r.table([{ key: "step", label: "Implementation step", weight: 3 }, { key: "owner_role", label: "Owner" }, { key: "due", label: "Due" }], d6.implementation_plan);
  r.field("Validation plan", d6.validation_plan); r.list("Evidence required", d6.evidence_required);
  r.table([{ key: "criterion", label: "Acceptance criterion", weight: 2 }, { key: "result", label: "Observed result", weight: 2 }, { key: "evidence_reference", label: "Evidence reference", weight: 1.5 }], d6.validation_results); r.end();
  const approvalColumns = [{ key: "role", label: "Closure approval role" }, { key: "name_placeholder", label: "Approver name", weight: 2 }];
  const listValues = (value: unknown) => {
    const items = array(value);
    return items.length ? items.map(item => typeof item === "string" ? item : JSON.stringify(item)) : ["None recorded"];
  };
  const d7Fields = [
    ...listValues(d7.standardization),
    ...listValues(d7.documents_to_update),
    ...listValues(d7.read_across),
    d7.lessons_learned,
  ];
  const closureHeight = r.estimate([d8.closure_summary, d8.recognition], [{ columns: approvalColumns, rows: d8.approvals }], W - M * 2);
  const d7Height = r.estimate(d7Fields, [], half);
  const writeD7 = () => {
    r.begin("D7", "Prevent recurrence / institutionalize learning");
    r.list("Standardization", d7.standardization); r.list("Documents to update", d7.documents_to_update); r.list("Read-across", d7.read_across); r.field("Lessons learned", d7.lessons_learned); r.end();
  };
  const writeD8 = () => {
    r.begin("D8", "Closure and recognition"); r.field("Closure summary", d8.closure_summary); r.field("Recognition", d8.recognition);
    r.table(approvalColumns, d8.approvals); r.end();
  };
  // Compact D7/D8 records share the available row. This prevents a short
  // closure block from becoming an almost-empty continuation page while long
  // records retain the existing full-width, overflow-safe layout.
  if (Math.max(d7Height, closureHeight) <= 125) r.pair(writeD7, writeD8, Math.max(d7Height, closureHeight));
  else { writeD7(); if (closureHeight <= 125) r.ensure(closureHeight); writeD8(); }
  return r.finish();
}
