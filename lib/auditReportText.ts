import type { jsPDF } from "jspdf";
import * as fontkit from "fontkit";
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";

export class AuditReportTextError extends Error {}
const cache = new Map<string, { bytes: Buffer; face: ReturnType<typeof fontkit.openSync> }>();
function asset(name: string) {
  const local = join(process.cwd(), "assets", name);
  const path = existsSync(local) ? local : join(process.cwd(), "vercel-export-service", "assets", name);
  if (!cache.has(name)) cache.set(name, { bytes: readFileSync(path), face: fontkit.openSync(path) });
  return cache.get(name)!;
}

/** Renders mixed scripts using the same bundled Noto fonts as Qhubio's 8D reports. */
export class AuditReportText {
  private fallbackAdded = false;
  private glyphs = new Map<string, boolean>();
  constructor(private readonly pdf: jsPDF) {
    for (const [file, style] of [["NotoSans-Regular.ttf", "normal"], ["NotoSans-Bold.ttf", "bold"]] as const) {
      pdf.addFileToVFS(file, asset(file).bytes.toString("base64"));
      pdf.addFont(file, "AuditReport", style);
    }
  }
  private runs(value: string, bold: boolean) {
    const runs: Array<{ text: string; fallback: boolean }> = [];
    const primary = asset(bold ? "NotoSans-Bold.ttf" : "NotoSans-Regular.ttf").face;
    for (const ch of value) {
      const point = ch.codePointAt(0)!;
      const key = `${bold ? "b" : "r"}:${point}`;
      let hasPrimary = this.glyphs.get(key);
      if (hasPrimary === undefined) { hasPrimary = primary.hasGlyphForCodePoint(point); this.glyphs.set(key, hasPrimary); }
      const fallback = !hasPrimary;
      if (fallback) {
        const secondary = asset("NotoSansSC-Report-Regular.ttf");
        if (point > 0xffff || !secondary.face.hasGlyphForCodePoint(point)) throw new AuditReportTextError("This report contains characters its PDF fonts cannot render. No partial report was exported.");
        if (!this.fallbackAdded) { this.pdf.addFileToVFS("AuditFallback.ttf", secondary.bytes.toString("base64")); this.pdf.addFont("AuditFallback.ttf", "AuditFallback", "normal"); this.fallbackAdded = true; }
      }
      const last = runs.at(-1);
      if (last?.fallback === fallback) last.text += ch; else runs.push({ text: ch, fallback });
    }
    return runs;
  }
  private select(fallback: boolean, bold: boolean, size: number) {
    this.pdf.setFont(fallback ? "AuditFallback" : "AuditReport", fallback ? "normal" : bold ? "bold" : "normal");
    this.pdf.setFontSize(size);
  }
  private width(text: string, bold: boolean, size: number) {
    return this.runs(text, bold).reduce((sum, run) => { this.select(run.fallback, bold, size); return sum + this.pdf.getTextWidth(run.text); }, 0);
  }
  lines(text: string, width: number, size: number, bold: boolean): string[] {
    const runs = this.runs(text, bold);
    if (runs.length <= 1) { this.select(runs[0]?.fallback ?? false, bold, size); return this.pdf.splitTextToSize(text, width) as string[]; }
    const lines: string[] = []; let line = "";
    for (const token of text.match(/\S+|\s+/gu) ?? []) {
      if (line && this.width(line + token, bold, size) > width) { lines.push(line.trimEnd()); line = ""; }
      if (!line && !token.trim()) continue;
      if (this.width(token, bold, size) <= width) line += token;
      else for (const ch of token) { if (line && this.width(line + ch, bold, size) > width) { lines.push(line); line = ""; } line += ch; }
    }
    if (line) lines.push(line.trimEnd());
    return lines;
  }
  draw(text: string, x: number, y: number, size: number, bold: boolean) {
    let left = x;
    for (const run of this.runs(text, bold)) { this.select(run.fallback, bold, size); this.pdf.text(run.text, left, y); left += this.pdf.getTextWidth(run.text); }
  }
}

