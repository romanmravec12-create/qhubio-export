declare module "fontkit" {
  export interface Font {
    hasGlyphForCodePoint(codePoint: number): boolean;
  }

  export function openSync(path: string): Font;
}
