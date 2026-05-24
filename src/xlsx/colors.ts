export interface ExcelColorRef {
  argb?: string;
  theme?: number;
  tint?: number;
  indexed?: number;
}

// Default Office theme color palette (OOXML clrScheme). The vast majority of
// workbooks use this palette; for workbooks with a custom theme the colors
// will be slightly off but the tint math still produces a sensible variant.
export const DEFAULT_THEME_RGB = [
  "FFFFFF", // 0: lt1 (Background 1)
  "000000", // 1: dk1 (Text 1)
  "E7E6E6", // 2: lt2 (Background 2)
  "44546A", // 3: dk2 (Text 2)
  "4472C4", // 4: Accent 1 (blue)
  "ED7D31", // 5: Accent 2 (orange)
  "A5A5A5", // 6: Accent 3 (gray)
  "FFC000", // 7: Accent 4 (gold)
  "5B9BD5", // 8: Accent 5 (light blue)
  "70AD47", // 9: Accent 6 (green)
  "0563C1", // 10: Hyperlink
  "954F72", // 11: Followed Hyperlink
];

// Standard Excel indexed color palette (xlsx legacy "indexed" attribute).
export const INDEXED_COLORS = [
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
  "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
  "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
  "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
  "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
  "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
];

export function argbToCss(argb: string): string {
  // ExcelJS hands us 8 hex chars (alpha first). CSS hex with alpha is RRGGBBAA.
  // Pad short inputs to a full 8-hex ARGB so we never emit malformed CSS.
  const padded = argb.length >= 8 ? argb : argb.padStart(8, "F");
  const a = padded.slice(-8, -6).toUpperCase();
  const rgb = padded.slice(-6);
  return a === "FF" ? `#${rgb}` : `#${rgb}${a}`;
}

export function resolveExcelColor(
  color: ExcelColorRef | undefined | null,
  theme: readonly string[] = DEFAULT_THEME_RGB,
): string | null {
  if (!color) return null;
  if (color.argb) return argbToCss(color.argb);
  if (typeof color.theme === "number") {
    const hex = theme[color.theme] ?? DEFAULT_THEME_RGB[color.theme];
    if (hex) return applyTint(hex, color.tint ?? 0);
    // Fall through: workbooks can carry both a theme ref and an indexed
    // fallback. If the theme index is out of range (custom themes can
    // exceed 11), try indexed before giving up.
  }
  if (typeof color.indexed === "number") {
    const hex = INDEXED_COLORS[color.indexed];
    if (hex) return `#${hex}`;
  }
  return null;
}

// OOXML tint is applied in HSL space against luminance. Positive tint lightens
// toward white, negative tint darkens toward black.
export function applyTint(hex6: string, tint: number): string {
  if (Math.abs(tint) < 0.001) return `#${hex6}`;
  const r = parseInt(hex6.slice(0, 2), 16) / 255;
  const g = parseInt(hex6.slice(2, 4), 16) / 255;
  const b = parseInt(hex6.slice(4, 6), 16) / 255;

  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  let h = 0;
  let s = 0;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h /= 6;
  }

  const newL = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;

  let nr: number;
  let ng: number;
  let nb: number;
  if (s === 0) {
    nr = ng = nb = newL;
  } else {
    const q = newL < 0.5 ? newL * (1 + s) : newL + s - newL * s;
    const p = 2 * newL - q;
    const hueToRgb = (t: number): number => {
      if (t < 0) t += 1;
      if (t > 1) t -= 1;
      if (t < 1 / 6) return p + (q - p) * 6 * t;
      if (t < 1 / 2) return q;
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
      return p;
    };
    nr = hueToRgb(h + 1 / 3);
    ng = hueToRgb(h);
    nb = hueToRgb(h - 1 / 3);
  }

  const toHex = (n: number): string => {
    const v = Math.round(Math.max(0, Math.min(255, n * 255)));
    return v.toString(16).padStart(2, "0").toUpperCase();
  };
  return `#${toHex(nr)}${toHex(ng)}${toHex(nb)}`;
}
