import { NS, elementChildren } from "./ooxml";
import { DEFAULT_SCHEME, type PptxTheme, type ClrMap } from "./themes";

// DrawingML color modifiers. Percentage-typed values are stored as 0..1 factors
// (OOXML carries them as thousandths of a percent, e.g. val="50000" => 0.5).
// hueMod is likewise a 0..1 factor, but hueOff is a fixed angle (60000ths of a
// degree) normalized here to turns (0..1) to match the 0..1 HSL hue scale.
export interface ColorMods {
  tint?: number;
  shade?: number;
  lumMod?: number;
  lumOff?: number;
  satMod?: number;
  satOff?: number;
  hueMod?: number;
  hueOff?: number;
  alpha?: number;
}

// A DrawingML color reference: exactly one of srgb/scheme/sys/prst is set.
export interface DrawingMlColor {
  srgb?: string; // "RRGGBB"
  scheme?: string; // scheme name or role (accent1, tx1, bg1, phClr, ...)
  sys?: string; // resolved lastClr "RRGGBB"
  prst?: string; // preset color name
  mods: ColorMods;
}

// A small subset of DrawingML preset colors. Anything missing falls through to
// null (callers then skip the fill rather than render a wrong color).
const PRESET: Record<string, string> = {
  black: "000000",
  white: "FFFFFF",
  red: "FF0000",
  green: "008000",
  blue: "0000FF",
  yellow: "FFFF00",
  cyan: "00FFFF",
  magenta: "FF00FF",
  gray: "808080",
  grey: "808080",
  darkGray: "A9A9A9",
  lightGray: "D3D3D3",
  orange: "FFA500",
  purple: "800080",
};

// Read the first color choice element under `parent` (e.g. a solidFill, gs, or
// bgRef) and its modifier children.
export function parseColorChoice(parent: Element): DrawingMlColor | null {
  for (const child of elementChildren(parent)) {
    if (child.namespaceURI !== NS.a) continue;
    switch (child.localName) {
      case "srgbClr": {
        const v = child.getAttribute("val");
        if (v && /^[0-9A-Fa-f]{6}$/.test(v)) return { srgb: v.toUpperCase(), mods: readMods(child) };
        break;
      }
      case "schemeClr": {
        const v = child.getAttribute("val");
        if (v) return { scheme: v, mods: readMods(child) };
        break;
      }
      case "sysClr": {
        const last = child.getAttribute("lastClr");
        if (last && /^[0-9A-Fa-f]{6}$/.test(last)) return { sys: last.toUpperCase(), mods: readMods(child) };
        break;
      }
      case "prstClr": {
        const v = child.getAttribute("val");
        if (v) return { prst: v, mods: readMods(child) };
        break;
      }
    }
  }
  return null;
}

function readMods(colorEl: Element): ColorMods {
  const mods: ColorMods = {};
  for (const m of elementChildren(colorEl)) {
    if (m.namespaceURI !== NS.a) continue;
    const v = parseInt(m.getAttribute("val") ?? "", 10);
    if (Number.isNaN(v)) continue;
    switch (m.localName) {
      case "tint": mods.tint = v / 100000; break;
      case "shade": mods.shade = v / 100000; break;
      case "lumMod": mods.lumMod = v / 100000; break;
      case "lumOff": mods.lumOff = v / 100000; break;
      case "satMod": mods.satMod = v / 100000; break;
      case "satOff": mods.satOff = v / 100000; break;
      case "hueMod": mods.hueMod = v / 100000; break;
      // hueOff is ST_FixedAngle (60000ths of a degree); normalize to turns
      // (deg/360) so it adds onto the 0..1 hue scale used below.
      case "hueOff": mods.hueOff = v / 21600000; break;
      case "alpha": mods.alpha = v / 100000; break;
    }
  }
  return mods;
}

// Resolve a DrawingMlColor to a CSS color string, or null if unresolvable.
export function resolveDrawingMlColor(
  c: DrawingMlColor | null,
  theme: PptxTheme | null,
  clrMap: ClrMap,
): string | null {
  if (!c) return null;
  let hex6: string | null = null;
  if (c.srgb) hex6 = c.srgb;
  else if (c.sys) hex6 = c.sys;
  else if (c.scheme) hex6 = schemeHex(c.scheme, theme, clrMap);
  else if (c.prst) hex6 = PRESET[c.prst] ?? null;
  if (!hex6) return null;

  hex6 = applyMods(hex6, c.mods);
  if (c.mods.alpha != null && c.mods.alpha < 1) {
    return hexToRgba(hex6, c.mods.alpha);
  }
  return `#${hex6}`;
}

// Resolve a schemeClr val (scheme name or logical role) to a hex via clrMap.
function schemeHex(name: string, theme: PptxTheme | null, clrMap: ClrMap): string | null {
  if (name === "phClr") return null; // placeholder color: only meaningful inside a style def
  let n = name;
  if (n === "bg1" || n === "tx1" || n === "bg2" || n === "tx2") {
    n = clrMap[n] ?? n;
  }
  const scheme = theme?.scheme ?? DEFAULT_SCHEME;
  return scheme[n] ?? DEFAULT_SCHEME[n] ?? null;
}

function applyMods(hex6: string, mods: ColorMods): string {
  let { r, g, b } = hexToRgb(hex6);
  if (mods.shade != null) {
    r *= mods.shade;
    g *= mods.shade;
    b *= mods.shade;
  }
  if (mods.tint != null) {
    r = r * mods.tint + 255 * (1 - mods.tint);
    g = g * mods.tint + 255 * (1 - mods.tint);
    b = b * mods.tint + 255 * (1 - mods.tint);
  }
  // HSL modifiers, applied in fixed order hue -> sat -> lum. All operate on the
  // 0..1 HSL scale (hue wraps into [0,1)).
  if (
    mods.hueMod != null ||
    mods.hueOff != null ||
    mods.satMod != null ||
    mods.satOff != null ||
    mods.lumMod != null ||
    mods.lumOff != null
  ) {
    const hsl = rgbToHsl(r, g, b);
    let h = hsl.h;
    let s = hsl.s;
    let l = hsl.l;
    if (mods.hueMod != null) h *= mods.hueMod;
    if (mods.hueOff != null) h += mods.hueOff;
    h -= Math.floor(h); // wrap into [0,1)
    if (mods.satMod != null) s *= mods.satMod;
    if (mods.satOff != null) s += mods.satOff;
    if (mods.lumMod != null) l *= mods.lumMod;
    if (mods.lumOff != null) l += mods.lumOff;
    const rgb = hslToRgb(h, clamp01(s), clamp01(l));
    r = rgb.r;
    g = rgb.g;
    b = rgb.b;
  }
  return rgbToHex(r, g, b);
}

function hexToRgb(hex6: string): { r: number; g: number; b: number } {
  return {
    r: parseInt(hex6.slice(0, 2), 16),
    g: parseInt(hex6.slice(2, 4), 16),
    b: parseInt(hex6.slice(4, 6), 16),
  };
}

function rgbToHex(r: number, g: number, b: number): string {
  const toHex = (n: number): string =>
    Math.round(Math.max(0, Math.min(255, n))).toString(16).padStart(2, "0").toUpperCase();
  return `${toHex(r)}${toHex(g)}${toHex(b)}`;
}

function hexToRgba(hex6: string, alpha: number): string {
  const { r, g, b } = hexToRgb(hex6);
  return `rgba(${r}, ${g}, ${b}, ${clamp01(alpha)})`;
}

function clamp01(n: number): number {
  return Math.max(0, Math.min(1, n));
}

function rgbToHsl(r: number, g: number, b: number): { h: number; s: number; l: number } {
  r /= 255;
  g /= 255;
  b /= 255;
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
  return { h, s, l };
}

function hslToRgb(h: number, s: number, l: number): { r: number; g: number; b: number } {
  if (s === 0) {
    const v = l * 255;
    return { r: v, g: v, b: v };
  }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const hue = (t: number): number => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return {
    r: hue(h + 1 / 3) * 255,
    g: hue(h) * 255,
    b: hue(h - 1 / 3) * 255,
  };
}
