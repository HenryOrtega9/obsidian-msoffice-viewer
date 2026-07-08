import { NS, elementChildren } from "./ooxml";
import { DEFAULT_SCHEME, type PptxTheme, type ClrMap } from "./themes";

// A DrawingML color modifier, kept in XML document order because the spec
// applies transforms sequentially in the order authored. Percentage-typed
// values arrive as thousandths of a percent (val="50000" => 0.5); hueOff is a
// fixed angle (60000ths of a degree) normalized to turns to match the 0..1
// HSL hue scale.
export interface ColorMod {
  kind:
    | "tint"
    | "shade"
    | "lumMod"
    | "lumOff"
    | "satMod"
    | "satOff"
    | "hueMod"
    | "hueOff"
    | "alpha";
  val: number;
}

// A DrawingML color reference: exactly one of srgb/scheme/sys/prst is set.
export interface DrawingMlColor {
  srgb?: string; // "RRGGBB"
  scheme?: string; // scheme name or role (accent1, tx1, bg1, phClr, ...)
  sys?: string; // resolved lastClr "RRGGBB"
  prst?: string; // preset color name
  mods: ColorMod[];
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

const MOD_NAMES = new Set([
  "tint", "shade", "lumMod", "lumOff", "satMod", "satOff", "hueMod", "hueOff", "alpha",
]);

function readMods(colorEl: Element): ColorMod[] {
  const mods: ColorMod[] = [];
  for (const m of elementChildren(colorEl)) {
    if (m.namespaceURI !== NS.a || !MOD_NAMES.has(m.localName)) continue;
    const raw = parseInt(m.getAttribute("val") ?? "", 10);
    if (Number.isNaN(raw)) continue;
    const kind = m.localName as ColorMod["kind"];
    // hueOff is ST_FixedAngle (60000ths of a degree); normalize to turns.
    const val = kind === "hueOff" ? raw / 21600000 : raw / 100000;
    mods.push({ kind, val });
  }
  return mods;
}

// Resolve a DrawingMlColor to a CSS color string, or null if unresolvable.
// `phClr` supplies the substitution color ("RRGGBB") for scheme val="phClr"
// references inside theme format-scheme style definitions.
export function resolveDrawingMlColor(
  c: DrawingMlColor | null,
  theme: PptxTheme | null,
  clrMap: ClrMap,
  phClr?: string | null,
): string | null {
  if (!c) return null;
  let hex6: string | null = null;
  if (c.srgb) hex6 = c.srgb;
  else if (c.sys) hex6 = c.sys;
  else if (c.scheme) hex6 = schemeHex(c.scheme, theme, clrMap, phClr);
  else if (c.prst) hex6 = PRESET[c.prst] ?? null;
  if (!hex6) return null;

  const { hex, alpha } = applyModList(hex6, c.mods);
  if (alpha != null && alpha < 1) {
    return hexToRgba(hex, alpha);
  }
  return `#${hex}`;
}

// Resolve a schemeClr val (scheme name or logical role) to a hex via clrMap.
function schemeHex(
  name: string,
  theme: PptxTheme | null,
  clrMap: ClrMap,
  phClr?: string | null,
): string | null {
  if (name === "phClr") return phClr ?? null; // placeholder color from a style ref
  let n = name;
  if (n === "bg1" || n === "tx1" || n === "bg2" || n === "tx2") {
    n = clrMap[n] ?? n;
  }
  const scheme = theme?.scheme ?? DEFAULT_SCHEME;
  return scheme[n] ?? DEFAULT_SCHEME[n] ?? null;
}

// Apply the DrawingML color modifiers carried as children of a color element
// (lumMod/lumOff/tint/shade/sat/hue) to a base "RRGGBB" hex. Shared with the
// xlsx chart renderer so series scheme colors get the same transforms Excel
// applies (e.g. accent1 lumMod 60% lumOff 40%).
export function applyColorMods(hex6: string, colorEl: Element): string {
  return applyModList(hex6, readMods(colorEl)).hex;
}

// sRGB <-> linear conversions (2.2 gamma approximation). Office applies
// tint/shade in linear RGB, so doing it on gamma-encoded channels renders
// shades markedly too dark and tints too light.
function srgbToLinear(c255: number): number {
  return Math.pow(Math.max(0, Math.min(255, c255)) / 255, 2.2);
}

function linearToSrgb(lin: number): number {
  return Math.pow(Math.max(0, Math.min(1, lin)), 1 / 2.2) * 255;
}

// Apply modifiers sequentially in document order. HSL-typed mods convert to
// HSL for that step only, so interleavings like shade->satMod->tint behave.
function applyModList(hex6: string, mods: ColorMod[]): { hex: string; alpha: number | null } {
  let { r, g, b } = hexToRgb(hex6);
  let alpha: number | null = null;

  const inLinear = (f: (lin: number) => number): void => {
    r = linearToSrgb(f(srgbToLinear(r)));
    g = linearToSrgb(f(srgbToLinear(g)));
    b = linearToSrgb(f(srgbToLinear(b)));
  };
  const inHsl = (f: (hsl: { h: number; s: number; l: number }) => void): void => {
    const hsl = rgbToHsl(r, g, b);
    f(hsl);
    hsl.h -= Math.floor(hsl.h); // wrap into [0,1)
    const rgb = hslToRgb(hsl.h, clamp01(hsl.s), clamp01(hsl.l));
    r = rgb.r;
    g = rgb.g;
    b = rgb.b;
  };

  for (const m of mods) {
    switch (m.kind) {
      case "shade": inLinear((lin) => lin * m.val); break;
      case "tint": inLinear((lin) => lin * m.val + (1 - m.val)); break;
      case "hueMod": inHsl((hsl) => { hsl.h *= m.val; }); break;
      case "hueOff": inHsl((hsl) => { hsl.h += m.val; }); break;
      case "satMod": inHsl((hsl) => { hsl.s *= m.val; }); break;
      case "satOff": inHsl((hsl) => { hsl.s += m.val; }); break;
      case "lumMod": inHsl((hsl) => { hsl.l *= m.val; }); break;
      case "lumOff": inHsl((hsl) => { hsl.l += m.val; }); break;
      case "alpha": alpha = clamp01(m.val); break;
    }
  }
  return { hex: rgbToHex(r, g, b), alpha };
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
