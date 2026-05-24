// EMU = English Metric Units. 914400 EMU/inch; at 96 px/inch that is 9525
// EMU/px. PresentationML is an ABSOLUTE coordinate space (the whole slide is
// sized in EMU and every shape carries an absolute a:off/a:ext), unlike the
// xlsx cell-anchor model — so the helpers here scale a slide to a target CSS
// width rather than walking cumulative grid sizes.
export const EMU_PER_PX = 9525;
export const EMU_PER_PT = 12700; // 914400 / 72

export interface Box {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface SlideScale {
  scale: number; // px per EMU
  widthPx: number;
  heightPx: number;
}

export function emuToPx(emu: number): number {
  return emu / EMU_PER_PX;
}

// Compute the px-per-EMU scale that fits a slide of the given EMU dimensions to
// `targetWidthPx`, plus the resulting stage pixel dimensions.
export function slideScaleFor(
  cxEmu: number,
  cyEmu: number,
  targetWidthPx: number,
): SlideScale {
  const scale = cxEmu > 0 ? targetWidthPx / cxEmu : 1 / EMU_PER_PX;
  return { scale, widthPx: targetWidthPx, heightPx: cyEmu * scale };
}

export function xfrmToBox(
  off: { x: number; y: number },
  ext: { cx: number; cy: number },
  scale: number,
): Box {
  return {
    left: off.x * scale,
    top: off.y * scale,
    width: Math.max(1, ext.cx * scale),
    height: Math.max(1, ext.cy * scale),
  };
}

// a:xfrm rot is in 60000ths of a degree.
export function degFromRot(rot: number): number {
  return rot / 60000;
}

// A coordinate frame maps EMU points in some coordinate space to stage pixels.
// At the slide root this is a uniform scale with zero origin; inside a p:grpSp
// the group's child coordinate system (a:chOff/a:chExt) yields a shifted,
// independently-scaled frame so nested shapes position correctly.
export interface Frame {
  scaleX: number; // px per EMU
  scaleY: number;
  originXEmu: number; // EMU origin of this coordinate space
  originYEmu: number;
  baseLeftPx: number; // px position of that origin within the stage
  baseTopPx: number;
}

export function rootFrame(scale: number): Frame {
  return { scaleX: scale, scaleY: scale, originXEmu: 0, originYEmu: 0, baseLeftPx: 0, baseTopPx: 0 };
}

export function boxInFrame(
  off: { x: number; y: number },
  ext: { cx: number; cy: number },
  frame: Frame,
): Box {
  return {
    left: frame.baseLeftPx + (off.x - frame.originXEmu) * frame.scaleX,
    top: frame.baseTopPx + (off.y - frame.originYEmu) * frame.scaleY,
    width: Math.max(1, ext.cx * frame.scaleX),
    height: Math.max(1, ext.cy * frame.scaleY),
  };
}

// Build the child frame for a group from the group's own box (already in parent
// frame coordinates) and its child coordinate system.
export function childFrame(
  groupBox: Box,
  chOff: { x: number; y: number },
  chExt: { cx: number; cy: number },
): Frame {
  return {
    scaleX: chExt.cx > 0 ? groupBox.width / chExt.cx : 0,
    scaleY: chExt.cy > 0 ? groupBox.height / chExt.cy : 0,
    originXEmu: chOff.x,
    originYEmu: chOff.y,
    baseLeftPx: groupBox.left,
    baseTopPx: groupBox.top,
  };
}
