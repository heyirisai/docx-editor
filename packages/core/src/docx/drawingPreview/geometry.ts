/**
 * DrawingML geometry (`a:prstGeom` / `a:custGeom`) as SVG path data.
 *
 * Only what can be drawn faithfully is drawn: a handful of presets plus custom
 * geometry, whose guide formulas (ECMA-376 §20.1.9.11) are evaluated here.
 * Anything else returns null and the caller paints a placeholder instead —
 * a wrong shape is worse than a marked one.
 */

import { getAttribute, getChildElements, getLocalName, type XmlElement } from '../xmlParser';

/** One drawable path, in the shape's own box (origin top-left, units EMU). */
export interface GeometryPath {
  d: string;
  /** `a:path@fill="none"` — the outline only. */
  fill: boolean;
  /** `a:path@stroke="0"` — the fill only. */
  stroke: boolean;
}

/** Hard bounds: the markup comes from an untrusted file. */
const MAX_GUIDES = 2000;
const MAX_PATH_COMMANDS = 20000;

const childrenNamed = (el: XmlElement | null | undefined, name: string): XmlElement[] =>
  getChildElements(el).filter((c) => getLocalName(c.name ?? '') === name);
const childNamed = (el: XmlElement | null | undefined, name: string): XmlElement | null =>
  childrenNamed(el, name)[0] ?? null;

const fmt = (n: number): string => {
  if (!Number.isFinite(n)) return '0';
  const r = Math.round(n * 100) / 100;
  return Object.is(r, -0) ? '0' : String(r);
};

// ============================================================================
// GUIDES
// ============================================================================

/** 60000ths of a degree → radians. */
const angToRad = (a: number): number => ((a / 60000) * Math.PI) / 180;
const radToAng = (r: number): number => ((r * 180) / Math.PI) * 60000;

/** The built-in guide values every shape gets (§20.1.9.11, "shape guides"). */
function builtinGuides(w: number, h: number): Map<string, number> {
  const ss = Math.min(w, h);
  const ls = Math.max(w, h);
  const g = new Map<string, number>([
    ['l', 0],
    ['t', 0],
    ['r', w],
    ['b', h],
    ['w', w],
    ['h', h],
    ['hc', w / 2],
    ['vc', h / 2],
    ['ss', ss],
    ['ls', ls],
    ['cd2', 10800000],
    ['cd4', 5400000],
    ['cd8', 2700000],
    ['3cd4', 16200000],
    ['3cd8', 8100000],
    ['5cd8', 13500000],
    ['7cd8', 18900000],
  ]);
  for (const d of [2, 3, 4, 5, 6, 8, 10, 12, 16, 24, 32]) {
    g.set(`wd${d}`, w / d);
    g.set(`hd${d}`, h / d);
    g.set(`ssd${d}`, ss / d);
  }
  return g;
}

/**
 * Evaluate one guide formula. Returns null for an unknown operator or an
 * argument that is neither a number nor a defined guide.
 */
function evalFormula(fmla: string, guides: Map<string, number>): number | null {
  const parts = fmla.trim().split(/\s+/);
  const op = parts[0];
  const args: number[] = [];
  for (const raw of parts.slice(1)) {
    const v = guides.has(raw) ? guides.get(raw)! : Number(raw);
    if (!Number.isFinite(v)) return null;
    args.push(v);
  }
  const [x = 0, y = 0, z = 0] = args;
  switch (op) {
    case 'val':
      return x;
    case '*/':
      return z === 0 ? 0 : (x * y) / z;
    case '+-':
      return x + y - z;
    case '+/':
      return z === 0 ? 0 : (x + y) / z;
    case '?:':
      return x > 0 ? y : z;
    case 'abs':
      return Math.abs(x);
    case 'at2':
      return radToAng(Math.atan2(y, x));
    case 'cat2':
      return x * Math.cos(Math.atan2(z, y));
    case 'cos':
      return x * Math.cos(angToRad(y));
    case 'max':
      return Math.max(x, y);
    case 'min':
      return Math.min(x, y);
    case 'mod':
      return Math.sqrt(x * x + y * y + z * z);
    case 'pin':
      return y < x ? x : y > z ? z : y;
    case 'sat2':
      return x * Math.sin(Math.atan2(z, y));
    case 'sin':
      return x * Math.sin(angToRad(y));
    case 'sqrt':
      return Math.sqrt(Math.max(0, x));
    case 'tan':
      return x * Math.tan(angToRad(y));
    default:
      return null;
  }
}

/** `a:avLst` / `a:gdLst` guides layered over `base`; null if one cannot be evaluated. */
function evalGuideList(
  list: XmlElement | null,
  guides: Map<string, number>
): Map<string, number> | null {
  const gds = childrenNamed(list, 'gd');
  if (gds.length > MAX_GUIDES) return null;
  for (const gd of gds) {
    const name = getAttribute(gd, null, 'name');
    const fmla = getAttribute(gd, null, 'fmla');
    if (!name || !fmla) continue;
    const v = evalFormula(fmla, guides);
    if (v === null) return null;
    guides.set(name, v);
  }
  return guides;
}

// ============================================================================
// PRESETS
// ============================================================================

/** Adjust value `name` from a preset's `a:avLst`, or the preset default. */
function presetAdjust(prstGeom: XmlElement, name: string, fallback: number): number {
  for (const gd of childrenNamed(childNamed(prstGeom, 'avLst'), 'gd')) {
    if (getAttribute(gd, null, 'name') !== name) continue;
    const m = /^val\s+(-?\d+(?:\.\d+)?)$/.exec((getAttribute(gd, null, 'fmla') ?? '').trim());
    if (m) return Number(m[1]);
  }
  return fallback;
}

const rectPath = (w: number, h: number): string => `M0 0H${fmt(w)}V${fmt(h)}H0Z`;

function roundRectPath(w: number, h: number, adj: number): string {
  const r = (Math.min(w, h) * Math.max(0, Math.min(adj, 50000))) / 100000;
  if (r <= 0) return rectPath(w, h);
  const [W, H, R] = [fmt(w), fmt(h), fmt(r)];
  return (
    `M${R} 0H${fmt(w - r)}A${R} ${R} 0 0 1 ${W} ${R}V${fmt(h - r)}` +
    `A${R} ${R} 0 0 1 ${fmt(w - r)} ${H}H${R}A${R} ${R} 0 0 1 0 ${fmt(h - r)}V${R}` +
    `A${R} ${R} 0 0 1 ${R} 0Z`
  );
}

function ellipsePath(w: number, h: number): string {
  const [rx, ry] = [fmt(w / 2), fmt(h / 2)];
  return `M0 ${ry}A${rx} ${ry} 0 1 1 ${fmt(w)} ${ry}A${rx} ${ry} 0 1 1 0 ${ry}Z`;
}

/** Preset geometry names this module draws; everything else is a placeholder. */
export const SUPPORTED_PRESETS: ReadonlySet<string> = new Set([
  'rect',
  'flowChartProcess',
  'roundRect',
  'flowChartAlternateProcess',
  'ellipse',
  'flowChartConnector',
  'line',
  'straightConnector1',
  'triangle',
  'rtTriangle',
  'diamond',
  'flowChartDecision',
  'parallelogram',
  'homePlate',
  'chevron',
]);

function presetPaths(prstGeom: XmlElement, w: number, h: number): GeometryPath[] | null {
  const prst = getAttribute(prstGeom, null, 'prst') ?? '';
  const closed = (d: string): GeometryPath[] => [{ d, fill: true, stroke: true }];
  switch (prst) {
    case 'rect':
    case 'flowChartProcess':
      return closed(rectPath(w, h));
    case 'roundRect':
      return closed(roundRectPath(w, h, presetAdjust(prstGeom, 'adj', 16667)));
    case 'flowChartAlternateProcess':
      return closed(roundRectPath(w, h, 16667));
    case 'ellipse':
    case 'flowChartConnector':
      return closed(ellipsePath(w, h));
    case 'line':
    case 'straightConnector1':
      return [{ d: `M0 0L${fmt(w)} ${fmt(h)}`, fill: false, stroke: true }];
    case 'triangle': {
      const x = (w * presetAdjust(prstGeom, 'adj', 50000)) / 100000;
      return closed(`M${fmt(x)} 0L${fmt(w)} ${fmt(h)}H0Z`);
    }
    case 'rtTriangle':
      return closed(`M0 0L${fmt(w)} ${fmt(h)}H0Z`);
    case 'diamond':
    case 'flowChartDecision':
      return closed(
        `M${fmt(w / 2)} 0L${fmt(w)} ${fmt(h / 2)}L${fmt(w / 2)} ${fmt(h)}L0 ${fmt(h / 2)}Z`
      );
    case 'parallelogram': {
      const x = (Math.min(w, h) * presetAdjust(prstGeom, 'adj', 25000)) / 100000;
      return closed(`M${fmt(x)} 0H${fmt(w)}L${fmt(w - x)} ${fmt(h)}H0Z`);
    }
    case 'homePlate': {
      const x = w - (Math.min(w, h) * presetAdjust(prstGeom, 'adj', 50000)) / 100000;
      return closed(`M0 0H${fmt(x)}L${fmt(w)} ${fmt(h / 2)}L${fmt(x)} ${fmt(h)}H0Z`);
    }
    case 'chevron': {
      const x = (Math.min(w, h) * presetAdjust(prstGeom, 'adj', 50000)) / 100000;
      return closed(
        `M0 0H${fmt(w - x)}L${fmt(w)} ${fmt(h / 2)}L${fmt(w - x)} ${fmt(h)}H0L${fmt(x)} ${fmt(h / 2)}Z`
      );
    }
    default:
      return null;
  }
}

// ============================================================================
// CUSTOM GEOMETRY
// ============================================================================

function customPaths(custGeom: XmlElement, w: number, h: number): GeometryPath[] | null {
  let guides: Map<string, number> | null = builtinGuides(w, h);
  guides = evalGuideList(childNamed(custGeom, 'avLst'), guides);
  if (guides) guides = evalGuideList(childNamed(custGeom, 'gdLst'), guides);
  if (!guides) return null;
  const g = guides;

  const out: GeometryPath[] = [];
  let commands = 0;
  for (const path of childrenNamed(childNamed(custGeom, 'pathLst'), 'path')) {
    // A path declares its own coordinate space; absent, it is the shape's.
    const pw = Number(getAttribute(path, null, 'w') ?? '') || w;
    const ph = Number(getAttribute(path, null, 'h') ?? '') || h;
    const sx = pw === 0 ? 1 : w / pw;
    const sy = ph === 0 ? 1 : h / ph;
    const val = (raw: string | null): number | null => {
      if (raw === null) return null;
      const v = g.has(raw) ? g.get(raw)! : Number(raw);
      return Number.isFinite(v) ? v : null;
    };
    const pt = (el: XmlElement | null): [number, number] | null => {
      const x = val(getAttribute(el, null, 'x'));
      const y = val(getAttribute(el, null, 'y'));
      return x === null || y === null ? null : [x, y];
    };

    let d = '';
    let cur: [number, number] = [0, 0];
    for (const cmd of getChildElements(path)) {
      if (++commands > MAX_PATH_COMMANDS) return null;
      const name = getLocalName(cmd.name ?? '');
      const pts = childrenNamed(cmd, 'pt').map(pt);
      if (pts.some((p) => p === null)) return null;
      const P = pts as Array<[number, number]>;
      const XY = (p: [number, number]) => `${fmt(p[0] * sx)} ${fmt(p[1] * sy)}`;
      switch (name) {
        case 'moveTo':
          if (!P[0]) return null;
          d += `M${XY(P[0])}`;
          cur = P[0];
          break;
        case 'lnTo':
          if (!P[0]) return null;
          d += `L${XY(P[0])}`;
          cur = P[0];
          break;
        case 'cubicBezTo':
          if (P.length < 3) return null;
          d += `C${XY(P[0])} ${XY(P[1])} ${XY(P[2])}`;
          cur = P[2];
          break;
        case 'quadBezTo':
          if (P.length < 2) return null;
          d += `Q${XY(P[0])} ${XY(P[1])}`;
          cur = P[1];
          break;
        case 'arcTo': {
          const wR = val(getAttribute(cmd, null, 'wR'));
          const hR = val(getAttribute(cmd, null, 'hR'));
          const stAng = val(getAttribute(cmd, null, 'stAng'));
          const swAng = val(getAttribute(cmd, null, 'swAng'));
          if (wR === null || hR === null || stAng === null || swAng === null) return null;
          // The current point sits on the ellipse at stAng; walk swAng from it.
          const st = angToRad(stAng);
          const en = angToRad(stAng + swAng);
          const cx = cur[0] - wR * Math.cos(st);
          const cy = cur[1] - hR * Math.sin(st);
          const end: [number, number] = [cx + wR * Math.cos(en), cy + hR * Math.sin(en)];
          const large = Math.abs(swAng) > 10800000 ? 1 : 0;
          const sweep = swAng > 0 ? 1 : 0;
          d += `A${fmt(wR * sx)} ${fmt(hR * sy)} 0 ${large} ${sweep} ${XY(end)}`;
          cur = end;
          break;
        }
        case 'close':
          d += 'Z';
          break;
        default:
          return null;
      }
    }
    if (!d) continue;
    out.push({
      d,
      fill: getAttribute(path, null, 'fill') !== 'none',
      stroke: !['0', 'false'].includes(getAttribute(path, null, 'stroke') ?? ''),
    });
  }
  return out.length > 0 ? out : null;
}

/**
 * The geometry declared in a shape's `spPr`, as paths in a `w`×`h` box, or
 * null when it is absent or not something this module can draw.
 */
export function geometryPaths(
  spPr: XmlElement | null,
  w: number,
  h: number
): GeometryPath[] | null {
  if (!spPr || !(w >= 0) || !(h >= 0)) return null;
  const prst = childNamed(spPr, 'prstGeom');
  if (prst) return presetPaths(prst, w, h);
  const cust = childNamed(spPr, 'custGeom');
  if (cust) return customPaths(cust, w, h);
  return null;
}
