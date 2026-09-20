/**
 * DrawingML shapes (`wps:wsp`) and text boxes — preset shape types,
 * fill, outline, shape text body, transform.
 */

import type { ColorValue } from '../colors';
import type { ImageSize, ImagePosition, ImageWrap, ImageTransform } from './image';
import type { Paragraph } from './paragraph';
import type { Table } from './table';

/**
 * What a shape's text body can hold. `w:txbxContent` is `EG_BlockLevelElts`,
 * so a text box may contain tables as well as paragraphs — the Iris proposal
 * template's "PROOF POINT" panel is a two-column table inside one.
 *
 * @public
 */
export type ShapeBlockContent = Paragraph | Table;

/**
 * Shape types
 */
export type ShapeType =
  // Basic shapes
  | 'rect'
  | 'roundRect'
  | 'ellipse'
  | 'triangle'
  | 'rtTriangle'
  | 'parallelogram'
  | 'trapezoid'
  | 'pentagon'
  | 'hexagon'
  | 'heptagon'
  | 'octagon'
  | 'decagon'
  | 'dodecagon'
  | 'star4'
  | 'star5'
  | 'star6'
  | 'star7'
  | 'star8'
  | 'star10'
  | 'star12'
  | 'star16'
  | 'star24'
  | 'star32'
  // Lines and connectors
  | 'line'
  | 'straightConnector1'
  | 'bentConnector2'
  | 'bentConnector3'
  | 'bentConnector4'
  | 'bentConnector5'
  | 'curvedConnector2'
  | 'curvedConnector3'
  | 'curvedConnector4'
  | 'curvedConnector5'
  // Arrows
  | 'rightArrow'
  | 'leftArrow'
  | 'upArrow'
  | 'downArrow'
  | 'leftRightArrow'
  | 'upDownArrow'
  | 'quadArrow'
  | 'leftRightUpArrow'
  | 'bentArrow'
  | 'uturnArrow'
  | 'leftUpArrow'
  | 'bentUpArrow'
  | 'curvedRightArrow'
  | 'curvedLeftArrow'
  | 'curvedUpArrow'
  | 'curvedDownArrow'
  | 'stripedRightArrow'
  | 'notchedRightArrow'
  | 'homePlate'
  | 'chevron'
  | 'rightArrowCallout'
  | 'downArrowCallout'
  | 'leftArrowCallout'
  | 'upArrowCallout'
  | 'leftRightArrowCallout'
  | 'quadArrowCallout'
  | 'circularArrow'
  // Flowchart
  | 'flowChartProcess'
  | 'flowChartAlternateProcess'
  | 'flowChartDecision'
  | 'flowChartInputOutput'
  | 'flowChartPredefinedProcess'
  | 'flowChartInternalStorage'
  | 'flowChartDocument'
  | 'flowChartMultidocument'
  | 'flowChartTerminator'
  | 'flowChartPreparation'
  | 'flowChartManualInput'
  | 'flowChartManualOperation'
  | 'flowChartConnector'
  | 'flowChartOffpageConnector'
  | 'flowChartPunchedCard'
  | 'flowChartPunchedTape'
  | 'flowChartSummingJunction'
  | 'flowChartOr'
  | 'flowChartCollate'
  | 'flowChartSort'
  | 'flowChartExtract'
  | 'flowChartMerge'
  | 'flowChartOnlineStorage'
  | 'flowChartDelay'
  | 'flowChartMagneticTape'
  | 'flowChartMagneticDisk'
  | 'flowChartMagneticDrum'
  | 'flowChartDisplay'
  // Callouts
  | 'wedgeRectCallout'
  | 'wedgeRoundRectCallout'
  | 'wedgeEllipseCallout'
  | 'cloudCallout'
  | 'borderCallout1'
  | 'borderCallout2'
  | 'borderCallout3'
  | 'accentCallout1'
  | 'accentCallout2'
  | 'accentCallout3'
  | 'callout1'
  | 'callout2'
  | 'callout3'
  | 'accentBorderCallout1'
  | 'accentBorderCallout2'
  | 'accentBorderCallout3'
  // Other
  | 'actionButtonBlank'
  | 'actionButtonHome'
  | 'actionButtonHelp'
  | 'actionButtonInformation'
  | 'actionButtonBackPrevious'
  | 'actionButtonForwardNext'
  | 'actionButtonBeginning'
  | 'actionButtonEnd'
  | 'actionButtonReturn'
  | 'actionButtonDocument'
  | 'actionButtonSound'
  | 'actionButtonMovie'
  | 'irregularSeal1'
  | 'irregularSeal2'
  | 'frame'
  | 'halfFrame'
  | 'corner'
  | 'diagStripe'
  | 'chord'
  | 'arc'
  | 'bracketPair'
  | 'bracePair'
  | 'leftBracket'
  | 'rightBracket'
  | 'leftBrace'
  | 'rightBrace'
  | 'can'
  | 'cube'
  | 'bevel'
  | 'donut'
  | 'noSmoking'
  | 'blockArc'
  | 'foldedCorner'
  | 'smileyFace'
  | 'heart'
  | 'lightningBolt'
  | 'sun'
  | 'moon'
  | 'cloud'
  | 'snip1Rect'
  | 'snip2SameRect'
  | 'snip2DiagRect'
  | 'snipRoundRect'
  | 'round1Rect'
  | 'round2SameRect'
  | 'round2DiagRect'
  | 'plaque'
  | 'teardrop'
  | 'mathPlus'
  | 'mathMinus'
  | 'mathMultiply'
  | 'mathDivide'
  | 'mathEqual'
  | 'mathNotEqual'
  | 'gear6'
  | 'gear9'
  | 'funnel'
  | 'pieWedge'
  | 'pie'
  | 'leftCircularArrow'
  | 'leftRightCircularArrow'
  | 'swooshArrow'
  | 'textBox';

/**
 * Shape fill type
 */
export interface ShapeFill {
  type: 'none' | 'solid' | 'gradient' | 'pattern' | 'picture';
  /** Solid fill color */
  color?: ColorValue;
  /** Gradient stops for gradient fill */
  gradient?: {
    type: 'linear' | 'radial' | 'rectangular' | 'path';
    angle?: number;
    stops: Array<{
      position: number; // 0-100000
      color: ColorValue;
    }>;
  };
}

/**
 * Shape outline/stroke
 */
export interface ShapeOutline {
  /** Line width in EMUs */
  width?: number;
  /** Line color */
  color?: ColorValue;
  /** Line style */
  style?:
    | 'solid'
    | 'dot'
    | 'dash'
    | 'lgDash'
    | 'dashDot'
    | 'lgDashDot'
    | 'lgDashDotDot'
    | 'sysDot'
    | 'sysDash'
    | 'sysDashDot'
    | 'sysDashDotDot';
  /** Line cap */
  cap?: 'flat' | 'round' | 'square';
  /** Line join */
  join?: 'bevel' | 'miter' | 'round';
  /** Head arrow */
  headEnd?: {
    type: 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow';
    width?: 'sm' | 'med' | 'lg';
    length?: 'sm' | 'med' | 'lg';
  };
  /** Tail arrow */
  tailEnd?: {
    type: 'none' | 'triangle' | 'stealth' | 'diamond' | 'oval' | 'arrow';
    width?: 'sm' | 'med' | 'lg';
    length?: 'sm' | 'med' | 'lg';
  };
}

/**
 * Text body inside a shape
 */
export interface ShapeTextBody {
  /** Text direction */
  vertical?: boolean;
  /** Rotation */
  rotation?: number;
  /** Anchor/vertical alignment */
  anchor?: 'top' | 'middle' | 'bottom' | 'distributed' | 'justified';
  /** Anchor center */
  anchorCenter?: boolean;
  /** Auto fit */
  autoFit?: 'none' | 'normal' | 'shape';
  /** Text margins */
  margins?: {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
  };
  /** Block content inside the shape — paragraphs and tables. */
  content: ShapeBlockContent[];
  /**
   * The source `<wps:bodyPr>` verbatim.
   *
   * The element carries a dozen attributes plus an autofit child, and the
   * fields above model five of them — rebuilding it from those alone dropped
   * `<a:spAutoFit/>` (so Word and LibreOffice stopped sizing the box to its
   * text) along with `wrap`, `vertOverflow`, `horzOverflow`, `anchor` and
   * `compatLnSpc`. Nothing here depends on the shape's size or its text, so
   * replaying it is safe across an edit. Re-validated before it is written.
   */
  bodyPrXml?: string;
}

/**
 * Shape/drawing object (wps:wsp)
 */
export interface Shape {
  type: 'shape';
  /** Shape type preset */
  shapeType: ShapeType;
  /** Unique ID */
  id?: string;
  /** Name */
  name?: string;
  /** Size in EMUs */
  size: ImageSize;
  /** Position for floating shapes */
  position?: ImagePosition;
  /** Wrap settings */
  wrap?: ImageWrap;
  /**
   * `wp:anchor relativeHeight` — z-order among overlapping anchored
   * objects (higher paints on top); see `Image.relativeHeight`.
   */
  relativeHeight?: number;
  /** Fill */
  fill?: ShapeFill;
  /** Outline/stroke */
  outline?: ShapeOutline;
  /** Transform */
  transform?: ImageTransform;
  /** Text content inside the shape */
  textBody?: ShapeTextBody;
  /** Custom geometry points */
  customGeometry?: string;
  /**
   * A stroke-only connector (`<wps:cNvCnPr>`, or `a:prstGeom` in the line
   * family). Word draws these as a single line between two corners of the
   * extent box, NOT as a rectangle — a footer rule is `prst="line"` with
   * `cy="0"`, and outlining its bounding box paints a full-width border
   * where the file asked for a hairline. `'down'` runs top-left to
   * bottom-right, `'up'` (from `a:xfrm/@flipV`) bottom-left to top-right.
   */
  lineShape?: 'down' | 'up';
  /**
   * Rounded outline from `a:prstGeom` (§20.1.9.18). Word draws the preset's
   * geometry, not its bounding box: an `ellipse` badge and a `roundRect`
   * pill button both painted as hard-cornered rectangles without this.
   * `'ellipse'` also covers the flow-chart connector preset (a circle);
   * `'roundRect'` takes its corner radius from {@link Shape.cornerAdj}.
   */
  geometry?: 'ellipse' | 'roundRect';
  /**
   * `a:avLst/a:gd[@name="adj"]` as a fraction of the shape's SHORTER side
   * (§20.1.9.11), so the corner radius is `cornerAdj * min(w, h)`. Word's
   * default is 0.16667; `0.5` is a full pill. Only read for
   * `geometry === 'roundRect'`.
   */
  cornerAdj?: number;
  /**
   * Canvas-only shape lifted out of preserved markup so the page paints its
   * fill; the serializer skips it. Set for decorative filled shapes with no
   * text (`isFilledShapeDrawing`), whose original `mc:AlternateContent` is
   * written back verbatim. Mirrors `Image.renderOnly`.
   */
  renderOnly?: boolean;
  /**
   * `<wps:spPr>` children this model has no field for, verbatim — today
   * `<a:ln>` and `<a:effectLst>`. A shape that explicitly declares "no
   * outline" (`<a:ln><a:noFill/></a:ln>`) parses to no `outline`, so
   * rebuilding spPr from the model alone silently swapped that for the
   * default outline. Only emitted when the model has nothing of its own to
   * say, so an edit through the UI still wins.
   */
  spPrExtraXml?: string;
}

/**
 * Text box (floating text container)
 */
export interface TextBox {
  type: 'textBox';
  /** See {@link ShapeTextBody.bodyPrXml}. */
  bodyPrXml?: string;
  /** See {@link Shape.spPrExtraXml}. */
  spPrExtraXml?: string;
  /** Unique ID */
  id?: string;
  /** Size */
  size: ImageSize;
  /** Position */
  position?: ImagePosition;
  /** Wrap settings */
  wrap?: ImageWrap;
  /**
   * `wp:anchor relativeHeight` — z-order among overlapping anchored
   * objects (higher paints on top); see `Image.relativeHeight`.
   */
  relativeHeight?: number;
  /** Fill */
  fill?: ShapeFill;
  /** Outline */
  outline?: ShapeOutline;
  /** Text content — paragraphs and tables (§`w:txbxContent`). */
  content: ShapeBlockContent[];
  /** Internal margins */
  margins?: {
    top?: number;
    bottom?: number;
    left?: number;
    right?: number;
  };
  /** See {@link Shape.lineShape}. */
  lineShape?: 'down' | 'up';
  /** See {@link Shape.geometry}. */
  geometry?: 'ellipse' | 'roundRect';
  /** See {@link Shape.cornerAdj}. */
  cornerAdj?: number;
  /** See {@link Shape.renderOnly}. */
  renderOnly?: boolean;
}
