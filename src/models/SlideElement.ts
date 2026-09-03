export interface Position {
  /** Horizontal position in EMUs (English Metric Units) */
  x: number;

  /** Vertical position in EMUs */
  y: number;
}

export interface Size {
  /** Width in EMUs */
  width: number;

  /** Height in EMUs */
  height: number;
}

export interface TextElement {
  /** Element type identifier */
  type: "text";
  /** Paint order within the slide. Larger values are painted in front. */
  zIndex?: number;

  /** Text content */
  content: string;

  /** Element position */
  position: Position;

  /** Element size */
  size: Size;

  /** Text font configuration */
  font: {
    /** Font family name */
    name: string;

    /** Font size in points */
    size: number;

    /** Font color as hexadecimal string (#RRGGBB) */
    color: string;

    /** Font weight (normal/bold) */
    weight?: "normal" | "bold";
  };

  /** Horizontal and vertical alignment inside its bounding box */
  align?: {
    horizontal?: "left" | "center" | "right" | "justify";
    vertical?: "top" | "middle" | "bottom";
  };

  /** Internal padding in px from bodyPr insets */
  padding?: { left: number; top: number; right: number; bottom: number };

  /** Optional rich HTML content (e.g., bullets/numbering) */
  html?: string;
}

export interface ImageElement {
  /** Element type identifier */
  type: "image";
  zIndex?: number;

  /** Relationship ID pointing to the image in /ppt/media */
  relId: string;

  /** Data URI or image source path */
  src: string;

  /** Image position */
  position: Position;

  /** Image size */
  size: Size;
}

export interface ShapeElement {
  /** Element type identifier */
  type: "shape";
  zIndex?: number;

  /** Shape type name (e.g. rectangle, ellipse) */
  shapeType: string;

  /** Shape position */
  position: Position;

  /** Shape size */
  size: Size;

  /** Fill color as hexadecimal string (#RRGGBB) */
  fillColor: string;

  /** Border color as hexadecimal string (#RRGGBB) */
  borderColor?: string;

  /** Stroke width in px for lines/connectors */
  strokeWidth?: number;

  /** Rotation in degrees applied around center */
  rotationDeg?: number;

  /** Arrowhead at the start of the line */
  headEnd?: { type?: string; w?: string; len?: string };

  /** Arrowhead at the end of the line */
  tailEnd?: { type?: string; w?: string; len?: string };
}

export interface BackgroundElement {
  /** Element type identifier */
  type: "background";
  zIndex?: number;

  /** Solid background color (hex) */
  fillColor?: string;

  /** Background image as data URI */
  imageSrc?: string;
}

/** Union type of any possible slide element */
export interface TableCell {
  text: string;
  font?: { name?: string; size?: number; color?: string };
  align?: { horizontal?: "left" | "center" | "right" | "justify"; vertical?: "top" | "middle" | "bottom" };
  padding?: { left: number; top: number; right: number; bottom: number };
  fillColor?: string;
  colSpan?: number;
  rowSpan?: number;
  borders?: {
    top?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" };
    right?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" };
    bottom?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" };
    left?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" };
  };
}

export interface TableRow {
  cells: TableCell[];
}

export interface TableElement {
  type: "table";
  zIndex?: number;
  position: Position; // EMUs
  size: Size; // EMUs
  columns: number[]; // column widths in EMUs
  rows: TableRow[];
  tableStyle?: { firstRow?: boolean; firstCol?: boolean; lastRow?: boolean; lastCol?: boolean; bandRow?: boolean; bandCol?: boolean };
  tableBorders?: {
    top?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" };
    right?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" };
    bottom?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" };
    left?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" };
    insideH?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" };
    insideV?: { color?: string; width?: number; style?: "solid" | "dashed" | "dotted" };
  };
  /** Optional table style id and resolved fills/font colors from theme */
  tableStyleId?: string;
  /** Optional table-level background fill color (from tblPr solidFill) */
  tableFillColor?: string;
  style?: {
    fills?: Partial<Record<
      "wholeTbl" | "band1H" | "band2H" | "band1V" | "band2V" | "firstRow" | "lastRow" | "firstCol" | "lastCol",
      string
    >>;
    fontColors?: Partial<Record<
      "wholeTbl" | "band1H" | "band2H" | "band1V" | "band2V" | "firstRow" | "lastRow" | "firstCol" | "lastCol",
      string
    >>;
  };
}

export type SlideElement =
  | TextElement
  | ImageElement
  | ShapeElement
  | BackgroundElement
  | TableElement
  | ChartElement;

export type ChartType = "bar" | "column" | "line" | "pie" | "area" | "scatter";

export interface ChartSeries {
  name?: string;
  values?: number[];
  points?: { x: number; y: number }[]; // for scatter
  color?: string;
  valueFormat?: string; // optional Excel/OOXML numFmt formatCode
  /** Per-datapoint colors from c:dPt elements (used by pie/doughnut) */
  ptColors?: string[];
}

export interface ChartElement {
  type: "chart";
  zIndex?: number;
  chartType: ChartType;
  position: Position; // EMUs
  size: Size; // EMUs
  categories: (string | number)[];
  series: ChartSeries[];
  /** Optional palette (e.g., theme accents) used for series when color missing */
  palette?: string[];
  /** Optional chart title */
  title?: string;
  /** Legend visibility (if present in chart) */
  showLegend?: boolean;
  /** Data labels visibility */
  showDataLabels?: boolean;
  /** For bar/column: stacked modes */
  stackedMode?: "none" | "stacked" | "percent";
  /** Default number format for labels/ticks (OOXML numFmt) */
  valueFormat?: string;
}
