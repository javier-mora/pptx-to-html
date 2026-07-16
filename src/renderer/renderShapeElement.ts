import { ShapeElement } from "../models/SlideElement";
import { getSvgPathForShape } from "./shapePathMap";
import { safeColor } from "./htmlSafety";

/**
 * Renders a shape element as an absolutely positioned HTML or SVG element.
 * Supports all recognized PPTX shape types using SVG when necessary.
 * @param el Shape element to render.
 * @returns HTML string representing the shape.
 */
export function renderShapeElement(el: ShapeElement, options: { scaleStrokes?: boolean } = {}): string {
    const nf = (n: number, fb = 0) => (Number.isFinite(n) ? n : fb);
    const x = nf(el.position?.x, 0) / 9525;
    const y = nf(el.position?.y, 0) / 9525;
    const width = nf(el.size?.width, 0) / 9525;
    const height = nf(el.size?.height, 0) / 9525;

    const rotation = el.rotationDeg && !isNaN(el.rotationDeg) ? el.rotationDeg : 0;
    const rotationStyle = rotation ? `transform: rotate(${rotation}deg); transform-origin: center;` : "";

    const fillColor = safeColor(el.fillColor);
    const borderColor = safeColor(el.borderColor);
    const borderWidth = el.strokeWidth && Number.isFinite(el.strokeWidth) ? el.strokeWidth : 1;
    const style = `
    position: absolute;
    left: ${x}px;
    top: ${y}px;
    width: ${width}px;
    height: ${height}px;
    z-index: ${Number.isFinite(el.zIndex) ? el.zIndex : 0};
    ${rotationStyle}
  `;

    // Basic HTML shapes
    if (el.shapeType === "rect") {
        return `<div style="${style}
      background-color: ${fillColor};
      border: ${borderWidth}px solid ${borderColor};
      box-sizing: border-box;"></div>`;
    }

    if (el.shapeType === "ellipse") {
        return `<div style="${style}
      background-color: ${fillColor};
      border: ${borderWidth}px solid ${borderColor};
      border-radius: 50%;
      box-sizing: border-box;"></div>`;
    }

    if (el.shapeType === "roundRect") {
        return `<div style="${style}
      background-color: ${fillColor};
      border: ${borderWidth}px solid ${borderColor};
      border-radius: 16px;
      box-sizing: border-box;"></div>`;
    }

    // SVG-based shapes using prefixed definition
    const raw = getSvgPathForShape(el.shapeType);
    return shapeSvg(
      x,
      y,
      width,
      height,
      fillColor,
      borderColor,
      raw,
      el.strokeWidth && Number.isFinite(el.strokeWidth) ? el.strokeWidth : undefined,
      rotation,
      el.headEnd,
      el.tailEnd,
      options.scaleStrokes === true,
      Number.isFinite(el.zIndex) ? el.zIndex : 0
    );
}

function shapeSvg(
  x: number,
  y: number,
  width: number,
  height: number,
  fill: string,
  stroke: string | undefined,
  raw: string,
  strokeWidthPx?: number,
  rotationDeg?: number,
  headEnd?: { type?: string; w?: string; len?: string },
  tailEnd?: { type?: string; w?: string; len?: string },
  scaleStrokes?: boolean,
  zIndex = 0
): string {
  const strokeColorOpt = resolveStrokeColor(stroke, fill);
  const [typeRaw, ...rest] = raw.trim().split(/\s+/);
  const type = typeRaw.toUpperCase().replace("_ARROW", "");
  const isArrow = typeRaw.endsWith("_ARROW");
  const data = rest.join(" ");

  const svgHeight = height;
  const svgWidth = width;
  const sw = strokeWidthPx && strokeWidthPx > 0 ? strokeWidthPx : 2;

  const rotationStyle = rotationDeg ? `transform: rotate(${rotationDeg}deg); transform-origin: center;` : "";
  const commonStyle = `
    position: absolute;
    left: ${x}px;
    top: ${y}px;
    width: ${svgWidth}px;
    height: ${svgHeight}px;
    z-index: ${zIndex};
    ${rotationStyle}
  `;

  switch (type) {
    case "PATH": {
      const defs = buildMarkerDefs(headEnd, tailEnd, strokeColorOpt || "#000");
      const markerStartAttr = defs.startId ? `marker-start=\"url(#${defs.startId})\"` : "";
      const markerEndAttr = defs.endId ? `marker-end=\"url(#${defs.endId})\"` : "";
      return `<svg viewBox="0 0 100 100" style="${commonStyle}" overflow="visible">
        ${defs.defs}
        <path d="${data}" fill="none" stroke="${strokeColorOpt || "#000"}" stroke-width="${sw}" ${scaleStrokes ? "" : "vector-effect=\"non-scaling-stroke\""} ${markerStartAttr} ${markerEndAttr} />
      </svg>`;
    }

    case "POLYLINE":
    case "LINE": {
      const coords = data
        .split(/[\s,]+/)
        .map((v) => parseFloat(v))
        .filter((v) => !isNaN(v));

      if (coords.length < 4 || coords.length % 2 !== 0) {
        if (typeof console !== "undefined" && console.warn) {
          console.warn(`[pptx-to-html] Invalid POLYLINE/LINE shape data: "${data}"`);
        }
        return "";
      }

      // "x1,y1 x2,y2 x3,y3 ..."
      const pointPairs: string[] = [];
      for (let i = 0; i < coords.length; i += 2) {
        pointPairs.push(`${coords[i]},${coords[i + 1]}`);
      }

      // Dimensiones mínimas para que el trazo no se pierda
      const effectiveWidth = width > 0 ? width : Math.max(sw * 2, 2);
      const effectiveHeight = height > 0 ? height : Math.max(sw * 2, 2);

      const defs = buildMarkerDefs(headEnd, tailEnd ?? (isArrow ? { type: "triangle", w: "med", len: "med" } : undefined), strokeColorOpt || "#000");
      const markerStartAttr = defs.startId ? `marker-start=\"url(#${defs.startId})\"` : "";
      const markerEndAttr = defs.endId ? `marker-end=\"url(#${defs.endId})\"` : "";

      // Escalar puntos a píxeles directamente para evitar distorsión
      const scaledPairs: string[] = [];
      for (let i = 0; i < coords.length; i += 2) {
        const px = (coords[i] / 100) * effectiveWidth;
        const py = (coords[i + 1] / 100) * effectiveHeight;
        scaledPairs.push(`${px},${py}`);
      }
      const scaledPoints = scaledPairs.join(" ");

      return `
        <svg viewBox="0 0 ${effectiveWidth} ${effectiveHeight}"
            style="
              position: absolute;
              left: ${x}px;
              top: ${y}px;
              width: ${effectiveWidth}px;
              height: ${effectiveHeight}px;
              z-index: ${zIndex};
              ${rotationStyle}
            "
            overflow="visible">
          ${defs.defs}
          <polyline points="${scaledPoints}"
                    fill="none"
                    stroke="${strokeColorOpt || "#000"}"
                    stroke-width="${sw}"
                    ${scaleStrokes ? "" : "vector-effect=\"non-scaling-stroke\""}
                    ${markerStartAttr} ${markerEndAttr} />
        </svg>`;
    }

    case "POLYGON":
    default:
      return `<svg viewBox="0 0 100 100" style="${commonStyle}">
        <polygon points="${data}" fill="${fill}" stroke="${strokeColorOpt ?? "none"}" stroke-width="${sw}" ${scaleStrokes ? "" : "vector-effect=\"non-scaling-stroke\""} />
      </svg>`;
  }
}

function buildMarkerDefs(
  headEnd: { type?: string; w?: string; len?: string } | undefined,
  tailEnd: { type?: string; w?: string; len?: string } | undefined,
  color: string
): { defs: string; startId?: string; endId?: string } {
  const parts: string[] = [];
  let startId: string | undefined;
  let endId: string | undefined;

  if (headEnd && headEnd.type && headEnd.type !== "none") {
    startId = `mstart-${Math.random().toString(36).slice(2, 8)}`;
    parts.push(markerDef(startId, headEnd, color));
  }
  if (tailEnd && tailEnd.type && tailEnd.type !== "none") {
    endId = `mend-${Math.random().toString(36).slice(2, 8)}`;
    parts.push(markerDef(endId, tailEnd, color));
  }

  return { defs: parts.length ? `<defs>${parts.join("\n")}</defs>` : "", startId, endId };
}

function markerDef(id: string, spec: { type?: string; w?: string; len?: string }, color: string): string {
  const sizeFactor = mapLen(spec.len);
  const base = 4 * sizeFactor; // base marker box size
  const refX = base; // tip at end
  const refY = base / 2;

  switch ((spec.type || "triangle").toLowerCase()) {
    case "diamond":
      return `<marker id="${id}" markerUnits="strokeWidth" markerWidth="${base}" markerHeight="${base}"
                      refX="${refX}" refY="${refY}" orient="auto-start-reverse">
                <polygon points="${base/2},0 ${base},${base/2} ${base/2},${base} 0,${base/2}" fill="${color}" />
              </marker>`;
    case "oval":
      return `<marker id="${id}" markerUnits="strokeWidth" markerWidth="${base}" markerHeight="${base}"
                      refX="${refX}" refY="${refY}" orient="auto-start-reverse">
                <circle cx="${base/2}" cy="${base/2}" r="${base/2}" fill="${color}" />
              </marker>`;
    case "stealth":
      return `<marker id="${id}" markerUnits="strokeWidth" markerWidth="${base}" markerHeight="${base}"
                      refX="${refX}" refY="${refY}" orient="auto-start-reverse">
                <polygon points="${base},${base/2} 0,0 0,${base}" fill="${color}" />
              </marker>`;
    case "arrow":
    case "triangle":
    default:
      return `<marker id="${id}" markerUnits="strokeWidth" markerWidth="${base}" markerHeight="${base}"
                      refX="${refX}" refY="${refY}" orient="auto-start-reverse">
                <polygon points="0,0 ${base},${base/2} 0,${base}" fill="${color}" />
              </marker>`;
  }
}

function mapLen(len?: string): number {
  switch ((len || "med").toLowerCase()) {
    case "sm":
    case "small":
      return 1.5;
    case "lg":
    case "large":
      return 2.5;
    case "med":
    case "medium":
    default:
      return 2;
  }
}

function resolveStrokeColor(stroke?: string, fill?: string): string | undefined {
  // Prefer explicit border color if present and not transparent
  if (stroke && stroke !== "transparent") return stroke;
  // Fallback to fill if it's a solid color
  if (fill && fill !== "transparent") return fill;
  return undefined;
}
