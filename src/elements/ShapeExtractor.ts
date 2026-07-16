import { ShapeElement } from "../models/SlideElement";
import { XmlHelper } from "../core/XmlHelper";

/**
 * Responsible for extracting shape elements (including connectors) from a slide XML node.
 */
export class ShapeExtractor {
  /**
   * Extracts shape and connector elements from the <spTree> element of the slide.
   * @param spTree The <spTree> element.
   * @param themeColors Theme color mapping.
   * @returns List of ShapeElement extracted.
   */
  static extract(spTree: Element | null, themeColors: Record<string, string>): ShapeElement[] {
    if (!spTree) return [];

    const elements: ShapeElement[] = [];

    const allShapes = [
      ...Array.from(spTree.getElementsByTagNameNS("*", "sp")),
      ...Array.from(spTree.getElementsByTagNameNS("*", "cxnSp"))
    ];

    for (const shape of allShapes) {
      const xfrm = shape.getElementsByTagNameNS("*", "xfrm")[0];
      const off = xfrm?.getElementsByTagNameNS("*", "off")[0];
      const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0];

      const x = off ? XmlHelper.getAttrAsNumber(off, "x") : 0;
      const y = off ? XmlHelper.getAttrAsNumber(off, "y") : 0;
      const cx = ext ? XmlHelper.getAttrAsNumber(ext, "cx") : 1000000;
      const cy = ext ? XmlHelper.getAttrAsNumber(ext, "cy") : 500000;
      const rotAttr = xfrm?.getAttribute("rot");
      const rotationDeg = rotAttr ? Number(rotAttr) / 60000 : undefined;

      const prstGeom = shape.getElementsByTagNameNS("*", "prstGeom")[0];
      const shapeType = prstGeom?.getAttribute("prst") ?? "rect";

      const spPr = shape.getElementsByTagNameNS("*", "spPr")[0];

      let fillColor = "transparent";
      let borderColor = "transparent";
      let strokeWidth: number | undefined = undefined;
      let headEnd: { type?: string; w?: string; len?: string } | undefined = undefined;
      let tailEnd: { type?: string; w?: string; len?: string } | undefined = undefined;

      if (spPr) {
        const solidFill = spPr.getElementsByTagNameNS("*", "solidFill")[0] ?? null;
        fillColor = XmlHelper.getColorFromElement(solidFill, themeColors) ?? "transparent";

        const ln = spPr.getElementsByTagNameNS("*", "ln")[0];
        const borderFill = ln?.getElementsByTagNameNS("*", "solidFill")[0] ?? null;
        borderColor = XmlHelper.getColorFromElement(borderFill, themeColors) ?? "transparent";

        // Extract line width (w) in EMUs and convert to px if present
        const wAttr = ln?.getAttribute("w");
        if (wAttr) {
          const w = Number(wAttr);
          if (!isNaN(w)) {
            strokeWidth = w / 9525; // EMUs to px (approx at 96dpi)
          }
        }

        // Arrowheads: <a:headEnd> and <a:tailEnd> with attributes type, w, len
        const headEndEl = ln?.getElementsByTagNameNS("*", "headEnd")[0] ?? null;
        const tailEndEl = ln?.getElementsByTagNameNS("*", "tailEnd")[0] ?? null;

        headEnd = headEndEl
          ? {
              type: headEndEl.getAttribute("type") || undefined,
              w: headEndEl.getAttribute("w") || undefined,
              len: headEndEl.getAttribute("len") || undefined,
            }
          : undefined;

        tailEnd = tailEndEl
          ? {
              type: tailEndEl.getAttribute("type") || undefined,
              w: tailEndEl.getAttribute("w") || undefined,
              len: tailEndEl.getAttribute("len") || undefined,
            }
          : undefined;
      }

      if (fillColor === "transparent") {
        const style = shape.getElementsByTagNameNS("*", "style")[0];
        const fillRef = style?.getElementsByTagNameNS("*", "fillRef")[0];
        const schemeClr = fillRef?.getElementsByTagNameNS("*", "schemeClr")[0];
        const val = schemeClr?.getAttribute("val");
        if (val && themeColors[val]) {
          fillColor = themeColors[val];
        }
      }

      const element: ShapeElement = {
        type: "shape",
        zIndex: XmlHelper.getZIndex(shape, spTree),
        shapeType,
        position: { x, y },
        size: { width: cx, height: cy },
        fillColor,
        borderColor,
        strokeWidth,
        rotationDeg,
        headEnd,
        tailEnd,
      };

      elements.push(element);
    }

    return elements;
  }
}
