import { TableElement, TableRow, TableCell } from "../models/SlideElement";
import { XmlHelper } from "../core/XmlHelper";

/**
 * Extract tables from <spTree> by scanning p:graphicFrame with a:tbl content.
 */
export class TableExtractor {
  static extract(
    spTree: Element | null,
    themeColors: Record<string, string>,
    themeTableStyles?: Record<string, { fills: Record<string, string>; fontColors: Record<string, string> }>
  ): TableElement[] {
    if (!spTree) return [];

    const tables: TableElement[] = [];
    const gFrames = spTree.getElementsByTagNameNS("*", "graphicFrame");
    for (const gf of Array.from(gFrames)) {
      const tbl = this.findTbl(gf);
      if (!tbl) continue;

      const { x, y, cx, cy, rotationDeg } = XmlHelper.getAbsoluteTransform(gf, spTree);

      const columns: number[] = [];
      const grid = tbl.getElementsByTagNameNS("*", "tblGrid")[0] ?? null;
      if (grid) {
        for (const col of Array.from(grid.getElementsByTagNameNS("*", "gridCol"))) {
          const w = Number(col.getAttribute("w") || 0);
          columns.push(w);
        }
      }

      // Read table style flags and borders
      const tblPr = tbl.getElementsByTagNameNS("*", "tblPr")[0] ?? null;
      const tableStyle = this.extractTableStyleFlags(tblPr);
      const tableBorders = this.extractTableBorders(tblPr, themeColors);
      const tableStyleId = this.extractTableStyleId(tblPr);
      const tableFillColor = this.extractTableFill(tblPr, themeColors);
      const resolvedStyle = tableStyleId && themeTableStyles ? themeTableStyles[tableStyleId] : undefined;
      const fallbackStyle = this.buildFallbackTableStyle(themeColors, tableStyle);
      const mergedStyle = fallbackStyle || resolvedStyle
        ? {
            fills: { ...(fallbackStyle?.fills || {}), ...(resolvedStyle?.fills || {}) },
            fontColors: { ...(fallbackStyle?.fontColors || {}), ...(resolvedStyle?.fontColors || {}) },
          }
        : undefined;

      const rows: TableRow[] = [];
      for (const tr of Array.from(tbl.getElementsByTagNameNS("*", "tr"))) {
        const cells: TableCell[] = [];
        for (const tc of Array.from(tr.getElementsByTagNameNS("*", "tc"))) {
          const tcPr = tc.getElementsByTagNameNS("*", "tcPr")[0] ?? null;
          // OOXML emits a physical continuation cell for merged ranges. It
          // must not become a second HTML <td>; the originating cell already
          // carries gridSpan/rowSpan.
          const isMergedContinuation =
            tcPr?.getAttribute("hMerge") === "1" || tcPr?.getAttribute("hMerge") === "true"
            || tcPr?.getAttribute("vMerge") === "1" || tcPr?.getAttribute("vMerge") === "true"
            || !!tcPr?.getElementsByTagNameNS("*", "hMerge")[0]
            || !!tcPr?.getElementsByTagNameNS("*", "vMerge")[0];
          if (isMergedContinuation) continue;
          const txBody = tc.getElementsByTagNameNS("*", "txBody")[0] ?? null;
          const { text, font, align, padding } = this.extractCellText(txBody, themeColors, tc);

          const fillColor = this.extractFillColor(tcPr, themeColors);
          const borders = this.extractCellBorders(tcPr, themeColors);
          const cell: TableCell = { text, font, align, padding, fillColor, borders };
          // Spans (basic)
          const gridSpan = tcPr?.getElementsByTagNameNS("*", "gridSpan")[0] ?? null;
          const rowSpan = tcPr?.getElementsByTagNameNS("*", "rowSpan")[0] ?? null;
          if (gridSpan) cell.colSpan = Number(gridSpan.getAttribute("val") || 1);
          if (rowSpan) cell.rowSpan = Number(rowSpan.getAttribute("val") || 1);
          cells.push(cell);
        }
        rows.push({ cells });
      }

      tables.push({
        type: "table",
        zIndex: XmlHelper.getZIndex(gf, spTree),
        rotationDeg,
        position: { x, y },
        size: { width: cx, height: cy },
        columns,
        rows,
        tableStyle,
        tableBorders,
        tableStyleId: tableStyleId,
        tableFillColor,
        style: mergedStyle ? { fills: mergedStyle.fills, fontColors: mergedStyle.fontColors } : undefined,
      });
    }

    return tables;
  }

  private static buildFallbackTableStyle(
    themeColors: Record<string, string>,
    tableStyleFlags: any
  ): { fills: Record<string, string>; fontColors: Record<string, string> } | undefined {
    // Use accent1 as primary color for header and bands if present.
    const accent = themeColors["accent1"] || themeColors["accent2"] || themeColors["dk1"];
    if (!accent) return undefined;
    const white = themeColors["lt1"] || "#FFFFFF";
    const fills: Record<string, string> = {};
    const fontColors: Record<string, string> = {};

    // Header (firstRow) typical: solid accent background, light text
    if (tableStyleFlags?.firstRow) {
      fills["firstRow"] = accent;
      fontColors["firstRow"] = white;
    }

    // First column emphasis (lighter accent)
    if (tableStyleFlags?.firstCol) {
      fills["firstCol"] = this.lightenHex(accent, 0.85);
      fontColors["firstCol"] = undefined as any;
    }

    // Banding: typical light shade on alternate rows/cols
    const bandShade = this.lightenHex(accent, 0.92);
    if (tableStyleFlags?.bandRow) {
      // mimic Light Style: band1H no fill, band2H shaded
      fills["band2H"] = bandShade;
    }
    if (tableStyleFlags?.bandCol) {
      fills["band2V"] = bandShade;
    }

    return { fills, fontColors };
  }

  private static lightenHex(hex: string, ratio: number): string {
    const m = (hex || "").replace("#", "");
    if (m.length !== 6 || /[^0-9a-fA-F]/.test(m)) return hex || "#FFFFFF";
    const r = parseInt(m.substring(0, 2), 16);
    const g = parseInt(m.substring(2, 4), 16);
    const b = parseInt(m.substring(4, 6), 16);
    const lr = Math.round(r + (255 - r) * ratio);
    const lg = Math.round(g + (255 - g) * ratio);
    const lb = Math.round(b + (255 - b) * ratio);
    const to2 = (n: number) => n.toString(16).padStart(2, "0");
    return `#${to2(lr)}${to2(lg)}${to2(lb)}`;
  }

  private static findTbl(gf: Element): Element | null {
    // a:graphic/a:graphicData/a:tbl (URI usually main/table)
    const graphicData = gf.getElementsByTagNameNS("*", "graphicData")[0] ?? null;
    if (!graphicData) return null;
    const tbl = graphicData.getElementsByTagNameNS("*", "tbl")[0] ?? null;
    return tbl;
  }

  private static extractCellText(txBody: Element | null, themeColors: Record<string, string>, tc?: Element | null) {
    if (!txBody) return { text: "", font: {}, align: {}, padding: { left: 0, top: 0, right: 0, bottom: 0 } };
    const bodyPr = txBody.getElementsByTagNameNS("*", "bodyPr")[0] ?? null;
    // Vertical alignment: prefer tcPr@anchor if present, else bodyPr@anchor
    let vertical: "top" | "middle" | "bottom" = "top";
    const tcPr = tc?.getElementsByTagNameNS("*", "tcPr")[0] ?? null;
    const tcAnchor = tcPr?.getAttribute("anchor") || tcPr?.getAttribute("vAlign") || undefined;
    const bpAnchor = bodyPr?.getAttribute("anchor") || undefined; // t|ctr|b
    const vSrc = tcAnchor || bpAnchor;
    if (vSrc === "ctr") vertical = "middle";
    else if (vSrc === "b") vertical = "bottom";
    else vertical = "top";
    const lIns = bodyPr?.getAttribute("lIns");
    const tIns = bodyPr?.getAttribute("tIns");
    const rIns = bodyPr?.getAttribute("rIns");
    const bIns = bodyPr?.getAttribute("bIns");
    const padding = {
      left: lIns ? Number(lIns) / 9525 : 6,
      top: tIns ? Number(tIns) / 9525 : 2,
      right: rIns ? Number(rIns) / 9525 : 6,
      bottom: bIns ? Number(bIns) / 9525 : 2,
    };
    let horiz: "left" | "center" | "right" | "justify" = "left";
    let fontName = "Arial";
    let fontSize = 14; // pt
    let color: string | undefined;
    let parts: string[] = [];

    const paragraphs = txBody.getElementsByTagNameNS("*", "p");
    for (const p of Array.from(paragraphs)) {
      const pPr = p.getElementsByTagNameNS("*", "pPr")[0] ?? null;
      const algn = pPr?.getAttribute("algn") || undefined;
      if (algn) {
        horiz = algn === "ctr" ? "center" : algn === "r" ? "right" : algn.startsWith("just") ? "justify" : "left";
      }
      // collect text including breaks
      for (const child of Array.from(p.childNodes) as any[]) {
        if (child.nodeType === 1) {
          const ln = (child as Element).localName;
          if (ln === "r") {
            const rPr = (child as Element).getElementsByTagNameNS("*", "rPr")[0] ?? null;
            if (rPr) {
              const latin = rPr.getElementsByTagNameNS("*", "latin")[0] ?? null;
              fontName = latin?.getAttribute("typeface") ?? fontName;
              const sz = rPr.getAttribute("sz");
              if (sz) fontSize = parseInt(sz, 10) / 100;
              const solidFill = rPr.querySelector("*|solidFill");
              const c = XmlHelper.getColorFromElement(solidFill || null, themeColors);
              if (c) color = c;
            }
            const t = (child as Element).getElementsByTagNameNS("*", "t")[0]?.textContent ?? "";
            parts.push(t);
          } else if (ln === "br") {
            parts.push("\n");
          } else if (ln === "fld") {
            const runs = (child as Element).getElementsByTagNameNS("*", "r");
            for (const r of Array.from(runs)) {
              const t = r.getElementsByTagNameNS("*", "t")[0]?.textContent ?? "";
              parts.push(t);
            }
          }
        }
      }
    }

    return {
      text: parts.join("").trim(),
      font: { name: fontName, size: fontSize, color: color },
      align: { horizontal: horiz, vertical },
      padding,
    };
  }

  private static extractFillColor(tcPr: Element | null, themeColors: Record<string, string>): string | undefined {
    if (!tcPr) return undefined;
    // If explicit noFill, honor transparency
    if (tcPr.getElementsByTagNameNS("*", "noFill")[0]) return undefined;

    // Prefer a direct child solidFill under tcPr (cell background)
    const direct = Array.from(tcPr.children).find((c) => (c as Element).localName === "solidFill") as Element | undefined;
    if (direct) return XmlHelper.getColorFromElement(direct, themeColors);

    // Otherwise, search for a solidFill not inside any line/border nodes
    const all = Array.from(tcPr.getElementsByTagNameNS("*", "solidFill"));
    for (const cand of all) {
      let p: Element | null = cand.parentElement;
      let insideLine = false;
      while (p && p !== tcPr) {
        const ln = p.localName;
        if (ln === "ln" || ln === "lnL" || ln === "lnR" || ln === "lnT" || ln === "lnB" || ln === "tcBorders") {
          insideLine = true; break;
        }
        p = p.parentElement;
      }
      if (!insideLine) {
        const col = XmlHelper.getColorFromElement(cand, themeColors);
        if (col) return col;
      }
    }
    return undefined;
  }

  private static extractCellBorders(tcPr: Element | null, themeColors: Record<string, string>) {
    const borders: any = {};
    if (!tcPr) return borders;
    const map: Record<string, keyof typeof borders> = {
      lnL: "left",
      lnR: "right",
      lnT: "top",
      lnB: "bottom",
    } as any;
    for (const key of Object.keys(map)) {
      const ln = tcPr.getElementsByTagNameNS("*", key)[0] ?? null;
      if (!ln) continue;
      const wAttr = ln.getAttribute("w");
      const w = wAttr ? Number(wAttr) / 9525 : undefined;
      const solidFill = ln.getElementsByTagNameNS("*", "solidFill")[0] ?? null;
      const color = XmlHelper.getColorFromElement(solidFill, themeColors);
      // Line style mapping (prstDash)
      const prstDash = ln.getElementsByTagNameNS("*", "prstDash")[0] ?? null;
      const dashVal = prstDash?.getAttribute("val") || "";
      const style = this.mapPrstDashToCss(dashVal);
      borders[map[key]] = { color, width: w, style };
    }
    return borders;
  }

  private static extractTableStyleFlags(tblPr: Element | null) {
    if (!tblPr) return {};
    const flags: any = {};
    for (const k of ["firstRow", "firstCol", "lastRow", "lastCol", "bandRow", "bandCol"]) {
      if (tblPr.getAttribute(k) === "1" || tblPr.getAttribute(k) === "true") flags[k] = true;
    }
    return flags;
  }

  private static extractTableStyleId(tblPr: Element | null): string | undefined {
    if (!tblPr) return undefined;
    // Try a:tblStyleId or a:tblStyle or direct attribute
    const idEl = tblPr.getElementsByTagNameNS("*", "tblStyleId")[0] || tblPr.getElementsByTagNameNS("*", "tblStyle")[0] || null;
    const idAttr = idEl?.getAttribute("val") || undefined;
    const text = idEl?.textContent?.trim() || undefined;
    const direct = tblPr.getAttribute("tblStyle") || undefined;
    return idAttr || text || direct || undefined;
  }

  private static extractTableBorders(tblPr: Element | null, themeColors: Record<string, string>) {
    if (!tblPr) return undefined;
    const borders: any = {};
    const tblBorders = tblPr.getElementsByTagNameNS("*", "tblBorders")[0] ?? null;
    if (!tblBorders) return undefined;
    const map: Record<string, string> = {
      top: "top",
      right: "right",
      bottom: "bottom",
      left: "left",
      insideH: "insideH",
      insideV: "insideV",
    };
    for (const tag of Object.keys(map)) {
      const node = tblBorders.getElementsByTagNameNS("*", tag)[0] ?? null;
      if (!node) continue;
      const ln = node.getElementsByTagNameNS("*", "ln")[0] ?? node; // sometimes border node directly has ln props
      const wAttr = ln.getAttribute("w");
      const w = wAttr ? Number(wAttr) / 9525 : undefined;
      const solidFill = ln.getElementsByTagNameNS("*", "solidFill")[0] ?? null;
      const color = XmlHelper.getColorFromElement(solidFill, themeColors);
      const prstDash = ln.getElementsByTagNameNS("*", "prstDash")[0] ?? null;
      const dashVal = prstDash?.getAttribute("val") || "";
      const style = this.mapPrstDashToCss(dashVal);
      borders[map[tag]] = { color, width: w, style };
    }
    return borders;
  }

  private static extractTableFill(tblPr: Element | null, themeColors: Record<string, string>): string | undefined {
    if (!tblPr) return undefined;
    // 1) Prefer a solidFill that is a direct child of tblPr (table-level background)
    const directSolid = Array.from(tblPr.children).find((c) => (c as Element).localName === "solidFill") as Element | undefined;
    if (directSolid) return XmlHelper.getColorFromElement(directSolid, themeColors);

    // 2) Otherwise, search for a solidFill under tblPr that is NOT inside tblBorders
    const allSolid = Array.from(tblPr.getElementsByTagNameNS("*", "solidFill"));
    for (const cand of allSolid) {
      let p: Element | null = cand.parentElement;
      let insideBorders = false;
      while (p && p !== tblPr) {
        if (p.localName === "tblBorders") { insideBorders = true; break; }
        p = p.parentElement;
      }
      if (!insideBorders) {
        const col = XmlHelper.getColorFromElement(cand, themeColors);
        if (col) return col;
      }
    }
    return undefined;
  }

  private static mapPrstDashToCss(val: string | undefined): "solid" | "dashed" | "dotted" | undefined {
    const v = (val || "").toLowerCase();
    if (!v) return undefined;
    if (v === "solid") return "solid";
    if (v.includes("dot")) return "dotted"; // dot, sysDot, dashDot, etc. -> dotted/dashed heuristic
    if (v.includes("dash")) return "dashed"; // dash, lgDash, sysDash, etc.
    return undefined;
  }
}
