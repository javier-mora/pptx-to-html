import { TableElement } from "../models/SlideElement";
import { safeColor, safeFontFamily } from "./htmlSafety";

export function renderTableElement(el: TableElement): string {
  const nf = (n: number, fb = 0) => (Number.isFinite(n) ? n : fb);
  const x = nf(el.position?.x, 0) / 9525;
  const y = nf(el.position?.y, 0) / 9525;
  const width = nf(el.size?.width, 0) / 9525;
  const height = nf(el.size?.height, 0) / 9525;

  const colWidthsPx = el.columns.map((w) => nf(w, 0) / 9525);
  const colTotal = colWidthsPx.reduce((a, b) => a + b, 0) || 1;
  const cols = colWidthsPx
    .map((w) => `<col style="width:${(w / colTotal) * 100}%">`)
    .join("");

  // Determine table-level background: prefer explicit table fill, then wholeTbl from style
  const tableBg = el.tableFillColor || el.style?.fills?.wholeTbl;

  let rowIndex = 0;
  const rowsHtml = el.rows
    .map((row) => {
      let colIndex = 0;
      const tds = row.cells
        .map((cell) => {
          const pad = cell.padding || { left: 6, top: 2, right: 6, bottom: 2 };
          const ta = cell.align?.horizontal || "left";
          const va = cell.align?.vertical || "top";
          const borderCss = computeCellBordersCSS(el, cell, rowIndex, colIndex);
          // Banding/background using theme table style if available; fall back to subtle grays
          const { bg, fontColor, emphasize } = computeCellStyleFromTableStyle(el, rowIndex, colIndex);
          const isHeaderCol = emphasize;
          const style = `
            padding:${pad.top}px ${pad.right}px ${pad.bottom}px ${pad.left}px;
            text-align:${ta};
            vertical-align:${va === "middle" ? "middle" : va};
            ${cell.fillColor ? `background-color:${safeColor(cell.fillColor)};` : ""}
            ${!cell.fillColor && bg ? `background-color:${safeColor(bg)};` : ""}
            ${cell.font?.color ? `color:${safeColor(cell.font.color, "#000")};` : (fontColor ? `color:${safeColor(fontColor, "#000")};` : "")}
            ${cell.font?.name ? `font-family:${safeFontFamily(cell.font.name)};` : ""}
            ${cell.font?.size ? `font-size:${cell.font.size}pt;` : ""}
            ${borderCss}
            ${isHeaderCol ? "font-weight:600;" : ""}
            overflow:hidden; word-break: break-word; white-space: pre-wrap;`;
          const span = `${cell.colSpan ? ` colspan=\"${cell.colSpan}\"` : ""}${cell.rowSpan ? ` rowspan=\"${cell.rowSpan}\"` : ""}`;
          const content = escape(cell.text).replace(/\n/g, "<br>");
          const html = `<td${span} style="${style}">${content}</td>`;
          colIndex += cell.colSpan || 1;
          return html;
        })
        .join("");
      const rowStyle = el.tableStyle?.firstRow && rowIndex === 0 ? ' style="font-weight:600;"' : "";
      const trHtml = `<tr${rowStyle}>${tds}</tr>`;
      rowIndex += 1;
      return trHtml;
    })
    .join("");

  return `<div style="position:absolute; z-index:${Number.isFinite(el.zIndex) ? el.zIndex : 0}; left:${x}px; top:${y}px; width:${width}px; height:${height}px;">
    <table style="border-collapse:collapse; width:100%; height:100%; table-layout:fixed;${tableBg ? ` background-color:${safeColor(tableBg)};` : ""}">
      <colgroup>${cols}</colgroup>
      <tbody>${rowsHtml}</tbody>
    </table>
  </div>`;
}

function computeCellBordersCSS(el: TableElement, cell: any, rowIndex: number, colIndex: number): string {
  const css: string[] = [];
  const sides: Array<"top" | "right" | "bottom" | "left"> = ["top", "right", "bottom", "left"];
  const apply = (side: string, b?: { color?: string; width?: number; style?: string }) => {
    if (!b) return;
    const w = b.width ?? 1;
    const c = safeColor(b.color, "#000");
    const st = b.style === "dashed" || b.style === "dotted" ? b.style : "solid";
    css.push(`border-${side}: ${Math.max(1, Math.round(w))}px ${st} ${c};`);
  };

  // Prefer explicit cell borders
  for (const s of sides) {
    const b = cell.borders?.[s];
    if (b) apply(s, b);
  }

  // Fallback to table borders
  const tb = el.tableBorders || {};
  const lastRow = rowIndex === el.rows.length - 1;
  const lastCol = colIndex === (el.columns.length - 1);
  // Outer borders
  if (!cell.borders?.top && rowIndex === 0) apply("top", tb.top);
  if (!cell.borders?.bottom && lastRow) apply("bottom", tb.bottom);
  if (!cell.borders?.left && colIndex === 0) apply("left", tb.left);
  if (!cell.borders?.right && lastCol) apply("right", tb.right);

  // Inside borders
  if (!cell.borders?.top && rowIndex > 0) apply("top", tb.insideH);
  if (!cell.borders?.left && colIndex > 0) apply("left", tb.insideV);

  return css.join(" ");
}

function escape(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function computeCellStyleFromTableStyle(
  el: TableElement,
  rowIndex: number,
  colIndex: number
): { bg?: string; fontColor?: string; emphasize: boolean } {
  const lastRow = rowIndex === el.rows.length - 1;
  const lastCol = colIndex === el.columns.length - 1;
  const s = el.style || {};
  const fills = s.fills || {};
  const fontColors = s.fontColors || {};
  let emphasize = false;

  // Precedence: firstRow/lastRow/firstCol/lastCol -> bandRow -> bandCol -> wholeTbl
  if (el.tableStyle?.firstRow && rowIndex === 0) {
    emphasize = true;
    return { bg: fills.firstRow || fills.wholeTbl || "#f0f0f0", fontColor: fontColors.firstRow || fontColors.wholeTbl, emphasize };
  }
  if (el.tableStyle?.lastRow && lastRow) {
    return { bg: fills.lastRow || fills.wholeTbl || "#f0f0f0", fontColor: fontColors.lastRow || fontColors.wholeTbl, emphasize };
  }
  if (el.tableStyle?.firstCol && colIndex === 0) {
    emphasize = true;
    return { bg: fills.firstCol || fills.wholeTbl || "#f0f0f0", fontColor: fontColors.firstCol || fontColors.wholeTbl, emphasize };
  }
  if (el.tableStyle?.lastCol && lastCol) {
    return { bg: fills.lastCol || fills.wholeTbl || "#f0f0f0", fontColor: fontColors.lastCol || fontColors.wholeTbl, emphasize };
  }

  if (el.tableStyle?.bandRow) {
    const baseIndex = el.tableStyle?.firstRow ? rowIndex - 1 : rowIndex;
    const band = baseIndex % 2 === 0 ? "band1H" : "band2H";
    return { bg: fills[band] || fills.wholeTbl || (rowIndex % 2 === 1 ? "#fafafa" : undefined), fontColor: fontColors[band] || fontColors.wholeTbl, emphasize };
  }
  if (el.tableStyle?.bandCol) {
    const band = colIndex % 2 === 0 ? "band1V" : "band2V";
    return { bg: fills[band] || fills.wholeTbl || (colIndex % 2 === 1 ? "#fafafa" : undefined), fontColor: fontColors[band] || fontColors.wholeTbl, emphasize };
  }

  return { bg: fills.wholeTbl, fontColor: fontColors.wholeTbl, emphasize };
}
