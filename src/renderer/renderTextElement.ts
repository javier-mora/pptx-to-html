import { TextElement } from "../models/SlideElement";
import { safeColor, safeFontFamily } from "./htmlSafety";

/**
 * Renders a text element as an absolutely positioned HTML <div>.
 * @param el Text element to render.
 * @returns HTML string representing the text element.
 */
export function renderTextElement(el: TextElement): string {
  const nf = (n: number, fb = 0) => (Number.isFinite(n) ? n : fb);
  const x = nf(el.position?.x, 0) / 9525;
  const y = nf(el.position?.y, 0) / 9525;
  const w = nf(el.size?.width, 0) / 9525;
  const h = nf(el.size?.height, 0) / 9525;
  const pad = el.padding || { left: 0, top: 0, right: 0, bottom: 0 };
  const textAlign = el.align?.horizontal || "left";
  const justify = el.align?.vertical === "middle" ? "center" : el.align?.vertical === "bottom" ? "flex-end" : "flex-start";
  const inner = el.html ? el.html : escape(el.content);
  return `<div style="
    position: absolute;
    left: ${x}px;
    top: ${y}px;
    width: ${w}px;
    height: ${h}px;
    box-sizing: border-box;
    z-index: ${Number.isFinite(el.zIndex) ? el.zIndex : 0};
    display: flex;
    flex-direction: column;
    justify-content: ${justify};
    text-align: ${textAlign};
    padding: ${pad.top}px ${pad.right}px ${pad.bottom}px ${pad.left}px;
    font-family: ${safeFontFamily(el.font?.name)};
    font-size: ${nf(Number(el.font?.size), 12)}pt;
    font-weight: ${el.font?.weight || "normal"};
    color: ${safeColor(el.font?.color, "#000")};
    overflow: hidden;
    white-space: pre-wrap;
  ">${inner}</div>`;
}

function escape(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
