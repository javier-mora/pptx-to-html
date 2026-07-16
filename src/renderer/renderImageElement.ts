import { ImageElement } from "../models/SlideElement";
import { escapeAttribute } from "./htmlSafety";

/**
 * Renders an image element as an absolutely positioned <img> tag.
 * @param el Image element to render.
 * @returns HTML string representing the image element.
 */
export function renderImageElement(el: ImageElement): string {
  const nf = (n: number, fb = 0) => (Number.isFinite(n) ? n : fb);
  return `<img src="${escapeAttribute(el.src)}" style="
    position: absolute;
    z-index: ${Number.isFinite(el.zIndex) ? el.zIndex : 0};
    left: ${nf(el.position?.x, 0) / 9525}px;
    top: ${nf(el.position?.y, 0) / 9525}px;
    width: ${nf(el.size?.width, 0) / 9525}px;
    height: ${nf(el.size?.height, 0) / 9525}px;
    object-fit: cover;
  " />`;
}
