import { ImageElement } from "../models/SlideElement";
import { escapeAttribute } from "./htmlSafety";

/**
 * Renders an image element as an absolutely positioned <img> tag.
 * @param el Image element to render.
 * @returns HTML string representing the image element.
 */
export function renderImageElement(el: ImageElement): string {
  const nf = (n: number, fb = 0) => (Number.isFinite(n) ? n : fb);
  const x = nf(el.position?.x, 0) / 9525;
  const y = nf(el.position?.y, 0) / 9525;
  const width = nf(el.size?.width, 0) / 9525;
  const height = nf(el.size?.height, 0) / 9525;
  const zIndex = Number.isFinite(el.zIndex) ? el.zIndex : 0;
  const rotation = Number.isFinite(el.rotationDeg) && el.rotationDeg ? `transform:rotate(${el.rotationDeg}deg); transform-origin:center;` : "";
  const src = escapeAttribute(el.src);
  const crop = el.crop;
  if (crop && crop.left + crop.right < 100000 && crop.top + crop.bottom < 100000) {
    const visibleX = 100000 - crop.left - crop.right;
    const visibleY = 100000 - crop.top - crop.bottom;
    const imageWidth = width * 100000 / visibleX;
    const imageHeight = height * 100000 / visibleY;
    return `<div style="position:absolute; z-index:${zIndex}; left:${x}px; top:${y}px; width:${width}px; height:${height}px; overflow:hidden; ${rotation}">
      <img src="${src}" style="position:absolute; left:${-imageWidth * crop.left / 100000}px; top:${-imageHeight * crop.top / 100000}px; width:${imageWidth}px; height:${imageHeight}px;" />
    </div>`;
  }
  return `<img src="${src}" style="
    position: absolute;
    z-index: ${zIndex};
    left: ${x}px;
    top: ${y}px;
    width: ${width}px;
    height: ${height}px;
    object-fit: fill;
    ${rotation}
  " />`;
}
