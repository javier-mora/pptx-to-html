import { SlideElement } from "../models/SlideElement";
import { renderTextElement } from "./renderTextElement";
import { renderImageElement } from "./renderImageElement";
import { renderShapeElement } from "./renderShapeElement";
import { renderTableElement } from "./renderTableElement";
import { renderChartElement } from "./renderChartElement";
import { escapeAttribute, safeColor } from "./htmlSafety";

/**
 * Converts a list of SlideElements into an HTML string with absolute positioning.
 */
export class HtmlRenderer {
  /**
   * Renders a slide to an HTML <div> with all elements positioned accordingly.
   * @param elements List of SlideElement (text, image, shape)
   * @param options Optional width and height (in px) for the slide container.
   *  - If scaleToFit=true, width/height define the outer container size and contents are scaled from the base 960x540.
   * @returns HTML string representing the slide
   */
  static render(
    elements: SlideElement[],
    options: { width?: number; height?: number; scaleToFit?: boolean; letterbox?: boolean; baseWidth?: number; baseHeight?: number } = {}
  ): string {
    // This is public JavaScript-facing API as well as TypeScript API. Avoid
    // emitting invalid CSS or divisions by zero when callers pass bad values.
    const dimension = (value: unknown, fallback: number) =>
      typeof value === "number" && Number.isFinite(value) && value > 0 ? value : fallback;
    const baseW = dimension(options.baseWidth, 960);
    const baseH = dimension(options.baseHeight, 540);
    const targetW = dimension(options.width, baseW);
    const targetH = dimension(options.height, baseH);
    const scaleToFit = options.scaleToFit === true;
    // Default letterbox to true when scaleToFit is enabled unless explicitly set to false
    const letterbox = scaleToFit ? options.letterbox !== false : options.letterbox === true;

    const htmlParts = elements.map((el) => {
      switch (el.type) {
        case "background": {
          const hasImg = Boolean((el as any).imageSrc);
          const styleBg = hasImg
            ? `background-image: url('${escapeAttribute((el as any).imageSrc)}'); background-size: cover; background-position: center; background-repeat: no-repeat;`
            : `background-color: ${safeColor((el as any).fillColor)};`;
          return `<div style="position:absolute; z-index:${Number.isFinite(el.zIndex) ? el.zIndex : -1}; left:0; top:0; width:${baseW}px; height:${baseH}px; ${styleBg}"></div>`;
        }
        case "text": return renderTextElement(el);
        case "image": return renderImageElement(el);
        case "shape": return renderShapeElement(el, { scaleStrokes: scaleToFit });
        case "table": return renderTableElement(el);
        case "chart": return renderChartElement(el as any);
        default:
          if (typeof console !== "undefined" && console.warn) {
            console.warn(`[pptx-to-html] Unsupported element type: ${(el as any)?.type}`);
          }
          return "";
      }
    });

    if (scaleToFit) {
      if (letterbox) {
        const s = Math.min(targetW / baseW, targetH / baseH);
        const offsetX = (targetW - baseW * s) / 2;
        const offsetY = (targetH - baseH * s) / 2;
        return (
          `<div class="slide-container" style="position: relative; width: ${targetW}px; height: ${targetH}px; overflow: hidden; background-color: #000;">
            <div class="slide" style="position: absolute; left: ${offsetX}px; top: ${offsetY}px; width: ${baseW}px; height: ${baseH}px; transform: scale(${s}); transform-origin: top left; background-color: #fff;">
              ${htmlParts.join("\n")}
            </div>
          </div>`
        );
      } else {
        const sx = targetW / baseW;
        const sy = targetH / baseH;
        return (
          `<div class="slide-container" style="position: relative; width: ${targetW}px; height: ${targetH}px; overflow: hidden;">
            <div class="slide" style="position: absolute; left: 0; top: 0; width: ${baseW}px; height: ${baseH}px; transform: scale(${sx}, ${sy}); transform-origin: top left; background-color: #fff;">
              ${htmlParts.join("\n")}
            </div>
          </div>`
        );
      }
    }

    return (
      `<div class="slide" style="position: relative; width: ${targetW}px; height: ${targetH}px; overflow: hidden; background-color: #fff;">
        ${htmlParts.join("\n")}
      </div>`
    );
  }
}
