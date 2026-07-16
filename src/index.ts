import { PptxReader } from "./core/PptxReader";
import { XmlHelper } from "./core/XmlHelper";
import { HtmlRenderer } from "./renderer/HtmlRenderer";

/**
 * Converts a PPTX file buffer into an array of HTML slides.
 * @param buffer ArrayBuffer representing the .pptx file.
 * @param config Optional rendering configuration object.
 * @param config.width Target container width in pixels (defaults to 960 when not provided).
 * @param config.height Target container height in pixels (defaults to 540 when not provided).
 * @param config.scaleToFit When true, scales the slide content to fit the container size.
 * @param config.letterbox When scaling, use black bars to preserve aspect ratio (defaults to true when scaleToFit is true).
 * @returns Array of HTML strings, each representing one slide.
 */
export async function pptxToHtml(
  buffer: ArrayBuffer,
  config?: { width?: number; height?: number; scaleToFit?: boolean; letterbox?: boolean; domParserFactory?: () => { parseFromString(xml: string, mime: string): Document } }
): Promise<string[]> {
  // Optional DOM parser injection for Node environments without global DOMParser
  if (config?.domParserFactory) {
    XmlHelper.setDomParser(config.domParserFactory as any);
  }
  const reader = new PptxReader();
  const slides = await reader.load(buffer);
  const base = await reader.getBaseSizePx();
  const opts = { ...(config || {}), baseWidth: base.width, baseHeight: base.height } as any;
  return slides.map((slideElements) => HtmlRenderer.render(slideElements, opts));
}
