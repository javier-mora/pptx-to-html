import { ImageElement } from "../models/SlideElement";
import { XmlHelper } from "../core/XmlHelper";
import JSZip from "jszip";
import { imageMimeType } from "../renderer/htmlSafety";

/**
 * Responsible for extracting image elements from a slide XML node.
 */
export class ImageExtractor {
  /**
   * Extracts image elements from the <spTree> element using rels from slide relationships.
   * @param spTree The <spTree> element of the slide.
   * @param rels XML Document for slide relationships (ppt/slides/_rels/slideX.xml.rels).
   * @param zip The JSZip archive of the entire .pptx file.
   * @returns List of ImageElement extracted.
   */
  static async extract(
    spTree: Element | null,
    rels: Document,
    zip: JSZip,
    basePath: string = "ppt/slides"
  ): Promise<ImageElement[]> {
    if (!spTree) return [];

    const elements: ImageElement[] = [];

    const pics = spTree.getElementsByTagNameNS("*", "pic");
    for (const pic of Array.from(pics)) {
      const blip = pic.getElementsByTagNameNS("*", "blip")[0];
      const embedId = blip?.getAttribute("r:embed") ?? "";
      if (!embedId) continue;

      const relEl = (rels && (rels as any).getElementsByTagName) ? (function(){
        const els = rels.getElementsByTagName("Relationship");
        for (const e of Array.from(els)) { if (e.getAttribute("Id") === embedId) return e as Element; }
        return null;
      })() : null;
      const relTarget = relEl?.getAttribute("Target");
      if (!relTarget) continue;

      const normalizedPath = this.normalizePath(relTarget, basePath);
      const imageFile = zip.file(normalizedPath);
      if (!imageFile) continue;

      const binary = await imageFile.async("base64");
      const dataUri = `data:${imageMimeType(normalizedPath)};base64,${binary}`;

      const xfrm = pic.getElementsByTagNameNS("*", "xfrm")[0];

      const off = xfrm?.getElementsByTagNameNS("*", "off")[0];
      const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0];

      const x = off ? XmlHelper.getAttrAsNumber(off, "x") : 0;
      const y = off ? XmlHelper.getAttrAsNumber(off, "y") : 0;

      const cx = ext ? XmlHelper.getAttrAsNumber(ext, "cx") : 1000000;
      const cy = ext ? XmlHelper.getAttrAsNumber(ext, "cy") : 500000;

      const element: ImageElement = {
        type: "image",
        zIndex: XmlHelper.getZIndex(pic, spTree),
        relId: embedId,
        src: dataUri,
        position: { x, y },
        size: { width: cx, height: cy }
      };

      elements.push(element);
    }

    return elements;
  }

  /**
   * Normalizes a relative path from a slide rels file.
   * @param target Path from the relationship XML (e.g. "../media/image1.png")
   * @param basePath Base folder (e.g. "ppt/slides")
   * @returns Normalized path inside the zip (e.g. "ppt/media/image1.png")
   */
  private static normalizePath(target: string, basePath: string): string {
    // Absolute targets (starting with "/") are relative to the package root,
    // not to the .rels file location. pptxgenjs writes absolute targets.
    if (target.startsWith("/")) {
      return target.slice(1);
    }
    const parts = (basePath + "/" + target).split("/");
    const resolved: string[] = [];
    for (const part of parts) {
      if (part === "..") resolved.pop();
      else if (part !== "." && part !== "") resolved.push(part);
    }
    return resolved.join("/");
  }
}
