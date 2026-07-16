import JSZip from "jszip";
import { XmlHelper } from "./XmlHelper";
import { TextExtractor } from "../elements/TextExtractor";
import { ImageExtractor } from "../elements/ImageExtractor";
import { ShapeExtractor } from "../elements/ShapeExtractor";
import { TableExtractor } from "../elements/TableExtractor";
import { ChartExtractor } from "../elements/ChartExtractor";
import { SlideElement } from "../models/SlideElement";

/**
 * Responsible for extracting all slides from the .pptx file as lists of SlideElement.
 */
export class SlideExtractor {
  constructor(private zip: JSZip) {}

  /**
   * Extracts all slides in order and parses their visual elements.
   * @returns An array of SlideElement lists (one per slide).
   */
  async extractSlides(): Promise<SlideElement[][]> {
    // Load theme colors
    const themeXmlStr = await this.zip.file("ppt/theme/theme1.xml")?.async("string");
    const themeXml = themeXmlStr ? XmlHelper.parseXml(themeXmlStr) : null;
    const themeColors = XmlHelper.extractThemeColors(themeXml);
    const themeTableStyles = XmlHelper.extractThemeTableStyles(themeXml);

    const slidePaths = Object.keys(this.zip.files)
      .filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f))
      .sort((a, b) => {
        const numA = parseInt(a.match(/slide(\d+)\.xml$/)?.[1] || "0", 10);
        const numB = parseInt(b.match(/slide(\d+)\.xml$/)?.[1] || "0", 10);
        return numA - numB;
      });

    const slides: SlideElement[][] = [];

    for (const slidePath of slidePaths) {
      const slideXmlStr = await this.zip.file(slidePath)!.async("string");
      const slideXml = XmlHelper.parseXml(slideXmlStr);

      // Intentar encontrar el spTree usando namespace-tolerancia
      const spTree =
        slideXml.querySelector("p\\:spTree") ||
        slideXml.getElementsByTagNameNS("*", "spTree")[0];

      if (!spTree) {
        console.warn(`Warning: no <spTree> found in ${slidePath}`);
        slides.push([]);
        continue;
      }

      const relsPath = slidePath.replace("slides/", "slides/_rels/") + ".rels";
      const relsXml = this.zip.file(relsPath)
        ? XmlHelper.parseXml(await this.zip.file(relsPath)!.async("string"))
        : XmlHelper.parseXml(`<Relationships/>`);

      // Resolve layout from slide rels
      const layoutRel = XmlHelper.findRelationshipByTypeSuffix(relsXml, "/slideLayout");
      const layoutTarget = layoutRel?.getAttribute("Target") || undefined;
      let layoutSpTree: Element | null = null;
      let layoutRelsXml: Document | null = null;
      if (layoutTarget) {
        const layoutPath = this.resolvePath(layoutTarget, "ppt/slides");
        const layoutXmlStr = await this.zip.file(layoutPath)?.async("string");
        if (layoutXmlStr) {
          const layoutXml = XmlHelper.parseXml(layoutXmlStr);
          layoutSpTree = layoutXml.querySelector("p\\:spTree") || layoutXml.getElementsByTagNameNS("*", "spTree")[0] || null;
          const layoutRelsPath = layoutPath.replace("slideLayouts/", "slideLayouts/_rels/") + ".rels";
          layoutRelsXml = this.zip.file(layoutRelsPath)
            ? XmlHelper.parseXml(await this.zip.file(layoutRelsPath)!.async("string"))
            : XmlHelper.parseXml(`<Relationships/>`);
        }
      }

      // Resolve master from layout rels
      let masterSpTree: Element | null = null;
      let masterRelsXml: Document | null = null;
      if (layoutRelsXml) {
        const masterRel = XmlHelper.findRelationshipByTypeSuffix(layoutRelsXml, "/slideMaster");
        const masterTarget = masterRel?.getAttribute("Target") || undefined;
        if (masterTarget) {
          const masterPath = this.resolvePath(masterTarget, "ppt/slideLayouts");
          const masterXmlStr = await this.zip.file(masterPath)?.async("string");
          if (masterXmlStr) {
            const masterXml = XmlHelper.parseXml(masterXmlStr);
            masterSpTree = masterXml.querySelector("p\\:spTree") || masterXml.getElementsByTagNameNS("*", "spTree")[0] || null;
            const masterRelsPath = masterPath.replace("slideMasters/", "slideMasters/_rels/") + ".rels";
            masterRelsXml = this.zip.file(masterRelsPath)
              ? XmlHelper.parseXml(await this.zip.file(masterRelsPath)!.async("string"))
              : XmlHelper.parseXml(`<Relationships/>`);
          }
        }
      }

      // Extract background from slide/layout/master (slide overrides layout overrides master)
      const slideBg = await this.extractBackground(slideXml, relsXml, "ppt/slides", this.zip, themeColors);
      const layoutBg = layoutRelsXml ? await this.extractBackground(layoutSpTree?.ownerDocument || null, layoutRelsXml, "ppt/slideLayouts", this.zip, themeColors) : null;
      const masterBg = masterRelsXml ? await this.extractBackground(masterSpTree?.ownerDocument || null, masterRelsXml, "ppt/slideMasters", this.zip, themeColors) : null;
      const bgElement = slideBg || layoutBg || masterBg;

      // Extract elements from master → layout → slide (respecting z-order: back to front)
      const masterText = masterSpTree ? TextExtractor.extract(masterSpTree, themeColors, { context: "master" }) : [];
      const masterImages = masterSpTree && masterRelsXml ? await ImageExtractor.extract(masterSpTree, masterRelsXml, this.zip, "ppt/slideMasters") : [];
      const masterShapes = masterSpTree ? ShapeExtractor.extract(masterSpTree, themeColors) : [];

      const layoutText = layoutSpTree ? TextExtractor.extract(layoutSpTree, themeColors, { context: "layout" }) : [];
      const layoutImages = layoutSpTree && layoutRelsXml ? await ImageExtractor.extract(layoutSpTree, layoutRelsXml, this.zip, "ppt/slideLayouts") : [];
      const layoutShapes = layoutSpTree ? ShapeExtractor.extract(layoutSpTree, themeColors) : [];

      const masterGeom = this.extractPlaceholderGeom(masterSpTree);
      const layoutGeom = this.extractPlaceholderGeom(layoutSpTree);
      const mergedGeom: Record<string, { x: number; y: number; cx: number; cy: number }> = { ...masterGeom, ...layoutGeom };
      const masterTypes = this.extractPlaceholderTypes(masterSpTree);
      const layoutTypes = this.extractPlaceholderTypes(layoutSpTree);
      const mergedTypes: Record<string, string> = { ...masterTypes, ...layoutTypes };
      const slideText = TextExtractor.extract(spTree, themeColors, {
        context: "slide",
        placeholderGeom: mergedGeom,
        placeholderTypes: mergedTypes,
      });
      const slideImages = await ImageExtractor.extract(spTree, relsXml, this.zip, "ppt/slides");
      const slideTables = TableExtractor.extract(spTree, themeColors, themeTableStyles);
      const slideCharts = await ChartExtractor.extract(spTree, relsXml, this.zip, themeColors);
      const slideShapes = ShapeExtractor.extract(spTree, themeColors);

      slides.push([
        ...(bgElement ? [bgElement] : []),
        ...masterShapes,
        ...masterImages,
        ...masterText,
        ...layoutShapes,
        ...layoutImages,
        ...layoutText,
        ...slideShapes,
        ...slideTables,
        ...slideCharts,
        ...slideImages,
        ...slideText,
      ]);
    }

    return slides;
  }

  /** Normalize a relative path against a base directory inside ppt folder */
  private resolvePath(target: string, baseDir: string): string {
    const parts = (baseDir + "/" + target).split("/");
    const resolved: string[] = [];
    for (const part of parts) {
      if (part === "..") {
        if (resolved.length) resolved.pop();
      } else if (part !== "." && part !== "") {
        resolved.push(part);
      }
    }
    return resolved.join("/");
  }

  /** Build an `idx → type` map from layout/master placeholders so slide-level
   *  shapes that only carry `idx` can resolve their inherited placeholder
   *  type (e.g. "body", "title"). */
  private extractPlaceholderTypes(spTree: Element | null): Record<string, string> {
    const map: Record<string, string> = {};
    if (!spTree) return map;
    const shapes = spTree.getElementsByTagNameNS("*", "sp");
    for (const shape of Array.from(shapes)) {
      const nvPr = shape.getElementsByTagNameNS("*", "nvPr")[0] ?? null;
      const ph = nvPr?.getElementsByTagNameNS("*", "ph")[0] ?? null;
      if (!ph) continue;
      const idx = ph.getAttribute("idx") || undefined;
      const type = ph.getAttribute("type") || undefined;
      if (idx && type) map[idx] = type;
    }
    return map;
  }

  private extractPlaceholderGeom(spTree: Element | null): Record<string, { x: number; y: number; cx: number; cy: number }> {
    const map: Record<string, { x: number; y: number; cx: number; cy: number }> = {};
    if (!spTree) return map;
    const shapes = spTree.getElementsByTagNameNS("*", "sp");
    for (const shape of Array.from(shapes)) {
      const nvPr = shape.getElementsByTagNameNS("*", "nvPr")[0] ?? null;
      const ph = nvPr?.getElementsByTagNameNS("*", "ph")[0] ?? null;
      const idx = ph?.getAttribute("idx") || undefined;
      const type = ph?.getAttribute("type") || undefined;
      // Title placeholders typically have only `type` (no `idx`), so key the
      // map by idx when present and fall back to a `type:<value>` key.
      const key = idx || (type ? "type:" + type : undefined);
      if (!key) continue;
      const xfrm = shape.getElementsByTagNameNS("*", "xfrm")[0] ?? null;
      const off = xfrm?.getElementsByTagNameNS("*", "off")[0] ?? null;
      const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0] ?? null;
      if (!off || !ext) continue;
      const geom = {
        x: XmlHelper.getAttrAsNumber(off, "x"),
        y: XmlHelper.getAttrAsNumber(off, "y"),
        cx: XmlHelper.getAttrAsNumber(ext, "cx"),
        cy: XmlHelper.getAttrAsNumber(ext, "cy"),
      };
      map[key] = geom;
      // Slides may use `title` while the layout uses `ctrTitle` for the same
      // box (and vice versa). Mirror so the slide-side lookup matches.
      if (type === "title") map["type:ctrTitle"] = geom;
      else if (type === "ctrTitle") map["type:title"] = geom;
    }
    return map;
  }

  private async extractBackground(
    doc: Document | null,
    rels: Document | null,
    baseDir: string,
    zip: JSZip,
    themeColors: Record<string, string>
  ): Promise<SlideElement | null> {
    if (!doc) return null;
    const bg = doc.getElementsByTagNameNS("*", "bg")[0];
    if (!bg) return null;

    // Try solid fill
    const bgPr = bg.getElementsByTagNameNS("*", "bgPr")[0] || null;
    const solidFill = bgPr?.getElementsByTagNameNS("*", "solidFill")[0] || null;
    const color = XmlHelper.getColorFromElement(solidFill, themeColors);
    if (color) {
      return { type: "background", fillColor: color } as SlideElement;
    }

    // Try scheme color via bgRef
    const bgRef = bg.getElementsByTagNameNS("*", "bgRef")[0] || null;
    const schemeClr = bgRef?.getElementsByTagNameNS("*", "schemeClr")[0] || null;
    const schemeVal = schemeClr?.getAttribute("val") || undefined;
    if (schemeVal && themeColors[schemeVal]) {
      return { type: "background", fillColor: themeColors[schemeVal] } as SlideElement;
    }

    // Try image fill
    const blipFill = bgPr?.getElementsByTagNameNS("*", "blipFill")[0] || null;
    const blip = blipFill?.getElementsByTagNameNS("*", "blip")[0] || null;
    const embedId = blip?.getAttribute("r:embed") || undefined;
    if (embedId && rels) {
      const rel = XmlHelper.findRelationshipById(rels, embedId);
      const target = rel?.getAttribute("Target") || undefined;
      if (target) {
        const fullPath = this.resolvePath(target, baseDir);
        const file = zip.file(fullPath);
        if (file) {
          const binary = await file.async("base64");
          const ext = fullPath.split(".").pop()?.toLowerCase() || "png";
          const dataUri = `data:image/${ext};base64,${binary}`;
          return { type: "background", imageSrc: dataUri } as SlideElement;
        }
      }
    }

    return null;
  }
}
