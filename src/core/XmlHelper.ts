// Lightweight logging helper to centralize library warnings
function libWarn(msg: string) {
  if (typeof console !== "undefined" && console.warn) {
    console.warn(`[pptx-to-html] ${msg}`);
  }
}

type DomParserLike = { parseFromString(xml: string, mimeType: string): Document };

export class XmlHelper {
  private static domParserFactory: (() => DomParserLike) | null = null;
  /**
   * Parses a string containing XML into a DOM Document
   * @param xmlString XML string to parse
   * @returns DOM Document
   */
  static parseXml(xmlString: string): Document {
    if (XmlHelper.domParserFactory) {
      return XmlHelper.domParserFactory().parseFromString(xmlString, "application/xml");
    }
    const anyGlobal = globalThis as any;
    const DP: any = anyGlobal?.DOMParser;
    if (typeof DP === "function") {
      return new DP().parseFromString(xmlString, "application/xml");
    }
    libWarn("No DOMParser available. Provide one through XmlHelper.setDomParser() or pptxToHtml({ domParserFactory }).");
    throw new Error("DOMParser not available in this environment");
  }

  /**
   * Gets a direct child element by local tag name
   */
  static getDirectChildrenByTagName(
    parent: Element,
    tag: string
  ): Element[] {
    return Array.from(parent.children).filter(
      (child) => child.localName === tag
    );
  }

  /**
   * Gets attribute value as number, defaulting to 0
   */
  static getAttrAsNumber(el: Element, name: string): number {
    const raw = el.getAttribute(name);
    if (raw == null || raw === "") return 0;
    const n = Number(raw);
    return Number.isFinite(n) ? n : 0;
  }

  /** Returns the paint order of an element's top-level ancestor in a shape tree. */
  static getZIndex(element: Element, spTree: Element): number {
    let top: Element | null = element;
    while (top?.parentElement && top.parentElement !== spTree) top = top.parentElement;
    if (!top || top.parentElement !== spTree) return 0;
    return Array.from(spTree.children).indexOf(top);
  }

  /**
   * Resolves a drawing object's transform into slide coordinates. Child objects
   * inside a `p:grpSp` are expressed in the group's child coordinate system,
   * so using their local `a:xfrm` directly places them incorrectly.
   */
  static getAbsoluteTransform(element: Element, spTree: Element): {
    x: number; y: number; cx: number; cy: number; rotationDeg?: number;
  } {
    const readTransform = (node: Element) => {
      const xfrm = node.getElementsByTagNameNS("*", "xfrm")[0] ?? null;
      const off = xfrm?.getElementsByTagNameNS("*", "off")[0] ?? null;
      const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0] ?? null;
      const chOff = xfrm?.getElementsByTagNameNS("*", "chOff")[0] ?? null;
      const chExt = xfrm?.getElementsByTagNameNS("*", "chExt")[0] ?? null;
      return {
        xfrm,
        x: off ? XmlHelper.getAttrAsNumber(off, "x") : 0,
        y: off ? XmlHelper.getAttrAsNumber(off, "y") : 0,
        cx: ext ? XmlHelper.getAttrAsNumber(ext, "cx") : 1000000,
        cy: ext ? XmlHelper.getAttrAsNumber(ext, "cy") : 500000,
        chX: chOff ? XmlHelper.getAttrAsNumber(chOff, "x") : 0,
        chY: chOff ? XmlHelper.getAttrAsNumber(chOff, "y") : 0,
        chCx: chExt ? XmlHelper.getAttrAsNumber(chExt, "cx") : 0,
        chCy: chExt ? XmlHelper.getAttrAsNumber(chExt, "cy") : 0,
        rotationDeg: xfrm?.getAttribute("rot") ? Number(xfrm.getAttribute("rot")) / 60000 : 0,
      };
    };

    const own = readTransform(element);
    let { x, y, cx, cy } = own;
    let rotationDeg = own.rotationDeg;
    let parent = element.parentElement;
    while (parent && parent !== spTree) {
      if (parent.localName === "grpSp") {
        const group = readTransform(parent);
        const sx = group.chCx > 0 ? group.cx / group.chCx : 1;
        const sy = group.chCy > 0 ? group.cy / group.chCy : 1;
        x = group.x + (x - group.chX) * sx;
        y = group.y + (y - group.chY) * sy;
        cx *= sx;
        cy *= sy;
        rotationDeg += group.rotationDeg;
      }
      parent = parent.parentElement;
    }
    return { x, y, cx, cy, rotationDeg: rotationDeg || undefined };
  }

  static getColorFromElement(el: Element | null, themeColors?: Record<string, string>): string | undefined {
    if (!el) return undefined;

    // 1. Try <srgbClr val="...">
    const srgb = el.getElementsByTagNameNS("*", "srgbClr")[0];
    if (srgb) {
      const val = srgb.getAttribute("val");
      return val ? `#${val}` : undefined;
    }

    // 2. Try <schemeClr val="..."> resolved via themeColors (including aliases)
    const scheme = el.getElementsByTagNameNS("*", "schemeClr")[0];
    if (scheme) {
      const val = scheme.getAttribute("val");
      if (val && themeColors) {
        // Resolve alias like bg1, bg2, tx1, tx2 to known theme keys
        const aliasMap: Record<string, string> = {
          bg1: "lt1",
          bg2: "lt2",
          tx1: "dk1",
          tx2: "dk2"
        };
        const resolvedKey = aliasMap[val] || val;
        return themeColors[resolvedKey];
      }
      return undefined;
    }

    // 3. Try <sysClr lastClr="...">
    const sys = el.getElementsByTagNameNS("*", "sysClr")[0];
    if (sys) {
      const lastClr = sys.getAttribute("lastClr");
      return lastClr ? `#${lastClr}` : undefined;
    }

    return undefined;
  }

  static extractThemeColors(themeDoc: Document | null): Record<string, string> {
    if (!themeDoc) return {};

    const NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
    const themeColors: Record<string, string> = {};

    const clrScheme = themeDoc.getElementsByTagNameNS(NS, "clrScheme")[0];
    if (!clrScheme) return {};

    for (const node of Array.from(clrScheme.children)) {
      const name = node.localName; // e.g., bg1, tx1, accent1...

      const srgbClr = node.getElementsByTagNameNS(NS, "srgbClr")[0];
      const sysClr = node.getElementsByTagNameNS(NS, "sysClr")[0];

      const hex = srgbClr?.getAttribute("val") ?? sysClr?.getAttribute("lastClr");
      if (hex) {
        themeColors[name] = `#${hex}`;
      }
    }

    return themeColors;
  }

  /**
   * Extracts table styles (fills and text colors per region) from theme XML.
   * Returns a map keyed by styleId (GUID or name), with region color maps.
   */
  static extractThemeTableStyles(themeDoc: Document | null): Record<string, { fills: Record<string, string>; fontColors: Record<string, string> }> {
    const styles: Record<string, { fills: Record<string, string>; fontColors: Record<string, string> }> = {};
    if (!themeDoc) return styles;

    const themeColors = XmlHelper.extractThemeColors(themeDoc);
    const tblStyleLst = themeDoc.getElementsByTagNameNS("*", "tblStyleLst")[0] || null;
    if (!tblStyleLst) return styles;

    const tblStyles = Array.from(tblStyleLst.getElementsByTagNameNS("*", "tblStyle"));
    for (const ts of tblStyles) {
      const id = ts.getAttribute("styleId") || ts.getAttribute("name") || "";
      if (!id) continue;
      const fills: Record<string, string> = {};
      const fontColors: Record<string, string> = {};

      const prNodes = Array.from(ts.getElementsByTagNameNS("*", "tblStylePr"));
      for (const pr of prNodes) {
        const type = pr.getAttribute("type") || pr.getAttribute("val") || ""; // wholeTbl, firstRow, band1H, band2H, band1V, band2V, firstCol, lastCol, lastRow
        if (!type) continue;

        // Resolve fill color: try tcStyle/tcPr/solidFill, then any solidFill under tcStyle, then fillRef, then direct solidFill
        const tcStyle = pr.getElementsByTagNameNS("*", "tcStyle")[0] || null;
        const tcPr = tcStyle?.getElementsByTagNameNS("*", "tcPr")[0] || null;
        const solidCandidates: (Element | null)[] = [
          tcPr?.getElementsByTagNameNS("*", "solidFill")[0] || null,
          tcStyle?.getElementsByTagNameNS("*", "solidFill")[0] || null,
          pr.getElementsByTagNameNS("*", "solidFill")[0] || null,
        ];
        let fillColor: string | undefined;
        for (const cand of solidCandidates) {
          if (cand && !fillColor) fillColor = XmlHelper.getColorFromElement(cand, themeColors);
        }
        if (!fillColor) {
          const fillRef = tcStyle?.getElementsByTagNameNS("*", "fillRef")[0] || pr.getElementsByTagNameNS("*", "fillRef")[0] || null;
          fillColor = XmlHelper.getColorFromElement(fillRef, themeColors);
        }
        if (fillColor) fills[type] = fillColor;

        // Resolve text color: try tcTxStyle/txFill/solidFill, any solidFill, then fontRef (schemeClr)
        const txStyle = pr.getElementsByTagNameNS("*", "tcTxStyle")[0] || null;
        const txFillSolid = txStyle?.getElementsByTagNameNS("*", "solidFill")[0] || null;
        let textColor = XmlHelper.getColorFromElement(txFillSolid, themeColors);
        if (!textColor) {
          const fontRef = txStyle?.getElementsByTagNameNS("*", "fontRef")[0] || null;
          textColor = XmlHelper.getColorFromElement(fontRef as any, themeColors);
        }
        if (!textColor) {
          const anyScheme = txStyle?.getElementsByTagNameNS("*", "schemeClr")[0] || null;
          textColor = XmlHelper.getColorFromElement(anyScheme as any, themeColors);
        }
        if (textColor) fontColors[type] = textColor;
      }

      styles[id] = { fills, fontColors };
    }

    return styles;
  }

  /** Allow a host application to provide a DOM parser instance factory. */
  static setDomParser(factory: () => DomParserLike) {
    XmlHelper.domParserFactory = factory;
  }

  /** Relationship lookup: by Type suffix (avoids querySelector CSS) */
  static findRelationshipByTypeSuffix(doc: Document, suffix: string): Element | null {
    const rels = doc.getElementsByTagName("Relationship");
    for (const el of Array.from(rels)) {
      const t = el.getAttribute("Type") || "";
      if (t.endsWith(suffix)) return el;
    }
    return null;
  }

  /** Relationship lookup: by Id */
  static findRelationshipById(doc: Document, id: string): Element | null {
    const rels = doc.getElementsByTagName("Relationship");
    for (const el of Array.from(rels)) {
      if (el.getAttribute("Id") === id) return el;
    }
    return null;
  }
}
