import { TextElement } from "../models/SlideElement";
import { XmlHelper } from "../core/XmlHelper";

/**
 * Responsible for extracting text elements from a slide XML node.
 */
export class TextExtractor {
  /**
   * Extracts all text elements from the slide's <spTree> node.
   * @param spTree The <spTree> element of the slide.
   * @param themeColors Theme color mapping (e.g. tx1, bg2).
   * @returns List of TextElement extracted.
   */
  static extract(
    spTree: Element | null,
    themeColors: Record<string, string>,
    opts: {
      context?: "slide" | "layout" | "master";
      placeholderGeom?: Record<string, { x: number; y: number; cx: number; cy: number }>;
      /** Effective ph type per idx, resolved from layout/master so slide-level
       *  placeholders that only carry idx still expose their inherited type. */
      placeholderTypes?: Record<string, string>;
    } = {}
  ): TextElement[] {
    if (!spTree) return [];

    const elements: TextElement[] = [];

    const shapes = spTree.getElementsByTagNameNS("*", "sp");
    for (const shape of Array.from(shapes)) {
      // Skip placeholder prompts in layout/master
      const nvPr = shape.getElementsByTagNameNS("*", "nvPr")[0] ?? null;
      const ph = nvPr?.getElementsByTagNameNS("*", "ph")[0] ?? null;
      const isPlaceholder = !!ph;
      if (opts.context && opts.context !== "slide" && isPlaceholder) {
        continue;
      }
      const txBody = shape.getElementsByTagNameNS("*", "txBody")[0];
      if (!txBody) continue;

      const paragraphs = txBody.getElementsByTagNameNS("*", "p");
      const bodyPr = txBody.getElementsByTagNameNS("*", "bodyPr")[0] ?? null;
      const anchor = bodyPr?.getAttribute("anchor") || undefined; // t|ctr|b
      const phIdx = ph?.getAttribute("idx") || undefined;
      // Slide-level placeholders often omit `type` and inherit it from the
      // layout/master via matching `idx`. Resolve the effective type so the
      // placeholder-aware defaults below also apply to inherited body/title.
      const phType =
        ph?.getAttribute("type")
        || (phIdx ? opts.placeholderTypes?.[phIdx] : undefined)
        || undefined;
      const isTitlePh = phType === "title" || phType === "ctrTitle";
      // The lib doesn't read txStyles from the slideMaster, so placeholder
      // text falls back to hard-coded defaults. PowerPoint's master usually
      // anchors title bodies to bottom, sets ~44pt bold and centers them, and
      // anchors subtitles to top at ~24pt centered. Apply those defaults here
      // so titles/subtitles aren't tiny, top-left and floating mid-box.
      const verticalAlign =
        anchor === "ctr" ? "middle"
        : anchor === "b" ? "bottom"
        : isTitlePh ? "bottom"
        : "top";
      const lIns = bodyPr?.getAttribute("lIns");
      const tIns = bodyPr?.getAttribute("tIns");
      const rIns = bodyPr?.getAttribute("rIns");
      const bIns = bodyPr?.getAttribute("bIns");
      const padding = {
        left: lIns ? Number(lIns) / 9525 : 0,
        top: tIns ? Number(tIns) / 9525 : 0,
        right: rIns ? Number(rIns) / 9525 : 0,
        bottom: bIns ? Number(bIns) / 9525 : 0,
      };

      const textRuns: string[] = [];
      let fontName = "Arial";
      let fontSize = isTitlePh ? 44 : phType === "subTitle" ? 24 : 18;
      let fontWeight: "normal" | "bold" = isTitlePh ? "bold" : "normal";
      const isCenteredPh = isTitlePh || phType === "subTitle";
      let color: string | undefined = undefined;
      let horizontalAlign: "left" | "center" | "right" | "justify" | undefined = undefined;

      type ParaItem = { kind: "p" | "ul" | "ol"; text: string; lvl: number; listStyle?: string };
      const paraItems: ParaItem[] = [];

      // Extract default list styles per level from lstStyle (if present)
      const lvlDefaults: Record<number, { kind: "p" | "ul" | "ol"; listStyle?: string }> = {};
      const lstStyle = txBody.querySelector("*|lstStyle");
      if (lstStyle) {
        for (const node of Array.from(lstStyle.children)) {
          const m = node.localName.match(/^lvl(\d+)pPr$/);
          if (!m) continue;
          const idx = parseInt(m[1], 10) - 1; // lvl1pPr => level 0
          let kind: "p" | "ul" | "ol" = "p";
          let listStyle: string | undefined = undefined;
          if (node.querySelector("*|buNone")) {
            kind = "p";
          } else if (node.querySelector("*|buAutoNum")) {
            kind = "ol";
            const auto = node.querySelector("*|buAutoNum");
            const typ = auto?.getAttribute("type") || "arabicPeriod";
            listStyle = mapAutoNumToCss(typ);
          } else if (node.querySelector("*|buChar")) {
            kind = "ul";
            listStyle = "disc";
          }
          lvlDefaults[idx] = { kind, listStyle };
        }
      }

      for (const p of Array.from(paragraphs)) {
        const pPr = p.getElementsByTagNameNS("*", "pPr")[0] ?? null;
        const algn = pPr?.getAttribute("algn") || undefined;
        if (algn && !horizontalAlign) {
          horizontalAlign = algn === "ctr" ? "center" : algn === "r" ? "right" : algn.startsWith("just") ? "justify" : "left";
        }

        const runs = p.getElementsByTagNameNS("*", "r");
        let paraText = getParagraphText(p);
        if (paraText) paraText.split(/\n+/).forEach((t) => { if (t) textRuns.push(t); });

        for (const r of Array.from(runs)) {
          const rPr = r.getElementsByTagNameNS("*", "rPr")[0];
          if (rPr) {
            const latin = rPr.getElementsByTagNameNS("*", "latin")[0];
            fontName = latin?.getAttribute("typeface") ?? fontName;

            const sz = rPr.getAttribute("sz");
            if (sz) {
              const n = parseInt(sz, 10);
              if (Number.isFinite(n)) fontSize = n / 100;
            }

            const bAttr = rPr.getAttribute("b");
            if (bAttr === "1") fontWeight = "bold";
            else if (bAttr === "0") fontWeight = "normal";

            const solidFill = rPr.querySelector("*|solidFill");
            const candidate = XmlHelper.getColorFromElement(solidFill || null, themeColors);
            if (candidate) color = candidate;
          }
        }

        // Paragraph-level color fallbacks if not set in runs
        if (!color && pPr) {
          const endParaRPr = pPr.getElementsByTagNameNS("*", "endParaRPr")[0] ?? null;
          const endFill = endParaRPr?.querySelector("*|solidFill") || null;
          const c1 = XmlHelper.getColorFromElement(endFill, themeColors);
          if (c1) color = c1;

          if (!color) {
            const defRPrP = pPr.getElementsByTagNameNS("*", "defRPr")[0] ?? null;
            const defFillP = defRPrP?.querySelector("*|solidFill") || null;
            const c2 = XmlHelper.getColorFromElement(defFillP, themeColors);
            if (c2) color = c2;
          }
        }

        // Determine bulleting for this paragraph
        const lvlAttr = pPr?.getAttribute("lvl");
        const lvl = lvlAttr ? parseInt(lvlAttr, 10) : 0;
        let kind: ParaItem["kind"] = "p";
        let listStyle: string | undefined = undefined;
        if (pPr) {
          if (pPr.querySelector("*|buNone")) {
            kind = "p";
          } else if (pPr.querySelector("*|buAutoNum")) {
            kind = "ol";
            const auto = pPr.querySelector("*|buAutoNum");
            const typ = auto?.getAttribute("type") || "arabicPeriod";
            listStyle = mapAutoNumToCss(typ);
          } else if (pPr.querySelector("*|buChar")) {
            kind = "ul";
            listStyle = "disc";
          } else if (lvlDefaults[lvl]) {
            kind = lvlDefaults[lvl].kind;
            listStyle = lvlDefaults[lvl].listStyle;
          } else if (lvl > 0) {
            // Heuristic: if level is set but no bu* and no defaults, treat as unordered list
            kind = "ul";
            listStyle = "disc";
          }
        } else if (lvlDefaults[lvl]) {
          kind = lvlDefaults[lvl].kind;
          listStyle = lvlDefaults[lvl].listStyle;
        } else if (lvl > 0) {
          kind = "ul";
          listStyle = "disc";
        }

        paraItems.push({ kind, text: paraText, lvl: isNaN(lvl) ? 0 : lvl, listStyle });
      }

      // Fallbacks for color if still undefined
      if (!color) {
        // 1. lstStyle defRPr (txBody > lstStyle)
        const lstDefRPr = txBody.querySelector("*|lstStyle *|defRPr");
        const lstFill = lstDefRPr?.querySelector("*|solidFill");
        const c0 = XmlHelper.getColorFromElement(lstFill || null, themeColors);
        if (c0) color = c0;

        // 2. defRPr from txBody
        const defRPr = txBody.querySelector("*|defRPr");
        const defFill = defRPr?.querySelector("*|solidFill");
        const fallback1 = XmlHelper.getColorFromElement(defFill || null, themeColors);
        if (fallback1) color = fallback1;

        // 3. spPr from shape
        if (!color) {
          const spPr = shape.querySelector("p\\:spPr, spPr");
          const shapeFill = spPr?.querySelector("*|solidFill");
          const fallback2 = XmlHelper.getColorFromElement(shapeFill || null, themeColors);
          if (fallback2) color = fallback2;
        }
      }

      const content = textRuns.join(" ").trim();
      // Filter default placeholder sample text for non-slide contexts
      if (opts.context && opts.context !== "slide") {
        const c = content.toLowerCase();
        const isDefault =
          c.includes("click to add") ||
          c.includes("click to edit") ||
          c.includes("haga clic para agregar") ||
          c.includes("haga clic para editar") ||
          c.includes("hacer clic para agregar") ||
          c.includes("hacer clic para editar");
        if (isDefault) continue;
      }
      if (content === "") continue;

      const xfrm = shape.getElementsByTagNameNS("*", "xfrm")[0];
      const off = xfrm?.getElementsByTagNameNS("*", "off")[0] ?? null;
      const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0] ?? null;

      // If xfrm missing at slide-level, we'll fallback using opts.placeholderGeom below

      // Compute final numbers, using placeholderGeom if needed
      let x: number, y: number, cx: number, cy: number, rotationDeg: number | undefined;
      if (off && ext) {
        ({ x, y, cx, cy, rotationDeg } = XmlHelper.getAbsoluteTransform(shape, spTree));
      } else if (opts.placeholderGeom) {
        // Title placeholders often lack idx but always have type, so fall
        // back to the type-keyed entry populated by SlideExtractor.
        const g = (phIdx ? opts.placeholderGeom[phIdx] : undefined)
          ?? (phType ? opts.placeholderGeom["type:" + phType] : undefined);
        x = g?.x ?? 0;
        y = g?.y ?? 0;
        cx = g?.cx ?? 1000000;
        cy = g?.cy ?? 500000;
      } else {
        x = 0; y = 0; cx = 1000000; cy = 500000;
      }

      // Body placeholders inherit bullet styling from the master's bodyStyle,
      // which the lib doesn't read. Promote plain paragraphs to disc bullets
      // when a body placeholder has multiple paragraphs.
      if (phType === "body" && paraItems.length > 1) {
        for (const it of paraItems) {
          if (it.kind === "p") {
            it.kind = "ul";
            it.listStyle = "disc";
          }
        }
      }

      // Build rich HTML for bullets/numbering, or whenever a shape has 2+
      // paragraphs — the plain content fallback joins runs with spaces and
      // collapses paragraph breaks, so multi-paragraph shapes need the rich
      // path even without explicit bullets.
      let richHtml: string | undefined = undefined;
      if (paraItems.length > 1 || paraItems.some((it) => it.kind !== "p")) {
        const parts: string[] = [];
        let open: { kind: "ul" | "ol"; listStyle?: string } | null = null;
        for (const it of paraItems) {
          if (it.kind === "p") {
            if (open) {
              parts.push(open.kind === "ul" ? "</ul>" : "</ol>");
              open = null;
            }
            if (it.text.trim()) {
              parts.push(`<div style="margin-left:${it.lvl * 24}px">${escapeHtml(it.text).replace(/\n/g, "<br>")}</div>`);
            }
            continue;
          }
          // it's a list item
          if (!open || open.kind !== it.kind) {
            if (open) parts.push(open.kind === "ul" ? "</ul>" : "</ol>");
            const commonListCss = `list-style-position: inside; padding-left: 0; margin: 0;`;
            const style = it.kind === "ol" ? ` style="${commonListCss} list-style-type: ${it.listStyle || "decimal"};"` : ` style="${commonListCss}"`;
            parts.push(it.kind === "ul" ? `<ul${style}>` : `<ol${style}>`);
            open = { kind: it.kind, listStyle: it.listStyle };
          }
          parts.push(`<li style="margin-left:${it.lvl * 24}px">${escapeHtml(it.text).replace(/\n/g, "<br>")}</li>`);
        }
        if (open) parts.push(open.kind === "ul" ? "</ul>" : "</ol>");
        richHtml = parts.join("");
      }

      const element: TextElement = {
        type: "text",
        zIndex: XmlHelper.getZIndex(shape, spTree),
        rotationDeg,
        content,
        position: { x, y },
        size: { width: cx, height: cy },
        font: {
          name: fontName,
          size: fontSize,
          color: color ?? "#000000", // fallback absoluto si quieres
          weight: fontWeight
        },
        align: {
          horizontal: horizontalAlign ?? (isCenteredPh ? "center" : "left"),
          vertical: verticalAlign
        },
        padding,
        html: richHtml
      };

      elements.push(element);
    }

    return elements;
  }
}

// Extract plain text from a paragraph, respecting runs and explicit breaks
function getParagraphText(p: Element): string {
  let out = "";
  // Iterate child nodes to preserve order of runs and breaks
  for (const child of Array.from(p.childNodes) as any[]) {
    // Do not use `instanceof Element`: DOM parsers supplied in Node do not
    // share the browser's global Element constructor.
    if (child.nodeType !== 1) {
      continue;
    }
    const ln = child.localName;
    if (ln === "r") {
      const t = child.getElementsByTagNameNS("*", "t")[0]?.textContent ?? "";
      out += t;
    } else if (ln === "br") {
      out += "\n";
    } else if (ln === "fld") {
      // Fields can contain runs inside
      const runs = child.getElementsByTagNameNS("*", "r");
      for (const r of Array.from(runs)) {
        const t = (r as Element).getElementsByTagNameNS("*", "t")[0]?.textContent ?? "";
        out += t;
      }
    } else if (ln === "tab") {
      out += "\t";
    }
  }
  return out;
}

function mapAutoNumToCss(typ: string): string {
  const t = typ.toLowerCase();
  if (t.includes("alphauc")) return "upper-alpha";
  if (t.includes("alphalc")) return "lower-alpha";
  if (t.includes("romanu")) return "upper-roman";
  if (t.includes("romanl")) return "lower-roman";
  return "decimal";
}

function escapeHtml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
