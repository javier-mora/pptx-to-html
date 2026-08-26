import JSZip from "jszip";
import { XmlHelper } from "../core/XmlHelper";
import { ChartElement, ChartSeries, ChartType } from "../models/SlideElement";

export class ChartExtractor {
  static async extract(
    spTree: Element | null,
    relsXml: Document,
    zip: JSZip,
    themeColors: Record<string, string>
  ): Promise<ChartElement[]> {
    if (!spTree) return [];

    const charts: ChartElement[] = [];
    const gFrames = spTree.getElementsByTagNameNS("*", "graphicFrame");
    for (const gf of Array.from(gFrames)) {
      const graphicData = gf.getElementsByTagNameNS("*", "graphicData")[0] ?? null;
      if (!graphicData) continue;
      const chartEl = graphicData.getElementsByTagNameNS("*", "chart")[0] ?? null;
      if (!chartEl) continue;

      const rId = chartEl.getAttribute("r:id") || chartEl.getAttribute("r:embed") || undefined;
      if (!rId) continue;

      const rel = XmlHelper.findRelationshipById(relsXml, rId);
      const target = rel?.getAttribute("Target") || undefined;
      if (!target) continue;

      const fullPath = this.resolvePath(target, "ppt/slides");
      const file = zip.file(fullPath);
      if (!file) continue;
      const xmlStr = await file.async("string");
      const doc = XmlHelper.parseXml(xmlStr);

      const parsed = this.parseChart(doc, themeColors);
      if (!parsed) continue;

      const xfrm = gf.getElementsByTagNameNS("*", "xfrm")[0] ?? null;
      const off = xfrm?.getElementsByTagNameNS("*", "off")[0] ?? null;
      const ext = xfrm?.getElementsByTagNameNS("*", "ext")[0] ?? null;
      const x = off ? XmlHelper.getAttrAsNumber(off, "x") : 0;
      const y = off ? XmlHelper.getAttrAsNumber(off, "y") : 0;
      const cx = ext ? XmlHelper.getAttrAsNumber(ext, "cx") : 1000000;
      const cy = ext ? XmlHelper.getAttrAsNumber(ext, "cy") : 600000;

      charts.push({
        type: "chart",
        zIndex: XmlHelper.getZIndex(gf, spTree),
        chartType: parsed.type,
        position: { x, y },
        size: { width: cx, height: cy },
        categories: parsed.categories,
        series: parsed.series,
        palette: parsed.palette,
        title: parsed.title,
        showLegend: parsed.showLegend,
        showDataLabels: parsed.showDataLabels,
        stackedMode: parsed.stackedMode,
        valueFormat: parsed.valueFormat,
      });
    }

    return charts;
  }

  private static resolvePath(target: string, baseDir: string): string {
    const parts = (baseDir + "/" + target).split("/");
    const resolved: string[] = [];
    for (const part of parts) {
      if (part === "..") resolved.pop();
      else if (part !== ".") resolved.push(part);
    }
    return resolved.join("/");
  }

  private static parseChart(doc: Document, themeColors: Record<string, string>):
    | { type: ChartType; categories: (string | number)[]; series: ChartSeries[]; palette?: string[]; title?: string; showLegend?: boolean; showDataLabels?: boolean; stackedMode?: "none" | "stacked" | "percent"; valueFormat?: string }
    | null {
    const plotArea = doc.getElementsByTagNameNS("*", "plotArea")[0] || null;
    if (!plotArea) return null;

    const titleText = this.extractTitle(doc);
    const showLegend = !!doc.getElementsByTagNameNS("*", "legend")[0];
    const showDataLabels = !!plotArea.getElementsByTagNameNS("*", "dLbls")[0];

    // Detect type order: bar/col, line, area, pie, scatter
    const bar = plotArea.getElementsByTagNameNS("*", "barChart")[0] || null;
    const line = plotArea.getElementsByTagNameNS("*", "lineChart")[0] || null;
    const area = plotArea.getElementsByTagNameNS("*", "areaChart")[0] || null;
    const pie = plotArea.getElementsByTagNameNS("*", "pieChart")[0] || null;
    const scatter = plotArea.getElementsByTagNameNS("*", "scatterChart")[0] || null;

    const chartNumFmt = plotArea.getElementsByTagNameNS("*", "dLbls")[0]?.getElementsByTagNameNS("*", "numFmt")[0]?.getAttribute("formatCode") || undefined;

    const palette = [
      themeColors["accent1"],
      themeColors["accent2"],
      themeColors["accent3"],
      themeColors["accent4"],
      themeColors["accent5"],
      themeColors["accent6"],
    ].filter(Boolean) as string[];

    if (bar) {
      const cat = this.extractCategories(bar) || [];
      const ser = this.extractSeries(bar, themeColors) || [];
      // barDir decides orientation: col vs bar
      const barDir = bar.getElementsByTagNameNS("*", "barDir")[0]?.getAttribute("val") || "col";
      const type: ChartType = barDir === "bar" ? "bar" : "column";
      const grouping = bar.getElementsByTagNameNS("*", "grouping")[0]?.getAttribute("val") || "clustered";
      const stackedMode = grouping === "stacked" ? "stacked" : grouping === "percentStacked" ? "percent" : "none";
      return { type, categories: cat, series: ser, palette, title: titleText, showLegend, showDataLabels, stackedMode, valueFormat: chartNumFmt };
    }
    if (line) {
      const cat = this.extractCategories(line) || [];
      const ser = this.extractSeries(line, themeColors) || [];
      const grouping = line.getElementsByTagNameNS("*", "grouping")[0]?.getAttribute("val") || "standard";
      const stackedMode = grouping === "stacked" ? "stacked" : grouping === "percentStacked" ? "percent" : "none";
      return { type: "line", categories: cat, series: ser, palette, title: titleText, showLegend, showDataLabels, stackedMode, valueFormat: chartNumFmt };
    }
    if (area) {
      const cat = this.extractCategories(area) || [];
      const ser = this.extractSeries(area, themeColors) || [];
      const grouping = area.getElementsByTagNameNS("*", "grouping")[0]?.getAttribute("val") || "standard";
      const stackedMode = grouping === "stacked" ? "stacked" : grouping === "percentStacked" ? "percent" : "none";
      return { type: "area", categories: cat, series: ser, palette, title: titleText, showLegend, showDataLabels, stackedMode, valueFormat: chartNumFmt };
    }
    if (pie) {
      const cat = this.extractCategories(pie) || [];
      const ser = this.extractSeries(pie, themeColors) || [];
      return { type: "pie", categories: cat, series: ser, palette, title: titleText, showLegend, showDataLabels, stackedMode: "none", valueFormat: chartNumFmt };
    }
    if (scatter) {
      const ser = this.extractScatterSeries(scatter, themeColors) || [];
      return { type: "scatter", categories: [], series: ser, palette, title: titleText, showLegend, showDataLabels, stackedMode: "none", valueFormat: chartNumFmt };
    }
    return null;
  }

  private static extractTitle(doc: Document): string | undefined {
    const title = doc.getElementsByTagNameNS("*", "title")[0] || null;
    if (!title) return undefined;
    const tx = title.getElementsByTagNameNS("*", "tx")[0] || null;
    const rich = tx?.getElementsByTagNameNS("*", "rich")[0] || null;
    if (rich) {
      const t = rich.getElementsByTagNameNS("*", "t")[0]?.textContent || undefined;
      return t || undefined;
    }
    const v = tx?.getElementsByTagNameNS("*", "v")[0]?.textContent || undefined;
    return v || undefined;
  }

  private static extractCategories(parent: Element): (string | number)[] | null {
    const cat = parent.getElementsByTagNameNS("*", "cat")[0] || null;
    if (!cat) return null;
    // Try string cache
    const strCache = cat.getElementsByTagNameNS("*", "strCache")[0] || null;
    if (strCache) {
      const pts = Array.from(strCache.getElementsByTagNameNS("*", "pt"));
      return pts.map((p) => p.getElementsByTagNameNS("*", "v")[0]?.textContent || "");
    }
    // Try numCache
    const numCache = cat.getElementsByTagNameNS("*", "numCache")[0] || null;
    if (numCache) {
      const pts = Array.from(numCache.getElementsByTagNameNS("*", "pt"));
      return pts.map((p) => Number(p.getElementsByTagNameNS("*", "v")[0]?.textContent || 0));
    }
    // Try multiLvlStrCache (pptxgenjs writes category labels this way)
    const multiLvlStrCache = cat.getElementsByTagNameNS("*", "multiLvlStrCache")[0] || null;
    if (multiLvlStrCache) {
      const firstLvl = multiLvlStrCache.getElementsByTagNameNS("*", "lvl")[0] || null;
      if (firstLvl) {
        const pts = Array.from(firstLvl.getElementsByTagNameNS("*", "pt"));
        return pts.map((p) => p.getElementsByTagNameNS("*", "v")[0]?.textContent || "");
      }
    }
    return null;
  }

  private static extractSeries(parent: Element, themeColors: Record<string, string>): ChartSeries[] | null {
    const series: ChartSeries[] = [];
    const sers = Array.from(parent.getElementsByTagNameNS("*", "ser"));
    let idx = 0;
    for (const s of sers) {
      const name = s.getElementsByTagNameNS("*", "tx")[0]?.getElementsByTagNameNS("*", "v")[0]?.textContent || undefined;
      const numCache = s.getElementsByTagNameNS("*", "numCache")[0] || null;
      let values: number[] = [];
      if (numCache) {
        const pts = Array.from(numCache.getElementsByTagNameNS("*", "pt"));
        values = pts.map((p) => Number(p.getElementsByTagNameNS("*", "v")[0]?.textContent || 0));
      }
      const valueFormat = s.getElementsByTagNameNS("*", "dLbls")[0]?.getElementsByTagNameNS("*", "numFmt")[0]?.getAttribute("formatCode") || undefined;
      // Series color from spPr/solidFill
      const spPr = s.getElementsByTagNameNS("*", "spPr")[0] || null;
      const solidFill = spPr?.getElementsByTagNameNS("*", "solidFill")[0] || null;
      const color = XmlHelper.getColorFromElement(solidFill, themeColors);
      // Per-datapoint colors from c:dPt (used especially by pie/doughnut charts)
      const ptColors: string[] = [];
      for (const dp of Array.from(s.getElementsByTagNameNS("*", "dPt"))) {
        const di = Number(dp.getElementsByTagNameNS("*", "idx")[0]?.getAttribute("val") || 0);
        const df = (dp.getElementsByTagNameNS("*", "spPr")[0] || null)?.getElementsByTagNameNS("*", "solidFill")[0] || null;
        const dc = XmlHelper.getColorFromElement(df, themeColors);
        if (dc) ptColors[di] = dc;
      }
      series.push({ name, values, color, valueFormat, ptColors: ptColors.length > 0 ? ptColors : undefined });
      idx += 1;
    }
    return series;
  }

  private static extractScatterSeries(parent: Element, themeColors: Record<string, string>): ChartSeries[] | null {
    const out: ChartSeries[] = [];
    const sers = Array.from(parent.getElementsByTagNameNS("*", "ser"));
    for (const s of sers) {
      const name = s.getElementsByTagNameNS("*", "tx")[0]?.getElementsByTagNameNS("*", "v")[0]?.textContent || undefined;
      const xCache = s.getElementsByTagNameNS("*", "xVal")[0]?.getElementsByTagNameNS("*", "numCache")[0] || null;
      const yCache = s.getElementsByTagNameNS("*", "yVal")[0]?.getElementsByTagNameNS("*", "numCache")[0] || null;
      const xPts = xCache ? Array.from(xCache.getElementsByTagNameNS("*", "pt")) : [];
      const yPts = yCache ? Array.from(yCache.getElementsByTagNameNS("*", "pt")) : [];
      const len = Math.min(xPts.length, yPts.length);
      const points: { x: number; y: number }[] = [];
      for (let i = 0; i < len; i++) {
        const xv = Number(xPts[i].getElementsByTagNameNS("*", "v")[0]?.textContent || 0);
        const yv = Number(yPts[i].getElementsByTagNameNS("*", "v")[0]?.textContent || 0);
        points.push({ x: xv, y: yv });
      }
      const spPr = s.getElementsByTagNameNS("*", "spPr")[0] || null;
      const solidFill = spPr?.getElementsByTagNameNS("*", "solidFill")[0] || null;
      const color = XmlHelper.getColorFromElement(solidFill, themeColors);
      const valueFormat = s.getElementsByTagNameNS("*", "dLbls")[0]?.getElementsByTagNameNS("*", "numFmt")[0]?.getAttribute("formatCode") || undefined;
      out.push({ name, points, color, valueFormat });
    }
    return out;
  }
}
