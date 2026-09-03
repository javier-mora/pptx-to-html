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

      const { x, y, cx, cy, rotationDeg } = XmlHelper.getAbsoluteTransform(gf, spTree);

      charts.push({
        type: "chart",
        zIndex: XmlHelper.getZIndex(gf, spTree),
        rotationDeg,
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
    // Absolute targets (starting with "/") are relative to the package root,
    // not to the .rels file location. pptxgenjs writes absolute targets.
    if (target.startsWith("/")) {
      return target.slice(1);
    }
    const parts = (baseDir + "/" + target).split("/");
    const resolved: string[] = [];
    for (const part of parts) {
      if (part === "..") resolved.pop();
      else if (part !== "." && part !== "") resolved.push(part);
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
    const doughnut = plotArea.getElementsByTagNameNS("*", "doughnutChart")[0] || null;
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
    if (pie || doughnut) {
      const chartEl = pie || doughnut;
      const cat = this.extractCategories(chartEl) || [];
      const ser = this.extractSeries(chartEl, themeColors) || [];
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
      const text = Array.from(rich.getElementsByTagNameNS("*", "t"))
        .map((t) => t.textContent || "")
        .join("");
      return text || undefined;
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
      return this.indexedValues(pts, (p) => p.getElementsByTagNameNS("*", "v")[0]?.textContent || "");
    }
    // Try numCache
    const numCache = cat.getElementsByTagNameNS("*", "numCache")[0] || null;
    if (numCache) {
      const pts = Array.from(numCache.getElementsByTagNameNS("*", "pt"));
      return this.indexedValues(pts, (p) => Number(p.getElementsByTagNameNS("*", "v")[0]?.textContent || 0));
    }
    // Try multiLvlStrCache (pptxgenjs writes category labels this way)
    const multiLvlStrCache = cat.getElementsByTagNameNS("*", "multiLvlStrCache")[0] || null;
    if (multiLvlStrCache) {
      const firstLvl = multiLvlStrCache.getElementsByTagNameNS("*", "lvl")[0] || null;
      if (firstLvl) {
        const pts = Array.from(firstLvl.getElementsByTagNameNS("*", "pt"));
        return this.indexedValues(pts, (p) => p.getElementsByTagNameNS("*", "v")[0]?.textContent || "");
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
        values = this.indexedValues(pts, (p) => Number(p.getElementsByTagNameNS("*", "v")[0]?.textContent || 0));
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
      const xValues = this.indexedValues(xPts, (p) => Number(p.getElementsByTagNameNS("*", "v")[0]?.textContent || 0));
      const yValues = this.indexedValues(yPts, (p) => Number(p.getElementsByTagNameNS("*", "v")[0]?.textContent || 0));
      const len = Math.max(xValues.length, yValues.length);
      const points: { x: number; y: number }[] = [];
      for (let i = 0; i < len; i++) {
        if (xValues[i] !== undefined && yValues[i] !== undefined) {
          points.push({ x: xValues[i], y: yValues[i] });
        }
      }
      const spPr = s.getElementsByTagNameNS("*", "spPr")[0] || null;
      const solidFill = spPr?.getElementsByTagNameNS("*", "solidFill")[0] || null;
      const color = XmlHelper.getColorFromElement(solidFill, themeColors);
      const valueFormat = s.getElementsByTagNameNS("*", "dLbls")[0]?.getElementsByTagNameNS("*", "numFmt")[0]?.getAttribute("formatCode") || undefined;
      out.push({ name, points, color, valueFormat });
    }
    return out;
  }

  /** OOXML caches address points by `idx`; XML order is not guaranteed. */
  private static indexedValues<T>(points: Element[], value: (point: Element) => T): T[] {
    let max = -1;
    const indexed: Array<{ index: number; value: T }> = [];
    points.forEach((point, fallbackIndex) => {
      const raw = point.getAttribute("idx");
      const index = raw !== null && /^\d+$/.test(raw) ? Number(raw) : fallbackIndex;
      max = Math.max(max, index);
      indexed.push({ index, value: value(point) });
    });
    const result = new Array<T>(max + 1);
    for (const point of indexed) result[point.index] = point.value;
    return result;
  }
}
