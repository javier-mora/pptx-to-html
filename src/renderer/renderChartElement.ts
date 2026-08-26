import { ChartElement } from "../models/SlideElement";

export function renderChartElement(el: ChartElement): string {
  const nf = (n: number, fb = 0) => (Number.isFinite(n) ? n : fb);
  const x = nf(el.position?.x, 0) / 9525;
  const y = nf(el.position?.y, 0) / 9525;
  const width = Math.max(1, nf(el.size?.width, 0) / 9525);
  const height = Math.max(1, nf(el.size?.height, 0) / 9525);
  const padding = 24; // basic margins for axes/labels
  const palette = el.palette && el.palette.length > 0 ? el.palette : ["#4e79a7", "#f28e2b", "#e15759", "#76b7b2", "#59a14f", "#edc949"];

  let svg = "";
  if (el.chartType === "column" || el.chartType === "bar") {
    svg = renderBarLike(el, width, height, padding, palette);
  } else if (el.chartType === "line") {
    svg = renderLine(el, width, height, padding, palette);
  } else if (el.chartType === "area") {
    svg = renderArea(el, width, height, padding, palette);
  } else if (el.chartType === "pie") {
    svg = renderPie(el, width, height, palette);
  } else if (el.chartType === "scatter") {
    svg = renderScatter(el, width, height, padding, palette);
  }

  const zIndex = Number.isFinite(el.zIndex) ? el.zIndex : 0;
  const title = el.title ? `<div style="position:absolute;z-index:${zIndex};left:${x}px;top:${y - 20}px;width:${width}px;text-align:center;font-weight:600;">${escape(el.title)}</div>` : "";

  return (
    `${title}<div style="position:absolute; z-index:${zIndex}; left:${x}px; top:${y}px; width:${width}px; height:${height}px;">
      <svg width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">
        ${svg}
      </svg>
    </div>`
  );
}

function renderBarLike(el: ChartElement, width: number, height: number, pad: number, palette: string[]): string {
  const catCount = el.categories.length || 1;
  const chartW = width - pad * 2;
  const chartH = height - pad * 2;
  const isHorizontal = el.chartType === "bar";
  const stacked = el.stackedMode && el.stackedMode !== "none";
  const percent = el.stackedMode === "percent";
  // seriesCount not required after stacked handling; compute widths per branch

  let maxVal = 1;
  if (stacked) {
    const sums = new Array(catCount).fill(0).map((_, i) => el.series.reduce((acc, s) => acc + (((s.values || [])[i]) || 0), 0));
    maxVal = percent ? 1 : Math.max(1, ...sums);
  } else {
    // Compute max across all series values without using flatMap for TS compatibility
    let mv = 1;
    for (const s of el.series) {
      const vals = s.values || [];
      for (const v of vals) mv = Math.max(mv, v);
    }
    maxVal = mv;
  }

  const parts: string[] = [];
  // Axes and ticks
  if (isHorizontal) {
    parts.push(`<line x1="${pad}" y1="${height - pad}" x2="${width - pad}" y2="${height - pad}" stroke="#999" stroke-width="1" />`);
    parts.push(`<line x1="${pad}" y1="${pad}" x2="${pad}" y2="${height - pad}" stroke="#999" stroke-width="1" />`);
    const ticks = computeTicks(0, maxVal, 4);
    ticks.forEach((t) => {
      const tx = pad + (t / maxVal) * chartW;
      const ty = height - pad;
      parts.push(`<line x1="${tx}" y1="${ty}" x2="${tx}" y2="${ty + 4}" stroke="#999" stroke-width="1" />`);
      parts.push(`<text x="${tx}" y="${ty + 16}" text-anchor="middle" font-size="10" fill="#666">${formatNumber(t, percent ? "0%" : el.valueFormat)}</text>`);
    });
  } else {
    parts.push(`<line x1="${pad}" y1="${height - pad}" x2="${width - pad}" y2="${height - pad}" stroke="#999" stroke-width="1" />`);
    parts.push(`<line x1="${pad}" y1="${pad}" x2="${pad}" y2="${height - pad}" stroke="#999" stroke-width="1" />`);
    const ticks = computeTicks(0, maxVal, 4);
    ticks.forEach((t) => {
      const ty = height - pad - (t / maxVal) * chartH;
      parts.push(`<line x1="${pad - 4}" y1="${ty}" x2="${pad}" y2="${ty}" stroke="#999" stroke-width="1" />`);
      parts.push(`<text x="${pad - 6}" y="${ty + 3}" text-anchor="end" font-size="10" fill="#666">${formatNumber(t, percent ? "0%" : el.valueFormat)}</text>`);
      parts.push(`<line x1="${pad}" y1="${ty}" x2="${width - pad}" y2="${ty}" stroke="#eee" stroke-width="1" />`);
    });
  }

  if (isHorizontal) {
    const catBand = chartH / catCount;
    const barH = Math.max(2, (catBand * 0.8) / (stacked ? 1 : el.series.length));
    el.series.forEach((s, si) => {
      const color = s.color || palette[si % palette.length];
      (s.values || []).forEach((v, ci) => {
        const baseY = pad + ci * catBand + (catBand - (stacked ? barH : el.series.length * barH)) / 2;
        if (stacked) {
          const prev = el.series.slice(0, si).reduce((acc, ss) => acc + (((ss.values || [])[ci]) || 0), 0);
          const sum = el.series.reduce((acc, ss) => acc + (((ss.values || [])[ci]) || 0), 0) || 1;
          const start = ((percent ? prev / sum : prev) / maxVal) * chartW;
          const w = ((percent ? (v || 0) / sum : v) / maxVal) * chartW;
          const x = pad + start;
          const y = baseY;
          parts.push(`<rect x="${x}" y="${y}" width="${w}" height="${barH}" fill="${color}" />`);
          if (el.showDataLabels) {
            const fmt = s.valueFormat || el.valueFormat;
            parts.push(`<text x="${x + w - 2}" y="${y + barH / 2 + 3}" text-anchor="end" font-size="10" fill="#000">${formatNumber(percent ? (v / sum) : v, percent ? "0%" : fmt)}</text>`);
          }
        } else {
          const w = (v / maxVal) * chartW;
          const y = baseY + si * barH;
          parts.push(`<rect x="${pad}" y="${y}" width="${w}" height="${barH}" fill="${color}" />`);
          if (el.showDataLabels) {
            const fmt = s.valueFormat || el.valueFormat;
            parts.push(`<text x="${pad + w + 2}" y="${y + barH / 2 + 3}" font-size="10" fill="#000">${formatNumber(v, fmt)}</text>`);
          }
        }
      });
    });
    const labelSize = Math.max(10, Math.min(12, chartH / (catCount * 2)));
    el.categories.forEach((c, i) => {
      const cy = pad + i * catBand + catBand / 2 + 4;
      parts.push(`<text x="${pad - 8}" y="${cy}" text-anchor="end" font-size="${labelSize}" fill="#333">${escape(String(c))}</text>`);
    });
  } else {
    const catBand = chartW / catCount;
    const barW = Math.max(2, (catBand * 0.8) / (stacked ? 1 : el.series.length));
    el.series.forEach((s, si) => {
      const color = s.color || palette[si % palette.length];
      (s.values || []).forEach((v, ci) => {
        const baseX = pad + ci * catBand + (catBand - (stacked ? barW : el.series.length * barW)) / 2;
        if (stacked) {
          const prev = el.series.slice(0, si).reduce((acc, ss) => acc + (((ss.values || [])[ci]) || 0), 0);
          const sum = el.series.reduce((acc, ss) => acc + (((ss.values || [])[ci]) || 0), 0) || 1;
          const start = ((percent ? prev / sum : prev) / maxVal) * chartH;
          const h = ((percent ? (v || 0) / sum : v) / maxVal) * chartH;
          const x = baseX;
          const y = height - pad - start - h;
          parts.push(`<rect x="${x}" y="${y}" width="${barW}" height="${h}" fill="${color}" />`);
          if (el.showDataLabels) {
            const fmt = s.valueFormat || el.valueFormat;
            parts.push(`<text x="${x + barW / 2}" y="${y - 2}" text-anchor="middle" font-size="10" fill="#000">${formatNumber(percent ? (v / sum) : v, percent ? "0%" : fmt)}</text>`);
          }
        } else {
          const h = (v / maxVal) * chartH;
          const x = baseX + si * barW;
          const y = height - pad - h;
          parts.push(`<rect x="${x}" y="${y}" width="${barW}" height="${h}" fill="${color}" />`);
          if (el.showDataLabels) {
            const fmt = s.valueFormat || el.valueFormat;
            parts.push(`<text x="${x + barW / 2}" y="${y - 2}" text-anchor="middle" font-size="10" fill="#000">${formatNumber(v, fmt)}</text>`);
          }
        }
      });
    });
    const labelSize = Math.max(10, Math.min(12, chartW / (catCount * 4)));
    el.categories.forEach((c, i) => {
      const cx = pad + i * catBand + catBand / 2;
      const cy = height - pad + 14;
      parts.push(`<text x="${cx}" y="${cy}" text-anchor="middle" font-size="${labelSize}" fill="#333">${escape(String(c))}</text>`);
    });
  }

  if (el.showLegend) {
    parts.push(renderLegend(el, width, pad, palette));
  }

  return parts.join("\n");
}

function renderLine(el: ChartElement, width: number, height: number, pad: number, palette: string[]): string {
  const catCount = el.categories.length || 1;
  const chartW = width - pad * 2;
  const chartH = height - pad * 2;
  let maxVal = 1;
  for (const s of el.series) {
    const vals = s.values || [];
    for (const v of vals) maxVal = Math.max(maxVal, v);
  }
  const xStep = chartW / Math.max(1, catCount - 1);
  const parts: string[] = [];
  parts.push(`<line x1="${pad}" y1="${height - pad}" x2="${width - pad}" y2="${height - pad}" stroke="#999" stroke-width="1" />`);
  parts.push(`<line x1="${pad}" y1="${pad}" x2="${pad}" y2="${height - pad}" stroke="#999" stroke-width="1" />`);
  const ticks = computeTicks(0, maxVal, 4);
  ticks.forEach((t) => {
    const ty = height - pad - (t / maxVal) * chartH;
    parts.push(`<line x1="${pad - 4}" y1="${ty}" x2="${pad}" y2="${ty}" stroke="#999" stroke-width="1" />`);
    parts.push(`<text x="${pad - 6}" y="${ty + 3}" text-anchor="end" font-size="10" fill="#666">${formatNumber(t, el.valueFormat)}</text>`);
    parts.push(`<line x1="${pad}" y1="${ty}" x2="${width - pad}" y2="${ty}" stroke="#eee" stroke-width="1" />`);
  });
  // Handle stacking if requested
  const stacked = el.stackedMode && el.stackedMode !== "none";
  const percent = el.stackedMode === "percent";
  const totals = percent ? new Array(catCount).fill(0).map((_, i) => el.series.reduce((acc, sr) => acc + ((sr.values || [])[i] || 0), 0)) : undefined;

  el.series.forEach((s, si) => {
    const color = s.color || palette[si % palette.length];
    let d = "";
    (s.values || []).forEach((v, i) => {
      const x = pad + i * xStep;
      let val = v;
      if (stacked) {
        const prev = el.series.slice(0, si).reduce((acc, ss) => acc + (((ss.values || [])[i]) || 0), 0);
        val = prev + v;
        if (percent && totals) val = totals[i] ? val / totals[i] : 0;
      }
      const y = height - pad - (val / (percent ? 1 : maxVal)) * chartH;
      d += i === 0 ? `M ${x} ${y}` : ` L ${x} ${y}`;
    });
    parts.push(`<path d="${d}" fill="none" stroke="${color}" stroke-width="2" />`);
    if (el.showDataLabels) {
      (s.values || []).forEach((v, i) => {
        const x = pad + i * xStep;
        const basePrev = stacked ? el.series.slice(0, si).reduce((acc, ss) => acc + (((ss.values || [])[i]) || 0), 0) : 0;
        const dispVal = percent && totals ? (v / (totals[i] || 1)) : v;
        const y = height - pad - ((stacked ? (basePrev + v) : v) / (percent ? 1 : maxVal)) * chartH;
        parts.push(`<circle cx="${x}" cy="${y}" r="2.5" fill="${color}" />`);
        const fmt = s.valueFormat || el.valueFormat;
        parts.push(`<text x="${x}" y="${y - 6}" text-anchor="middle" font-size="10" fill="#000">${formatNumber(dispVal, percent ? "0%" : fmt)}</text>`);
      });
    }
  });
  // Category labels on x-axis
  if (el.categories.length > 0) {
    const labelSize = Math.max(10, Math.min(12, chartW / (catCount * 4)));
    el.categories.forEach((c, i) => {
      const cx = pad + i * xStep;
      parts.push(`<text x="${cx}" y="${height - pad + 14}" text-anchor="middle" font-size="${labelSize}" fill="#333">${escape(String(c))}</text>`);
    });
  }
  return parts.join("\n");
}

function renderArea(el: ChartElement, width: number, height: number, pad: number, palette: string[]): string {
  const catCount = el.categories.length || 1;
  const chartW = width - pad * 2;
  const chartH = height - pad * 2;
  let maxVal = 1;
  for (const s of el.series) {
    const vals = s.values || [];
    for (const v of vals) maxVal = Math.max(maxVal, v);
  }
  const xStep = chartW / Math.max(1, catCount - 1);
  const parts: string[] = [];
  parts.push(`<line x1="${pad}" y1="${height - pad}" x2="${width - pad}" y2="${height - pad}" stroke="#999" stroke-width="1" />`);
  parts.push(`<line x1="${pad}" y1="${pad}" x2="${pad}" y2="${height - pad}" stroke="#999" stroke-width="1" />`);

  const stacked = el.stackedMode && el.stackedMode !== "none";
  const percent = el.stackedMode === "percent";
  const totals = percent ? new Array(catCount).fill(0).map((_, i) => el.series.reduce((acc, sr) => acc + ((sr.values || [])[i] || 0), 0)) : undefined;
  const baseline = new Array(catCount).fill(0);

  el.series.forEach((s, si) => {
    const color = s.color || palette[si % palette.length];
    const topY: number[] = [];
    const botY: number[] = [];
    (s.values || []).forEach((v, i) => {
      const prev = stacked ? baseline[i] : 0;
      const val = stacked ? (prev + (percent && totals ? (totals[i] ? v / totals[i] : 0) : v)) : v;
      const top = height - pad - (val / (percent ? 1 : maxVal)) * chartH;
      const bottom = height - pad - (prev / (percent ? 1 : maxVal)) * chartH;
      topY.push(top);
      botY.push(bottom);
      if (stacked) baseline[i] = percent && totals ? val : prev + v;
    });
    // Build polygon path
    let d = "";
    for (let i = 0; i < topY.length; i++) {
      const x = pad + i * xStep;
      d += i === 0 ? `M ${x} ${topY[i]}` : ` L ${x} ${topY[i]}`;
    }
    for (let i = botY.length - 1; i >= 0; i--) {
      const x = pad + i * xStep;
      d += ` L ${x} ${botY[i]}`;
    }
    d += " Z";
    parts.push(`<path d="${d}" fill="${color}" fill-opacity="0.6" stroke="none" />`);
  });

  return parts.join("\n");
}

function renderPie(el: ChartElement, width: number, height: number, palette: string[]): string {
  const s0 = el.series[0];
  const values: number[] = (s0 && s0.values) ? s0.values : [];
  const total = values.reduce((a, b) => a + Math.max(0, b), 0) || 1;
  const cx = width / 2;
  const cy = height / 2;
  const r = Math.min(width, height) * 0.35;
  let start = 0;
  const parts: string[] = [];
  values.forEach((v, i) => {
    const frac = Math.max(0, v) / total;
    const end = start + frac * 2 * Math.PI;
    const x1 = cx + r * Math.cos(start);
    const y1 = cy + r * Math.sin(start);
    const x2 = cx + r * Math.cos(end);
    const y2 = cy + r * Math.sin(end);
    const large = end - start > Math.PI ? 1 : 0;
    const color = (s0 && s0.ptColors && s0.ptColors[i]) || (s0 && s0.color) || palette[i % palette.length];
    const d = `M ${cx} ${cy} L ${x1} ${y1} A ${r} ${r} 0 ${large} 1 ${x2} ${y2} Z`;
    parts.push(`<path d="${d}" fill="${color}" />`);
    if (el.showDataLabels && frac > 0) {
      const mid = (start + end) / 2;
      const lx = cx + (r + 12) * Math.cos(mid);
      const ly = cy + (r + 12) * Math.sin(mid);
      const catName = String(el.categories[i] || "");
      const pieLabel = catName ? `${escape(catName)} ${(frac * 100).toFixed(0)}%` : `${(frac * 100).toFixed(0)}%`;
      parts.push(`<text x="${lx}" y="${ly}" text-anchor="middle" font-size="10" fill="#000">${pieLabel}</text>`);
    }
    start = end;
  });
  return parts.join("\n");
}

function renderScatter(el: ChartElement, width: number, height: number, pad: number, palette: string[]): string {
  const chartW = width - pad * 2;
  const chartH = height - pad * 2;
  const allPoints: { x: number; y: number }[] = [];
  for (const s of el.series) {
    const pts = s.points || [];
    for (const p of pts) allPoints.push(p);
  }
  const minX = Math.min(...allPoints.map((p) => p.x), 0);
  const maxX = Math.max(...allPoints.map((p) => p.x), 1);
  const minY = Math.min(...allPoints.map((p) => p.y), 0);
  const maxY = Math.max(...allPoints.map((p) => p.y), 1);
  const parts: string[] = [];
  // Axes
  parts.push(`<line x1="${pad}" y1="${height - pad}" x2="${width - pad}" y2="${height - pad}" stroke="#999" stroke-width="1" />`);
  parts.push(`<line x1="${pad}" y1="${pad}" x2="${pad}" y2="${height - pad}" stroke="#999" stroke-width="1" />`);
  const xticks = computeTicks(minX, maxX, 4);
  xticks.forEach((t) => {
    const tx = pad + ((t - minX) / Math.max(1e-9, maxX - minX)) * chartW;
    const ty = height - pad;
    parts.push(`<line x1="${tx}" y1="${ty}" x2="${tx}" y2="${ty + 4}" stroke="#999" stroke-width="1" />`);
    parts.push(`<text x="${tx}" y="${ty + 16}" text-anchor="middle" font-size="10" fill="#666">${formatNumber(t, el.valueFormat)}</text>`);
  });
  const yticks = computeTicks(minY, maxY, 4);
  yticks.forEach((t) => {
    const ty = height - pad - ((t - minY) / Math.max(1e-9, maxY - minY)) * chartH;
    parts.push(`<line x1="${pad - 4}" y1="${ty}" x2="${pad}" y2="${ty}" stroke="#999" stroke-width="1" />`);
    parts.push(`<text x="${pad - 6}" y="${ty + 3}" text-anchor="end" font-size="10" fill="#666">${formatNumber(t, el.valueFormat)}</text>`);
    parts.push(`<line x1="${pad}" y1="${ty}" x2="${width - pad}" y2="${ty}" stroke="#eee" stroke-width="1" />`);
  });
  el.series.forEach((s, si) => {
    const color = s.color || palette[si % palette.length];
    (s.points || []).forEach((p) => {
      const x = pad + ((p.x - minX) / Math.max(1e-9, maxX - minX)) * chartW;
      const y = height - pad - ((p.y - minY) / Math.max(1e-9, maxY - minY)) * chartH;
      parts.push(`<circle cx="${x}" cy="${y}" r="3" fill="${color}" />`);
      if (el.showDataLabels) {
        const fmt = s.valueFormat || el.valueFormat;
        parts.push(`<text x="${x + 5}" y="${y - 5}" font-size="10" fill="#000">${formatNumber(p.y, fmt)}</text>`);
      }
    });
  });
  return parts.join("\n");
}

function escape(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function computeTicks(min: number, max: number, count: number): number[] {
  const span = max - min;
  if (span <= 0) return [min, max];
  const step = niceNum(span / count, true);
  const ticks: number[] = [];
  let v = Math.ceil(min / step) * step;
  while (v <= max + 1e-9) {
    ticks.push(Number(v.toFixed(10)));
    v += step;
  }
  return ticks;
}

function niceNum(range: number, round: boolean): number {
  const exponent = Math.floor(Math.log10(range));
  const fraction = range / Math.pow(10, exponent);
  let niceFraction: number;
  if (round) {
    if (fraction < 1.5) niceFraction = 1;
    else if (fraction < 3) niceFraction = 2;
    else if (fraction < 7) niceFraction = 5;
    else niceFraction = 10;
  } else {
    if (fraction <= 1) niceFraction = 1;
    else if (fraction <= 2) niceFraction = 2;
    else if (fraction <= 5) niceFraction = 5;
    else niceFraction = 10;
  }
  return niceFraction * Math.pow(10, exponent);
}

function formatNumber(v: number, formatCode?: string): string {
  if (!formatCode) {
    if (Math.abs(v) >= 1000) return `${Math.round(v)}`;
    if (Math.abs(v) >= 10) return v.toFixed(0);
    if (Math.abs(v) >= 1) return v.toFixed(1);
    return v.toFixed(2);
  }
  // very small subset of Excel-like formats: handle %, currency symbols, thousand separator, decimals
  let isPercent = /%/.test(formatCode);
  let decimals = 0;
  const decMatch = formatCode.match(/\.([0#]+)/);
  if (decMatch) decimals = decMatch[1].length;
  const currencyMatch = formatCode.match(/([$€£¥])/);
  const currency = currencyMatch ? currencyMatch[1] : "";
  const useThousands = /#,##0/.test(formatCode);
  let n = v;
  if (isPercent) n = n * 100;
  let str = n.toFixed(decimals);
  if (useThousands) {
    const [int, frac] = str.split(".");
    const withSep = int.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    str = frac !== undefined ? `${withSep}.${frac}` : withSep;
  }
  if (currency) str = currency + str;
  if (isPercent) str = str + "%";
  return str;
}

function renderLegend(el: ChartElement, width: number, pad: number, palette: string[]): string {
  let x = pad;
  const y = pad - 12;
  const parts: string[] = [];
  el.series.forEach((s, i) => {
    const color = s.color || palette[i % palette.length];
    const label = s.name || `Series ${i + 1}`;
    parts.push(`<rect x="${x}" y="${y}" width="10" height="10" fill="${color}" />`);
    parts.push(`<text x="${x + 14}" y="${y + 9}" font-size="10" fill="#333">${escape(label)}</text>`);
    x += 14 + label.length * 6 + 10;
  });
  return parts.join("\n");
}
