/** Escapes content placed in an HTML attribute. */
export function escapeAttribute(value: unknown): string {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/** PPTX colors are expected to be hex values. Reject arbitrary CSS from input XML. */
export function safeColor(value: unknown, fallback = "transparent"): string {
  const color = String(value ?? "").trim();
  return /^#[0-9a-f]{3,8}$/i.test(color) || /^(transparent|currentcolor)$/i.test(color)
    ? color
    : fallback;
}

/** Keeps font names usable while preventing an XML value from breaking a style attribute. */
export function safeFontFamily(value: unknown): string {
  const font = String(value ?? "").trim();
  return /^[\p{L}\p{N} ,._-]{1,100}$/u.test(font) ? font : "Arial";
}

export function imageMimeType(path: string): string {
  const ext = path.split(".").pop()?.toLowerCase();
  switch (ext) {
    case "jpg":
    case "jpeg": return "image/jpeg";
    case "gif": return "image/gif";
    case "svg": return "image/svg+xml";
    case "webp": return "image/webp";
    case "bmp": return "image/bmp";
    case "tif":
    case "tiff": return "image/tiff";
    default: return "image/png";
  }
}
