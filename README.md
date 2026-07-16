# pptx-to-html

Convert PowerPoint (`.pptx`) files into HTML slides using TypeScript. The library parses the OOXML structures inside a presentation and renders each slide as a self-contained HTML snippet with absolutely positioned elements.

- Zero-runtime CSS required; styles are inlined.
- Theme colors, basic table styling, and common chart rendering are supported.
- Works in modern browsers. Node is supported via an injectable DOM parser.

> Based on the ISO/IEC 29500:2012 "Office Open XML File Formats - Fundamentals And Markup Language Reference".

**[Live demo](https://javier-mora.github.io/pptx-to-html/?path=/docs/library-pptxtohtml--overview)**

## Getting Started

With [npm](https://www.npmjs.com/package/@jvmr/pptx-to-html):
```bash
npm install @jvmr/pptx-to-html
```
Import library:
```javascript
import { pptxToHtml } from "@jvmr/pptx-to-html";
```

## Quick Start (Browser)

```ts
import { pptxToHtml } from "@jvmr/pptx-to-html";

async function handleFile(file: File) {
  const arrayBuffer = await file.arrayBuffer();
  const slidesHtml = await pptxToHtml(arrayBuffer, {
    width: 960,
    height: 540,
    scaleToFit: true,
    letterbox: true,
  });

  const container = document.getElementById("slides")!;
  container.innerHTML = slidesHtml.join("\n");
}

document.getElementById("file")!.addEventListener("change", (event) => {
  const input = event.target as HTMLInputElement;
  if (input.files?.[0]) {
    void handleFile(input.files[0]);
  }
});
```

## Quick Start (Node)

In Node, there is no built-in `DOMParser`. Install a DOM parser in your app and provide it through `domParserFactory`.

```bash
npm install @xmldom/xmldom
```

```ts
import { readFile } from "node:fs/promises";
import { DOMParser } from "@xmldom/xmldom";
import { pptxToHtml } from "@jvmr/pptx-to-html";

async function main() {
  const file = await readFile("./example.pptx");
  const slidesHtml = await pptxToHtml(file.buffer, {
    width: 960,
    height: 540,
    scaleToFit: true,
    domParserFactory: () => new DOMParser(),
  });

  console.log(slidesHtml.join("\n\n"));
}

void main();
```

## API

### `pptxToHtml(buffer, config?) => Promise<string[]>`

- `buffer`: `ArrayBuffer` containing the `.pptx` file contents.
- Returns an array of HTML strings, one per slide.

### Options

| Option | Type | Default | Description |
| --- | --- | --- | --- |
| `width` | `number` | PPT base width or `960` | Target viewport width in pixels; use `scaleToFit` to resize slide content |
| `height` | `number` | PPT base height or `540` | Target viewport height in pixels; use `scaleToFit` to resize slide content |
| `scaleToFit` | `boolean` | `false` | Scales the slide viewport into the target container |
| `letterbox` | `boolean` | `true` when `scaleToFit` is `true` | Preserves aspect ratio with bars instead of stretching |
| `domParserFactory` | `() => { parseFromString(xml: string, mime: string): Document }` | `undefined` | Optional parser factory for Node environments |

Notes:

- The library reads the slide base size from `ppt/presentation.xml` (`sldSz`). When unavailable, it defaults to `960x540`.
- All coordinates and sizes are normalized from EMUs to pixels (`EMU / 9525`).

## What It Renders

- Text boxes, including basic paragraphs, bullets, and numbering.
- Images.
- Shapes and connectors.
- Tables with common fills and borders.
- Charts for column, bar, line, area, pie, and scatter types.

## Limitations

- Fidelity is intentionally practical, not pixel-perfect.
- Animations, transitions, SmartArt, embedded audio or video, 3D, and advanced chart options are not supported.
- Fonts are not embedded; rendering depends on the fonts available in the runtime environment.
- Very large presentations may require substantial memory because ZIP contents and images are decoded in memory.

## Environment Support

- Browsers: modern evergreen browsers with ES module support.
- Node: supported through `domParserFactory`; for example, install `@xmldom/xmldom` in the host app and return a new `DOMParser` from the factory.

## License

MIT
