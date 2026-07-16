import { LitElement, css, html, nothing } from "lit";
import { unsafeHTML } from "lit/directives/unsafe-html.js";
import { pptxToHtml } from "../../../../dist/index.js";
import { createSampleDeck } from "./createSampleDeck";

export class PptxBrowserDemo extends LitElement {
  static properties = {
    slideWidth: { type: Number, attribute: "slide-width" },
    slideHeight: { type: Number, attribute: "slide-height" },
    scaleToFit: { type: Boolean, attribute: "scale-to-fit", reflect: true },
    letterbox: { type: Boolean, reflect: true },
    autoloadSample: { type: Boolean, attribute: "autoload-sample" },
  };

  static styles = css`
    :host {
      display: block;
      color: #102040;
      font-family: "IBM Plex Sans", "Segoe UI", sans-serif;
    }

    .shell {
      border: 1px solid #d6dee8;
      border-radius: 18px;
      background: linear-gradient(180deg, #ffffff 0%, #f7f9fc 100%);
      box-shadow: 0 18px 48px rgba(16, 32, 64, 0.08);
      overflow: hidden;
    }

    .toolbar {
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 10px;
      padding: 14px 16px;
      border-bottom: 1px solid #e1e7f0;
      background: #ffffff;
    }

    .meta {
      margin-right: auto;
      min-width: 220px;
    }

    .meta strong {
      display: block;
      font-size: 0.95rem;
      font-weight: 600;
    }

    .meta span {
      display: block;
      color: #55657f;
      font-size: 0.8rem;
      margin-top: 2px;
    }

    button,
    .file-label {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      min-height: 36px;
      padding: 0 14px;
      border-radius: 999px;
      border: 1px solid #c7d3e2;
      background: #ffffff;
      color: #102040;
      font: inherit;
      font-size: 0.85rem;
      font-weight: 600;
      cursor: pointer;
    }

    button.primary {
      border-color: #2f6feb;
      background: #2f6feb;
      color: #ffffff;
    }

    button:disabled {
      opacity: 0.6;
      cursor: wait;
    }

    .file-label {
      position: relative;
      overflow: hidden;
    }

    .file-label input {
      position: absolute;
      inset: 0;
      opacity: 0;
      cursor: pointer;
    }

    .body {
      display: grid;
      gap: 12px;
      padding: 16px;
      background:
        radial-gradient(circle at top right, rgba(47, 111, 235, 0.08), transparent 30%),
        linear-gradient(180deg, #f7f9fc 0%, #f2f5fa 100%);
    }

    .status {
      display: flex;
      align-items: center;
      gap: 10px;
      padding: 10px 12px;
      border: 1px solid #d9e3ef;
      border-radius: 12px;
      background: rgba(255, 255, 255, 0.82);
      color: #415168;
      font-size: 0.8rem;
    }

    .status strong {
      color: #102040;
    }

    .error {
      border-color: #f0c6ce;
      background: #fff5f7;
      color: #8a233a;
    }

    .slides {
      display: grid;
      gap: 16px;
    }

    .slide-card {
      display: grid;
      gap: 10px;
      padding: 14px;
      border: 1px solid #d9e3ef;
      border-radius: 16px;
      background: rgba(255, 255, 255, 0.92);
    }

    .slide-head {
      display: flex;
      justify-content: space-between;
      align-items: baseline;
      gap: 12px;
      font-size: 0.82rem;
      color: #5b6a80;
    }

    .slide-head strong {
      color: #102040;
      font-size: 0.86rem;
    }

    .slide-surface {
      overflow-x: auto;
      padding-bottom: 4px;
    }

    .slide-surface > * {
      min-width: fit-content;
    }

    .empty {
      padding: 24px 18px;
      border: 1px dashed #c7d3e2;
      border-radius: 14px;
      background: rgba(255, 255, 255, 0.72);
      color: #55657f;
      font-size: 0.85rem;
    }
  `;

  // `declare` prevents native class fields from shadowing Lit's reactive
  // accessors. Defaults must be assigned after `super()` in the constructor.
  declare slideWidth: number;
  declare slideHeight: number;
  declare scaleToFit: boolean;
  declare letterbox: boolean;
  declare autoloadSample: boolean;

  private isLoading = false;
  private errorMessage = "";
  private statusMessage = "Load the sample deck or upload a .pptx file.";
  private slides: string[] = [];
  private sourceName = "No deck loaded";
  private sourceBuffer: ArrayBuffer | null = null;
  private hasAutoloaded = false;

  constructor() {
    super();
    this.slideWidth = 960;
    this.slideHeight = 540;
    this.scaleToFit = true;
    this.letterbox = false;
    this.autoloadSample = false;
  }

  protected firstUpdated(): void {
    if (this.autoloadSample && !this.hasAutoloaded) {
      this.hasAutoloaded = true;
      void this.loadSampleDeck();
    }
  }

  protected updated(changedProperties: Map<string, unknown>): void {
    const renderOptionChanged =
      changedProperties.has("slideWidth") ||
      changedProperties.has("slideHeight") ||
      changedProperties.has("scaleToFit") ||
      changedProperties.has("letterbox");

    if (renderOptionChanged && this.sourceBuffer) {
      void this.renderDeck(this.sourceBuffer, this.sourceName, false);
    }
  }

  private async loadSampleDeck() {
    const buffer = await createSampleDeck();
    await this.renderDeck(buffer, "Generated sample deck");
  }

  private async handleFileChange(event: Event) {
    const target = event.currentTarget as HTMLInputElement;
    const file = target.files?.[0];

    if (!file) {
      return;
    }

    await this.renderDeck(await file.arrayBuffer(), file.name);
    target.value = "";
  }

  private clearDeck() {
    this.sourceBuffer = null;
    this.sourceName = "No deck loaded";
    this.slides = [];
    this.errorMessage = "";
    this.statusMessage = "Load the sample deck or upload a .pptx file.";
    this.requestUpdate();
  }

  private async renderDeck(buffer: ArrayBuffer, sourceName: string, keepStatus = true) {
    this.isLoading = true;
    this.errorMessage = "";
    if (!keepStatus) {
      this.statusMessage = "Refreshing slide rendering…";
    } else {
      this.statusMessage = "Rendering presentation…";
    }
    this.requestUpdate();

    try {
      const slides = await pptxToHtml(buffer, {
        width: this.slideWidth,
        height: this.slideHeight,
        scaleToFit: this.scaleToFit,
        letterbox: this.letterbox,
      });

      this.sourceBuffer = buffer;
      this.sourceName = sourceName;
      this.slides = slides;
      this.statusMessage = `${slides.length} slide${slides.length === 1 ? "" : "s"} rendered`;
    } catch (error) {
      this.errorMessage = error instanceof Error ? error.message : "The presentation could not be rendered.";
      this.slides = [];
      this.statusMessage = "Rendering failed";
    } finally {
      this.isLoading = false;
      this.requestUpdate();
    }
  }

  render() {
    return html`
      <section class="shell">
        <div class="toolbar">
          <div class="meta">
            <strong>${this.sourceName}</strong>
            <span>${this.slideWidth} × ${this.slideHeight}px · ${this.scaleToFit ? "scaled output" : "native output"}</span>
          </div>
          <button class="primary" ?disabled=${this.isLoading} @click=${this.loadSampleDeck}>
            ${this.isLoading ? "Rendering…" : "Load sample deck"}
          </button>
          <label class="file-label">
            Upload .pptx
            <input type="file" accept=".pptx" @change=${this.handleFileChange} />
          </label>
          <button ?disabled=${this.isLoading || !this.slides.length} @click=${this.clearDeck}>Clear</button>
        </div>

        <div class="body">
          <div class="status ${this.errorMessage ? "error" : ""}">
            <strong>${this.statusMessage}</strong>
            <span>${this.errorMessage || "The sample deck is generated in-memory to exercise the real library API."}</span>
          </div>

          ${this.slides.length
            ? html`
                <div class="slides">
                  ${this.slides.map(
                    (slideHtml, index) => html`
                      <article class="slide-card">
                        <div class="slide-head">
                          <strong>Slide ${index + 1}</strong>
                          <span>HTML output from <code>pptxToHtml()</code></span>
                        </div>
                        <div class="slide-surface">${unsafeHTML(slideHtml)}</div>
                      </article>
                    `,
                  )}
                </div>
              `
            : html`
                <div class="empty">
                  No slides yet. Use the sample deck for a deterministic browser-first example, or upload a presentation from your machine.
                </div>
              `}
        </div>
      </section>
    `;
  }
}

customElements.define("pptx-browser-demo", PptxBrowserDemo);

declare global {
  interface HTMLElementTagNameMap {
    "pptx-browser-demo": PptxBrowserDemo;
  }
}
