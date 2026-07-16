import assert from "node:assert/strict";
import test from "node:test";
import { pptxToHtml } from "../dist/index.js";

test("expone el convertidor público", () => {
  assert.equal(typeof pptxToHtml, "function");
});

test("rechaza una entrada que no es un archivo PPTX", async () => {
  await assert.rejects(() => pptxToHtml(new ArrayBuffer(0)));
});
