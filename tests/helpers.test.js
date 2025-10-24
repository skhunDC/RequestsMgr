import { strict as assert } from "node:assert";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { test } from "node:test";
import vm from "node:vm";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

function loadHelpers() {
  const scriptsPath = join(__dirname, "..", "scripts.html");
  const html = readFileSync(scriptsPath, "utf8");
  const startMarker = "function sanitizeText";
  const endMarker = "window.RequestsAppHelpers = {";
  const startIndex = html.indexOf(startMarker);
  const endIndex = html.indexOf(endMarker);
  if (startIndex === -1 || endIndex === -1) {
    throw new Error("Unable to locate helper functions in scripts.html");
  }
  const closingIndex = html.indexOf("};", endIndex);
  const block = html.slice(startIndex, closingIndex + 2);
  const sanitized = block.replace(/window\.RequestsApp\s*=\s*app;\s*/g, "");
  const script = `${sanitized}\nmodule.exports = window.RequestsAppHelpers;`;
  const context = { window: {}, module: {}, console };
  vm.runInNewContext(script, context);
  return context.module.exports;
}

const helpers = loadHelpers();

test("sanitizeText trims input and handles non-strings", () => {
  assert.equal(helpers.sanitizeText("  hello  "), "hello");
  assert.equal(helpers.sanitizeText(42), "");
});

test("parseQty enforces positive integers", () => {
  assert.equal(helpers.parseQty("3.7"), 3);
  assert.equal(helpers.parseQty("-1"), 0);
});

test("validatePayload flags missing fields and returns normalized values", () => {
  const result = helpers.validatePayload({
    description: "  mop  ",
    qty: "2",
    location: "  Main  ",
    notes: "  ok "
  });
  assert.equal(result.valid, true);
  assert.equal(result.value.description, "mop");
  assert.equal(result.value.qty, 2);
  assert.equal(result.value.location, "Main");
  assert.equal(result.value.notes, "ok");
  const invalid = helpers.validatePayload({ description: "", qty: "0" });
  assert.equal(invalid.valid, false);
  assert.ok(invalid.fields.description.length > 0);
  assert.ok(invalid.fields.qty.length > 0);
});

test("buildClientRequestId provides unique-ish identifiers", () => {
  const idA = helpers.buildClientRequestId();
  const idB = helpers.buildClientRequestId();
  assert.ok(typeof idA === "string" && idA.length > 0);
  assert.notEqual(idA, idB);
});

test("formatDate outputs readable strings", () => {
  const formatted = helpers.formatDate("2024-04-01T12:30:00.000Z");
  assert.ok(/Apr/.test(formatted));
});
