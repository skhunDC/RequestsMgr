import { strict as assert } from "node:assert";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { test } from "node:test";
import vm from "node:vm";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

function loadServerFunctions() {
  const scriptPath = join(__dirname, "..", "Code.gs");
  const contents = readFileSync(scriptPath, "utf8");
  const script = `${contents}\nmodule.exports = { summarizeSuppliesByLocation_, sanitizeString_, parsePositiveInteger_ };`;
  const context = { module: {}, console };
  vm.runInNewContext(script, context);
  return context.module.exports;
}

const { summarizeSuppliesByLocation_ } = loadServerFunctions();

test("summarizeSuppliesByLocation_ ranks all supply requests", () => {
  const records = [
    { fields: { location: "Plant", description: "Gloves", qty: 5 } },
    { fields: { location: "Plant", description: "Gloves", qty: 3 } },
    { fields: { location: "Plant", description: "Masks", qty: 9 } },
    { fields: { location: "Morse Rd.", description: "Gloves", qty: 10 } },
    { fields: { location: "South Dublin", description: "Soap", qty: 2 } }
  ];
  const results = summarizeSuppliesByLocation_(records);
  assert.equal(results.length, 4);
  assert.equal(results[0].location, "Morse Rd.");
  assert.equal(results[0].item, "Gloves");
  assert.equal(results[0].quantity, 10);
  assert.equal(results[1].location, "Plant");
  assert.equal(results[1].item, "Masks");
  assert.equal(results[1].quantity, 9);
  assert.equal(results[2].location, "Plant");
  assert.equal(results[2].item, "Gloves");
  assert.equal(results[2].quantity, 8);
  assert.equal(results[2].requestCount, 2);
  assert.equal(results[3].location, "South Dublin");
  assert.equal(results[3].item, "Soap");
  assert.equal(results[3].quantity, 2);
});
