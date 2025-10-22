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
  const script = `${contents}\nmodule.exports = { summarizeSuppliesByLocation_, summarizeTechnicalByLocation_, computeDashboardTopInsights_, sanitizeString_, parsePositiveInteger_ };`;
  const context = { module: {}, console };
  vm.runInNewContext(script, context);
  return context.module.exports;
}

const { summarizeSuppliesByLocation_, summarizeTechnicalByLocation_, computeDashboardTopInsights_ } = loadServerFunctions();

test("summarizeSuppliesByLocation_ ranks all supply requests", () => {
  const records = [
    { id: "REQ-1", fields: { location: "Plant", description: "Gloves", qty: 5 } },
    { id: "REQ-2", fields: { location: "Plant", description: "Gloves", qty: 3 } },
    { fields: { location: "Plant", description: "Masks", qty: 9 } },
    { fields: { location: "Morse Rd.", description: "Gloves", qty: 10 } },
    { fields: { location: "Frantz Rd.", description: "Soap", qty: 2 } }
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
  assert.equal(results[3].location, "Frantz Rd.");
  assert.equal(results[3].item, "Soap");
  assert.equal(results[3].quantity, 2);
});

test("summarizeSuppliesByLocation_ normalizes legacy location labels", () => {
  const records = [
    { id: "REQ-LEGACY-1", fields: { location: "Short North", description: "Gloves", qty: 2 } },
    { id: "REQ-LEGACY-2", fields: { location: "South Dublin", description: "Gloves", qty: 1 } }
  ];
  const results = summarizeSuppliesByLocation_(records);
  assert.equal(results.length, 2);
  assert.equal(results[0].location, "Short N.");
  assert.equal(results[1].location, "Frantz Rd.");
});

test("summarizeSuppliesByLocation_ aggregates quantities by SKU", () => {
  const records = [
    {
      id: "REQ-100",
      fields: { location: "Plant", description: "Glass Cleaner", catalogSku: "GL-100", qty: 2 }
    },
    {
      id: "REQ-101",
      fields: { location: "Plant", description: "Glass Cleaner (Case)", catalogSku: "GL-100", qty: 4 }
    },
    {
      id: "REQ-102",
      fields: { location: "Plant", description: "Glass Cleaner", catalogSku: "GL-100", qty: 1 }
    }
  ];
  const results = summarizeSuppliesByLocation_(records);
  assert.equal(results.length, 1);
  assert.equal(results[0].location, "Plant");
  assert.equal(results[0].catalogSku, "GL-100");
  assert.equal(results[0].quantity, 7);
  assert.equal(results[0].requestCount, 3);
  assert.equal(results[0].item, "Glass Cleaner");
});

test("summarizeSuppliesByLocation_ avoids double-counting the same request", () => {
  const records = [
    { id: "REQ-500", fields: { location: "Plant", description: "Soap", qty: 2 } },
    { id: "req-500", fields: { location: "Plant", description: "Soap", qty: 3 } },
    { id: "REQ-501", fields: { location: "Plant", description: "Soap", qty: 1 } }
  ];
  const results = summarizeSuppliesByLocation_(records);
  assert.equal(results.length, 1);
  assert.equal(results[0].quantity, 6);
  assert.equal(results[0].requestCount, 2);
});

test("summarizeTechnicalByLocation_ normalizes legacy location labels", () => {
  const itRecords = [{ fields: { location: "Short North" } }];
  const maintenanceRecords = [{ fields: { location: "South Dublin" } }];
  const results = summarizeTechnicalByLocation_(itRecords, maintenanceRecords);
  assert.equal(results.length, 2);
  assert.equal(results[0].location, "Short N.");
  assert.equal(results[0].itCount, 1);
  assert.equal(results[1].location, "Frantz Rd.");
  assert.equal(results[1].maintenanceCount, 1);
});

test("computeDashboardTopInsights_ returns all-time top five entries", () => {
  const recordsByType = {
    supplies: [
      { id: "REQ-1", fields: { location: "Plant", description: "Gloves", qty: 10 } },
      { id: "REQ-2", fields: { location: "Plant", description: "Masks", qty: 9 } },
      { id: "REQ-3", fields: { location: "Plant", description: "Cleaner", qty: 8 } },
      { id: "REQ-4", fields: { location: "Plant", description: "Buckets", qty: 7 } },
      { id: "REQ-5", fields: { location: "Plant", description: "Towels", qty: 6 } },
      { id: "REQ-6", fields: { location: "Plant", description: "Bags", qty: 5 } }
    ],
    it: [
      { fields: { location: "Plant" } },
      { fields: { location: "Plant" } },
      { fields: { location: "Short N." } },
      { fields: { location: "Morse Rd." } },
      { fields: { location: "Granville" } },
      { fields: { location: "Newark" } }
    ],
    maintenance: [
      { fields: { location: "Plant" } },
      { fields: { location: "Frantz Rd." } },
      { fields: { location: "Frantz Rd." } },
      { fields: { location: "Muirfield" } }
    ]
  };
  const insights = computeDashboardTopInsights_(recordsByType);
  assert.equal(insights.suppliesTopByLocation.length, 5);
  const topQuantities = insights.suppliesTopByLocation.map(entry => entry.quantity);
  assert.equal(topQuantities.join(","), "10,9,8,7,6");
  assert.equal(insights.suppliesAllByLocation.length, 6);
  assert.equal(insights.itMaintenanceTopByLocation.length, 5);
  assert.equal(insights.itMaintenanceAllByLocation.length, 7);
  const [plantInsight] = insights.itMaintenanceTopByLocation;
  assert.equal(plantInsight.location, "Plant");
  assert.equal(plantInsight.count, 3);
  assert.equal(plantInsight.itCount, 2);
  assert.equal(plantInsight.maintenanceCount, 1);
});
