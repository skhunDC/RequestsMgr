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
  const script = `${contents}\nmodule.exports = { buildClientRequest_, DEFAULT_STATUS_APPROVER_EMAILS };`;
  const context = { module: {}, console };
  vm.runInNewContext(script, context);
  return context.module.exports;
}

const { buildClientRequest_, DEFAULT_STATUS_APPROVER_EMAILS } = loadServerFunctions();

test("buildClientRequest_ normalizes legacy supplies locations", () => {
  const row = {
    id: "SUP-1",
    ts: "2023-01-01T00:00:00Z",
    requester: "short@example.com",
    status: "pending",
    approver: "",
    description: "",
    qty: 2,
    location: "Short North",
    notes: "",
    eta: ""
  };
  const record = buildClientRequest_("supplies", row);
  assert.equal(record.fields.location, "Short N.");
  assert(record.details.includes("Location: Short N."));
});

test("buildClientRequest_ normalizes legacy maintenance locations", () => {
  const row = {
    id: "MAIN-1",
    ts: "2023-01-01T00:00:00Z",
    requester: "maint@example.com",
    status: "pending",
    approver: "",
    location: "South Dublin",
    issue: "Fix light",
    urgency: "normal",
    accessNotes: ""
  };
  const record = buildClientRequest_("maintenance", row);
  assert.equal(record.fields.location, "Frantz Rd.");
  assert(record.details.includes("Location: Frantz Rd."));
});

test("DEFAULT_STATUS_APPROVER_EMAILS includes the full manager allowlist", () => {
  assert.ok(DEFAULT_STATUS_APPROVER_EMAILS.includes("rbrown@dublincleaners.com"));
  assert.ok(!DEFAULT_STATUS_APPROVER_EMAILS.includes("rbown@dublincleaners.com"));
});
