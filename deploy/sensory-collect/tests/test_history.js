// Node assertion harness for sensory_history.js. No framework: prints ok lines
// and exits non-zero on the first failed assertion (throw). Run with `node`.
const assert = require("assert");
const H = require("../static/sensory_history.js");

let pass = 0;
function t(name, fn) { fn(); pass++; console.log("ok - " + name); }

const HOUR = 3600 * 1000, DAY = 24 * HOUR, NOW = 1000000000000;

t("prune keeps <24h, drops >=24h and malformed", () => {
  const recs = [
    { ts: NOW - 1 },            // fresh -> keep
    { ts: NOW - DAY + 1 },      // just under 24h -> keep
    { ts: NOW - DAY },          // exactly 24h -> drop
    { ts: NOW - DAY - 1 },      // older -> drop
    { ts: "x" },                // malformed -> drop
    null                        // malformed -> drop
  ];
  assert.strictEqual(H.prune(recs, NOW).length, 2);
  assert.deepStrictEqual(H.prune("nope", NOW), []);
});

t("group folds into tests->files->samples with correct order + labels", () => {
  const recs = [
    { ts: NOW - 5, test_title: "T1", tester: "Al", round: "1", sample: {} },
    { ts: NOW - 4, test_title: "T1", tester: "Al", round: "1", sample: {} },
    { ts: NOW - 3, test_title: "T1", tester: "Bo", round: "2", sample: {} },
    { ts: NOW - 2, test_title: "T2", tester: "Al", round: "N/A", sample: {} }
  ];
  const g = H.group(recs);
  assert.strictEqual(g.length, 2);
  assert.strictEqual(g[0].test, "T1");
  assert.strictEqual(g[0].files.length, 2);
  assert.strictEqual(g[0].files[0].label, "Al R1");
  assert.strictEqual(g[0].files[0].samples.length, 2);
  assert.strictEqual(g[0].files[1].label, "Bo R2");
  assert.strictEqual(g[1].files[0].label, "Al");   // N/A round -> no suffix
});

t("seriesColorHex matches the curated 0..19 palette exactly", () => {
  assert.strictEqual(H.seriesColorHex(0), "#1f77b4");
  assert.strictEqual(H.seriesColorHex(1), "#d62728");
  assert.strictEqual(H.seriesColorHex(8), "#7f7f7f");
  assert.strictEqual(H.seriesColorHex(19), "#6a5acd");
});

t("seriesColorHex golden-angle beyond curated is a valid, distinct hex", () => {
  const a = H.seriesColorHex(20), b = H.seriesColorHex(21);
  assert.ok(/^#[0-9a-f]{6}$/.test(a) && /^#[0-9a-f]{6}$/.test(b));
  assert.notStrictEqual(a, b);
});

t("axisPointXY: score 1 at center, score 9 at the i=0 top vertex", () => {
  const c = H.axisPointXY(0, 1, 5, 100, 100, 80);
  assert.ok(Math.abs(c.x - 100) < 1e-9 && Math.abs(c.y - 100) < 1e-9);
  const top = H.axisPointXY(0, 9, 5, 100, 100, 80);   // 270deg -> straight up
  assert.ok(Math.abs(top.x - 100) < 1e-6 && Math.abs(top.y - 20) < 1e-6);
});

console.log("\n" + pass + " passed");
