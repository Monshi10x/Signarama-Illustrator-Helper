const assert = require('node:assert/strict');
const test = require('node:test');
const logic = require('../js/preflight-logic');

test('straight overlap detects partial and reversed coincident geometry only', () => {
  const a = {a: [0, 0], b: [100, 0]};
  const partial = logic.lineOverlap(a, {a: [130, 0], b: [40, 0]}, 0.03, 0.1);
  assert.ok(partial);
  assert.equal(partial.length, 60);
  assert.equal(logic.lineOverlap(a, {a: [40, 1], b: [130, 1]}, 0.03, 0.1), null);
  assert.equal(logic.lineOverlap(a, {a: [50, -10], b: [50, 10]}, 0.03, 0.1), null);
});

test('Bezier curves flatten and compare asynchronously through the spatial grid', async () => {
  const segment = {a: [0, 0], c1: [25, 40], c2: [75, 40], b: [100, 0]};
  const reverse = {a: [100, 0], c1: [75, 40], c2: [25, 40], b: [0, 0]};
  assert.ok(logic.flattenCurve([segment.a, segment.c1, segment.c2, segment.b], 0.1).length > 1);
  const result = await logic.findOverlapsAsync([
    {pathIndex: 0, segments: [segment]},
    {pathIndex: 1, segments: [reverse]}
  ], {tolerancePt: 0.1, minimumPt: 0.05, gridCellPt: 15});
  assert.ok(result.issues.length > 0);
  assert.ok(result.comparedPairs > 0);
});
