const assert = require('node:assert/strict');
const test = require('node:test');
const logic = require('../js/dimensions-logic');
const fs = require('node:fs');
const path = require('node:path');
const host = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'hostscript.jsx'), 'utf8');

test('rotation is finite and normalized to (-180, 180]', () => {
  assert.equal(logic.normalizeRotationDegrees(270), -90);
  assert.equal(logic.normalizeRotationDegrees(-270), 90);
  assert.equal(logic.normalizeRotationDegrees(Infinity), 0);
});

test('finite validation rejects malformed values without arbitrary clamping', () => {
  assert.equal(logic.finite('12.5', 1), 12.5);
  assert.equal(logic.finite(NaN, 7), 7);
  assert.equal(logic.finite(-1, 7, {positive: true}), 7);
});

test('payload normalization repairs invalid persisted settings and preserves zero gaps', () => {
  const payload = logic.normalizePayload({textPt: 0, strokePt: 'bad', scaleAppearance: Infinity, labelGapMm: 0, arrowheadSizePt: -2});
  assert.equal(payload.textPt, 10);
  assert.equal(payload.strokePt, 1);
  assert.equal(payload.scaleAppearance, 1);
  assert.equal(payload.labelGapMm, 0);
  assert.equal(payload.arrowheadSizePt, 0);
});

test('multi-side failures stop with side and completed-success count', () => {
  assert.deepEqual(logic.inspectMultiResult('RIGHT', 'Added 1 measure on 1 object.', 1), {ok: true});
  const failure = logic.inspectMultiResult('RIGHT', 'Error: stage=translateText | number=1346458189', 1);
  assert.equal(failure.ok, false);
  assert.match(failure.message, /side=RIGHT after 1 completed measure/);
  assert.match(failure.message, /stage=translateText/);
});

test('host text measurement is guarded and failed measurement groups are cleaned', () => {
  assert.match(host, /readLiveVisibleBounds/);
  assert.match(host, /createOutlineFallback/);
  assert.match(host, /finally \{[\s\S]{0,180}outlineGroup[\s\S]{0,120}measureCopy/);
  assert.match(host, /function _dim_drawHorizontalDim[\s\S]{0,1800}catch\(e\) \{[\s\S]{0,100}g\.remove\(\)/);
  assert.match(host, /rotation: -90/);
  assert.doesNotMatch(host, /txt\.rotate\(angle, true, true, true, true\)/);
});
