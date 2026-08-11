const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const main = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const host = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'hostscript.jsx'), 'utf8');

test('cutfile labels have no recurring host polling or obsolete switch', () => {
  assert.doesNotMatch(main, /cutfileDisableActiveTick|cutfileTickBusy|cutfile_tickFilePathLabels/);
  assert.doesNotMatch(html, /cutfileDisableActiveTick/);
  assert.doesNotMatch(main, /set(?:Interval|Timeout)[\s\S]{0,300}cutfile_refreshFilePathLabels/);
  assert.match(html, /id="btnRefreshCutfilePathLabel"/);
});

test('cutfile refresh is limited to its button and successful creation callback', () => {
  const calls = main.match(/signarama_helper_cutfile_refreshFilePathLabels\(\)/g) || [];
  assert.equal(calls.length, 1, 'one shared explicit refresh helper should contain the only host call');
  assert.match(main, /btnRefreshCutfilePathLabel[\s\S]{0,120}refreshCutfilePathLabels/);
  assert.match(main, /function makeCutfile[\s\S]{0,350}refreshCutfilePathLabels\(\)/);
});

test('all panel evalScript requests use one FIFO queue', () => {
  assert.match(main, /const jsxRequestQueue = \[\]/);
  assert.match(main, /if\(jsxRequestInFlight \|\| !jsxRequestQueue\.length\) return/);
  assert.match(main, /const request = jsxRequestQueue\.shift\(\)/);
  assert.match(main, /finally \{[\s\S]{0,120}jsxRequestInFlight = false;[\s\S]{0,80}drainJSXQueue\(\)/);
  assert.equal((main.match(/cs\.evalScript\(/g) || []).length, 1, 'only the queue drain may invoke CEP evalScript');
});

test('no always-on timer invokes the cutfile refresh', () => {
  const timerBodies = main.match(/setInterval\([\s\S]{0,800}?\},\s*\d+\s*\)/g) || [];
  timerBodies.forEach(body => assert.doesNotMatch(body, /cutfile|activeDocument/));
});

test('toast UI suppresses routine informational and success messages', () => {
  assert.match(main, /if\(type !== 'error' && type !== 'warn'\) return null/);
});

test('transform offers all-artboard sizing and move controls', () => {
  assert.match(html, /option value="artboardsAll"/);
  assert.match(html, /id="transformMoveMode"/);
  assert.match(html, /id="btnTransformMove"/);
  assert.match(main, /atlas_transform_move/);
});

test('transform Y follows Illustrator native downward-positive coordinates', () => {
  assert.match(host, /absolute \? -yPt - r\[1\] : -yPt/);
  assert.match(host, /absolute \? -yPt - b\[1\] : -yPt/);
});

test('Fixings tab supports corners, spacing limits, and quantities', () => {
  assert.match(html, /data-tab="tab-fixings"/);
  assert.match(html, /id="fixingIncludeCorners"/);
  assert.match(html, /option value="corners">Corners only/);
  assert.match(html, /option value="maximum">Maximum spacing/);
  assert.match(html, /option value="minimum">Minimum spacing/);
  assert.match(html, /option value="quantity">Quantity per edge/);
  assert.match(main, /signarama_helper_createFixingHoles/);
  assert.match(host, /function _srh_fixings_create_impl\(json\)/);
  assert.match(host, /circle\.name = 'FIXING_HOLE'/);
  assert.match(host, /if\(mode !== 'corners'\)/);
});

test('Testing console does not log fetched Corebridge result payloads', () => {
  assert.doesNotMatch(main, /corebridgeDebugLog\('fetch raw response text'/);
  assert.doesNotMatch(main, /corebridgeDebugLog\('fetch parsed response'/);
  assert.doesNotMatch(main, /corebridgeDebugLog\('extracted primary rows'/);
  assert.doesNotMatch(main, /corebridgeDebugLog\('filter comparison rows'/);
  assert.doesNotMatch(main, /corebridgeDebugLog\('filtered result'/);
});

test('Preflight presents ten sequential checks and runs double-cut geometry asynchronously', () => {
  assert.match(html, /data-tab="tab-preflight"/);
  assert.equal((html.match(/data-preflight-step="\d+"/g) || []).length, 10);
  assert.match(html, /Double Cutlines/);
  assert.match(main, /PreflightLogic\.findOverlapsAsync/);
  assert.match(main, /collectGeometryPage/);
  assert.match(main, /batchSize: 25/);
  assert.match(main, /Function\('return ' \+ rawText\)\(\)/);
  assert.match(main, /signarama_helper_preflight_extractCutGeometry/);
  assert.match(host, /path\.pathPoints/);
  assert.match(host, /colour\.typename === 'SpotColor'/);
  assert.match(host, /Preflight - Double Cuts/);
  assert.match(host, /nextIndex: i, done: i >= doc\.pathItems\.length/);
});
