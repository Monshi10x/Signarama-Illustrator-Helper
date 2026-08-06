const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const main = fs.readFileSync(path.join(__dirname, '..', 'js', 'main.js'), 'utf8');
const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

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
