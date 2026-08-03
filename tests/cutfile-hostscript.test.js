const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const hostscript = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'hostscript.jsx'), 'utf8');

test('cutfile documents are created with millimetre document units', () => {
  assert.match(hostscript, /docPreset\.units\s*=\s*RulerUnits\.Millimeters/);
  assert.match(hostscript, /addDocument\(startupPreset, docPreset, false\)/);
  assert.doesNotMatch(hostscript, /targetDoc\.rulerUnits\s*=/);
});

test('cutfile artboards are centered on the initial maximum-size canvas', () => {
  assert.match(hostscript, /canvasCenterX\s*=\s*\(Number\(initialArtboard\[0\]\)\s*\+\s*Number\(initialArtboard\[2\]\)\)\s*\/\s*2/);
  assert.match(hostscript, /artboardRect\s*=\s*\[artLeft, artTop, artRight, artBottom\]/);
});

test('the active cutfile tick never changes a clean saved document', () => {
  const tickStart = hostscript.indexOf('function signarama_helper_cutfile_tickFilePathLabels()');
  const tickEnd = hostscript.indexOf('function signarama_helper_makeRouterCutfile()', tickStart);
  const tickFunction = hostscript.slice(tickStart, tickEnd);

  assert.match(tickFunction, /if\(doc\.saved\) return/);
  assert.ok(
    tickFunction.indexOf('if(doc.saved) return') < tickFunction.indexOf('doc.layers.getByName'),
    'saved guard must run before a cutfile label is accessed or changed'
  );
});
