const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const hostscript = fs.readFileSync(path.join(__dirname, '..', 'jsx', 'hostscript.jsx'), 'utf8');

function functionSource(name, nextName) {
  const start = hostscript.indexOf('function ' + name + '(');
  const end = hostscript.indexOf('function ' + nextName + '(', start);
  assert.notEqual(start, -1, name + ' must exist');
  return hostscript.slice(start, end);
}

test('cutfile documents are created with millimetre document units', () => {
  assert.match(hostscript, /docPreset\.units\s*=\s*RulerUnits\.Millimeters/);
  assert.match(hostscript, /addDocument\(startupPreset, docPreset, false\)/);
  assert.doesNotMatch(hostscript, /targetDoc\.rulerUnits\s*=/);
});

test('cutfile artboards are centered on the initial maximum-size canvas', () => {
  assert.match(hostscript, /canvasCenterX\s*=\s*\(Number\(initialArtboard\[0\]\)\s*\+\s*Number\(initialArtboard\[2\]\)\)\s*\/\s*2/);
  assert.match(hostscript, /artboardRect\s*=\s*\[artLeft, artTop, artRight, artBottom\]/);
});

test('cutfile refresh returns before changing an already-current label', () => {
  const resize = functionSource('_srh_cutfileResizeFilePathTextToArtboard', 'signarama_helper_cutfile_refreshFilePathLabels');
  assert.match(resize, /if\(currentContents === filePath\) return false/);
  assert.ok(resize.indexOf('currentContents === filePath') < resize.indexOf('tf.contents = filePath'));
});

test('cutfile refresh is narrowly scoped and never saves or changes Illustrator context', () => {
  const refresh = functionSource('signarama_helper_cutfile_refreshFilePathLabels', 'signarama_helper_makeRouterCutfile');
  assert.match(refresh, /doc\.layers\.getByName\(_SRH_CUTFILE_LABEL_LAYER_NAME\)/);
  assert.match(refresh, /nm === _SRH_CUTFILE_FILE_PATH_TEXT_NAME \|\| note === _SRH_CUTFILE_FILE_PATH_TEXT_NAME/);
  assert.doesNotMatch(refresh, /doc\.save\s*\(/);
  assert.doesNotMatch(refresh, /doc\.selection\s*=/);
  assert.doesNotMatch(refresh, /doc\.activeLayer\s*=/);
  assert.doesNotMatch(refresh, /setActiveArtboardIndex|getActiveArtboardIndex/);
  assert.doesNotMatch(refresh, /app\.redraw\s*\(|app\.executeMenuCommand\s*\(/);
  assert.match(refresh, /'Updated ' \+ updated \+ ' cutfile file path label\(s\)\.'/);
  assert.match(refresh, /'No cutfile file path labels updated\.'/);
});

test('cutfile creation fits the view to the fitted artboard', () => {
  const create = functionSource('_srh_cutfileCreateFromSelection', '_srh_cutfileResizeFilePathTextToArtboard');
  assert.match(create, /targetDoc\.artboards\.setActiveArtboardIndex\(0\)/);
  assert.match(create, /app\.executeMenuCommand\('fitin'\)/);
});
