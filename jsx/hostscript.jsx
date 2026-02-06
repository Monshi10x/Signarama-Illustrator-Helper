#target illustrator
app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

/* ---------------- JSON polyfill (ExtendScript) ---------------- */
if (typeof JSON === 'undefined') { JSON = {}; }
if (!JSON.parse) { JSON.parse = function (s) { return eval('(' + s + ')'); }; }
if (!JSON.stringify) {
  JSON.stringify = function (o) { return o && o.toSource ? o.toSource() : String(o); };
}

/**
 * Fits the active artboard to the bounding box of *document* artwork.
 * - Ignores selection.
 * - Skips locked/hidden items and guides.
 */
function signarama_helper_fitArtboardToArtwork() {
  if (!app.documents.length) return 'No open document.';

  var doc = app.activeDocument;

  var b = _srh_getDocumentArtworkBounds(doc);
  if (!b) return 'No eligible artwork found (all items locked/hidden/guides).';

  var idx = doc.artboards.getActiveArtboardIndex();
  doc.artboards[idx].artboardRect = [b.left, b.top, b.right, b.bottom];

  return 'Artboard fitted to artwork bounds: ' + _srh_fmtBounds(b);
}

function _srh_fmtBounds(b){
  function r(n){ return Math.round(n*1000)/1000; }
  return '[L ' + r(b.left) + ', T ' + r(b.top) + ', R ' + r(b.right) + ', B ' + r(b.bottom) + ']';
}

function _srh_getDocumentArtworkBounds(doc){
  var left = null, top = null, right = null, bottom = null;

  // pageItems returns a flat collection of all items in the document (incl. nested)
  var items = doc.pageItems;
  for (var i = 0; i < items.length; i++){
    var it = items[i];

    if (!it) continue;

    // Skip hidden/locked items early (some types may not expose these cleanly)
    try { if (it.hidden) continue; } catch(_e0){}
    try { if (it.locked) continue; } catch(_e1){}
    try { if (it.layer && (it.layer.locked || !it.layer.visible)) continue; } catch(_e2){}

    // Skip guides
    try { if (it.guides) continue; } catch(_e3){}
    // Skip guide PathItems (some versions use .guides on PathItem)
    try { if (it.typename === 'PathItem' && it.guides) continue; } catch(_e4){}

    // Some items can throw on bounds if empty
    var gb = null;
    try { gb = it.geometricBounds; } catch(_e5){}
    if (!gb) {
      try { gb = it.visibleBounds; } catch(_e6){}
    }
    if (!gb || gb.length !== 4) continue;

    var l = gb[0], t = gb[1], r = gb[2], b = gb[3];

    if (left === null){
      left = l; top = t; right = r; bottom = b;
    } else {
      if (l < left) left = l;
      if (t > top) top = t;
      if (r > right) right = r;
      if (b < bottom) bottom = b;
    }
  }

  if (left === null) return null;
  return { left:left, top:top, right:right, bottom:bottom };
}


/* ---------------- Units ---------------- */
function _srh_mm2pt(mm){ return mm * 2.834645669291339; } // 72 / 25.4
function _srh_clamp(n, a, b){ return n < a ? a : (n > b ? b : n); }

/**
 * Creates an artboard for each selected item, fitted to its visible bounds.
 */
function signarama_helper_createArtboardsFromSelection() {
  if (!app.documents.length) return 'No open document.';
  var doc = app.activeDocument;
  if (!doc.selection || doc.selection.length === 0) return 'No selection. Select one or more items.';

  var sel = doc.selection;
  var created = 0;
  for (var i = 0; i < sel.length; i++){
    var it = sel[i];
    if (!it) continue;
    var b = null;
    try { b = it.visibleBounds; } catch(_e0){}
    if (!b || b.length !== 4) { try { b = it.geometricBounds; } catch(_e1){} }
    if (!b || b.length !== 4) continue;

    // visibleBounds/geometricBounds: [L, T, R, B]
    var rect = [b[0], b[1], b[2], b[3]];
    try {
      doc.artboards.add(rect);
      created++;
    } catch(_e2){}
  }

  return created ? ('Created ' + created + ' artboard(s) from selection.') : 'No artboards created (could not read bounds).';
}

/**
 * Duplicates selection, converts text to outlines, outlines strokes, then scales DOWN
 * so the duplicate fits within 297×210mm (A4 landscape) by either dimension.
 */
function signarama_helper_duplicateOutlineScaleA4() {
  if (!app.documents.length) return 'No open document.';
  var doc = app.activeDocument;
  if (!doc.selection || doc.selection.length === 0) return 'No selection. Select one or more items.';

  var layer = doc.activeLayer;
  var grp = layer.groupItems.add();
  grp.name = 'SRH_A4_Fit_Copy';

  // Duplicate selection into group
  var sel = doc.selection;
  for (var i = 0; i < sel.length; i++){
    try { sel[i].duplicate(grp, ElementPlacement.PLACEATEND); } catch(_e0){}
  }

  // Select the group for menu commands
  doc.selection = null;
  grp.selected = true;

  // Convert text to outlines (within group)
  try {
    var tfs = grp.textFrames;
    for (var t = tfs.length - 1; t >= 0; t--){
      try { tfs[t].createOutline(); } catch(_e1){}
    }
  } catch(_e2){}

  // Outline strokes (menu command operates on selection)
  try { app.executeMenuCommand('outline'); } catch(_e3){}

  // Scale down to fit within 297x210mm
  var targetW = _srh_mm2pt(297);
  var targetH = _srh_mm2pt(210);

  var b = null;
  try { b = grp.visibleBounds; } catch(_e4){}
  if (!b || b.length !== 4) { try { b = grp.geometricBounds; } catch(_e5){} }
  if (!b || b.length !== 4) return 'Copy created, but could not read bounds for scaling.';

  var w = b[2] - b[0];
  var h = b[1] - b[3];
  if (w <= 0 || h <= 0) return 'Copy created, but bounds were empty.';

  var factor = Math.min(targetW / w, targetH / h);
  if (factor > 1) factor = 1; // scale DOWN only

  if (factor < 0.999) {
    var pct = factor * 100;
    try { grp.resize(pct, pct, true, true, true, true, true, Transformation.CENTER); } catch(_e6){}
    return 'Duplicate outlined + scaled to fit within 297×210mm (scale ' + (Math.round(pct*100)/100) + '%).';
  }

  return 'Duplicate outlined. No scaling needed (already within 297×210mm).';
}

/**
 * Expands selected items by bleed amounts (mm) on each side.
 * Args: topMm, leftMm, bottomMm, rightMm
 */
function signarama_helper_applyBleed(topMm, leftMm, bottomMm, rightMm) {
  if (!app.documents.length) return 'No open document.';
  var doc = app.activeDocument;
  if (!doc.selection || doc.selection.length === 0) return 'No selection. Select one or more items.';

  var t = _srh_mm2pt(Number(topMm) || 0);
  var l = _srh_mm2pt(Number(leftMm) || 0);
  var b = _srh_mm2pt(Number(bottomMm) || 0);
  var r = _srh_mm2pt(Number(rightMm) || 0);

  var sel = doc.selection;
  var changed = 0;

  for (var i = 0; i < sel.length; i++){
    var it = sel[i];
    if (!it) continue;

    var vb = null;
    try { vb = it.visibleBounds; } catch(_e0){}
    if (!vb || vb.length !== 4) { try { vb = it.geometricBounds; } catch(_e1){} }
    if (!vb || vb.length !== 4) continue;

    var w = vb[2] - vb[0];
    var h = vb[1] - vb[3];
    if (w <= 0 || h <= 0) continue;

    var newW = w + l + r;
    var newH = h + t + b;
    if (newW <= 0 || newH <= 0) continue;

    var sx = (newW / w) * 100;
    var sy = (newH / h) * 100;

    try { it.resize(sx, sy, true, true, true, true, true, Transformation.CENTER); } catch(_e2){ continue; }

    // Shift for asymmetric bleeds
    var dx = (r - l) / 2;
    var dy = (t - b) / 2;
    try { it.translate(dx, dy); } catch(_e3){}

    changed++;
  }

  return changed ? ('Applied bleed to ' + changed + ' item(s).') : 'No items updated (could not read bounds).';
}

/**
 * Adds a text element to the top of each artboard with the current file path.
 * Text box width: 60% of artboard width. Centered horizontally. Small offset from top.
 */

/**
 * Apply Offset Path bleed to selected items (paths/letters/shapes).
 * - Moves target items into layer "bleed"
 * - Optionally duplicates original into top layer "cutline" with no fill and #ff00ff stroke
 * Args: JSON string {"offsetMm":number, "createCutline":boolean}
 */
function signarama_helper_applyPathBleed(jsonStr){
  var doc = app.activeDocument;
  if (!doc) { return "No document."; }
  if (!doc.selection || doc.selection.length === 0) { return "Select at least one item."; }

  var args = {};
  try { args = JSON.parse(String(jsonStr)); } catch(e){ args = {}; }
  var offsetMm = Number(args.offsetMm || 0);
  var createCutline = !!args.createCutline;

  var offsetPt = _srh_mm2pt(offsetMm);
  if (!(offsetPt > 0)) { return "Offset amount must be > 0 mm."; }

  function _getOrCreateLayer(name){
    for (var i=0;i<doc.layers.length;i++){
      if (doc.layers[i].name === name) return doc.layers[i];
    }
    var l = doc.layers.add();
    l.name = name;
    return l;
  }

  function _bringLayerToFront(layer){
    try { layer.zOrder(ZOrderMethod.BRINGTOFRONT); } catch(e) {}
  }

  function _getOrCreateCutContourSpot(){
    var name = "Cut Contour";
    for (var i=0;i<doc.spots.length;i++){
      if (doc.spots[i].name === name) return doc.spots[i];
    }
    var spot = doc.spots.add();
    spot.name = name;
    try { spot.colorType = ColorModel.SPOT; } catch(e){}
    var cmyk = new CMYKColor();
    cmyk.cyan = 0; cmyk.magenta = 100; cmyk.yellow = 0; cmyk.black = 0;
    spot.color = cmyk;
    return spot;
  }


  function _setCutlineStyleOnItem(it){
    // Apply to all path items within groups/compounds.
    try {
      if (it.typename === "PathItem") {
        it.filled = false;
        it.stroked = true;
        var spot = _getOrCreateCutContourSpot();
        var sc = new SpotColor();
        sc.spot = spot;
        sc.tint = 100;
        it.strokeColor = sc;
      } else if (it.typename === "CompoundPathItem") {
        for (var i=0;i<it.pathItems.length;i++){ _setCutlineStyleOnItem(it.pathItems[i]); }
      } else if (it.typename === "GroupItem") {
        for (var j=0;j<it.pageItems.length;j++){ _setCutlineStyleOnItem(it.pageItems[j]); }
      } else if (it.typename === "TextFrame") {
        // Convert then style
        var outlined = it.createOutline();
        it.remove();
        _setCutlineStyleOnItem(outlined);
      } else {
        // Ignore unsupported
      }
    } catch(e){}
  }

  function _offsetOnePath(path){
    // Returns a new PathItem (offset result) if supported.
    if (!path || path.typename !== "PathItem") return null;

    // Preserve appearance
    var wasFilled = path.filled;
    var wasStroked = path.stroked;
    var fillCol = path.fillColor;
    var strokeCol = path.strokeColor;
    var sw = path.strokeWidth;

    var out = null;

    // Preferred: PathItem.offsetPath (not present in all Illustrator builds)
    if (typeof path.offsetPath === "function") {
      try {
        out = path.offsetPath(offsetPt, JoinType.MITERENDJOIN, 180);
      } catch(e1) {
        // Some versions use different signature.
        try { out = path.offsetPath(offsetPt); } catch(e2) {}
      }
    }

    // Fallback: apply the native "Adobe Offset Path" live effect and expand.
    if (!out) {
      try {
        var dup = path.duplicate();
        // Keep style on the duplicate; apply effect to change geometry.
        // jntp: 0=miter, 1=round, 2=bevel. mlim: miter limit.
        var ofst = Number(offsetPt);
        // Illustrator is picky about the effect Dict string across versions, so try a couple of common encodings.
        var fx1 = '<LiveEffect name="Adobe Offset Path"><Dict data="R ofst ' + ofst + ' I jntp 0 R mlim 180"/></LiveEffect>';
        var fx2 = '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 180 R ofst ' + ofst + ' I jntp 0"/></LiveEffect>';
        try {
          dup.applyEffect(fx1);
        } catch(_eFx1) {
          dup.applyEffect(fx2);
        }

        // Expand appearance so the offset becomes real paths.
        // Different versions expose different APIs; try both.
        try {
          dup.expandAppearance();
        } catch(_eExp1) {
          try {
            var prevSel = doc.selection;
            doc.selection = null;
            dup.selected = true;
            app.executeMenuCommand('expandStyle');
            app.executeMenuCommand('expand');
            doc.selection = null;
            // restore selection best-effort
            doc.selection = prevSel;
          } catch(_eExp2) {}
        }

        // After expanding, the duplicate may become a group/compound.
        // We return the top-level expanded item (dup) and remove the original.
        out = dup;
      } catch(eFx) {
        out = null;
      }
    }

    if (out) {
      try {
        // Restore appearance if the operation stripped it.
        out.filled = wasFilled;
        if (wasFilled) out.fillColor = fillCol;
        out.stroked = wasStroked;
        if (wasStroked) { out.strokeColor = strokeCol; out.strokeWidth = sw; }
      } catch(e3){}
      try { path.remove(); } catch(e4){}
      return out;
    }
    return null;
  }

  function _offsetItem(it){
    // Offsets paths inside item; returns count of paths offset.
    var count = 0;
    if (!it) return 0;

    if (it.typename === "TextFrame") {
      var outlined = it.createOutline();
      it.remove();
      return _offsetItem(outlined);
    }

    if (it.typename === "PathItem") {
      return _offsetOnePath(it) ? 1 : 0;
    }

    if (it.typename === "CompoundPathItem") {
      // Offset each child path; compound will be destroyed and replaced by offset paths
      var paths = [];
      for (var i=0;i<it.pathItems.length;i++){ paths.push(it.pathItems[i]); }
      for (var j=0;j<paths.length;j++){ if (_offsetOnePath(paths[j])) count++; }
      return count;
    }

    if (it.typename === "GroupItem") {
      // Offset all contained pageItems
      var items = [];
      for (var k=0;k<it.pageItems.length;k++){ items.push(it.pageItems[k]); }
      for (var m=0;m<items.length;m++){ count += _offsetItem(items[m]); }
      return count;
    }

    // Unsupported types ignored
    return 0;
  }

  var bleedLayer = _getOrCreateLayer("bleed");
  var cutLayer = null;
  if (createCutline) {
    cutLayer = _getOrCreateLayer("cutline");
// _bringLayerToFront(cutLayer); // moved to after bleed to keep cutline above bleed
  }
  _bringLayerToFront(bleedLayer);
  if (createCutline && cutLayer) { _bringLayerToFront(cutLayer); }

  // Snapshot selection because we'll move items to layers.
  var sel = [];
  for (var s=0;s<doc.selection.length;s++){ sel.push(doc.selection[s]); }

  var cutCount = 0;
  var offsetCount = 0;

  for (var n=0;n<sel.length;n++){
    var item = sel[n];
    if (!item || item.locked || item.hidden) continue;

    // Cutline: duplicate original before bleed changes
    if (createCutline && cutLayer) {
      try {
        var dup = item.duplicate(cutLayer, ElementPlacement.PLACEATBEGINNING);
        _setCutlineStyleOnItem(dup);
        cutCount++;
      } catch(eDup){}
    }

    // Move target into bleed layer and offset it
    try { item.move(bleedLayer, ElementPlacement.PLACEATBEGINNING); } catch(eMv){}
    offsetCount += _offsetItem(item);
  }

  // Refresh selection: set to bleed layer items if possible
  try { doc.selection = null; } catch(eSel){}

  return "Path bleed applied. Offset paths: " + offsetCount + (createCutline ? (", cutlines created: " + cutCount) : "");
}



function signarama_helper_addFilePathTextToArtboards() {
  if (!app.documents.length) return 'No open document.';
  var doc = app.activeDocument;

  var filePath = 'Unsaved document';
  try {
    if (doc.fullName) filePath = doc.fullName.fsName;
  } catch(_e0){}

  var abCount = doc.artboards.length;
  if (!abCount) return 'No artboards found.';

  var created = 0;

  for (var i = 0; i < abCount; i++){
    var rect = doc.artboards[i].artboardRect; // [L,T,R,B]
    var left = rect[0], top = rect[1], right = rect[2], bottom = rect[3];
    var aw = right - left;
    var ah = top - bottom;
    if (aw <= 0 || ah <= 0) continue;

    var boxW = aw * 0.60;
    var boxH = _srh_mm2pt(12); // ~12mm tall
    var offset = _srh_mm2pt(5); // 5mm from top

    var boxLeft = left + (aw - boxW) / 2;
    var boxTop = top - offset;

    // Create an invisible rectangle to host area text
    var rectPath = doc.pathItems.rectangle(boxTop, boxLeft, boxW, boxH);
    rectPath.stroked = false;
    rectPath.filled = false;

    var tf = doc.textFrames.areaText(rectPath);
    tf.contents = filePath;

    try { tf.textRange.paragraphAttributes.justification = Justification.CENTER; } catch(_e1){}

    // Font size heuristic based on artboard width; clamp to sane range
    var sizePt = _srh_clamp(aw * 0.03, 8, 24);
    try { tf.textRange.characterAttributes.size = sizePt; } catch(_e2){}

    created++;
  }

  return created ? ('Added file path labels to ' + created + ' artboard(s).') : 'No labels created.';
}