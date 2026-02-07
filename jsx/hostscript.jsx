#target illustrator;
app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;

/* ---------------- JSON polyfill (ExtendScript) ---------------- */
if(typeof JSON === 'undefined') {JSON = {};}
if(!JSON.parse) {JSON.parse = function(s) {return eval('(' + s + ')');};}
if(!JSON.stringify) {
  JSON.stringify = function(o) {return o && o.toSource ? o.toSource() : String(o);};
}

/**
 * Fits the active artboard to the bounding box of *document* artwork.
 * - Ignores selection.
 * - Skips locked/hidden items and guides.
 */
function signarama_helper_fitArtboardToArtwork() {
  if(!app.documents.length) return 'No open document.';

  var doc = app.activeDocument;

  var b = _srh_getDocumentArtworkBounds(doc);
  if(!b) return 'No eligible artwork found (all items locked/hidden/guides).';

  var idx = doc.artboards.getActiveArtboardIndex();
  doc.artboards[idx].artboardRect = [b.left, b.top, b.right, b.bottom];

  return 'Artboard fitted to artwork bounds: ' + _srh_fmtBounds(b);
}

function _srh_fmtBounds(b) {
  function r(n) {return Math.round(n * 1000) / 1000;}
  return '[L ' + r(b.left) + ', T ' + r(b.top) + ', R ' + r(b.right) + ', B ' + r(b.bottom) + ']';
}

function _srh_getDocumentArtworkBounds(doc) {
  var left = null, top = null, right = null, bottom = null;

  // pageItems returns a flat collection of all items in the document (incl. nested)
  var items = doc.pageItems;
  for(var i = 0; i < items.length; i++) {
    var it = items[i];

    if(!it) continue;

    // Skip hidden/locked items early (some types may not expose these cleanly)
    try {if(it.hidden) continue;} catch(_e0) { }
    try {if(it.locked) continue;} catch(_e1) { }
    try {if(it.layer && (it.layer.locked || !it.layer.visible)) continue;} catch(_e2) { }

    // Skip guides
    try {if(it.guides) continue;} catch(_e3) { }
    // Skip guide PathItems (some versions use .guides on PathItem)
    try {if(it.typename === 'PathItem' && it.guides) continue;} catch(_e4) { }

    // Some items can throw on bounds if empty
    var gb = null;
    try {gb = it.geometricBounds;} catch(_e5) { }
    if(!gb) {
      try {gb = it.visibleBounds;} catch(_e6) { }
    }
    if(!gb || gb.length !== 4) continue;

    var l = gb[0], t = gb[1], r = gb[2], b = gb[3];

    if(left === null) {
      left = l; top = t; right = r; bottom = b;
    } else {
      if(l < left) left = l;
      if(t > top) top = t;
      if(r > right) right = r;
      if(b < bottom) bottom = b;
    }
  }

  if(left === null) return null;
  return {left: left, top: top, right: right, bottom: bottom};
}


/* ---------------- Units ---------------- */
function _srh_mm2pt(mm) {return mm * 2.834645669291339;} // 72 / 25.4
function _srh_clamp(n, a, b) {return n < a ? a : (n > b ? b : n);}

function _srh_getLayerByName(doc, name) {
  for(var i = 0; i < doc.layers.length; i++) {
    if(doc.layers[i].name === name) return doc.layers[i];
  }
  return null;
}

function _srh_getOrCreateLayer(doc, name) {
  var l = _srh_getLayerByName(doc, name);
  if(l) return l;
  l = doc.layers.add();
  l.name = name;
  return l;
}

function _srh_bringLayerToFront(layer) {
  try {layer.zOrder(ZOrderMethod.BRINGTOFRONT);} catch(e) { }
}

function _srh_setBleedLayerOrder(cutLayer, originalLayer, bleedLayer) {
  if(bleedLayer) _srh_bringLayerToFront(bleedLayer);
  if(originalLayer) _srh_bringLayerToFront(originalLayer);
  if(cutLayer) _srh_bringLayerToFront(cutLayer);
}

/**
 * Creates an artboard for each selected item, fitted to its visible bounds.
 */
function signarama_helper_createArtboardsFromSelection() {
  if(!app.documents.length) return 'No open document.';
  var doc = app.activeDocument;
  if(!doc.selection || doc.selection.length === 0) return 'No selection. Select one or more items.';

  var sel = doc.selection;
  var created = 0;
  for(var i = 0; i < sel.length; i++) {
    var it = sel[i];
    if(!it) continue;
    var b = null;
    try {b = it.visibleBounds;} catch(_e0) { }
    if(!b || b.length !== 4) {try {b = it.geometricBounds;} catch(_e1) { } }
    if(!b || b.length !== 4) continue;

    // visibleBounds/geometricBounds: [L, T, R, B]
    var rect = [b[0], b[1], b[2], b[3]];
    try {
      doc.artboards.add(rect);
      created++;
    } catch(_e2) { }
  }

  return created ? ('Created ' + created + ' artboard(s) from selection.') : 'No artboards created (could not read bounds).';
}

/**
 * Duplicates selection, converts text to outlines, outlines strokes, then scales DOWN
 * so the duplicate fits within 297×210mm (A4 landscape) by either dimension.
 */
function signarama_helper_duplicateOutlineScaleA4() {
  if(!app.documents.length) return 'No open document.';
  var doc = app.activeDocument;
  if(!doc.selection || doc.selection.length === 0) return 'No selection. Select one or more items.';

  var layer = doc.activeLayer;
  var grp = layer.groupItems.add();
  grp.name = 'SRH_A4_Fit_Copy';

  // Duplicate selection into group
  var sel = doc.selection;
  for(var i = 0; i < sel.length; i++) {
    try {sel[i].duplicate(grp, ElementPlacement.PLACEATEND);} catch(_e0) { }
  }

  // Select the group for menu commands
  doc.selection = null;
  grp.selected = true;

  // Convert text to outlines (within group)
  try {
    var tfs = grp.textFrames;
    for(var t = tfs.length - 1; t >= 0; t--) {
      try {tfs[t].createOutline();} catch(_e1) { }
    }
  } catch(_e2) { }

  // Outline strokes (menu command operates on selection)
  try {app.executeMenuCommand('outline');} catch(_e3) { }

  // Scale down to fit within 297x210mm
  var targetW = _srh_mm2pt(297);
  var targetH = _srh_mm2pt(210);

  var b = null;
  try {b = grp.visibleBounds;} catch(_e4) { }
  if(!b || b.length !== 4) {try {b = grp.geometricBounds;} catch(_e5) { } }
  if(!b || b.length !== 4) return 'Copy created, but could not read bounds for scaling.';

  var w = b[2] - b[0];
  var h = b[1] - b[3];
  if(w <= 0 || h <= 0) return 'Copy created, but bounds were empty.';

  var factor = Math.min(targetW / w, targetH / h);
  if(factor > 1) factor = 1; // scale DOWN only

  if(factor < 0.999) {
    var pct = factor * 100;
    try {grp.resize(pct, pct, true, true, true, true, true, Transformation.CENTER);} catch(_e6) { }
    return 'Duplicate outlined + scaled to fit within 297×210mm (scale ' + (Math.round(pct * 100) / 100) + '%).';
  }

  return 'Duplicate outlined. No scaling needed (already within 297×210mm).';
}

/**
 * Expands selected items by bleed amounts (mm) on each side.
 * Args: topMm, leftMm, bottomMm, rightMm
 */
function signarama_helper_applyBleed(topMm, leftMm, bottomMm, rightMm) {
  if(!app.documents.length) return 'No open document.';
  var doc = app.activeDocument;
  if(!doc.selection || doc.selection.length === 0) return 'No selection. Select one or more items.';

  var t = _srh_mm2pt(Number(topMm) || 0);
  var l = _srh_mm2pt(Number(leftMm) || 0);
  var b = _srh_mm2pt(Number(bottomMm) || 0);
  var r = _srh_mm2pt(Number(rightMm) || 0);

  var sel = doc.selection;

  var bleedLayer = _srh_getOrCreateLayer(doc, "bleed");
  var originalLayer = _srh_getOrCreateLayer(doc, "original");
  var cutLayer = _srh_getLayerByName(doc, "cutline");
  _srh_setBleedLayerOrder(cutLayer, originalLayer, bleedLayer);
  var changed = 0;

  for(var i = 0; i < sel.length; i++) {
    var it = sel[i];
    if(!it) continue;

    var vb = null;
    try {vb = it.visibleBounds;} catch(_e0) { }
    if(!vb || vb.length !== 4) {try {vb = it.geometricBounds;} catch(_e1) { } }
    if(!vb || vb.length !== 4) continue;

    var w = vb[2] - vb[0];
    var h = vb[1] - vb[3];
    if(w <= 0 || h <= 0) continue;

    try {it.duplicate(originalLayer, ElementPlacement.PLACEATBEGINNING);} catch(_eDup) { }

    var newW = w + l + r;
    var newH = h + t + b;
    if(newW <= 0 || newH <= 0) continue;

    var sx = (newW / w) * 100;
    var sy = (newH / h) * 100;

    try {it.resize(sx, sy, true, true, true, true, true, Transformation.CENTER);} catch(_e2) {continue;}

    // Shift for asymmetric bleeds
    var dx = (r - l) / 2;
    var dy = (t - b) / 2;
    try {it.translate(dx, dy);} catch(_e3) { }

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
 * - Optionally duplicates original into top layer "cutline" with no fill and 'Cut Contour' spot stroke
 * Args: JSON string {"offsetMm":number, "createCutline":boolean}
 */
function signarama_helper_applyPathBleed(jsonStr) {
  var doc = app.activeDocument;
  if(!doc) {return "No document.";}
  if(!doc.selection || doc.selection.length === 0) {return "Select at least one item.";}

  var args = {};
  try {args = JSON.parse(String(jsonStr));} catch(e) {args = {};}
  var offsetMm = Number(args.offsetMm || 0);
  var createCutline = !!args.createCutline;

  var offsetPt = _srh_mm2pt(offsetMm);
  if(!(offsetPt > 0)) {return "Offset amount must be > 0 mm.";}

  function _getOrCreateCutContourSpot() {
    var spot = null;
    try {spot = doc.spots.getByName("Cut Contour");} catch(eGet) { }
    if(!spot) {
      try {
        spot = doc.spots.add();
        spot.name = "Cut Contour";
      } catch(eAdd) { }
    }
    if(spot) {
      try {spot.colorType = ColorModel.SPOT;} catch(eType) { }
      var cmyk = new CMYKColor();
      cmyk.cyan = 0;
      cmyk.magenta = 100;
      cmyk.yellow = 0;
      cmyk.black = 0;
      try {spot.color = cmyk;} catch(eColor) { }
    }
    return spot;
  }

  function _getCutContourSpotColor() {
    var spot = _getOrCreateCutContourSpot();
    if(!spot) return null;
    var sc = new SpotColor();
    sc.spot = spot;
    sc.tint = 100;
    return sc;
  }

  function _setCutlineStyleOnItem(it) {
    // Apply to all path items within groups/compounds.
    try {
      if(it.typename === "PathItem") {
        it.filled = false;
        it.stroked = true;
        var sc = _getCutContourSpotColor();
        if(sc) {
          it.strokeColor = sc;
        } else {
          var c = new RGBColor(); c.red = 255; c.green = 0; c.blue = 255;
          it.strokeColor = c;
        }
      } else if(it.typename === "CompoundPathItem") {
        for(var i = 0; i < it.pathItems.length; i++) {_setCutlineStyleOnItem(it.pathItems[i]);}
      } else if(it.typename === "GroupItem") {
        for(var j = 0; j < it.pageItems.length; j++) {_setCutlineStyleOnItem(it.pageItems[j]);}
      } else if(it.typename === "TextFrame") {
        // Convert then style
        var outlined = it.createOutline();
        it.remove();
        _setCutlineStyleOnItem(outlined);
      } else {
        // Ignore unsupported
      }
    } catch(e) { }
  }

  function _offsetOnePath(path) {
    // Returns a new PathItem (offset result) if supported.
    if(!path || path.typename !== "PathItem") return null;

    // Preserve appearance
    var wasFilled = path.filled;
    var wasStroked = path.stroked;
    var fillCol = path.fillColor;
    var strokeCol = path.strokeColor;
    var sw = path.strokeWidth;

    var out = null;

    // Preferred: PathItem.offsetPath (not present in all Illustrator builds)
    if(typeof path.offsetPath === "function") {
      try {
        out = path.offsetPath(offsetPt, JoinType.MITERENDJOIN, 180);
      } catch(e1) {
        // Some versions use different signature.
        try {out = path.offsetPath(offsetPt);} catch(e2) { }
      }
    }

    // Fallback: apply the native "Adobe Offset Path" live effect and expand.
    if(!out) {
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
          } catch(_eExp2) { }
        }

        // After expanding, the duplicate may become a group/compound.
        // We return the top-level expanded item (dup) and remove the original.
        out = dup;
      } catch(eFx) {
        out = null;
      }
    }

    if(out) {
      try {
        // Restore appearance if the operation stripped it.
        out.filled = wasFilled;
        if(wasFilled) out.fillColor = fillCol;
        out.stroked = wasStroked;
        if(wasStroked) {out.strokeColor = strokeCol; out.strokeWidth = sw;}
      } catch(e3) { }
      try {path.remove();} catch(e4) { }
      return out;
    }
    return null;
  }

  function _offsetItem(it) {
    // Offsets paths inside item; returns count of paths offset.
    var count = 0;
    if(!it) return 0;

    if(it.typename === "TextFrame") {
      var outlined = it.createOutline();
      it.remove();
      return _offsetItem(outlined);
    }

    if(it.typename === "PathItem") {
      return _offsetOnePath(it) ? 1 : 0;
    }

    if(it.typename === "CompoundPathItem") {
      // Offset each child path; compound will be destroyed and replaced by offset paths
      var paths = [];
      for(var i = 0; i < it.pathItems.length; i++) {paths.push(it.pathItems[i]);}
      for(var j = 0; j < paths.length; j++) {if(_offsetOnePath(paths[j])) count++;}
      return count;
    }

    if(it.typename === "GroupItem") {
      // Offset all contained pageItems
      var items = [];
      for(var k = 0; k < it.pageItems.length; k++) {items.push(it.pageItems[k]);}
      for(var m = 0; m < items.length; m++) {count += _offsetItem(items[m]);}
      return count;
    }

    // Unsupported types ignored
    return 0;
  }

  var bleedLayer = _srh_getOrCreateLayer(doc, "bleed");
  var originalLayer = _srh_getOrCreateLayer(doc, "original");
  var cutLayer = createCutline ? _srh_getOrCreateLayer(doc, "cutline") : _srh_getLayerByName(doc, "cutline");
  _srh_setBleedLayerOrder(cutLayer, originalLayer, bleedLayer);

  // Snapshot selection because we'll move items to layers.
  var sel = [];
  for(var s = 0; s < doc.selection.length; s++) {sel.push(doc.selection[s]);}

  var cutCount = 0;
  var offsetCount = 0;

  for(var n = 0; n < sel.length; n++) {
    var item = sel[n];
    if(!item || item.locked || item.hidden) continue;

    // Original: duplicate before any changes
    try {item.duplicate(originalLayer, ElementPlacement.PLACEATBEGINNING);} catch(eDupOrig) { }


    // Cutline: duplicate original before bleed changes
    if(createCutline && cutLayer) {
      try {
        var dup = item.duplicate(cutLayer, ElementPlacement.PLACEATBEGINNING);
        _setCutlineStyleOnItem(dup);
        cutCount++;
      } catch(eDup) { }
    }

    // Move target into bleed layer and offset it
    try {item.move(bleedLayer, ElementPlacement.PLACEATBEGINNING);} catch(eMv) { }
    offsetCount += _offsetItem(item);
  }

  // Refresh selection: set to bleed layer items if possible
  try {doc.selection = null;} catch(eSel) { }

  return "Path bleed applied. Offset paths: " + offsetCount + (createCutline ? (", cutlines created: " + cutCount) : "");
}



function signarama_helper_addFilePathTextToArtboards() {
  if(!app.documents.length) return 'No open document.';
  var doc = app.activeDocument;

  var filePath = 'Unsaved document';
  try {
    if(doc.fullName) filePath = doc.fullName.fsName;
  } catch(_e0) { }

  var abCount = doc.artboards.length;
  if(!abCount) return 'No artboards found.';

  var created = 0;

  for(var i = 0; i < abCount; i++) {
    var rect = doc.artboards[i].artboardRect; // [L,T,R,B]
    var left = rect[0], top = rect[1], right = rect[2], bottom = rect[3];
    var aw = right - left;
    var ah = top - bottom;
    if(aw <= 0 || ah <= 0) continue;

    var boxW = aw * 0.60;
    var boxH = _srh_mm2pt(24); // ~24mm tall
    var offset = _srh_mm2pt(5); // 5mm from top

    var boxLeft = left + (aw - boxW) / 2;
    var boxTop = top - offset;

    // Create point text (single-line) centered within the box
    var tf = doc.textFrames.pointText([boxLeft + (boxW / 2), boxTop]);
    tf.contents = filePath;

    try {tf.textRange.paragraphAttributes.justification = Justification.CENTER;} catch(_e1) { }

    // Font size heuristic based on artboard width; clamp to sane range
    var sizePt = _srh_clamp(aw * 0.03, 8, 400);
    try {tf.textRange.characterAttributes.size = sizePt;} catch(_e2) { }

    // Scale font size so the text fills the target width (single-line)
    try {
      var pass = 0;
      while(pass < 2) {
        var tb = tf.visibleBounds; // [L, T, R, B]
        if(tb && tb.length === 4) {
          var textW = tb[2] - tb[0];
          var textH = tb[1] - tb[3];
          if(textW > 0) {
            var scale = boxW / textW;
            var newSize = _srh_clamp(tf.textRange.characterAttributes.size * scale, 6, 96);
            tf.textRange.characterAttributes.size = newSize;
          }
          // Ensure height fits the box to avoid clipping
          if(textH > boxH && textH > 0) {
            var hScale = boxH / textH;
            var hSize = _srh_clamp(tf.textRange.characterAttributes.size * hScale, 6, 96);
            tf.textRange.characterAttributes.size = hSize;
          }
        }
        pass++;
      }
    } catch(_e3) { }

    // Anchor to artboard: top of text bounds should be 10mm below artboard top
    try {
      var desiredTop = top - _srh_mm2pt(10);
      var bb = tf.visibleBounds;
      if(bb && bb.length === 4) {
        var dy = desiredTop - bb[1];
        if(dy !== 0) tf.translate(0, dy);
      }
    } catch(_e4) { }

    created++;
  }

  return created ? ('Added file path labels to ' + created + ' artboard(s).') : 'No labels created.';
}

function signarama_helper_outlineAllText() {
  if(!app.documents.length) return 'No open document.';
  var doc = app.activeDocument;
  var tfs = doc.textFrames;
  if(!tfs || tfs.length === 0) return 'No text frames found.';

  var outlined = 0;
  for(var i = tfs.length - 1; i >= 0; i--) {
    try {
      var tf = tfs[i];
      if(!tf) continue;
      tf.createOutline();
      tf.remove();
      outlined++;
    } catch(_e0) { }
  }

  return outlined ? ('Outlined ' + outlined + ' text frame(s).') : 'No text outlined.';
}

function signarama_helper_setAllFillsStrokes() {
  if(!app.documents.length) return 'No open document.';
  var doc = app.activeDocument;

  var black = new RGBColor();
  black.red = 0; black.green = 0; black.blue = 0;
  var noFill = new NoColor();

  function _applyToPathItem(it) {
    try {
      it.filled = false;
      it.stroked = true;
      it.strokeWidth = 1;
      it.strokeColor = black;
    } catch(_e0) { }
  }

  function _applyToTextFrame(tf) {
    try {
      var ca = tf.textRange.characterAttributes;
      ca.fillColor = noFill;
      ca.strokeColor = black;
      ca.strokeWeight = 1;
    } catch(_e1) { }
  }

  var changed = 0;
  var items = doc.pageItems;
  for(var i = 0; i < items.length; i++) {
    var it = items[i];
    if(!it) continue;
    try {
      if(it.typename === "PathItem") {
        _applyToPathItem(it);
        changed++;
      } else if(it.typename === "CompoundPathItem") {
        for(var j = 0; j < it.pathItems.length; j++) {
          _applyToPathItem(it.pathItems[j]);
        }
        changed++;
      } else if(it.typename === "TextFrame") {
        _applyToTextFrame(it);
        changed++;
      }
    } catch(_e2) { }
  }

  return changed ? ('Updated fills/strokes on ' + changed + ' item(s).') : 'No items updated.';
}

/* ---------------- Dimensions (Atlas) ---------------- */
function _dim_mm2pt(mm) {return mm * 72.0 / 25.4;}
function _dim_pt2mm(pt) {return pt * 25.4 / 72.0;}
function _dim_fmtMm(pt, decimals) {
  var mm = _dim_pt2mm(Math.abs(pt));
  var f = Math.pow(10, decimals | 0);
  return (Math.round(mm * f) / f).toFixed(decimals) + ' mm';
}
function _dim_fmtMmScaled(ptDistance, decimals, scaleFactor) {
  return _dim_fmtMm(ptDistance * (scaleFactor || 1.0), decimals);
}
function _dim_parseHexColorToRGBColor(hex) {
  try {
    if(!hex) return null;
    var s = String(hex).trim();
    if(s[0] === '#') s = s.slice(1);
    if(s.length === 3) {
      s = s.split('').map(function(c) {return c + c;}).join('');
    }
    if(s.length !== 6) return null;
    var r = parseInt(s.slice(0, 2), 16);
    var g = parseInt(s.slice(2, 4), 16);
    var b = parseInt(s.slice(4, 6), 16);
    var c = new RGBColor();
    c.red = r; c.green = g; c.blue = b;
    return c;
  } catch(e) {return null;}
}
function _dim_hexToRGB(hex) {
  var s = String(hex || '').replace('#', '');
  if(s.length === 3) {
    s = s.split('').map(function(h) {return h + h;}).join('');
  }
  if(s.length !== 6) return null;
  var r = parseInt(s.substring(0, 2), 16);
  var g = parseInt(s.substring(2, 4), 16);
  var b = parseInt(s.substring(4, 6), 16);
  var c = new RGBColor();
  c.red = r; c.green = g; c.blue = b;
  return c;
}
function _dim_getRGBFromOpts(opts, key, fallbackFn) {
  var c = _dim_parseHexColorToRGBColor(opts && opts[key]);
  if(c) return c;
  return (fallbackFn && typeof fallbackFn === 'function') ? fallbackFn() : null;
}
function _dim_colorMagenta() {
  var c = new RGBColor();
  c.red = 255; c.green = 0; c.blue = 255;
  return c;
}
function _dim_makeLine(group, x1, y1, x2, y2, strokePt, opts) {
  var p = group.pathItems.add();
  p.setEntirePath([[x1, y1], [x2, y2]]);
  p.stroked = true;
  p.strokeWidth = Math.max(strokePt || 1, 2);
  var lc = null;
  try {
    if(opts && opts.lineColor) {
      if(typeof opts.lineColor === 'object' && typeof opts.lineColor.red !== 'undefined') {
        lc = opts.lineColor;
      } else if(typeof opts.lineColor === 'string') {
        lc = _dim_parseHexColorToRGBColor(opts.lineColor);
      }
    }
  } catch(_eLc) { lc = null; }
  if(!lc) lc = _dim_parseHexColorToRGBColor('#000000');
  try {p.strokeColor = lc;} catch(_eSc) { }
  p.filled = false;
  p.opacity = 100;
  p.blendingMode = BlendModes.NORMAL;
  return p;
}

function _dim_addArrowhead(group, x, y, dir, sizePt, strokePt, lineColor) {
  if(!(sizePt > 0)) return;
  var s = sizePt;
  var a = s * 0.5;
  // Create two lines forming a hollow V
  if(dir === 'LEFT') {
    _dim_makeLine(group, x, y, x + s, y + a, strokePt, {lineColor: lineColor});
    _dim_makeLine(group, x, y, x + s, y - a, strokePt, {lineColor: lineColor});
  } else if(dir === 'RIGHT') {
    _dim_makeLine(group, x, y, x - s, y + a, strokePt, {lineColor: lineColor});
    _dim_makeLine(group, x, y, x - s, y - a, strokePt, {lineColor: lineColor});
  } else if(dir === 'UP') {
    _dim_makeLine(group, x, y, x - a, y - s, strokePt, {lineColor: lineColor});
    _dim_makeLine(group, x, y, x + a, y - s, strokePt, {lineColor: lineColor});
  } else if(dir === 'DOWN') {
    _dim_makeLine(group, x, y, x - a, y + s, strokePt, {lineColor: lineColor});
    _dim_makeLine(group, x, y, x + a, y + s, strokePt, {lineColor: lineColor});
  }
}
function _dim_addArrowheadAlongLine(group, x, y, dx, dy, sizePt, strokePt, lineColor) {
  if(!(sizePt > 0)) return;
  var len = Math.sqrt(dx * dx + dy * dy);
  if(!(len > 0)) return;
  var ux = dx / len;
  var uy = dy / len;
  var nx = -uy;
  var ny = ux;
  var s = sizePt;
  var a = s * 0.5;
  _dim_makeLine(group, x, y, x - ux * s + nx * a, y - uy * s + ny * a, strokePt, {lineColor: lineColor});
  _dim_makeLine(group, x, y, x - ux * s - nx * a, y - uy * s - ny * a, strokePt, {lineColor: lineColor});
}
function _dim_createAnchoredText(text, anchor, x, y, opts, layer) {
  if(app.documents.length === 0) throw new Error("No open document.");
  var doc = app.activeDocument;
  var lyr = layer || doc.activeLayer;
  opts = opts || {};

  var txt = lyr.textFrames.add();
  txt.contents = text;

  if(opts.font) {
    try {
      txt.textRange.characterAttributes.textFont = app.textFonts.getByName(opts.font);
    } catch(e) { }
  }

  txt.textRange.characterAttributes.size = opts.size || 12;

  if(opts.textColor) {
    try {
      var fillCol = _dim_hexToRGB(opts.textColor);
      if(fillCol) {
        var ca = txt.textRange.characterAttributes;
        ca.fillColor = fillCol;
        try {ca.fillOverPrint = false;} catch(_eOP) { }
        try {ca.strokeWeight = 0;} catch(_eSW) { }
        try {ca.strokeColor = new NoColor();} catch(_eSC) { }
        try {ca.stroked = false;} catch(_eSt) { }
      }
    } catch(_e1) { }
  }
  var angle = opts.rotation || 0;
  if(angle) {
    txt.rotate(angle, true, true, true, true);
  }

  var measureCopy = txt.duplicate();
  var outlineGroup = measureCopy.createOutline();
  var b = outlineGroup.visibleBounds; // [x1, y1, x2, y2]
  var midX = (b[0] + b[2]) / 2;
  var midY = (b[1] + b[3]) / 2;

  var ax, ay;
  switch(anchor) {
    case "TL": ax = b[0]; ay = b[1]; break;
    case "TM": ax = midX; ay = b[1]; break;
    case "TR": ax = b[2]; ay = b[1]; break;
    case "L": ax = b[0]; ay = midY; break;
    case "C": ax = midX; ay = midY; break;
    case "R": ax = b[2]; ay = midY; break;
    case "BL": ax = b[0]; ay = b[3]; break;
    case "BM": ax = midX; ay = b[3]; break;
    case "BR": ax = b[2]; ay = b[3]; break;
    default: ax = midX; ay = midY; break;
  }

  var dx = x - ax;
  var dy = y - ay;
  txt.translate(dx, dy);

  outlineGroup.remove();

  return txt;
}

function _dim_ensureLayer(name) {
  var doc = app.activeDocument;
  var lyr;
  try {lyr = doc.layers.getByName(name);}
  catch(_) {
    try {
      lyr = doc.layers.add();
      lyr.name = name;
    } catch(_eIso) {
      // Isolation mode: cannot add sibling layers; fall back to active layer.
      try {lyr = doc.activeLayer;} catch(_eAct) { }
    }
  }
  if(lyr) {
    try {lyr.visible = true; lyr.locked = false; lyr.printable = true;} catch(_eSet) { }
    try {lyr.zOrder(ZOrderMethod.BRINGTOFRONT);} catch(_eZ) { }
  }
  return lyr;
}
function _dim_captureSelection() {
  var out = [];
  try {
    var sel = app.activeDocument.selection;
    if(sel && sel.length) {
      for(var i = 0; i < sel.length; i++) out.push(sel[i]);
    }
  } catch(_) { }
  if(!out.length) {
    try {
      var items = app.activeDocument.pageItems;
      if(items && items.length) {
        for(var j = 0; j < items.length; j++) {
          var it = items[j];
          if(!it) continue;
          try {
            if(it.selected) out.push(it);
          } catch(_eSel) { }
        }
      }
    } catch(_) { }
  }
  return out;
}
function _dim_restoreSelection(list) {
  try {app.activeDocument.selection = list || [];} catch(_) { }
}
function _dim_getMetricsFor(item, measureClippedContent) {
  if(!item) throw new Error('Invalid selection item.');

  // If this is a clipping group, use the clipping path bounds (mask) unless measuring clipped content.
  try {
    if(item.typename === "GroupItem" && item.clipped && item.pageItems && item.pageItems.length > 0) {
      if(measureClippedContent) {
        // Use bounds of the clipped contents (exclude the clipping mask itself).
        var vb = null;
        for(var gi = 0; gi < item.pageItems.length; gi++) {
          var child = item.pageItems[gi];
          if(!child) continue;
          try {
            if(child.typename === "PathItem" && child.clipping) continue;
            if(child.typename === "CompoundPathItem" && child.pathItems && child.pathItems.length && child.pathItems[0].clipping) continue;
          } catch(_eSkip) { }

          var cb = null;
          try {cb = child.visibleBounds;} catch(_eC0) { }
          if(!cb || cb.length !== 4) {try {cb = child.geometricBounds;} catch(_eC1) { cb = null; } }
          if(!cb || cb.length !== 4) continue;

          if(!vb) {
            vb = [cb[0], cb[1], cb[2], cb[3]];
          } else {
            if(cb[0] < vb[0]) vb[0] = cb[0];
            if(cb[1] > vb[1]) vb[1] = cb[1];
            if(cb[2] > vb[2]) vb[2] = cb[2];
            if(cb[3] < vb[3]) vb[3] = cb[3];
          }
        }
        if(vb && vb.length === 4) {
          return {
            left: vb[0],
            top: vb[1],
            right: vb[2],
            bottom: vb[3],
            width: vb[2] - vb[0],
            height: vb[1] - vb[3],
            x: vb[0],
            y: vb[1]
          };
        }
      }
      var maskItem = null;
      for(var mi = 0; mi < item.pageItems.length; mi++) {
        var pi = item.pageItems[mi];
        if(!pi) continue;
        try {
          if(pi.typename === "PathItem" && pi.clipping) {maskItem = pi; break;}
          if(pi.typename === "CompoundPathItem" && pi.pathItems && pi.pathItems.length && pi.pathItems[0].clipping) {maskItem = pi; break;}
        } catch(_eMask) { }
      }
      if(maskItem) {
        try {return _dim_getMetricsFor(maskItem);} catch(_eMaskM) { }
        var mb = null;
        try {mb = maskItem.visibleBounds;} catch(_eM0) { }
        if(!mb || mb.length !== 4) {try {mb = maskItem.geometricBounds;} catch(_eM1) { } }
        if(mb && mb.length === 4) {
          return {
            left: mb[0],
            top: mb[1],
            right: mb[2],
            bottom: mb[3],
            width: mb[2] - mb[0],
            height: mb[1] - mb[3],
            x: mb[0],
            y: mb[1]
          };
        }
      }
    }
  } catch(_eClip) { }

  if(typeof item.position === 'undefined' ||
    typeof item.width === 'undefined' ||
    typeof item.height === 'undefined') {
    var vb = item.visibleBounds; // [top, left, bottom, right]
    var leftVB = vb[1];
    var topVB = vb[0];
    var rightVB = vb[3];
    var bottomVB = vb[2];
    return {
      left: leftVB,
      top: topVB,
      right: rightVB,
      bottom: bottomVB,
      width: rightVB - leftVB,
      height: topVB - bottomVB,
      x: leftVB,
      y: topVB
    };
  }

  var pos = item.position; // [left, top]
  var w = item.width;
  var h = item.height;
  var left = pos[0];
  var top = pos[1];
  return {
    left: left,
    top: top,
    right: left + w,
    bottom: top - h,
    width: w,
    height: h,
    x: left,
    y: top
  };
}
function _dim_drawHorizontalDim(lyr, left, right, yLine, ticLenPt, textPt, strokePt, decimals, side, textOffsetPt, scaleFactor, textColor, lineColor, includeArrowhead, arrowheadSizePt) {
  var g = lyr.groupItems.add();
  _dim_makeLine(g, left, yLine, right, yLine, strokePt, {lineColor: lineColor});

  var half = ticLenPt * 0.5;
  _dim_makeLine(g, left, yLine - half, left, yLine + half, strokePt, {lineColor: lineColor});
  _dim_makeLine(g, right, yLine - half, right, yLine + half, strokePt, {lineColor: lineColor});

  if(includeArrowhead) {
    _dim_addArrowhead(g, left, yLine, 'LEFT', arrowheadSizePt, strokePt, lineColor);
    _dim_addArrowhead(g, right, yLine, 'RIGHT', arrowheadSizePt, strokePt, lineColor);
  }

  var tx = (left + right) * 0.5;

  var txt = null;
  if(side === 'TOP') txt = _dim_createAnchoredText(_dim_fmtMmScaled(right - left, decimals, scaleFactor), "BM", tx, yLine + textOffsetPt, {size: textPt, textColor: textColor}, lyr);
  else if(side === "BOTTOM") txt = _dim_createAnchoredText(_dim_fmtMmScaled(right - left, decimals, scaleFactor), "TM", tx, yLine - textOffsetPt, {size: textPt, textColor: textColor}, lyr);
  if(txt) {
    try {txt.move(g, ElementPlacement.PLACEATEND);} catch(_eMvTxt) { }
  }

  return g;
}
function _dim_drawVerticalDim(lyr, top, bottom, xLine, ticLenPt, textPt, strokePt, decimals, rotation, textOffsetPt, side, scaleFactor, textColor, lineColor, includeArrowhead, arrowheadSizePt) {
  var g = lyr.groupItems.add();
  _dim_makeLine(g, xLine, bottom, xLine, top, strokePt, {lineColor: lineColor});

  var half = ticLenPt * 0.5;
  _dim_makeLine(g, xLine - half, top, xLine + half, top, strokePt, {lineColor: lineColor});
  _dim_makeLine(g, xLine - half, bottom, xLine + half, bottom, strokePt, {lineColor: lineColor});

  if(includeArrowhead) {
    _dim_addArrowhead(g, xLine, top, 'UP', arrowheadSizePt, strokePt, lineColor);
    _dim_addArrowhead(g, xLine, bottom, 'DOWN', arrowheadSizePt, strokePt, lineColor);
  }

  var ty = (top + bottom) * 0.5;
  var txt = null;
  if(side === 'LEFT') txt = _dim_createAnchoredText(_dim_fmtMmScaled(top - bottom, decimals, scaleFactor), "R", xLine - textOffsetPt, ty, {size: textPt, rotation: 90, textColor: textColor}, lyr);
  else if(side === "RIGHT") txt = _dim_createAnchoredText(_dim_fmtMmScaled(top - bottom, decimals, scaleFactor), "L", xLine + textOffsetPt, ty, {size: textPt, rotation: 270, textColor: textColor}, lyr);
  if(txt) {
    try {txt.move(g, ElementPlacement.PLACEATEND);} catch(_eMvTxt) { }
  }

  return g;
}
function _dim_addCenterText(lyr, b, decimals, textPt, scaleFactor, unitText, textColor) {
  try {
    var w = b.right - b.left;
    var h = b.top - b.bottom;
    var txt = String(Math.round(_dim_pt2mm(w * (scaleFactor || 1)))) + unitText + ' x ' + String(Math.round(_dim_pt2mm(h * (scaleFactor || 1)))) + unitText;
    var cx = b.left + w / 2;
    var cy = b.bottom + h / 2;
    _dim_createAnchoredText(txt, "C", cx, cy, {size: textPt, textColor: textColor}, lyr);
    return 1;
  } catch(e) {
    return 0;
  }
}
function _dim_drawForBounds(b, lyr, opts) {
  var added = 0;

  switch(opts.side) {
    case 'TOP': {
      var yLineT = b.top + opts.offsetPt;
      _dim_drawHorizontalDim(lyr, b.left, b.right, yLineT, opts.ticLenPt, opts.textPt, opts.strokePt, opts.decimals, 'TOP', opts.textOffsetPt, opts.scaleFactor, opts.textColor, opts.lineColor, opts.includeArrowhead, opts.arrowheadSizePt);
      added++; break;
    }
    case 'BOTTOM': {
      var yLineB = b.bottom - opts.offsetPt;
      _dim_drawHorizontalDim(lyr, b.left, b.right, yLineB, opts.ticLenPt, opts.textPt, opts.strokePt, opts.decimals, 'BOTTOM', opts.textOffsetPt, opts.scaleFactor, opts.textColor, opts.lineColor, opts.includeArrowhead, opts.arrowheadSizePt);
      added++; break;
    }
    case 'LEFT': {
      var xLineL = b.left - opts.offsetPt;
      _dim_drawVerticalDim(lyr, b.top, b.bottom, xLineL, opts.ticLenPt, opts.textPt, opts.strokePt, opts.decimals, 90, opts.textOffsetPt, opts.side, opts.scaleFactor, opts.textColor, opts.lineColor, opts.includeArrowhead, opts.arrowheadSizePt);
      added++; break;
    }
    case 'RIGHT': {
      var xLineR = b.right + opts.offsetPt;
      _dim_drawVerticalDim(lyr, b.top, b.bottom, xLineR, opts.ticLenPt, opts.textPt, opts.strokePt, opts.decimals, 270, opts.textOffsetPt, opts.side, opts.scaleFactor, opts.textColor, opts.lineColor, opts.includeArrowhead, opts.arrowheadSizePt);
      added++; break;
    }
    case 'H_BOTH': {
      var yLineTop = b.top + opts.offsetPt;
      var yLineBot = b.bottom - opts.offsetPt;
      _dim_drawHorizontalDim(lyr, b.left, b.right, yLineTop, opts.ticLenPt, opts.textPt, opts.strokePt, opts.decimals, 'TOP', opts.textOffsetPt, opts.scaleFactor, opts.textColor, opts.lineColor, opts.includeArrowhead, opts.arrowheadSizePt);
      _dim_drawHorizontalDim(lyr, b.left, b.right, yLineBot, opts.ticLenPt, opts.textPt, opts.strokePt, opts.decimals, 'BOTTOM', opts.textOffsetPt, opts.scaleFactor, opts.textColor, opts.lineColor, opts.includeArrowhead, opts.arrowheadSizePt);
      added += 2; break;
    }
    case 'V_BOTH': {
      var xLineLeft = b.left - opts.offsetPt;
      var xLineRight = b.right + opts.offsetPt;
      _dim_drawVerticalDim(lyr, b.top, b.bottom, xLineLeft, opts.ticLenPt, opts.textPt, opts.strokePt, opts.decimals, 90, opts.textOffsetPt, opts.side, opts.scaleFactor, opts.textColor, opts.lineColor, opts.includeArrowhead, opts.arrowheadSizePt);
      _dim_drawVerticalDim(lyr, b.top, b.bottom, xLineRight, opts.ticLenPt, opts.textPt, opts.strokePt, opts.decimals, 270, opts.textOffsetPt, opts.side, opts.scaleFactor, opts.textColor, opts.lineColor, opts.includeArrowhead, opts.arrowheadSizePt);
      added += 2; break;
    }
    case 'CENTER_TEXT': {
      added += _dim_addCenterText(lyr, b, opts.decimals, opts.textPt, opts.scaleFactor, 'mm', opts.textColor);
      break;
    }
    default:
      break;
  }

  return added;
}
function _dim_run(opts) {
  var doc = app.activeDocument;
  if(!doc) return "No document open.";

  var originalSel = _dim_captureSelection();
  if(!originalSel || !originalSel.length) return "Nothing selected.";

  var offsetPt = _dim_mm2pt(opts.offsetMm || 10);
  var ticLenPt = _dim_mm2pt(opts.ticLenMm || 2);
  var textPt = opts.textPt || 10;
  var strokePt = opts.strokePt || 1;
  var decimals = (opts.decimals | 0);
  var textOffsetPt = _dim_mm2pt(opts.labelGapMm || 0);
  var lineOffsetPt = _dim_mm2pt(opts.offsetMm || 0);
  var measureClippedContent = !!opts.measureClippedContent;

  var lyr = _dim_ensureLayer('Dimensions');
  var textColor = opts.textColor;
  var lineColor = null;
  try {
    lineColor = _dim_hexToRGB(opts.lineColor);
    if(!lineColor) lineColor = _dim_parseHexColorToRGBColor(opts.lineColor);
    if(!lineColor && opts.lineColor && typeof opts.lineColor === 'object' && typeof opts.lineColor.red !== 'undefined') {
      lineColor = opts.lineColor;
    }
  } catch(_eLc2) { lineColor = null; }
  if(!lineColor) lineColor = _dim_hexToRGB('#000000');
  var includeArrowhead = !!opts.includeArrowhead;
  var arrowheadSizePt = opts.arrowheadSizePt || 0;
  var scaleAppearance = opts.scaleAppearance || 1;

  var scaleFactor = 1.0;
  try {
    var sfDoc = app.activeDocument.scaleFactor;
    if(sfDoc && sfDoc > 0) scaleFactor = sfDoc;
  } catch(e) {scaleFactor = 1.0;}

  var dOpts = {
    side: opts.side,
    offsetPt: offsetPt * scaleAppearance,
    ticLenPt: ticLenPt * scaleAppearance,
    textPt: textPt * scaleAppearance,
    strokePt: strokePt,
    decimals: decimals,
    scaleFactor: scaleFactor,
    textOffsetPt: textOffsetPt * scaleAppearance,
    textColor: textColor,
    lineColor: lineColor,
    includeArrowhead: includeArrowhead,
    arrowheadSizePt: arrowheadSizePt
  };

  var objectsProcessed = 0;
  var measuresAdded = 0;

  try {
    for(var i = 0; i < originalSel.length; i++) {
      var item = originalSel[i];
      try {if(item.locked || item.hidden) continue;} catch(_) { }
      var b; try {b = _dim_getMetricsFor(item, measureClippedContent);} catch(e) {continue;}
      measuresAdded += _dim_drawForBounds(b, lyr, dOpts);
      objectsProcessed++;
    }
  } catch(e) {
    _dim_restoreSelection(originalSel);
    return 'Error: ' + e.message;
  }

  _dim_restoreSelection(originalSel);

  if(objectsProcessed === 0) return 'No measurable objects in selection.';
  var msg = 'Added ' + measuresAdded + ' measure' + (measuresAdded === 1 ? '' : 's') +
    ' on ' + objectsProcessed + ' object' + (objectsProcessed === 1 ? '' : 's') + '.';
  return msg;
}

function _dim_runLine(opts) {
  var doc = app.activeDocument;
  if(!doc) return "No document open.";

  var originalSel = _dim_captureSelection();
  if(!originalSel || !originalSel.length) return "Nothing selected.";

  var textPt = opts.textPt || 10;
  var strokePt = opts.strokePt || 1;
  var textOffsetPt = _dim_mm2pt(opts.labelGapMm || 0);
  var lineOffsetPt = _dim_mm2pt(opts.offsetMm || 0);
  var ticLenPt = _dim_mm2pt(opts.ticLenMm || 2);
  var includeArrowhead = !!opts.includeArrowhead;
  var arrowheadSizePt = opts.arrowheadSizePt || 0;
  var scaleAppearance = opts.scaleAppearance || 1;

  var lineColor = _dim_hexToRGB(opts.lineColor) || _dim_parseHexColorToRGBColor(opts.lineColor) || _dim_hexToRGB('#000000');
  var textColor = opts.textColor;

  var lyr = _dim_ensureLayer('Dimensions');

  function _getPathFromItem(it) {
    if(!it) return null;
    if(it.typename === "PathItem") return it;
    if(it.typename === "CompoundPathItem" && it.pathItems && it.pathItems.length) return it.pathItems[0];
    if(it.typename === "GroupItem" && it.pathItems && it.pathItems.length) return it.pathItems[0];
    return null;
  }

  function _getLengthForPath(p) {
    if(!p) return 0;
    try {if(p.length && p.length > 0) return p.length;} catch(_eL0) { }
    try {
      var pts = p.pathPoints;
      if(!pts || pts.length < 2) return 0;
      var total = 0;
      for(var i = 1; i < pts.length; i++) {
        var a = pts[i - 1].anchor;
        var b = pts[i].anchor;
        var dx = b[0] - a[0];
        var dy = b[1] - a[1];
        total += Math.sqrt(dx * dx + dy * dy);
      }
      return total;
    } catch(_eL1) { }
    return 0;
  }

  var added = 0;
  for(var s = 0; s < originalSel.length; s++) {
    var item = originalSel[s];
    var path = _getPathFromItem(item);
    if(!path) continue;

    var pts = path.pathPoints;
    if(!pts || pts.length < 2) continue;

    function _addSegmentMeasure(a0, a1) {
      var dx = a1[0] - a0[0];
      var dy = a1[1] - a0[1];
      var len = Math.sqrt(dx * dx + dy * dy);
      if(!(len > 0)) return 0;

      var nx = -dy / len;
      var ny = dx / len;
      var lineOff = lineOffsetPt * scaleAppearance;
      var textOff = textOffsetPt * scaleAppearance;
      var off = lineOff;

      var sx = a0[0] + nx * off;
      var sy = a0[1] + ny * off;
      var ex = a1[0] + nx * off;
      var ey = a1[1] + ny * off;

      var g = lyr.groupItems.add();
      _dim_makeLine(g, sx, sy, ex, ey, strokePt, {lineColor: lineColor});
      var halfTick = (ticLenPt * scaleAppearance) * 0.5;
      _dim_makeLine(g, sx - nx * halfTick, sy - ny * halfTick, sx + nx * halfTick, sy + ny * halfTick, strokePt, {lineColor: lineColor});
      _dim_makeLine(g, ex - nx * halfTick, ey - ny * halfTick, ex + nx * halfTick, ey + ny * halfTick, strokePt, {lineColor: lineColor});
      if(includeArrowhead) {
        _dim_addArrowheadAlongLine(g, sx, sy, -dx, -dy, arrowheadSizePt, strokePt, lineColor);
        _dim_addArrowheadAlongLine(g, ex, ey, dx, dy, arrowheadSizePt, strokePt, lineColor);
      }

      var midX = (sx + ex) * 0.5 + nx * textOff;
      var midY = (sy + ey) * 0.5 + ny * textOff;
      var angle = Math.atan2(dy, dx) * 180 / Math.PI;
      var label = _dim_fmtMmScaled(len, opts.decimals | 0, (doc.scaleFactor || 1.0));

      var txt = _dim_createAnchoredText(label, "C", midX, midY, {size: textPt * scaleAppearance, rotation: angle, textColor: textColor}, lyr);
      if(txt) {
        try {txt.move(g, ElementPlacement.PLACEATEND);} catch(_eMv) { }
      }

      return 1;
    }

    for(var i = 1; i < pts.length; i++) {
      try {
        added += _addSegmentMeasure(pts[i - 1].anchor, pts[i].anchor);
      } catch(_eSeg) { }
    }
    try {
      if(path.closed && pts.length > 2) {
        added += _addSegmentMeasure(pts[pts.length - 1].anchor, pts[0].anchor);
      }
    } catch(_eClosed) { }
  }

  return added ? ('Added ' + added + ' line measure' + (added === 1 ? '' : 's') + '.') : 'No measurable paths in selection.';
}

this.atlas_dimensions_run = function(json) {
  try {
    var opts;
    if(typeof json === 'string') {
      try {opts = JSON.parse(json);}
      catch(_) {opts = eval('(' + json + ')');}
    } else {
      opts = json;
    }
    return _dim_run(opts);
  } catch(e) {
    return 'Error: ' + e.message;
  }
};

this.atlas_dimensions_runMulti = function(json) {
  var opts;
  try {
    if(typeof json === 'string') {
      try {opts = JSON.parse(json);}
      catch(_) {opts = eval('(' + json + ')');}
    } else {
      opts = json;
    }
    var sides = opts && opts.sides ? opts.sides : [];
    if(!sides || !sides.length) return 'No sides provided.';

    try {app.beginUndoGroup('Add Dimensions');} catch(_e0) { }

    var totalMeasures = 0;
    var totalObjects = 0;
    var lastMsg = '';
    for(var i = 0; i < sides.length; i++) {
      var side = sides[i];
      var runOpts = {};
      for(var k in opts) {if(opts.hasOwnProperty(k) && k !== 'sides') runOpts[k] = opts[k];}
      runOpts.side = side;
      lastMsg = _dim_run(runOpts);
      // Best-effort parse of counts
      var m = /Added\s+(\d+)\s+measure(?:s)?\s+on\s+(\d+)/i.exec(lastMsg);
      if(m) {
        totalMeasures += parseInt(m[1], 10);
        totalObjects = Math.max(totalObjects, parseInt(m[2], 10));
      }
    }

    try {app.endUndoGroup();} catch(_e1) { }

    if(totalMeasures > 0) {
      return 'Added ' + totalMeasures + ' measure' + (totalMeasures === 1 ? '' : 's') +
        ' on ' + totalObjects + ' object' + (totalObjects === 1 ? '' : 's') + '.';
    }
    return lastMsg || 'No measures added.';
  } catch(e) {
    try {app.endUndoGroup();} catch(_e2) { }
    return 'Error: ' + e.message;
  }
};

this.atlas_dimensions_runLine = function(json) {
  try {
    var opts;
    if(typeof json === 'string') {
      try {opts = JSON.parse(json);}
      catch(_) {opts = eval('(' + json + ')');}
    } else {
      opts = json;
    }
    return _dim_runLine(opts);
  } catch(e) {
    return 'Error: ' + e.message;
  }
};

function signarama_helper_createLightbox(jsonStr) {
  if(!app.documents.length) return 'No open document.';
  var doc = app.activeDocument;
  var opts = {};
  try {opts = JSON.parse(String(jsonStr));} catch(e) {opts = {};}

  var wMm = Number(opts.widthMm || 0);
  var hMm = Number(opts.heightMm || 0);
  var dMm = Number(opts.depthMm || 0);
  var type = String(opts.type || 'Acrylic');
  var supports = parseInt(opts.supportCount, 10);
  if(!supports || supports < 0) supports = 0;
  var ledOffsetMm = Number(opts.ledOffsetMm || 0);

  if(!(wMm > 0) || !(hMm > 0)) return 'Width and height must be > 0.';

  var w = _dim_mm2pt(wMm);
  var h = _dim_mm2pt(hMm);
  var ledOffset = _dim_mm2pt(ledOffsetMm);
  var supportW = _dim_mm2pt(25);

  var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; // [L,T,R,B]
  var centerX = (ab[0] + ab[2]) / 2;
  var centerY = (ab[1] + ab[3]) / 2;
  var left = centerX - (w / 2);
  var top = centerY + (h / 2);

  var frameLayer = _srh_getOrCreateLayer(doc, 'lightbox frame');
  var panelLayer = _srh_getOrCreateLayer(doc, 'acm led panel');

  var black = new RGBColor();
  black.red = 0; black.green = 0; black.blue = 0;

  // Frame rectangle + inner offset as a minus-front compound shape
  var inset = _dim_mm2pt(25);
  var innerW = w - (2 * inset);
  var innerH = h - (2 * inset);
  var outer = frameLayer.pathItems.rectangle(top, left, w, h);
  var inner = null;
  if(innerW > 0 && innerH > 0) {
    inner = frameLayer.pathItems.rectangle(top - inset, left + inset, innerW, innerH);
  }

  function _lb_applyStroke(it) {
    try {
      if(it.typename === "PathItem") {
        it.filled = false;
        it.stroked = true;
        it.strokeWidth = 1;
        it.strokeColor = black;
      } else if(it.typename === "CompoundPathItem") {
        for(var i = 0; i < it.pathItems.length; i++) _lb_applyStroke(it.pathItems[i]);
      } else if(it.typename === "GroupItem") {
        for(var j = 0; j < it.pageItems.length; j++) _lb_applyStroke(it.pageItems[j]);
      }
    } catch(_eS) { }
  }

  if(inner) {
    var grp = frameLayer.groupItems.add();
    outer.move(grp, ElementPlacement.PLACEATBEGINNING);
    inner.move(grp, ElementPlacement.PLACEATBEGINNING);
    try {inner.zOrder(ZOrderMethod.BRINGTOFRONT);} catch(_eZ) { }

    var prevSel = doc.selection;
    try {doc.selection = null;} catch(_eSel0) { }
    try {grp.selected = true;} catch(_eSel1) { }
    try {app.executeMenuCommand('Live Pathfinder Subtract');} catch(_ePf0) { }
    try {app.executeMenuCommand('expandStyle');} catch(_ePf1) { }
    try {app.executeMenuCommand('expand');} catch(_ePf2) { }

    var result = null;
    try {
      if(doc.selection && doc.selection.length) result = doc.selection[0];
    } catch(_eSel2) { }
    if(!result) result = grp;
    _lb_applyStroke(result);

    try {doc.selection = prevSel;} catch(_eSel3) { }
  } else {
    _lb_applyStroke(outer);
  }

  // Supports
  if(supports > 0 && w > supportW) {
    var gap = (w - (supports * supportW)) / (supports + 1);
    if(gap < 0) gap = 0;
    for(var i = 0; i < supports; i++) {
      var sx = left + gap * (i + 1) + supportW * i;
      var s = frameLayer.pathItems.rectangle(top, sx, supportW, h);
      try {s.filled = false;} catch(_eS0) { }
      try {s.stroked = true; s.strokeWidth = 1; s.strokeColor = black;} catch(_eS1) { }
    }
  }

  // LED panel
  if(ledOffset > 0) {
    var pw = w - (2 * ledOffset);
    var ph = h - (2 * ledOffset);
    if(pw > 0 && ph > 0) {
      var pLeft = left + ledOffset;
      var pTop = top - ledOffset;
      var panel = panelLayer.pathItems.rectangle(pTop, pLeft, pw, ph);
      try {panel.filled = false;} catch(_eP0) { }
      try {panel.stroked = true; panel.strokeWidth = 1; panel.strokeColor = black;} catch(_eP1) { }
    }
  }

  return 'Lightbox created: ' + wMm + ' x ' + hMm + ' mm, depth ' + dMm + ' mm, type ' + type + ', supports ' + supports + '.';
}

this.atlas_dimensions_clear = function() {
  try {
    var doc = app.activeDocument;
    var lyr = doc.layers.getByName('Dimensions');
    lyr.remove();
    return "Cleared 'Dimensions' layer.";
  } catch(e) {
    return "Nothing to clear (no 'Dimensions' layer).";
  }
};
