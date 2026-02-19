//#target illustrator;
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

function _srh_hasClippingAncestor(item) {
  var p = null;
  try {p = item.parent;} catch(_ePa0) { p = null; }
  while(p) {
    try {
      if(p.typename === "GroupItem" && p.clipped) return true;
    } catch(_ePa1) { }
    try {p = p.parent;} catch(_ePa2) { p = null; }
  }
  return false;
}

function _srh_hasHiddenOrLockedAncestor(item) {
  var p = null;
  try {p = item.parent;} catch(_ePaL0) { p = null; }
  while(p) {
    try {if(p.hidden) return true;} catch(_ePaL1) { }
    try {if(p.locked) return true;} catch(_ePaL2) { }
    try {
      if(p.layer && (p.layer.locked || !p.layer.visible)) return true;
    } catch(_ePaL3) { }
    try {p = p.parent;} catch(_ePaL4) { p = null; }
  }
  return false;
}

function _srh_getClippingPathBounds(groupItem) {
  if(!groupItem) return null;
  try {
    if(!(groupItem.typename === "GroupItem" && groupItem.clipped)) return null;
  } catch(_eGp0) { return null; }
  try {
    var items = groupItem.pageItems;
    for(var i = 0; i < items.length; i++) {
      var pi = items[i];
      try {
        if(pi.typename === "PathItem" && pi.clipping) {
          var b0 = null;
          try {b0 = pi.visibleBounds;} catch(_eCp0) { }
          if(!b0 || b0.length !== 4) {try {b0 = pi.geometricBounds;} catch(_eCp1) { } }
          if(b0 && b0.length === 4) return b0;
        }
      } catch(_eCp2) { }
    }
  } catch(_eGp1) { }
  return null;
}

function _srh_getBoundsExcludeClipped(item) {
  if(!item) return null;
  if(_srh_hasClippingAncestor(item)) return null;
  var b = _srh_getClippingPathBounds(item);
  if(!b || b.length !== 4) {
    try {b = item.visibleBounds;} catch(_eB0) { }
  }
  if(!b || b.length !== 4) {
    try {b = item.geometricBounds;} catch(_eB1) { }
  }
  return (b && b.length === 4) ? b : null;
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
    if(_srh_hasHiddenOrLockedAncestor(it)) continue;

    // Skip guides
    try {if(it.guides) continue;} catch(_e3) { }
    // Skip guide PathItems (some versions use .guides on PathItem)
    try {if(it.typename === 'PathItem' && it.guides) continue;} catch(_e4) { }

    var gb = _srh_getBoundsExcludeClipped(it);
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
function _srh_pt2mm(pt) {return pt / 2.834645669291339;}
function _srh_clamp(n, a, b) {return n < a ? a : (n > b ? b : n);}

var _srh_scaleFactor = 1.0;
var _srh_isLargeArtboard = false;

function _srh_getScaleFactor() {
  var scaleFactor = 1.0;
  try {
    var sfDoc = app.activeDocument.scaleFactor;
    var sfNum = Number(sfDoc);
    if(sfNum && sfNum > 0) scaleFactor = sfNum;
  } catch(_eSf) {scaleFactor = 1.0;}
  _srh_scaleFactor = scaleFactor;
  _srh_isLargeArtboard = scaleFactor > 1.0001;
  return scaleFactor;
}

function _srh_isLargeArtboardDoc() {
  return _srh_getScaleFactor() > 1.0001;
}

function _srh_mm2ptDoc(mm) {
  return _srh_mm2pt(mm) / _srh_getScaleFactor();
}

function _srh_ptDoc(pt) {
  return (pt || 0) / _srh_getScaleFactor();
}

function _srh_pxStrokeDoc(px) {
  return px / _srh_getScaleFactor();
}

function signarama_helper_getArtboardScaleState() {
  if(!app.documents.length) return '{"isLargeArtboard":false,"scaleFactor":1}';
  var sf = _srh_getScaleFactor();
  return JSON.stringify({
    isLargeArtboard: _srh_isLargeArtboard,
    scaleFactor: sf
  });
}

function _srh_round(n, d) {
  var p = Math.pow(10, d || 2);
  return Math.round(n * p) / p;
}

function _srh_colorToCmyk(color) {
  if(!color) return null;
  try {
    if(color.typename === "CMYKColor") {
      return {c: color.cyan, m: color.magenta, y: color.yellow, k: color.black};
    }
    if(color.typename === "SpotColor" && color.spot && color.spot.color) {
      return _srh_colorToCmyk(color.spot.color);
    }
    if(color.typename === "RGBColor") {
      try {
        var cmyk = app.convertSampleColor(ColorSpace.RGB, [color.red, color.green, color.blue], ColorSpace.CMYK, ColorConvertPurpose.defaultpurpose);
        return {c: cmyk[0], m: cmyk[1], y: cmyk[2], k: cmyk[3]};
      } catch(_eConv) {
        // Manual RGB -> CMYK fallback (0-255 to 0-100)
        var r = color.red / 255;
        var g = color.green / 255;
        var b = color.blue / 255;
        var k = 1 - Math.max(r, g, b);
        if(k >= 1) return {c: 0, m: 0, y: 0, k: 100};
        var c = (1 - r - k) / (1 - k);
        var m = (1 - g - k) / (1 - k);
        var y = (1 - b - k) / (1 - k);
        return {c: c * 100, m: m * 100, y: y * 100, k: k * 100};
      }
    }
    if(color.typename === "GrayColor") {
      try {
        var cmykG = app.convertSampleColor(ColorSpace.GRAY, [color.gray], ColorSpace.CMYK, ColorConvertPurpose.defaultpurpose);
        return {c: cmykG[0], m: cmykG[1], y: cmykG[2], k: cmykG[3]};
      } catch(_eConvG) {
        var k2 = 100 - color.gray;
        return {c: 0, m: 0, y: 0, k: k2};
      }
    }
  } catch(e) { }
  return null;
}

function _srh_colorToRgb(color) {
  if(!color) return null;
  try {
    if(color.typename === "RGBColor") {
      return {r: color.red, g: color.green, b: color.blue};
    }
    if(color.typename === "SpotColor" && color.spot && color.spot.color) {
      return _srh_colorToRgb(color.spot.color);
    }
    if(color.typename === "CMYKColor") {
      var rgb = app.convertSampleColor(ColorSpace.CMYK, [color.cyan, color.magenta, color.yellow, color.black], ColorSpace.RGB, ColorConvertPurpose.defaultpurpose);
      return {r: rgb[0], g: rgb[1], b: rgb[2]};
    }
    if(color.typename === "GrayColor") {
      var rgbG = app.convertSampleColor(ColorSpace.GRAY, [color.gray], ColorSpace.RGB, ColorConvertPurpose.defaultpurpose);
      return {r: rgbG[0], g: rgbG[1], b: rgbG[2]};
    }
  } catch(e) { }
  return null;
}

function _srh_cmykToHex(cmyk) {
  if(!cmyk) return "#000000";
  try {
    var rgb = app.convertSampleColor(ColorSpace.CMYK, [cmyk.c, cmyk.m, cmyk.y, cmyk.k], ColorSpace.RGB, ColorConvertPurpose.defaultpurpose);
    function h(v) {v = Math.max(0, Math.min(255, Math.round(v))); var s = v.toString(16); return s.length === 1 ? "0" + s : s;}
    return "#" + h(rgb[0]) + h(rgb[1]) + h(rgb[2]);
  } catch(e) {
    // Manual CMYK -> RGB fallback
    var c = (cmyk.c || 0) / 100;
    var m = (cmyk.m || 0) / 100;
    var y = (cmyk.y || 0) / 100;
    var k = (cmyk.k || 0) / 100;
    var r = 255 * (1 - c) * (1 - k);
    var g = 255 * (1 - m) * (1 - k);
    var b = 255 * (1 - y) * (1 - k);
    function h2(v) {v = Math.max(0, Math.min(255, Math.round(v))); var s2 = v.toString(16); return s2.length === 1 ? "0" + s2 : s2;}
    return "#" + h2(r) + h2(g) + h2(b);
  }
  return "#000000";
}

function _srh_rgbToHex(rgb) {
  if(!rgb) return "#000000";
  function h(v) {v = Math.max(0, Math.min(255, Math.round(v))); var s = v.toString(16); return s.length === 1 ? "0" + s : s;}
  return "#" + h(rgb.r) + h(rgb.g) + h(rgb.b);
}

function _srh_colorKey(cmyk) {
  if(!cmyk) return null;
  return _srh_round(cmyk.c, 2) + "," + _srh_round(cmyk.m, 2) + "," + _srh_round(cmyk.y, 2) + "," + _srh_round(cmyk.k, 2);
}

function _srh_rgbKey(rgb) {
  if(!rgb) return null;
  return _srh_round(rgb.r, 2) + "," + _srh_round(rgb.g, 2) + "," + _srh_round(rgb.b, 2);
}

function _srh_rgbLabel(rgb) {
  if(!rgb) return "";
  return "R " + _srh_round(rgb.r, 2) + " G " + _srh_round(rgb.g, 2) + " B " + _srh_round(rgb.b, 2);
}

function _srh_cmykLabel(cmyk) {
  if(!cmyk) return "";
  return "C " + _srh_round(cmyk.c, 2) + " M " + _srh_round(cmyk.m, 2) + " Y " + _srh_round(cmyk.y, 2) + " K " + _srh_round(cmyk.k, 2);
}

function _srh_walkPageItems(container, cb) {
  if(!container) return;
  var items = null;
  try {items = container.pageItems;} catch(_eItems) {items = null;}
  if(!items) return;
  for(var i = 0; i < items.length; i++) {
    var it = items[i];
    if(!it) continue;
    if(it.typename === "GroupItem") {
      _srh_walkPageItems(it, cb);
      continue;
    }
    if(it.typename === "CompoundPathItem") {
      for(var j = 0; j < it.pathItems.length; j++) {
        cb(it.pathItems[j]);
      }
      continue;
    }
    cb(it);
  }
}

function signarama_helper_getDocumentColors() {
  if(!app.documents.length) return "[]";
  var doc = app.activeDocument;
  var mode = doc.documentColorSpace === DocumentColorSpace.RGB ? "RGB" : "CMYK";
  var seen = {};
  var list = [];
  var debug = {
    totalItems: 0,
    scanned: 0,
    pathItems: 0,
    textFrames: 0,
    skippedRaster: 0,
    skippedPlaced: 0,
    skippedHidden: 0,
    skippedLocked: 0,
    sampleTypes: [],
    totalPathItems: 0,
    totalTextFrames: 0,
    fallbackUsed: false
  };

  function addColor(color, typeLabel) {
    if(!color || color.typename === "NoColor") return;
    if(color.typename === "GradientColor" || color.typename === "PatternColor") return;
    var typeKey = (typeLabel || "fill");
    if(mode === "RGB") {
      var rgb = _srh_colorToRgb(color);
      if(!rgb) return;
      var rKey = _srh_rgbKey(rgb);
      var rgbCompositeKey = rKey + "|" + typeKey;
      if(!rKey || seen[rgbCompositeKey]) return;
      seen[rgbCompositeKey] = true;
      list.push({
        key: rKey,
        type: typeKey,
        hex: _srh_rgbToHex(rgb),
        r: _srh_round(rgb.r, 2),
        g: _srh_round(rgb.g, 2),
        b: _srh_round(rgb.b, 2),
        label: (typeKey.toUpperCase() + "  " + _srh_rgbLabel(rgb))
      });
      return;
    }
    var cmyk = _srh_colorToCmyk(color);
    if(!cmyk) return;
    var key = _srh_colorKey(cmyk);
    var compositeKey = key + "|" + typeKey;
    if(!key || seen[compositeKey]) return;
    seen[compositeKey] = true;
    list.push({
      key: key,
      type: typeKey,
      hex: _srh_cmykToHex(cmyk),
      c: _srh_round(cmyk.c, 2),
      m: _srh_round(cmyk.m, 2),
      y: _srh_round(cmyk.y, 2),
      k: _srh_round(cmyk.k, 2),
      label: (typeKey.toUpperCase() + "  " + _srh_cmykLabel(cmyk))
    });
  }

  _srh_walkPageItems(doc, function(it) {
    if(!it) return;
    debug.totalItems++;
    if(debug.sampleTypes.length < 10) debug.sampleTypes.push(it.typename);
    if(it.typename === "RasterItem") {debug.skippedRaster++; return;}
    if(it.typename === "PlacedItem") {debug.skippedPlaced++; return;}
    if(it.locked) {debug.skippedLocked++; return;}
    if(it.hidden) {debug.skippedHidden++; return;}
    try {
      if(it.layer && (it.layer.locked || !it.layer.visible)) return;
    } catch(_eLayer) { }

    if(it.typename === "PathItem") {
      debug.pathItems++;
      try {if(it.filled) addColor(it.fillColor, "fill");} catch(_eF) { }
      try {if(it.stroked) addColor(it.strokeColor, "stroke");} catch(_eS) { }
      debug.scanned++;
      return;
    }
    if(it.typename === "TextFrame") {
      debug.textFrames++;
      try {
        var tr = it.textRange.characterAttributes;
        addColor(tr.fillColor, "fill");
        addColor(tr.strokeColor, "stroke");
      } catch(_eT) { }
      debug.scanned++;
      return;
    }
  });

  // Fallback scan using direct collections if nothing found
  try {
    debug.totalPathItems = doc.pathItems.length;
    debug.totalTextFrames = doc.textFrames.length;
    if(list.length === 0 && (debug.totalPathItems > 0 || debug.totalTextFrames > 0)) {
      debug.fallbackUsed = true;
      for(var p = 0; p < doc.pathItems.length; p++) {
        var pi = doc.pathItems[p];
        if(!pi || pi.locked || pi.hidden) continue;
        try {if(pi.filled) addColor(pi.fillColor, "fill");} catch(_ePF2) { }
        try {if(pi.stroked) addColor(pi.strokeColor, "stroke");} catch(_ePS2) { }
      }
      for(var t = 0; t < doc.textFrames.length; t++) {
        var tf = doc.textFrames[t];
        if(!tf || tf.locked || tf.hidden) continue;
        try {
          var tr = tf.textRange.characterAttributes;
          addColor(tr.fillColor, "fill");
          addColor(tr.strokeColor, "stroke");
        } catch(_eTF2) { }
      }
    }
  } catch(_eFallback) { }

  return JSON.stringify({colors: list, debug: debug, mode: mode});
}

function signarama_helper_getDocumentColorMode() {
  if(!app.documents.length) return "";
  var doc = app.activeDocument;
  return doc.documentColorSpace === DocumentColorSpace.RGB ? "RGB" : "CMYK";
}

function signarama_helper_replaceColor(jsonStr) {
  if(!app.documents.length) return "No document.";
  var doc = app.activeDocument;
  var args = {};
  try {args = JSON.parse(String(jsonStr));} catch(e) {args = {};}
  var fromKey = String(args.fromKey || "");
  var fromType = String(args.fromType || "");
  var fromMode = String(args.fromMode || "CMYK");
  var fromHex = String(args.fromHex || "");
  var toHex = args.toHex != null ? String(args.toHex) : null;
  var toCmyk = args.toCmyk || null;
  var fromRgb = null;

  if(!fromKey || (!toHex && !toCmyk)) return "Missing arguments.";

  // Convert hex to RGB
  function _hexToRgb(hex) {
    var h = hex.replace('#', '');
    if(h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
    var r = parseInt(h.substring(0, 2), 16);
    var g = parseInt(h.substring(2, 4), 16);
    var b = parseInt(h.substring(4, 6), 16);
    var c = new RGBColor();
    c.red = r; c.green = g; c.blue = b;
    return c;
  }

  function _parseRgbKey(key) {
    if(!key) return null;
    var parts = key.split(",");
    if(parts.length !== 3) return null;
    var r = Number(parts[0]);
    var g = Number(parts[1]);
    var b = Number(parts[2]);
    if(isNaN(r) || isNaN(g) || isNaN(b)) return null;
    return {r: r, g: g, b: b};
  }

  function _rgbClose(a, b, tol) {
    if(!a || !b) return false;
    var t = (tol != null) ? tol : 0.5;
    return Math.abs(a.r - b.r) <= t && Math.abs(a.g - b.g) <= t && Math.abs(a.b - b.b) <= t;
  }

  if(fromMode === "RGB") {
    fromRgb = _parseRgbKey(fromKey);
  }

  var newCmyk = null;
  if(fromMode !== "RGB") {
    newCmyk = new CMYKColor();
    if(toCmyk) {
      newCmyk.cyan = Number(toCmyk.c || 0);
      newCmyk.magenta = Number(toCmyk.m || 0);
      newCmyk.yellow = Number(toCmyk.y || 0);
      newCmyk.black = Number(toCmyk.k || 0);
    } else {
      var rgb = _hexToRgb(toHex);
      var cmykArr = app.convertSampleColor(ColorSpace.RGB, [rgb.red, rgb.green, rgb.blue], ColorSpace.CMYK, ColorConvertPurpose.defaultpurpose);
      newCmyk.cyan = cmykArr[0];
      newCmyk.magenta = cmykArr[1];
      newCmyk.yellow = cmykArr[2];
      newCmyk.black = cmykArr[3];
    }
  }

  var newRgb = null;
  if(toHex) {
    newRgb = _hexToRgb(toHex);
  }

  var updated = 0;
  _srh_walkPageItems(doc, function(it) {
    if(!it) return;
    if(it.typename === "RasterItem" || it.typename === "PlacedItem") return;
    if(it.locked || it.hidden) return;
    try {
      if(it.layer && (it.layer.locked || !it.layer.visible)) return;
    } catch(_eLayer2) { }

    if(it.typename === "PathItem") {
      if(it.filled && (fromType === "" || fromType === "fill")) {
        var k1 = fromMode === "RGB" ? _srh_rgbKey(_srh_colorToRgb(it.fillColor)) : _srh_colorKey(_srh_colorToCmyk(it.fillColor));
        var k1Hex = fromMode === "RGB" ? _srh_rgbToHex(_srh_colorToRgb(it.fillColor)) : "";
        var k1Rgb = fromMode === "RGB" ? _srh_colorToRgb(it.fillColor) : null;
        if(k1 === fromKey || (fromMode === "RGB" && fromHex && k1Hex === fromHex) || (fromMode === "RGB" && fromRgb && _rgbClose(k1Rgb, fromRgb, 0.5))) {
          it.fillColor = fromMode === "RGB" && newRgb ? newRgb : newCmyk;
          updated++;
        }
      }
      if(it.stroked && (fromType === "" || fromType === "stroke")) {
        var k2 = fromMode === "RGB" ? _srh_rgbKey(_srh_colorToRgb(it.strokeColor)) : _srh_colorKey(_srh_colorToCmyk(it.strokeColor));
        var k2Hex = fromMode === "RGB" ? _srh_rgbToHex(_srh_colorToRgb(it.strokeColor)) : "";
        var k2Rgb = fromMode === "RGB" ? _srh_colorToRgb(it.strokeColor) : null;
        if(k2 === fromKey || (fromMode === "RGB" && fromHex && k2Hex === fromHex) || (fromMode === "RGB" && fromRgb && _rgbClose(k2Rgb, fromRgb, 0.5))) {
          it.strokeColor = fromMode === "RGB" && newRgb ? newRgb : newCmyk;
          updated++;
        }
      }
      return;
    }
    if(it.typename === "TextFrame") {
      try {
        var tr = it.textRange.characterAttributes;
        if(fromType === "" || fromType === "fill") {
          var kf = fromMode === "RGB" ? _srh_rgbKey(_srh_colorToRgb(tr.fillColor)) : _srh_colorKey(_srh_colorToCmyk(tr.fillColor));
          var kfHex = fromMode === "RGB" ? _srh_rgbToHex(_srh_colorToRgb(tr.fillColor)) : "";
          var kfRgb = fromMode === "RGB" ? _srh_colorToRgb(tr.fillColor) : null;
          if(kf === fromKey || (fromMode === "RGB" && fromHex && kfHex === fromHex) || (fromMode === "RGB" && fromRgb && _rgbClose(kfRgb, fromRgb, 0.5))) {
            tr.fillColor = fromMode === "RGB" && newRgb ? newRgb : newCmyk;
            updated++;
          }
        }
        if(fromType === "" || fromType === "stroke") {
          var ks = fromMode === "RGB" ? _srh_rgbKey(_srh_colorToRgb(tr.strokeColor)) : _srh_colorKey(_srh_colorToCmyk(tr.strokeColor));
          var ksHex = fromMode === "RGB" ? _srh_rgbToHex(_srh_colorToRgb(tr.strokeColor)) : "";
          var ksRgb = fromMode === "RGB" ? _srh_colorToRgb(tr.strokeColor) : null;
          if(ks === fromKey || (fromMode === "RGB" && fromHex && ksHex === fromHex) || (fromMode === "RGB" && fromRgb && _rgbClose(ksRgb, fromRgb, 0.5))) {
            tr.strokeColor = fromMode === "RGB" && newRgb ? newRgb : newCmyk;
            updated++;
          }
        }
      } catch(_eT) { }
      return;
    }
  });

  return "Updated " + updated + " items.";
}

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
  var seenClippedParents = [];
  var created = 0;
  for(var i = 0; i < sel.length; i++) {
    var it = sel[i];
    if(!it) continue;

    var p = null;
    try {p = it.parent;} catch(_eP0) { p = null; }
    while(p) {
      var isClippedParent = false;
      try {isClippedParent = (p.typename === "GroupItem" && p.clipped);} catch(_eP1) { isClippedParent = false; }
      if(isClippedParent) {
        it = p;
        break;
      }
      try {p = p.parent;} catch(_eP2) { p = null; }
    }

    var isClippedGroup = false;
    try {isClippedGroup = (it.typename === "GroupItem" && it.clipped);} catch(_eP3) { isClippedGroup = false; }
    if(isClippedGroup) {
      var seen = false;
      for(var s = 0; s < seenClippedParents.length; s++) {
        if(seenClippedParents[s] === it) {seen = true; break;}
      }
      if(seen) continue;
      seenClippedParents.push(it);
    }

    var b = _srh_getBoundsExcludeClipped(it);
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
 * Duplicates selection, then scales DOWN
 * so the duplicate fits within 297×210mm (A4 landscape) by either dimension.
 * Keeps artwork live (no outlining). Workflow: scale geometry first, then scale strokes.
 */
function signarama_helper_duplicateOutlineScaleA4(jsonStr) {
  if(!app.documents.length) return 'No open document.';
  var doc = app.activeDocument;
  if(!doc.selection || doc.selection.length === 0) return 'No selection. Select one or more items.';
  var args = {};
  try {args = jsonStr ? JSON.parse(String(jsonStr)) : {};} catch(_eArgs) { args = {}; }
  var shouldRasterize = !!args.rasterize;
  var rasterizeQuality = String(args.rasterizeQuality || 'high').toLowerCase();
  var rasterizeDpi = 300;
  if(rasterizeQuality === 'draft') rasterizeDpi = 72;
  else if(rasterizeQuality === 'medium') rasterizeDpi = 150;
  else if(rasterizeQuality === 'print') rasterizeDpi = 600;
  else if(rasterizeQuality === 'ultra') rasterizeDpi = 1200;
  else if(rasterizeQuality === 'max') rasterizeDpi = 2400;
  else if(rasterizeQuality === 'super') rasterizeDpi = 7200;
  else if(rasterizeQuality === 'insane') rasterizeDpi = 30000;

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

  // Keep all objects live (fills/strokes/text); no outlining/expanding.

  function _srh_scaleAllStrokesInContainer(container, factorScale) {
    if(!container || !(factorScale > 0)) return 0;
    var scaled = 0;
    var minStroke = 0.01; // keep valid non-zero stroke values after heavy downscales
    var seen = [];
    function _scaleStrokeNumber(value) {
      var n = Number(value);
      if(!(isFinite(n) && n > 0)) return null;
      var out = n * factorScale;
      if(out < minStroke) out = minStroke;
      return out;
    }
    function _scaleTextFrameStrokes(tf) {
      if(!tf) return;
      try {
        var tr = tf.textRange;
        if(tr && tr.characterAttributes) {
          try {
            var tw = _scaleStrokeNumber(tr.characterAttributes.strokeWeight);
            if(tw !== null) {tr.characterAttributes.strokeWeight = tw; scaled++;}
          } catch(_eSiTr0) { }
        }
      } catch(_eSiTr1) { }
      try {
        if(tf.textRange && tf.textRange.characters) {
          var chars = tf.textRange.characters;
          for(var ci = 0; ci < chars.length; ci++) {
            try {
              var ca = chars[ci].characterAttributes;
              if(!ca) continue;
              var cw = _scaleStrokeNumber(ca.strokeWeight);
              if(cw !== null) {ca.strokeWeight = cw; scaled++;}
            } catch(_eSiCh0) { }
          }
        }
      } catch(_eSiCh1) { }
    }
    function _scaleItemStroke(it) {
      if(!it) return;
      for(var siSeen = 0; siSeen < seen.length; siSeen++) {
        if(seen[siSeen] === it) return;
      }
      seen.push(it);

      try {
        var hasStroke = true;
        try {if(it.stroked === false) hasStroke = false;} catch(_eSiSt0) { }
        if(hasStroke) {
          var sw = _scaleStrokeNumber(it.strokeWidth);
          if(sw !== null) {
            try {it.strokeWidth = sw; scaled++;} catch(_eSiSw0) { }
          }
        }
      } catch(_eSiSw1) { }
      try {
        if(it.typename === "TextFrame") _scaleTextFrameStrokes(it);
      } catch(_eSiTf0) { }
      try {
        if(it.typename === "CompoundPathItem" && it.pathItems) {
          for(var pi = 0; pi < it.pathItems.length; pi++) {
            _scaleItemStroke(it.pathItems[pi]);
          }
        }
      } catch(_eSiCp0) { }
      try {
        if(it.typename === "GroupItem" && it.pageItems) {
          for(var gi = 0; gi < it.pageItems.length; gi++) {
            _scaleItemStroke(it.pageItems[gi]);
          }
        }
      } catch(_eSiGp0) { }
    }

    _scaleItemStroke(container);
    return scaled;
  }

  // Scale down to fit within 297x210mm.
  // Large artboards use a document scale factor, so convert the A4 target into doc-space points.
  var targetW = _srh_mm2ptDoc(297);
  var targetH = _srh_mm2ptDoc(210);

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
    try {grp.resize(pct, pct, true, true, true, true, 100, Transformation.CENTER);} catch(_e6) { }
    var strokeCount = _srh_scaleAllStrokesInContainer(grp, factor);
    if(shouldRasterize) {
      try {
        // Rasterize each selected item using visible bounds.
        doc.selection = null;
        var gi = null;
        try {gi = grp.pageItems;} catch(_eSelA) { gi = null; }
        if(!gi || gi.length === 0) {
          try {grp.selected = true;} catch(_eSelB) { }
        } else {
          for(var s = 0; s < gi.length; s++) {
            try {gi[s].selected = true;} catch(_eSelC) { }
          }
        }

        var selection = doc.selection;
        if(!selection || selection.length === 0) {
          return 'Duplicate scaled to fit within 297×210mm (scale ' + (Math.round(pct * 100) / 100) + '%, strokes scaled: ' + strokeCount + '). Rasterize skipped (no selectable items).';
        }

        var itemsToRasterize = [];
        for(var rs = 0; rs < selection.length; rs++) {
          itemsToRasterize.push(selection[rs]);
        }

        var options = new RasterizeOptions();
        var rasterSf = _srh_getScaleFactor();
        if(!(rasterSf > 0)) rasterSf = 1;
        // Workaround for large-canvas docs: rasterize an upscaled copy, then scale raster back down.
        var useTempUpscale = rasterSf > 1.0001;
        var baseDpi = rasterizeDpi;
        if(baseDpi < 72) baseDpi = 72;
        if(baseDpi > 2400) baseDpi = 2400;
        var upscalePct = rasterSf * 100;
        var downscalePct = 100 / rasterSf;
        var effectiveDpi = baseDpi * rasterSf;
        try {options.antiAliasingMethod = AntiAliasingMethod.ARTOPTIMIZED;} catch(_eRo0) { }
        try {options.colorModel = RasterizationColorModel.DEFAULTCOLORMODEL;} catch(_eRo1) { }
        try {options.resolution = baseDpi;} catch(_eRo2) { }
        try {options.transparency = true;} catch(_eRo3) { }

        var rasterizedCount = 0;
        for(var r = 0; r < itemsToRasterize.length; r++) {
          var item = itemsToRasterize[r];
          if(!item) continue;
          if(useTempUpscale) {
            // Temporary upscale must include stroke widths so raster detail matches normal-canvas output.
            try {item.resize(upscalePct, upscalePct, true, true, true, true, upscalePct, Transformation.CENTER);} catch(_eUp0) { }
          }
          var bounds = null;
          try {bounds = item.visibleBounds;} catch(_eRb0) { bounds = null; }
          if(!bounds || bounds.length !== 4) continue;
          try {
            var rasterItem = doc.rasterize(item, bounds, options);
            if(useTempUpscale && rasterItem) {
              try {rasterItem.resize(downscalePct, downscalePct, true, true, true, true, 100, Transformation.CENTER);} catch(_eDown0) { }
            }
            rasterizedCount++;
          } catch(_eROne) { }
        }

        if(rasterizedCount > 0) {
          return 'Duplicate scaled to fit within 297×210mm (scale ' + (Math.round(pct * 100) / 100) + '%, strokes scaled: ' + strokeCount + ', rasterized ' + rasterizedCount + ' item(s) at base ' + baseDpi + ' dpi' + (useTempUpscale ? (' with temp x' + (Math.round(rasterSf * 1000) / 1000) + ' upscale (effective ~' + (Math.round(effectiveDpi * 100) / 100) + ' dpi)') : '') + ').';
        }
        return 'Duplicate scaled to fit within 297×210mm (scale ' + (Math.round(pct * 100) / 100) + '%, strokes scaled: ' + strokeCount + '). Rasterize skipped (no valid visible bounds).';
      } catch(_eRast0) {
        return 'Duplicate scaled to fit within 297×210mm (scale ' + (Math.round(pct * 100) / 100) + '%, strokes scaled: ' + strokeCount + '). Rasterize failed: ' + _eRast0;
      }
    }
    return 'Duplicate scaled to fit within 297×210mm (scale ' + (Math.round(pct * 100) / 100) + '%, strokes scaled: ' + strokeCount + ').';
  }

  return 'Duplicate created. No scaling needed (already within 297×210mm).';
}

/**
 * Expands selected items by bleed amounts (mm) on each side.
 * Args: topMm, leftMm, bottomMm, rightMm, excludeClippedContent, keepOriginal, expandArtboards
 */
function signarama_helper_applyBleed(topMm, leftMm, bottomMm, rightMm, excludeClippedContent, keepOriginal, expandArtboards) {
  if(!app.documents.length) return 'No open document.';
  var doc = app.activeDocument;
  if(!doc.selection || doc.selection.length === 0) return 'No selection. Select one or more items.';
  var excludeClipped = (excludeClippedContent === undefined) ? true : !!excludeClippedContent;
  var preserveOriginal = (keepOriginal === undefined) ? true : !!keepOriginal;
  var expandAb = (expandArtboards === undefined) ? false : !!expandArtboards;
  var _dbgLines = [];
  var artboardBoundsByIndex = {};

  function _dbg(msg) {
    try {$.writeln('[SRH][applyBleed] ' + msg);} catch(_eDbg0) { }
    try {_dbgLines.push('[SRH][applyBleed] ' + msg);} catch(_eDbgPush) { }
  }
  function _dbgBounds(b) {
    if(!b || b.length !== 4) return 'null';
    function r(n) {return Math.round(Number(n) * 1000) / 1000;}
    return '[' + r(b[0]) + ', ' + r(b[1]) + ', ' + r(b[2]) + ', ' + r(b[3]) + ']';
  }
  function _dbgType(item) {
    if(!item) return 'null';
    try {return String(item.typename || 'unknown');} catch(_eDbg1) { return 'unknown'; }
  }
  function _dbgName(item) {
    if(!item) return '';
    try {return String(item.name || '');} catch(_eDbg2) { return ''; }
  }
  function _boundsUnion(a, b2) {
    if(!a || a.length !== 4) return b2;
    if(!b2 || b2.length !== 4) return a;
    return [
      Math.min(a[0], b2[0]),
      Math.max(a[1], b2[1]),
      Math.max(a[2], b2[2]),
      Math.min(a[3], b2[3])
    ];
  }
  function _rectIntersects(a, b2) {
    if(!a || a.length !== 4 || !b2 || b2.length !== 4) return false;
    if(a[2] < b2[0]) return false;
    if(a[0] > b2[2]) return false;
    if(a[1] < b2[3]) return false;
    if(a[3] > b2[1]) return false;
    return true;
  }
  function _pointInRect(x, y, r2) {
    if(!r2 || r2.length !== 4) return false;
    return (x >= r2[0] && x <= r2[2] && y <= r2[1] && y >= r2[3]);
  }
  function _findArtboardIndexForBounds(bounds) {
    if(!bounds || bounds.length !== 4) return -1;
    var cx = (bounds[0] + bounds[2]) / 2;
    var cy = (bounds[1] + bounds[3]) / 2;
    var i;
    try {
      for(i = 0; i < doc.artboards.length; i++) {
        var r3 = doc.artboards[i].artboardRect;
        if(_pointInRect(cx, cy, r3)) return i;
      }
      for(i = 0; i < doc.artboards.length; i++) {
        var r4 = doc.artboards[i].artboardRect;
        if(_rectIntersects(bounds, r4)) return i;
      }
      return doc.artboards.getActiveArtboardIndex();
    } catch(_eAbi0) {
      return -1;
    }
  }

  var t = _srh_mm2ptDoc(Number(topMm) || 0);
  var l = _srh_mm2ptDoc(Number(leftMm) || 0);
  var b = _srh_mm2ptDoc(Number(bottomMm) || 0);
  var r = _srh_mm2ptDoc(Number(rightMm) || 0);
  var sf = _srh_getScaleFactor();
  _dbg('START sel=' + doc.selection.length + ', excludeClipped=' + excludeClipped + ', keepOriginal=' + preserveOriginal + ', expandArtboards=' + expandAb + ', scaleFactor=' + sf);
  _dbg('bleed input mm T/L/B/R=' + [topMm, leftMm, bottomMm, rightMm].join('/') + ', doc-pt=' + [t, l, b, r].join('/'));

  var sel = doc.selection;

  var bleedLayer = null;
  var originalLayer = null;
  if(preserveOriginal) {
    bleedLayer = _srh_getOrCreateLayer(doc, "bleed");
    originalLayer = _srh_getOrCreateLayer(doc, "original");
    var cutLayer = _srh_getLayerByName(doc, "cutline");
    _srh_setBleedLayerOrder(cutLayer, originalLayer, bleedLayer);
    try {bleedLayer.visible = true; bleedLayer.locked = false;} catch(_eBL0) { }
    try {originalLayer.visible = true; originalLayer.locked = false;} catch(_eOL0) { }
  }
  var changed = 0;
  var originalsMoved = 0;

  function _getItemBounds(item) {
    if(!item) return null;
    var b = null;
    try {b = item.visibleBounds;} catch(_eIb0) { }
    if(!b || b.length !== 4) {try {b = item.geometricBounds;} catch(_eIb1) { } }
    return (b && b.length === 4) ? b : null;
  }

  function _getClippingMaskItem(groupItem) {
    if(!groupItem) return null;
    try {
      if(!(groupItem.typename === "GroupItem" && groupItem.clipped)) return null;
    } catch(_eGi0) { return null; }
    try {
      var items = groupItem.pageItems;
      for(var i = 0; i < items.length; i++) {
        var pi = items[i];
        try {
          if((pi.typename === "PathItem" || pi.typename === "CompoundPathItem") && pi.clipping) return pi;
        } catch(_eGi1) { }
        try {
          if(pi.typename === "CompoundPathItem" && pi.pathItems && pi.pathItems.length) {
            for(var j = 0; j < pi.pathItems.length; j++) {
              try {
                if(pi.pathItems[j].clipping) return pi;
              } catch(_eGi2) { }
            }
          }
        } catch(_eGi3) { }
      }
    } catch(_eGi4) { }
    return null;
  }

  function _getNearestClippedGroup(item) {
    var p = item;
    while(p) {
      var isClippedGroup = false;
      try {isClippedGroup = (p.typename === "GroupItem" && p.clipped);} catch(_eCg0) { isClippedGroup = false; }
      if(isClippedGroup) return p;
      try {p = p.parent;} catch(_eCg1) { p = null; }
    }
    return null;
  }

  function _scalePathPoints(pathItem, sx, sy, cx, cy) {
    if(!pathItem || pathItem.typename !== "PathItem") return false;
    try {
      var pts = pathItem.pathPoints;
      for(var i = 0; i < pts.length; i++) {
        var p = pts[i];
        var a = p.anchor;
        var l = p.leftDirection;
        var r = p.rightDirection;
        p.anchor = [cx + (a[0] - cx) * sx, cy + (a[1] - cy) * sy];
        p.leftDirection = [cx + (l[0] - cx) * sx, cy + (l[1] - cy) * sy];
        p.rightDirection = [cx + (r[0] - cx) * sx, cy + (r[1] - cy) * sy];
      }
      return true;
    } catch(_eSpp0) { return false; }
  }

  function _scaleMaskGeometry(maskItem, sx, sy, cx, cy) {
    if(!maskItem) return false;
    try {
      if(maskItem.typename === "PathItem") {
        return _scalePathPoints(maskItem, sx, sy, cx, cy);
      }
      if(maskItem.typename === "CompoundPathItem") {
        var ok = false;
        var pItems = maskItem.pathItems;
        for(var i = 0; i < pItems.length; i++) {
          ok = _scalePathPoints(pItems[i], sx, sy, cx, cy) || ok;
        }
        return ok;
      }
    } catch(_eSmg0) { }
    return false;
  }

  var processedItems = [];

  for(var i = 0; i < sel.length; i++) {
    var it = sel[i];
    if(!it) {
      _dbg('sel[' + i + '] skipped: null item');
      continue;
    }
    _dbg('sel[' + i + '] item type=' + _dbgType(it) + ', name="' + _dbgName(it) + '"');

    var sourceItem = it;
    if(excludeClipped) {
      var nearestClipped = _getNearestClippedGroup(it);
      if(nearestClipped) {
        sourceItem = nearestClipped;
        _dbg('sel[' + i + '] nearest clipped parent found -> type=' + _dbgType(sourceItem) + ', name="' + _dbgName(sourceItem) + '"');
      } else {
        _dbg('sel[' + i + '] no clipped parent found');
      }
    }

    var seen = false;
    for(var si = 0; si < processedItems.length; si++) {
      if(processedItems[si] === sourceItem) { seen = true; break; }
    }
    if(seen) {
      _dbg('sel[' + i + '] skipped: source already processed');
      continue;
    }
    processedItems.push(sourceItem);

    it = sourceItem;
    var workItem = it;
    if(preserveOriginal) {
      try {
        it.move(originalLayer, ElementPlacement.PLACEATBEGINNING);
        originalsMoved++;
        _dbg('sel[' + i + '] moved source to original layer');
      } catch(_eMvOrig0) {
        _dbg('sel[' + i + '] FAILED move to original layer: ' + _eMvOrig0);
        continue;
      }
      try {
        workItem = it.duplicate(bleedLayer, ElementPlacement.PLACEATBEGINNING);
        _dbg('sel[' + i + '] duplicated source to bleed layer');
      } catch(_eDupBleed0) {
        _dbg('sel[' + i + '] FAILED duplicate to bleed layer: ' + _eDupBleed0);
        continue;
      }
    }

    var vb = null;
    var target = workItem;
    var clippingMask = null;
    if(excludeClipped) {
      clippingMask = _getClippingMaskItem(workItem);
      if(clippingMask) {
        target = clippingMask;
        _dbg('sel[' + i + '] clipping mask target found: type=' + _dbgType(target) + ', name="' + _dbgName(target) + '"');
      } else {
        _dbg('sel[' + i + '] clipping mask target NOT found. Using workItem type=' + _dbgType(workItem));
      }
    }
    vb = _getItemBounds(target);
    _dbg('sel[' + i + '] target bounds before=' + _dbgBounds(vb));
    if(!vb || vb.length !== 4) {
      _dbg('sel[' + i + '] skipped: invalid target bounds');
      continue;
    }
    var origW = vb[2] - vb[0];
    var origH = vb[1] - vb[3];
    var expectW = origW + l + r;
    var expectH = origH + t + b;
    _dbg('sel[' + i + '] size before w/h=' + origW + '/' + origH + ', expected after w/h=' + expectW + '/' + expectH);

    var prevStrokeW = null;
    try {prevStrokeW = target.strokeWidth;} catch(_eSw0) { prevStrokeW = null; }

    var resized = false;
    var targetType = "";
    try {targetType = target.typename || "";} catch(_eTy0) { targetType = ""; }
    if(excludeClipped && clippingMask && (targetType === "PathItem" || targetType === "CompoundPathItem")) {
      // For clipped groups, explicitly extend mask bounds by bleed each side.
      // This avoids scale/anchor quirks on large-document and grouped clipping cases.
      var desiredBounds = [vb[0] - l, vb[1] + t, vb[2] + r, vb[3] - b];
      _dbg('sel[' + i + '] applying clipping-mask bounds set to=' + _dbgBounds(desiredBounds));
      var wMask = vb[2] - vb[0];
      var hMask = vb[1] - vb[3];
      if(wMask > 0 && hMask > 0) {
        var sxMask = (wMask + l + r) / wMask;
        var syMask = (hMask + t + b) / hMask;
        var cxMask = (vb[0] + vb[2]) / 2;
        var cyMask = (vb[1] + vb[3]) / 2;
        _dbg('sel[' + i + '] clipping-mask point-scale sx/sy=' + sxMask + '/' + syMask + ', center=' + cxMask + '/' + cyMask);
        if(_scaleMaskGeometry(target, sxMask, syMask, cxMask, cyMask)) {
          resized = true;
          _dbg('sel[' + i + '] clipping-mask point-scale OK');
        } else {
          _dbg('sel[' + i + '] clipping-mask point-scale FAILED; trying bounds assignment');
        }
      }
      if(resized) {
        var dxb = (r - l) / 2;
        var dyb = (t - b) / 2;
        try {target.translate(dxb, dyb);} catch(_eTrMask0) { }
        _dbg('sel[' + i + '] clipping-mask translate dx/dy=' + dxb + '/' + dyb);
      }
    }
    if(excludeClipped && clippingMask && !resized && (targetType === "PathItem" || targetType === "CompoundPathItem")) {
      try {
        target.geometricBounds = desiredBounds;
        resized = true;
        _dbg('sel[' + i + '] geometricBounds set OK');
      } catch(_eGb0) { resized = false; }
      if(!resized) {
        _dbg('sel[' + i + '] geometricBounds set FAILED: ' + _eGb0);
        try {
          target.visibleBounds = desiredBounds;
          resized = true;
          _dbg('sel[' + i + '] visibleBounds set OK');
        } catch(_eGb1) { resized = false; }
        if(!resized) _dbg('sel[' + i + '] visibleBounds set FAILED: ' + _eGb1);
      }
    }

    if(!resized) {
      var w = vb[2] - vb[0];
      var h = vb[1] - vb[3];
      if(w <= 0 || h <= 0) {
        _dbg('sel[' + i + '] skipped: non-positive size w=' + w + ', h=' + h);
        continue;
      }

      var newW = w + l + r;
      var newH = h + t + b;
      if(newW <= 0 || newH <= 0) {
        _dbg('sel[' + i + '] skipped: non-positive new size newW=' + newW + ', newH=' + newH);
        continue;
      }

      var sx = (newW / w) * 100;
      var sy = (newH / h) * 100;
      _dbg('sel[' + i + '] fallback resize sx/sy=' + sx + '/' + sy);
      try {
        target.resize(sx, sy, true, true, true, true, 100, Transformation.CENTER);
        resized = true;
        _dbg('sel[' + i + '] resize with lineScale=100 OK');
      } catch(_e2a) {
        _dbg('sel[' + i + '] resize with lineScale=100 FAILED: ' + _e2a);
        try {
          // Fallback for object types/builds that reject numeric line-width scaling param.
          target.resize(sx, sy, true, true, true, true, true, Transformation.CENTER);
          resized = true;
          _dbg('sel[' + i + '] resize fallback boolean lineScale OK');
        } catch(_e2b) { resized = false; }
        if(!resized) _dbg('sel[' + i + '] resize fallback boolean lineScale FAILED: ' + _e2b);
      }
      if(!resized) {
        _dbg('sel[' + i + '] skipped: unable to transform target');
        continue;
      }

      // Shift for asymmetric bleeds
      var dx = (r - l) / 2;
      var dy = (t - b) / 2;
      try {target.translate(dx, dy);} catch(_e3) { }
      _dbg('sel[' + i + '] translate dx/dy=' + dx + '/' + dy);
    }
    try {if(prevStrokeW !== null && isFinite(prevStrokeW)) target.strokeWidth = prevStrokeW;} catch(_eSw1) { }
    var vbAfter = _getItemBounds(target);
    _dbg('sel[' + i + '] target bounds after=' + _dbgBounds(vbAfter));
    if(vbAfter && vbAfter.length === 4) {
      var afterW = vbAfter[2] - vbAfter[0];
      var afterH = vbAfter[1] - vbAfter[3];
      _dbg('sel[' + i + '] size after w/h=' + afterW + '/' + afterH + ', delta w/h=' + (afterW - origW) + '/' + (afterH - origH));
      if(expandAb) {
        var abIdx = _findArtboardIndexForBounds(vb);
        if(abIdx >= 0) {
          var existing = null;
          try {existing = artboardBoundsByIndex[abIdx];} catch(_eAbG0) {existing = null;}
          artboardBoundsByIndex[abIdx] = _boundsUnion(existing, vbAfter);
          _dbg('sel[' + i + '] mapped to artboard ' + abIdx + ' for expansion');
        }
      }
    }

    changed++;
    _dbg('sel[' + i + '] changed OK');
  }

  _dbg('END changed=' + changed + ', originalsMoved=' + originalsMoved);
  var expandedArtboards = 0;
  if(expandAb) {
    for(var abKey in artboardBoundsByIndex) {
      if(!artboardBoundsByIndex.hasOwnProperty(abKey)) continue;
      var idx = parseInt(abKey, 10);
      if(!(idx >= 0)) continue;
      try {
        var abRect = doc.artboards[idx].artboardRect;
        var u = _boundsUnion(abRect, artboardBoundsByIndex[abKey]);
        doc.artboards[idx].artboardRect = u;
        expandedArtboards++;
      } catch(_eAbSet0) { }
    }
  }
  if(preserveOriginal && originalLayer) {
    try {originalLayer.visible = false;} catch(_eOL1) { }
  }
  var dbgText = '';
  try {dbgText = _dbgLines.join('\n');} catch(_eDbgJoin) { dbgText = ''; }
  function _withDbg(msg) {
    if(!dbgText) return msg;
    return msg + '\n' + dbgText;
  }

  if(!changed) return _withDbg('No items updated (could not read bounds).');
  if(preserveOriginal) {
    var msgKeep = 'Applied bleed to ' + changed + ' item(s) on bleed layer. Originals moved: ' + originalsMoved + '.';
    if(expandAb) msgKeep += ' Artboards expanded: ' + expandedArtboards + '.';
    return _withDbg(msgKeep);
  }
  var msg = 'Applied bleed to ' + changed + ' item(s).';
  if(expandAb) msg += ' Artboards expanded: ' + expandedArtboards + '.';
  return _withDbg(msg);
}

/**
 * Adds a text element to the top of each artboard with the current file path.
 * Text box width: 60% of artboard width. Centered horizontally. Small offset from top.
 */

/**
 * Apply Offset Path bleed to selected items (paths/letters/shapes).
 * - Moves target items into layer "bleed"
 * - Optionally duplicates original into top layer "cutline" with no fill and 'Cut Contour' spot stroke
 * Args: JSON string {"offsetMm":number, "createCutline":boolean, "outlineText":boolean, "outlineStroke":boolean, "autoWeld":boolean}
 */
function signarama_helper_applyPathBleed(jsonStr) {
  var doc = app.activeDocument;
  if(!doc) {return "No document.";}
  if(!doc.selection || doc.selection.length === 0) {return "Select at least one item.";}

  var args = {};
  try {args = JSON.parse(String(jsonStr));} catch(e) {args = {};}
  var offsetMm = Number(args.offsetMm || 0);
  var createCutline = !!args.createCutline;
  var outlineText = (typeof args.outlineText === 'undefined') ? true : !!args.outlineText;
  var outlineStroke = (typeof args.outlineStroke === 'undefined') ? true : !!args.outlineStroke;
  var autoWeld = (typeof args.autoWeld === 'undefined') ? true : !!args.autoWeld;

  var offsetPt = _srh_mm2ptDoc(offsetMm);
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

  function _setCutlineStyleOnItem(it, opts) {
    opts = opts || {};
    var doOutlineText = (typeof opts.outlineText === 'undefined') ? true : !!opts.outlineText;
    var doOutlineStroke = (typeof opts.outlineStroke === 'undefined') ? false : !!opts.outlineStroke;
    try {
      if(it.typename === "PathItem") {
        if(doOutlineStroke && it.stroked) {
          try {
            var prevSel0 = doc.selection;
            doc.selection = null;
            it.selected = true;
            app.executeMenuCommand("OffsetPath v22");
            doc.selection = prevSel0;
          } catch(_eOlStroke0) { }
        }
        try {it.filled = false;} catch(_ePf0) { }
        try {it.stroked = true;} catch(_ePs0) { }
        try {it.strokeWidth = _srh_pxStrokeDoc(1);} catch(_eSw0) { }
        var sc = _getCutContourSpotColor();
        if(sc) {
          try {it.strokeColor = sc;} catch(_eSc0) { }
        } else {
          try {
            var c = new RGBColor();
            c.red = 255; c.green = 0; c.blue = 255;
            it.strokeColor = c;
          } catch(_eRgb0) { }
        }
      } else if(it.typename === "CompoundPathItem") {
        for(var i = 0; i < it.pathItems.length; i++) {_setCutlineStyleOnItem(it.pathItems[i], opts);}
      } else if(it.typename === "GroupItem") {
        var items = [];
        for(var j = 0; j < it.pageItems.length; j++) items.push(it.pageItems[j]);
        for(var k = 0; k < items.length; k++) _setCutlineStyleOnItem(items[k], opts);
      } else if(it.typename === "TextFrame") {
        if(doOutlineText) {
          try {
            var outlined = it.createOutline();
            try {it.remove();} catch(_eTrm0) { }
            _setCutlineStyleOnItem(outlined, opts);
          } catch(_eTxtOutline0) { }
        } else {
          try {it.filled = false;} catch(_eTf0) { }
          try {it.stroked = true;} catch(_eTs0) { }
          try {it.strokeWidth = _srh_pxStrokeDoc(1);} catch(_eTw0) { }
          var tsc = _getCutContourSpotColor();
          if(tsc) {
            try {it.strokeColor = tsc;} catch(_eTsc0) { }
          }
        }
      } else {
        // Ignore unsupported
      }
    } catch(e) { }
  }

  function _runOnSelection(items, fn) {
    if(!items) return false;
    var list = items.length ? items : [items];
    if(!list.length) return false;
    var prevSel = null;
    try {prevSel = doc.selection;} catch(_eS0) {prevSel = null;}
    try {doc.selection = null;} catch(_eS1) { }
    try {
      for(var i = 0; i < list.length; i++) {
        try {list[i].selected = true;} catch(_eSelOne) { }
      }
      fn();
      return true;
    } catch(_eRun) {
      return false;
    } finally {
      try {doc.selection = null;} catch(_eS2) { }
      try {if(prevSel) doc.selection = prevSel;} catch(_eS3) { }
    }
  }

  function _expandSelection() {
    try {app.executeMenuCommand('expandStyle');} catch(_eEx1) { }
    try {app.executeMenuCommand('expand');} catch(_eEx2) { }
  }

  function _outlineTextInContainer(container) {
    if(!container) return;
    try {
      var tfs = container.textFrames;
      for(var i = tfs.length - 1; i >= 0; i--) {
        try {tfs[i].createOutline();} catch(_eTxt0) { }
      }
    } catch(_eTxt1) { }
  }

  function _outlineStrokeInContainer(container) {
    if(!container) return;
    _runOnSelection(container, function() {
      try {app.executeMenuCommand("OffsetPath v22");} catch(_eOs0) { }
      _expandSelection();
    });
  }

  function _applyOffsetToContainer(container, ofst) {
    if(!container) return false;
    var fx1 = '<LiveEffect name="Adobe Offset Path"><Dict data="R ofst ' + Number(ofst) + ' I jntp 0 R mlim 180"/></LiveEffect>';
    var fx2 = '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 180 R ofst ' + Number(ofst) + ' I jntp 0"/></LiveEffect>';
    var applied = false;
    try {
      container.applyEffect(fx1);
      applied = true;
    } catch(_eFxA) {
      try {
        container.applyEffect(fx2);
        applied = true;
      } catch(_eFxB) { }
    }
    if(!applied) {
      var kids = [];
      try {for(var i = 0; i < container.pageItems.length; i++) kids.push(container.pageItems[i]);} catch(_eKid0) { }
      for(var k = 0; k < kids.length; k++) {
        var it = kids[k];
        try {
          it.applyEffect(fx1);
          applied = true;
        } catch(_eFxK1) {
          try {
            it.applyEffect(fx2);
            applied = true;
          } catch(_eFxK2) { }
        }
      }
    }
    if(applied) {
      _runOnSelection(container, function() {_expandSelection();});
    }
    return applied;
  }

  function _countPathLikeItems(container) {
    if(!container) return 0;
    var count = 0;
    try {count += (container.pathItems ? container.pathItems.length : 0);} catch(_eC0) { }
    try {
      var compounds = container.compoundPathItems;
      if(compounds && compounds.length) {
        for(var i = 0; i < compounds.length; i++) {
          try {count += compounds[i].pathItems.length;} catch(_eC1) {count++;}
        }
      }
    } catch(_eC2) { }
    return count;
  }

  function _uniteContainer(container) {
    if(!container) return false;
    var ok = _runOnSelection(container, function() {
      try {app.executeMenuCommand('Live Pathfinder Add');} catch(_eU0) { }
      _expandSelection();
    });
    return !!ok;
  }

  var bleedLayer = _srh_getOrCreateLayer(doc, "bleed");
  var originalLayer = _srh_getOrCreateLayer(doc, "original");
  var cutLayer = createCutline ? _srh_getOrCreateLayer(doc, "cutline") : _srh_getLayerByName(doc, "cutline");
  _srh_setBleedLayerOrder(cutLayer, originalLayer, bleedLayer);
  try {originalLayer.visible = true; originalLayer.locked = false;} catch(_eOlShow0) { }

  // Snapshot selection because we'll move items to layers.
  var sel = [];
  for(var s = 0; s < doc.selection.length; s++) {sel.push(doc.selection[s]);}

  var bleedGroup = null;
  var originalGroup = null;
  var cutGroup = null;
  try {
    bleedGroup = bleedLayer.groupItems.add();
    bleedGroup.name = 'SRH_BLEED_WORK';
  } catch(_eG0) { }
  try {
    originalGroup = originalLayer.groupItems.add();
    originalGroup.name = 'SRH_ORIGINAL_SNAPSHOT';
  } catch(_eG1) {originalGroup = originalLayer;}
  if(createCutline && cutLayer) {
    try {
      cutGroup = cutLayer.groupItems.add();
      cutGroup.name = 'SRH_CUTLINE_WORK';
    } catch(_eG2) {cutGroup = cutLayer;}
  }

  var originalCount = 0;
  var cutCount = 0;
  var movedCount = 0;
  var originalItems = [];

  // First operation: move selected source artwork to the original layer/group.
  for(var n0 = 0; n0 < sel.length; n0++) {
    var src = sel[n0];
    if(!src || src.locked || src.hidden) continue;
    try {
      if(originalGroup) {
        src.move(originalGroup, ElementPlacement.PLACEATEND);
      } else {
        src.move(originalLayer, ElementPlacement.PLACEATBEGINNING);
      }
      originalCount++;
      originalItems.push(src);
    } catch(_eMoveOrig0) { }
  }

  if(!originalCount) {
    try {doc.selection = null;} catch(_eSelOrig0) { }
    try {if(originalLayer) originalLayer.visible = false;} catch(_eOlHide0) { }
    return "No eligible items to move into original layer.";
  }

  // Second pass: duplicate from originals into cutline/bleed work layers.
  for(var n = 0; n < originalItems.length; n++) {
    var item = originalItems[n];
    if(!item) continue;

    if(createCutline && cutGroup) {
      try {item.duplicate(cutGroup, ElementPlacement.PLACEATEND); cutCount++;} catch(_eDupCut) { }
    }

    if(bleedGroup) {
      try {item.duplicate(bleedGroup, ElementPlacement.PLACEATEND); movedCount++;} catch(_eDupBleed0) { }
    } else {
      try {item.duplicate(bleedLayer, ElementPlacement.PLACEATBEGINNING); movedCount++;} catch(_eDupBleed1) { }
    }
  }

  if(!movedCount) {
    try {doc.selection = null;} catch(_eSel0) { }
    try {if(originalLayer) originalLayer.visible = false;} catch(_eOlHide1) { }
    return "No eligible items to offset.";
  }

  // Hard isolate processing from originals: never restore commands onto original selection.
  try {doc.selection = null;} catch(_eSelIso0) { }

  if(outlineText) _outlineTextInContainer(bleedGroup || bleedLayer);
  if(outlineStroke) _outlineStrokeInContainer(bleedGroup || bleedLayer);
  var offsetApplied = _applyOffsetToContainer(bleedGroup || bleedLayer, offsetPt);

  if(createCutline && cutGroup) {
    if(outlineText) _outlineTextInContainer(cutGroup);
    if(outlineStroke) _outlineStrokeInContainer(cutGroup);
    _setCutlineStyleOnItem(cutGroup, {outlineText: outlineText, outlineStroke: false});
    if(autoWeld) {
      _uniteContainer(cutGroup);
      // Re-apply style to welded result in cutline container only (never selection-based).
      _setCutlineStyleOnItem(cutGroup, {outlineText: false, outlineStroke: false});
    }
  }

  var offsetCount = _countPathLikeItems(bleedGroup || bleedLayer);
  try {doc.selection = null;} catch(_eSel1) { }

  var msg = "Path bleed applied (grouped). Originals: " + originalCount + ", targets: " + movedCount + ", bleed paths: " + offsetCount;
  if(createCutline) msg += ", cutlines created: " + cutCount + (autoWeld ? " (auto-weld on)" : " (auto-weld off)");
  if(!offsetApplied) msg += ". Note: Offset effect fallback was limited on this selection.";
  try {if(originalLayer) originalLayer.visible = false;} catch(_eOlHide2) { }
  return msg;
}

function _srh_transform_getBounds(item, includeStroke) {
  if(!item) return null;
  var b = null;
  if(includeStroke) {
    try {b = item.visibleBounds;} catch(_eTb0) { }
    if(!b || b.length !== 4) {try {b = item.geometricBounds;} catch(_eTb1) {b = null;} }
  } else {
    try {b = item.geometricBounds;} catch(_eTb2) { }
    if(!b || b.length !== 4) {try {b = item.visibleBounds;} catch(_eTb3) {b = null;} }
  }
  return (b && b.length === 4) ? b : null;
}
function _srh_transform_getClippingMaskItem(groupItem) {
  if(!groupItem) return null;
  try {
    if(!(groupItem.typename === "GroupItem" && groupItem.clipped && groupItem.pageItems && groupItem.pageItems.length)) return null;
  } catch(_eCk0) { return null; }
  try {
    for(var i = 0; i < groupItem.pageItems.length; i++) {
      var pi = groupItem.pageItems[i];
      if(!pi) continue;
      try {
        if(pi.typename === "PathItem" && pi.clipping) return pi;
      } catch(_eCk1) { }
      try {
        if(pi.typename === "CompoundPathItem" && pi.pathItems && pi.pathItems.length && pi.pathItems[0].clipping) return pi;
      } catch(_eCk2) { }
    }
  } catch(_eCk3) { }
  return null;
}
function _srh_transform_getTargetBounds(item, includeStroke) {
  // For clipped groups, size should be derived from the clipping mask only.
  try {
    if(item && item.typename === "GroupItem" && item.clipped) {
      var mask = _srh_transform_getClippingMaskItem(item);
      if(mask) {
        var mb = _srh_transform_getBounds(mask, includeStroke);
        if(mb) return mb;
      }
    }
  } catch(_eTbM0) { }
  return _srh_transform_getBounds(item, includeStroke);
}
function _srh_transform_originToTransformation(origin) {
  switch(String(origin || 'C')) {
    case 'TL': return Transformation.TOPLEFT;
    case 'TM': return Transformation.TOP;
    case 'TR': return Transformation.TOPRIGHT;
    case 'LM': return Transformation.LEFT;
    case 'RM': return Transformation.RIGHT;
    case 'BL': return Transformation.BOTTOMLEFT;
    case 'BM': return Transformation.BOTTOM;
    case 'BR': return Transformation.BOTTOMRIGHT;
    case 'C':
    default: return Transformation.CENTER;
  }
}
function _srh_transform_originFractions(origin) {
  switch(String(origin || 'C')) {
    case 'TL': return {fx: 0, fy: 1};
    case 'TM': return {fx: 0.5, fy: 1};
    case 'TR': return {fx: 1, fy: 1};
    case 'LM': return {fx: 0, fy: 0.5};
    case 'RM': return {fx: 1, fy: 0.5};
    case 'BL': return {fx: 0, fy: 0};
    case 'BM': return {fx: 0.5, fy: 0};
    case 'BR': return {fx: 1, fy: 0};
    case 'C':
    default: return {fx: 0.5, fy: 0.5};
  }
}
function _srh_transform_makeSize_impl(json) {
  if(!app.documents.length) return 'No open document.';
  var doc = app.activeDocument;
  var opts = {};
  try {
    if(typeof json === 'string') {
      try {opts = JSON.parse(json);}
      catch(_eParse0) {opts = eval('(' + json + ')');}
    } else {
      opts = json || {};
    }
  } catch(_eParse1) {
    opts = {};
  }

  var mode = String(opts.mode || 'selection');
  var requestedArtboardIndices = [];
  try {
    if(opts.artboardIndices && opts.artboardIndices.length) {
      for(var ari = 0; ari < opts.artboardIndices.length; ari++) {
        var aiv = Number(opts.artboardIndices[ari]);
        if(!(aiv >= 0)) continue;
        var ai = Math.round(aiv);
        if(ai < 0 || ai >= doc.artboards.length) continue;
        var existsAi = false;
        for(var aie = 0; aie < requestedArtboardIndices.length; aie++) {
          if(requestedArtboardIndices[aie] === ai) {existsAi = true; break;}
        }
        if(!existsAi) requestedArtboardIndices.push(ai);
      }
    }
  } catch(_eReqAb0) { requestedArtboardIndices = []; }
  var origin = String(opts.origin || 'C');
  var excludeStroke = (opts.excludeStroke === undefined) ? true : !!opts.excludeStroke;
  var includeStroke = !excludeStroke;
  var widthSpec = (typeof opts.widthSpec !== 'undefined') ? opts.widthSpec : opts.widthMm;
  var heightSpec = (typeof opts.heightSpec !== 'undefined') ? opts.heightSpec : opts.heightMm;
  var hasWidthSpec = !(widthSpec == null || String(widthSpec).replace(/^\s+|\s+$/g, '') === '');
  var hasHeightSpec = !(heightSpec == null || String(heightSpec).replace(/^\s+|\s+$/g, '') === '');
  var sf = _srh_getScaleFactor();
  function _ptDocToMmScaled(ptDoc) {
    return _dim_pt2mm(ptDoc * (sf || 1));
  }
  function _resolveSizeSpecMm(spec, currentMm) {
    if(spec == null) return currentMm;
    var raw = String(spec).replace(/^\s+|\s+$/g, '');
    if(!raw) return currentMm;
    var op = raw.charAt(0);
    var hasOp = (op === '+' || op === '-' || op === '*' || op === '/');
    var out = null;
    if(hasOp) {
      var nRaw = String(raw.substring(1)).replace(/^\s+|\s+$/g, '');
      var n = Number(nRaw);
      if(!(n >= 0)) return null;
      if(op === '+') out = currentMm + n;
      else if(op === '-') out = currentMm - n;
      else if(op === '*') out = currentMm * n;
      else if(op === '/') out = (n === 0) ? null : (currentMm / n);
    } else {
      out = Number(raw);
    }
    if(!(out > 0)) return null;
    return out;
  }
  if(!hasWidthSpec && !hasHeightSpec) return 'No size provided.';

  if(mode === 'artboard' || mode === 'artboards') {
    function _rectsEqual(a, b) {
      if(!a || !b || a.length !== 4 || b.length !== 4) return false;
      var eps = 1.0;
      return (Math.abs(a[0] - b[0]) <= eps &&
              Math.abs(a[1] - b[1]) <= eps &&
              Math.abs(a[2] - b[2]) <= eps &&
              Math.abs(a[3] - b[3]) <= eps);
    }
    function _findArtboardIndexByRect(rect) {
      if(!rect || rect.length !== 4) return -1;
      for(var ai = 0; ai < doc.artboards.length; ai++) {
        var ar = null;
        try {ar = doc.artboards[ai].artboardRect;} catch(_eAbIdx0) { ar = null; }
        if(_rectsEqual(rect, ar)) return ai;
      }
      return -1;
    }
    function _rectFromUnknown(obj) {
      if(!obj) return null;
      var r = null;
      try {r = obj.artboardRect;} catch(_eRu0) { r = null; }
      if(r && r.length === 4) return r;
      try {r = obj.visibleBounds;} catch(_eRu1) { r = null; }
      if(r && r.length === 4) return r;
      try {r = obj.geometricBounds;} catch(_eRu2) { r = null; }
      if(r && r.length === 4) return r;
      try {
        if(typeof obj.left !== 'undefined' && typeof obj.top !== 'undefined' &&
           typeof obj.right !== 'undefined' && typeof obj.bottom !== 'undefined') {
          return [Number(obj.left), Number(obj.top), Number(obj.right), Number(obj.bottom)];
        }
      } catch(_eRu3) { }
      return null;
    }
    function _pushUniqueIndex(arr, idx) {
      if(!(idx >= 0)) return;
      for(var iU = 0; iU < arr.length; iU++) {
        if(arr[iU] === idx) return;
      }
      arr.push(idx);
    }
    function _getSelectedArtboardIndices() {
      var indices = [];
      // Attempt 1: direct artboard selected flags (if available in this host/version).
      for(var iA = 0; iA < doc.artboards.length; iA++) {
        try {
          if(doc.artboards[iA].selected) _pushUniqueIndex(indices, iA);
        } catch(_eAbSel0) { }
      }
      // Attempt 2: selection may contain artboard-like objects in some host versions.
      var selAny = null;
      try {selAny = doc.selection;} catch(_eAbSel1) { selAny = null; }
      if(selAny && selAny.length) {
        for(var sA = 0; sA < selAny.length; sA++) {
          var si = selAny[sA];
          if(!si) continue;
          var selIdx = -1;
          try {selIdx = Number(si.artboardIndex);} catch(_eAbSelIdx0) { selIdx = -1; }
          if(!(selIdx >= 0)) {
            try {selIdx = Number(si.index);} catch(_eAbSelIdx1) { selIdx = -1; }
          }
          if(selIdx >= 0 && selIdx < doc.artboards.length) {
            _pushUniqueIndex(indices, selIdx);
            continue;
          }
          var sRect = _rectFromUnknown(si);
          var matchIdx = _findArtboardIndexByRect(sRect);
          if(matchIdx >= 0) _pushUniqueIndex(indices, matchIdx);
        }
      }
      return indices;
    }

    var targetIndices = [];
    if(mode === 'artboard') {
      _pushUniqueIndex(targetIndices, doc.artboards.getActiveArtboardIndex());
    } else {
      if(requestedArtboardIndices.length) {
        targetIndices = requestedArtboardIndices.slice(0);
      } else {
        targetIndices = _getSelectedArtboardIndices();
        if(!targetIndices.length) {
          _pushUniqueIndex(targetIndices, doc.artboards.getActiveArtboardIndex());
        }
      }
    }
    if(!targetIndices.length) return 'No artboards available to transform.';

    var changedAb = 0;
    var lastW = null;
    var lastH = null;
    for(var ti = 0; ti < targetIndices.length; ti++) {
      var idx = targetIndices[ti];
      var rect = doc.artboards[idx].artboardRect; // [L,T,R,B]
      var left = rect[0], top = rect[1], right = rect[2], bottom = rect[3];
      var curW = right - left;
      var curH = top - bottom;
      if(!(curW > 0) || !(curH > 0)) continue;

      var newW = curW;
      var newH = curH;
      if(hasWidthSpec) {
        var curWmm = _ptDocToMmScaled(curW);
        var resolvedWmm = _resolveSizeSpecMm(widthSpec, curWmm);
        if(resolvedWmm == null) return 'Invalid width expression.';
        newW = _srh_mm2ptDoc(resolvedWmm);
      }
      if(hasHeightSpec) {
        var curHmm = _ptDocToMmScaled(curH);
        var resolvedHmm = _resolveSizeSpecMm(heightSpec, curHmm);
        if(resolvedHmm == null) return 'Invalid height expression.';
        newH = _srh_mm2ptDoc(resolvedHmm);
      }
      var o = _srh_transform_originFractions(origin);
      var ax = left + (curW * o.fx);
      var ay = bottom + (curH * o.fy);
      var newLeft = ax - (newW * o.fx);
      var newBottom = ay - (newH * o.fy);
      var newRect = [newLeft, newBottom + newH, newLeft + newW, newBottom];
      doc.artboards[idx].artboardRect = newRect;
      changedAb++;
      lastW = newW;
      lastH = newH;
    }
    if(!changedAb) return (mode === 'artboard') ? 'Active artboard has invalid size.' : 'No selected artboards could be transformed.';
    if(mode === 'artboard') {
      return 'Transformed active artboard to ' + Math.round(_dim_pt2mm(lastW) * 100) / 100 + ' x ' + Math.round(_dim_pt2mm(lastH) * 100) / 100 + ' mm.';
    }
    return 'Transformed ' + changedAb + ' artboard' + (changedAb === 1 ? '' : 's') + '.';
  }

  var sel = doc.selection;
  if(!sel || !sel.length) return 'No selection. Select one or more items.';
  var xformOrigin = _srh_transform_originToTransformation(origin);
  function _getNearestClippedGroup(item) {
    var p = item;
    while(p) {
      var isClippedGroup = false;
      try {isClippedGroup = (p.typename === "GroupItem" && p.clipped);} catch(_eCg0) { isClippedGroup = false; }
      if(isClippedGroup) return p;
      try {p = p.parent;} catch(_eCg1) { p = null; }
    }
    return null;
  }
  var targets = [];
  for(var s = 0; s < sel.length; s++) {
    var raw = sel[s];
    if(!raw) continue;
    var target = _getNearestClippedGroup(raw) || raw;
    var already = false;
    for(var t = 0; t < targets.length; t++) {
      if(targets[t] === target) {already = true; break;}
    }
    if(!already) targets.push(target);
  }
  var changed = 0;
  for(var i = 0; i < targets.length; i++) {
    var it = targets[i];
    if(!it) continue;
    try {if(it.locked || it.hidden) continue;} catch(_eLock) { }
    var b = _srh_transform_getTargetBounds(it, includeStroke);
    if(!b) continue;
    var curItemW = b[2] - b[0];
    var curItemH = b[1] - b[3];
    if(!(curItemW > 0) || !(curItemH > 0)) continue;
    var sx = 100;
    var sy = 100;
    if(hasWidthSpec) {
      var curItemWmm = _ptDocToMmScaled(curItemW);
      var resolvedItemWmm = _resolveSizeSpecMm(widthSpec, curItemWmm);
      if(resolvedItemWmm == null) continue;
      var targetWpt = _srh_mm2ptDoc(resolvedItemWmm);
      sx = (targetWpt / curItemW) * 100;
    }
    if(hasHeightSpec) {
      var curItemHmm = _ptDocToMmScaled(curItemH);
      var resolvedItemHmm = _resolveSizeSpecMm(heightSpec, curItemHmm);
      if(resolvedItemHmm == null) continue;
      var targetHpt = _srh_mm2ptDoc(resolvedItemHmm);
      sy = (targetHpt / curItemH) * 100;
    }
    try {
      it.resize(sx, sy, true, true, true, true, 100, xformOrigin);
      changed++;
    } catch(_eResize0) { }
  }
  if(!changed) return 'No eligible selection items to transform.';
  return 'Transformed ' + changed + ' selection item' + (changed === 1 ? '' : 's') + '.';
}

function signarama_helper_transform_listArtboards() {
  if(!app.documents.length) return '[]';
  var doc = app.activeDocument;
  var out = [];
  function _normalizeRect4(r) {
    if(!r) return null;
    try {
      if(typeof r.length !== 'undefined' && r.length >= 4) {
        return [Number(r[0]), Number(r[1]), Number(r[2]), Number(r[3])];
      }
    } catch(_eNr0) { }
    try {
      if(typeof r.left !== 'undefined' && typeof r.top !== 'undefined' &&
         typeof r.right !== 'undefined' && typeof r.bottom !== 'undefined') {
        return [Number(r.left), Number(r.top), Number(r.right), Number(r.bottom)];
      }
    } catch(_eNr1) { }
    return null;
  }
  var activeIdx = 0;
  try {activeIdx = doc.artboards.getActiveArtboardIndex();} catch(_eAbL0) { activeIdx = 0; }
  for(var i = 0; i < doc.artboards.length; i++) {
    var ab = doc.artboards[i];
    var r = null;
    try {r = _normalizeRect4(ab.artboardRect);} catch(_eAbL1) { r = null; }
    if(!r) continue;
    var w = r[2] - r[0];
    var h = r[1] - r[3];
    var name = '';
    try {name = String(ab.name || '');} catch(_eAbL2) { name = ''; }
    out.push({
      index: i,
      name: name,
      widthMm: Math.round(_dim_pt2mm(w) * 100) / 100,
      heightMm: Math.round(_dim_pt2mm(h) * 100) / 100,
      isActive: i === activeIdx
    });
  }
  return JSON.stringify(out);
}

function signarama_helper_transform_debugArtboards() {
  var debug = {
    hasDocument: false,
    artboardCount: 0,
    activeIndex: -1,
    items: []
  };
  if(!app.documents.length) return JSON.stringify(debug);
  var doc = app.activeDocument;
  debug.hasDocument = true;
  try {debug.artboardCount = doc.artboards.length;} catch(_eDbg0) { debug.artboardCount = -1; }
  try {debug.activeIndex = doc.artboards.getActiveArtboardIndex();} catch(_eDbg1) { debug.activeIndex = -1; }

  for(var i = 0; i < doc.artboards.length; i++) {
    var row = {index: i, name: '', rectType: 'none', rectRaw: '', rectParsed: null, selectedFlag: null};
    var ab = doc.artboards[i];
    try {row.name = String(ab.name || '');} catch(_eDbg2) { row.name = ''; }
    try {row.selectedFlag = !!ab.selected;} catch(_eDbg3) { row.selectedFlag = null; }
    try {
      var r = ab.artboardRect;
      row.rectType = Object.prototype.toString.call(r);
      try {row.rectRaw = String(r);} catch(_eDbg4) { row.rectRaw = '[unprintable]'; }
      try {
        if(r && typeof r.length !== 'undefined') {
          row.rectParsed = [Number(r[0]), Number(r[1]), Number(r[2]), Number(r[3])];
        } else if(r && typeof r.left !== 'undefined') {
          row.rectParsed = [Number(r.left), Number(r.top), Number(r.right), Number(r.bottom)];
        }
      } catch(_eDbg5) { row.rectParsed = null; }
    } catch(_eDbg6) { }
    debug.items.push(row);
  }
  return JSON.stringify(debug);
}

function signarama_helper_corebridge_createLink() {
  if(!app.documents.length) return 'No open document.';
  var doc = app.activeDocument;
  var view = null;
  try {view = doc.activeView;} catch(_eCbV0) { view = null; }
  if(!view) return 'No active view.';
  var center = null;
  try {center = view.centerPoint;} catch(_eCbV1) { center = null; }
  if(!center || center.length < 2) return 'Could not determine viewport center.';

  var cx = Number(center[0]);
  var cy = Number(center[1]);
  var w = _srh_mm2ptDoc(70);
  var h = _srh_mm2ptDoc(20);
  var r = _srh_mm2ptDoc(4);
  var left = cx - (w / 2);
  var top = cy + (h / 2);

  var grp = doc.activeLayer.groupItems.add();
  grp.name = 'SRH_COREBRIDGE_LINK';

  var rect = grp.pathItems.roundedRectangle(top, left, w, h, r, r);
  rect.name = 'SRH_COREBRIDGE_LINK_BG';
  try {rect.stroked = true;} catch(_eCbR0) { }
  try {rect.strokeWidth = _srh_pxStrokeDoc(1);} catch(_eCbR1) { }
  try {rect.filled = true;} catch(_eCbR2) { }
  try {
    var fill = new RGBColor();
    fill.red = 6; fill.green = 82; fill.blue = 152;
    rect.fillColor = fill;
  } catch(_eCbR3) { }
  try {
    var stroke = new RGBColor();
    stroke.red = 255; stroke.green = 255; stroke.blue = 255;
    rect.strokeColor = stroke;
  } catch(_eCbR4) { }

  var tf = grp.textFrames.pointText([cx, cy - _srh_mm2ptDoc(1)]);
  tf.name = 'SRH_COREBRIDGE_LINK_TEXT';
  tf.contents = 'corebridge';
  try {tf.textRange.paragraphAttributes.justification = Justification.CENTER;} catch(_eCbT0) { }
  try {tf.textRange.characterAttributes.size = _srh_ptDoc(14);} catch(_eCbT1) { }
  try {
    var txt = new RGBColor();
    txt.red = 255; txt.green = 255; txt.blue = 255;
    tf.textRange.characterAttributes.fillColor = txt;
  } catch(_eCbT2) { }

  // Prevent immediate click-detection trigger after creation.
  try {grp.selected = false;} catch(_eCbSel0) { }
  try {rect.selected = false;} catch(_eCbSel1) { }
  try {tf.selected = false;} catch(_eCbSel2) { }
  try {doc.selection = null;} catch(_eCbSel3) { }
  try {app.executeMenuCommand('deselectall');} catch(_eCbSel4) { }

  return 'Corebridge link created.';
}

function signarama_helper_corebridge_isLinkSelected() {
  if(!app.documents.length) return '0';
  var doc = app.activeDocument;
  var sel = null;
  try {sel = doc.selection;} catch(_eCbS0) { sel = null; }
  if(!sel || !sel.length) return '0';

  function _hasCorebridgeAncestor(item) {
    var p = item;
    while(p) {
      try {
        if(String(p.name || '') === 'SRH_COREBRIDGE_LINK') return true;
      } catch(_eCbS1) { }
      try {p = p.parent;} catch(_eCbS2) { p = null; }
    }
    return false;
  }

  for(var i = 0; i < sel.length; i++) {
    if(_hasCorebridgeAncestor(sel[i])) return '1';
  }
  return '0';
}
this.signarama_helper_transform_makeSize = function(json) {
  return _srh_transform_makeSize_impl(json);
};
this.atlas_transform_makeSize = function(json) {
  return _srh_transform_makeSize_impl(json);
};
try { if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_transform_makeSize = this.signarama_helper_transform_makeSize; } catch(_eTg0) { }
try { if(typeof $ !== 'undefined' && $.global) $.global.atlas_transform_makeSize = this.atlas_transform_makeSize; } catch(_eTg2) { }
try { signarama_helper_transform_makeSize = this.signarama_helper_transform_makeSize; } catch(_eTg1) { }
try { if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_transform_listArtboards = this.signarama_helper_transform_listArtboards; } catch(_eTg3) { }
try { signarama_helper_transform_listArtboards = this.signarama_helper_transform_listArtboards; } catch(_eTg4) { }
try { if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_transform_debugArtboards = this.signarama_helper_transform_debugArtboards; } catch(_eTg5) { }
try { signarama_helper_transform_debugArtboards = this.signarama_helper_transform_debugArtboards; } catch(_eTg6) { }
try { if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_corebridge_createLink = this.signarama_helper_corebridge_createLink; } catch(_eTg7) { }
try { signarama_helper_corebridge_createLink = this.signarama_helper_corebridge_createLink; } catch(_eTg8) { }
try { if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_corebridge_isLinkSelected = this.signarama_helper_corebridge_isLinkSelected; } catch(_eTg9) { }
try { signarama_helper_corebridge_isLinkSelected = this.signarama_helper_corebridge_isLinkSelected; } catch(_eTg10) { }



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
    var boxH = _srh_mm2ptDoc(24); // ~24mm tall
    var offset = _srh_mm2ptDoc(5); // 5mm from top

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
      var desiredTop = top - _srh_mm2ptDoc(10);
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
function _dim_mm2ptDoc(mm, scaleFactor) {
  var sf = Number(scaleFactor);
  if(!(sf > 0)) sf = _srh_getScaleFactor();
  return _dim_mm2pt(mm) / sf;
}
function _dim_ptDoc(pt, scaleFactor) {
  var sf = Number(scaleFactor);
  if(!(sf > 0)) sf = _srh_getScaleFactor();
  return (pt || 0) / sf;
}
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
  var strokeWidth = (typeof strokePt === 'number' && isFinite(strokePt)) ? strokePt : 1;
  p.strokeWidth = Math.max(strokeWidth, 0.1);
  var lc = null;
  try {
    if(opts && opts.lineColor) {
      if(typeof opts.lineColor === 'object' && typeof opts.lineColor.red !== 'undefined') {
        lc = opts.lineColor;
      } else if(typeof opts.lineColor === 'string') {
        lc = _dim_parseHexColorToRGBColor(opts.lineColor);
      }
    }
  } catch(_eLc) {lc = null;}
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
function _dim_getBoundsArray(item, includeStroke) {
  var b = null;
  if(includeStroke) {
    try {b = item.visibleBounds;} catch(_eV0) { }
    if(!b || b.length !== 4) {try {b = item.geometricBounds;} catch(_eG0) {b = null;} }
  } else {
    try {b = item.geometricBounds;} catch(_eG1) { }
    if(!b || b.length !== 4) {try {b = item.visibleBounds;} catch(_eV1) {b = null;} }
  }
  return (b && b.length === 4) ? b : null;
}
function _dim_getMetricsFor(item, measureClippedContent, includeStroke) {
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

          var cb = _dim_getBoundsArray(child, includeStroke);
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
        try {return _dim_getMetricsFor(maskItem, false, includeStroke);} catch(_eMaskM) { }
        var mb = _dim_getBoundsArray(maskItem, includeStroke);
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

  var vb = _dim_getBoundsArray(item, includeStroke);
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

  if(typeof item.position === 'undefined' ||
    typeof item.width === 'undefined' ||
    typeof item.height === 'undefined') {
    throw new Error('Unable to read item bounds.');
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
  var hasSelection = !!(originalSel && originalSel.length);

  var scaleFactor = _srh_getScaleFactor();
  var offsetPt = _dim_mm2ptDoc(opts.offsetMm || 10, scaleFactor);
  var ticLenPt = _dim_mm2ptDoc(opts.ticLenMm || 2, scaleFactor);
  var textPt = _dim_ptDoc(opts.textPt || 10, scaleFactor);
  var strokePt = _dim_ptDoc(opts.strokePt || 1, scaleFactor);
  var decimals = (opts.decimals | 0);
  var textOffsetPt = _dim_mm2ptDoc(opts.labelGapMm || 0, scaleFactor);
  var measureClippedContent = !!opts.measureClippedContent;
  var measureIncludeStroke = !!opts.measureIncludeStroke;

  var lyr = _dim_ensureLayer('Dimensions');
  var textColor = opts.textColor;
  var lineColor = null;
  try {
    lineColor = _dim_hexToRGB(opts.lineColor);
    if(!lineColor) lineColor = _dim_parseHexColorToRGBColor(opts.lineColor);
    if(!lineColor && opts.lineColor && typeof opts.lineColor === 'object' && typeof opts.lineColor.red !== 'undefined') {
      lineColor = opts.lineColor;
    }
  } catch(_eLc2) {lineColor = null;}
  if(!lineColor) lineColor = _dim_hexToRGB('#000000');
  var includeArrowhead = !!opts.includeArrowhead;
  var arrowheadSizePt = _dim_ptDoc(opts.arrowheadSizePt || 0, scaleFactor);
  var scaleAppearance = opts.scaleAppearance || 1;

  var dOpts = {
    side: opts.side,
    offsetPt: offsetPt * scaleAppearance,
    ticLenPt: ticLenPt * scaleAppearance,
    textPt: textPt * scaleAppearance,
    strokePt: strokePt * scaleAppearance,
    decimals: decimals,
    scaleFactor: scaleFactor,
    textOffsetPt: textOffsetPt * scaleAppearance,
    textColor: textColor,
    lineColor: lineColor,
    includeArrowhead: includeArrowhead,
    arrowheadSizePt: arrowheadSizePt * scaleAppearance
  };

  var objectsProcessed = 0;
  var measuresAdded = 0;

  if(hasSelection) {
    try {
      for(var i = 0; i < originalSel.length; i++) {
        var item = originalSel[i];
        try {if(item.locked || item.hidden) continue;} catch(_) { }
        var b; try {b = _dim_getMetricsFor(item, measureClippedContent, measureIncludeStroke);} catch(e) {continue;}
        measuresAdded += _dim_drawForBounds(b, lyr, dOpts);
        objectsProcessed++;
      }
    } catch(e) {
      _dim_restoreSelection(originalSel);
      return 'Error: ' + e.message;
    }
  } else {
    try {
      var abCount = doc.artboards.length;
      for(var a = 0; a < abCount; a++) {
        var rect = doc.artboards[a].artboardRect; // [L,T,R,B]
        if(!rect || rect.length !== 4) continue;
        var abBounds = {left: rect[0], top: rect[1], right: rect[2], bottom: rect[3]};
        measuresAdded += _dim_drawForBounds(abBounds, lyr, dOpts);
        objectsProcessed++;
      }
    } catch(e2) {
      return 'Error: ' + e2.message;
    }
  }

  if(hasSelection) _dim_restoreSelection(originalSel);

  if(objectsProcessed === 0) {
    if(hasSelection) return 'No measurable objects in selection.';
    return 'No artboards found.';
  }
  var targetLabel = hasSelection ? 'object' : 'artboard';
  var msg = 'Added ' + measuresAdded + ' measure' + (measuresAdded === 1 ? '' : 's') +
    ' on ' + objectsProcessed + ' ' + targetLabel + (objectsProcessed === 1 ? '' : 's') + '.';
  return msg;
}

function _dim_runLine(opts) {
  var doc = app.activeDocument;
  if(!doc) return "No document open.";

  var originalSel = _dim_captureSelection();
  if(!originalSel || !originalSel.length) return "Nothing selected.";

  var scaleFactor = _srh_getScaleFactor();
  var textPt = _dim_ptDoc(opts.textPt || 10, scaleFactor);
  var strokePt = _dim_ptDoc(opts.strokePt || 1, scaleFactor);
  var textOffsetPt = _dim_mm2ptDoc(opts.labelGapMm || 0, scaleFactor);
  var lineOffsetPt = _dim_mm2ptDoc(opts.offsetMm || 0, scaleFactor);
  var ticLenPt = _dim_mm2ptDoc(opts.ticLenMm || 2, scaleFactor);
  var includeArrowhead = !!opts.includeArrowhead;
  var arrowheadSizePt = _dim_ptDoc(opts.arrowheadSizePt || 0, scaleFactor);
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
      var strokeScaled = strokePt * scaleAppearance;
      var tickScaled = ticLenPt * scaleAppearance;
      var arrowScaled = arrowheadSizePt * scaleAppearance;

      var g = lyr.groupItems.add();
      _dim_makeLine(g, sx, sy, ex, ey, strokeScaled, {lineColor: lineColor});
      var halfTick = tickScaled * 0.5;
      _dim_makeLine(g, sx - nx * halfTick, sy - ny * halfTick, sx + nx * halfTick, sy + ny * halfTick, strokeScaled, {lineColor: lineColor});
      _dim_makeLine(g, ex - nx * halfTick, ey - ny * halfTick, ex + nx * halfTick, ey + ny * halfTick, strokeScaled, {lineColor: lineColor});
      if(includeArrowhead) {
        _dim_addArrowheadAlongLine(g, sx, sy, -dx, -dy, arrowScaled, strokeScaled, lineColor);
        _dim_addArrowheadAlongLine(g, ex, ey, dx, dy, arrowScaled, strokeScaled, lineColor);
      }

      var midX = (sx + ex) * 0.5 + nx * textOff;
      var midY = (sy + ey) * 0.5 + ny * textOff;
      var angle = Math.atan2(dy, dx) * 180 / Math.PI;
      var label = _dim_fmtMmScaled(len, opts.decimals | 0, scaleFactor);

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

this.atlas_dimensions_hasSelection = function() {
  try {
    if(!app.documents.length) return '0';
    var sel = _dim_captureSelection();
    return (sel && sel.length) ? '1' : '0';
  } catch(_eSelCheck) {
    return '0';
  }
};

function _srh_panelSettingsFile() {
  var candidates = _srh_panelSettingsFileCandidates();
  return (candidates && candidates.length) ? candidates[0] : null;
}
function _srh_panelSettingsFileCandidates() {
  var out = [];
  var dir = null;
  try {
    var extRoot = new Folder($.fileName).parent.parent; // .../Signarama-Illustrator-Helper
    var localDir = new Folder(extRoot.fsName + '/.local');
    try {if(!localDir.exists) localDir.create();} catch(_eMk0) { }
    var localFile = new File(localDir.fsName + '/panel-settings.json');
    try {if(localFile) out.push(localFile);} catch(_eTryLocal) { }
  } catch(_eLocalRoot) { }

  // Fallback for restricted installs (e.g. Program Files without write permission).
  var root = null;
  try {root = Folder.userData;} catch(_eRoot) {root = null;}
  if(root) {
    dir = new Folder(root.fsName + '/Signarama-Illustrator-Helper');
    try {if(!dir.exists) dir.create();} catch(_eMk1) { }
    try {out.push(new File(dir.fsName + '/panel-settings.json'));} catch(_ePushU) { }
  }
  return out;
}

this.signarama_helper_panelSettingsSave = function(json) {
  var f = null;
  try {
    var txt = '';
    if(typeof json === 'string') txt = json;
    else txt = String(json || '{}');
    // Validate JSON payload before writing; preserve strict JSON text from panel.
    JSON.parse(txt);
    var out = txt;
    var files = _srh_panelSettingsFileCandidates();
    if(!files || !files.length) return 'Error: Could not resolve settings path.';
    var lastErr = '';
    for(var i = 0; i < files.length; i++) {
      f = files[i];
      if(!f) continue;
      try {
        f.encoding = 'UTF-8';
        if(!f.open('w')) {lastErr = 'Could not open settings file for write.'; continue;}
        f.write(out);
        f.close();
        return 'OK';
      } catch(_eWriteOne) {
        try {if(f && f.opened) f.close();} catch(_eCloseOne) { }
        lastErr = String(_eWriteOne);
      }
    }
    return 'Error: ' + (lastErr || 'Could not write settings file.');
  } catch(e) {
    try {if(f && f.opened) f.close();} catch(_eClose) { }
    return 'Error: ' + e.message;
  }
};
try { if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_panelSettingsSave = this.signarama_helper_panelSettingsSave; } catch(_ePss0) { }
try { signarama_helper_panelSettingsSave = this.signarama_helper_panelSettingsSave; } catch(_ePss1) { }

this.signarama_helper_panelSettingsLoad = function() {
  var f = null;
  try {
    var files = _srh_panelSettingsFileCandidates();
    if(!files || !files.length) return 'NO_SETTINGS';
    for(var i = 0; i < files.length; i++) {
      f = files[i];
      if(!f || !f.exists) continue;
      try {
        f.encoding = 'UTF-8';
        if(!f.open('r')) continue;
        var txt = f.read();
        f.close();
        if(txt) return String(txt);
      } catch(_eReadOne) {
        try {if(f && f.opened) f.close();} catch(_eCloseOne2) { }
      }
    }
    return 'NO_SETTINGS';
  } catch(e) {
    try {if(f && f.opened) f.close();} catch(_eClose2) { }
    return 'NO_SETTINGS';
  }
};
try { if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_panelSettingsLoad = this.signarama_helper_panelSettingsLoad; } catch(_ePsl0) { }
try { signarama_helper_panelSettingsLoad = this.signarama_helper_panelSettingsLoad; } catch(_ePsl1) { }

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

var _SRH_LIGHTBOX_OUTER_NAME = 'SRH_LIGHTBOX_OUTER';
var _SRH_LIGHTBOX_SUPPORT_NAME = 'SRH_LIGHTBOX_SUPPORT';
var _SRH_LIGHTBOX_MEASURE_NOTE = 'SRH_LIGHTBOX_MEASURE';
var _SRH_LIGHTBOX_MEASURE_NAME = 'SRH_LIGHTBOX_MEASURE_GROUP';
var _srh_lightboxLiveSignatures = {};
var _srh_lightboxLiveEnabled = {};

function _srh_nameStartsWith(value, prefix) {
  if(value == null || prefix == null) return false;
  value = String(value);
  prefix = String(prefix);
  return value.slice(0, prefix.length) === prefix;
}

function _srh_generateLightboxId() {
  var ts = (new Date()).getTime();
  var r = Math.floor(Math.random() * 1000000);
  return String(ts) + '_' + String(r);
}

function _srh_getLightboxTag(item) {
  if(!item) return null;
  var note = '';
  try {note = String(item.note || '');} catch(_eN) {note = '';}
  if(note) {
    var noteParts = note.split('|');
    if(noteParts.length >= 3 && noteParts[0] === 'SRH_LIGHTBOX') {
      return {id: noteParts[1], role: noteParts[2]};
    }
  }
  var nm = '';
  try {nm = String(item.name || '');} catch(_eNm) {nm = '';}
  if(!nm) return null;
  var m = /^SRH_LIGHTBOX_(OUTER|SUPPORT)_(.+)$/.exec(nm);
  if(m && m.length >= 3) {
    return {id: m[2], role: m[1].toLowerCase()};
  }
  return null;
}

function _srh_setLightboxItemTag(item, role, lightboxId) {
  if(!item) return;
  var id = String(lightboxId || '');
  if(!id) return;
  var roleNorm = String(role || '').toLowerCase();
  if(roleNorm !== 'outer' && roleNorm !== 'support') return;
  try {
    if(roleNorm === 'outer') item.name = _SRH_LIGHTBOX_OUTER_NAME + '_' + id;
    else item.name = _SRH_LIGHTBOX_SUPPORT_NAME + '_' + id;
  } catch(_eNm0) { }
  try {item.note = 'SRH_LIGHTBOX|' + id + '|' + roleNorm;} catch(_eNt0) { }
}

function _srh_getItemBounds(item) {
  if(!item) return null;
  var b = null;
  try {b = item.visibleBounds;} catch(_e0) { }
  if(!b || b.length !== 4) {
    try {b = item.geometricBounds;} catch(_e1) {b = null;}
  }
  if(!b || b.length !== 4) return null;
  return {left: b[0], top: b[1], right: b[2], bottom: b[3]};
}

function _srh_buildLightboxSignature(bounds, supportCenters) {
  if(!bounds) return '';
  function _r(n) {return Math.round(Number(n || 0) * 1000) / 1000;}
  var p = [_r(bounds.left), _r(bounds.top), _r(bounds.right), _r(bounds.bottom)];
  if(supportCenters && supportCenters.length) {
    for(var i = 0; i < supportCenters.length; i++) p.push(_r(supportCenters[i]));
  }
  return p.join('|');
}

function _srh_findAllLightboxData(doc) {
  if(!doc) return null;
  var frameLayer = null;
  try {frameLayer = doc.layers.getByName('lightbox frame');} catch(_e0) {return null;}
  if(!frameLayer) return null;

  var outerCandidates = [];
  var supportCandidates = [];
  var items = frameLayer.pageItems;
  for(var i = 0; i < items.length; i++) {
    var it = items[i];
    if(!it) continue;
    var nm = '';
    try {nm = String(it.name || '');} catch(_eNm) {nm = '';}
    var tag = _srh_getLightboxTag(it);
    if(tag && tag.role === 'outer') outerCandidates.push(it);
    else if(tag && tag.role === 'support') supportCandidates.push(it);
    else {
      if(_srh_nameStartsWith(nm, _SRH_LIGHTBOX_OUTER_NAME)) outerCandidates.push(it);
      if(_srh_nameStartsWith(nm, _SRH_LIGHTBOX_SUPPORT_NAME)) supportCandidates.push(it);
    }
  }
  if(!outerCandidates.length) return [];

  var lightboxes = [];
  var byId = {};
  var legacyIdx = 0;
  for(var o = 0; o < outerCandidates.length; o++) {
    var outer = outerCandidates[o];
    var ob = _srh_getItemBounds(outer);
    if(!ob) continue;
    var otag = _srh_getLightboxTag(outer);
    var lid = (otag && otag.id) ? String(otag.id) : ('legacy_' + (legacyIdx++));
    var lb = {
      id: lid,
      bounds: ob,
      supportCenters: [],
      _centerX: (ob.left + ob.right) * 0.5
    };
    lightboxes.push(lb);
    byId[lid] = lb;
  }

  for(var s = 0; s < supportCandidates.length; s++) {
    var support = supportCandidates[s];
    var sb = _srh_getItemBounds(support);
    if(!sb) continue;
    var cx = (sb.left + sb.right) * 0.5;

    var stag = _srh_getLightboxTag(support);
    var target = null;
    if(stag && stag.id && byId[String(stag.id)]) {
      target = byId[String(stag.id)];
    } else {
      // Fallback: closest outer that spans this x position.
      var bestDist = 1e20;
      for(var j = 0; j < lightboxes.length; j++) {
        var cand = lightboxes[j];
        if(cx < cand.bounds.left - 1 || cx > cand.bounds.right + 1) continue;
        var d = Math.abs(cx - cand._centerX);
        if(d < bestDist) {
          bestDist = d;
          target = cand;
        }
      }
    }
    if(target) target.supportCenters.push(cx);
  }

  for(var k = 0; k < lightboxes.length; k++) {
    lightboxes[k].supportCenters.sort(function(a, b) {return a - b;});
    lightboxes[k].signature = _srh_buildLightboxSignature(lightboxes[k].bounds, lightboxes[k].supportCenters);
  }
  return lightboxes;
}

function _srh_tagLightboxMeasureGroup(groupItem, lightboxId) {
  if(!groupItem) return;
  var id = String(lightboxId || '');
  if(id) {
    try {groupItem.note = _SRH_LIGHTBOX_MEASURE_NOTE + '|' + id;} catch(_e0) { }
    try {groupItem.name = _SRH_LIGHTBOX_MEASURE_NAME + '_' + id;} catch(_e1) { }
  } else {
    try {groupItem.note = _SRH_LIGHTBOX_MEASURE_NOTE;} catch(_e2) { }
    try {groupItem.name = _SRH_LIGHTBOX_MEASURE_NAME;} catch(_e3) { }
  }
}

function _srh_clearTaggedLightboxMeasures(doc, lightboxId) {
  if(!doc) return 0;
  var lyr = null;
  try {lyr = doc.layers.getByName('Dimensions');} catch(_e0) {return 0;}
  if(!lyr) return 0;

  var id = String(lightboxId || '');
  var removed = 0;
  for(var i = lyr.groupItems.length - 1; i >= 0; i--) {
    var g = lyr.groupItems[i];
    if(!g) continue;
    var isTagged = false;
    var note = '';
    var name = '';
    try {note = String(g.note || '');} catch(_eN0) {note = '';}
    try {name = String(g.name || '');} catch(_eN1) {name = '';}
    if(id) {
      if(note === (_SRH_LIGHTBOX_MEASURE_NOTE + '|' + id)) isTagged = true;
      if(!isTagged && name === (_SRH_LIGHTBOX_MEASURE_NAME + '_' + id)) isTagged = true;
    } else {
      if(note === _SRH_LIGHTBOX_MEASURE_NOTE) isTagged = true;
      if(!isTagged && _srh_nameStartsWith(name, _SRH_LIGHTBOX_MEASURE_NAME)) isTagged = true;
    }
    if(!isTagged) {
      try {
        if(!id && _srh_nameStartsWith(name, _SRH_LIGHTBOX_MEASURE_NAME)) isTagged = true;
      } catch(_eNm) { }
    }
    if(isTagged) {
      try {g.remove(); removed++;} catch(_eR) { }
    }
  }
  return removed;
}

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
  var addMeasures = !!opts.addMeasures;
  var updateMeasuresLive = !!opts.updateMeasuresLive;
  var measureOptions = opts.measureOptions || null;

  if(!(wMm > 0) || !(hMm > 0)) return 'Width and height must be > 0.';

  var w = _srh_mm2ptDoc(wMm);
  var h = _srh_mm2ptDoc(hMm);
  var ledOffset = _srh_mm2ptDoc(ledOffsetMm);
  var supportW = _srh_mm2ptDoc(25);
  var stroke1px = _srh_pxStrokeDoc(1);
  var lightboxId = _srh_generateLightboxId();

  var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; // [L,T,R,B]
  var centerX = (ab[0] + ab[2]) / 2;
  var centerY = (ab[1] + ab[3]) / 2;
  var left = centerX - (w / 2);
  var top = centerY + (h / 2);

  var frameLayer = _srh_getOrCreateLayer(doc, 'lightbox frame');
  var panelLayer = _srh_getOrCreateLayer(doc, 'acm led panel');

  var black = new RGBColor();
  black.red = 0; black.green = 0; black.blue = 0;

  var frameFill = _dim_hexToRGB('#848484');

  // Frame rectangle + inner offset as a minus-front compound shape
  var inset = _srh_mm2ptDoc(25);
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
        it.filled = true;
        if(frameFill) it.fillColor = frameFill;
        it.opacity = 50;
        it.stroked = true;
        it.strokeWidth = stroke1px;
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
    _srh_setLightboxItemTag(result, 'outer', lightboxId);

    try {doc.selection = prevSel;} catch(_eSel3) { }
  } else {
    _lb_applyStroke(outer);
    _srh_setLightboxItemTag(outer, 'outer', lightboxId);
  }

  // Supports
  var supportCenters = [];
  if(supports > 0 && w > supportW) {
    var gap = (w - (supports * supportW)) / (supports + 1);
    if(gap < 0) gap = 0;
    for(var i = 0; i < supports; i++) {
      var sx = left + gap * (i + 1) + supportW * i;
      var s = frameLayer.pathItems.rectangle(top, sx, supportW, h);
      try {s.filled = true; if(frameFill) s.fillColor = frameFill; s.opacity = 50;} catch(_eS0) { }
      try {s.stroked = true; s.strokeWidth = stroke1px; s.strokeColor = black;} catch(_eS1) { }
      _srh_setLightboxItemTag(s, 'support', lightboxId);
      supportCenters.push(sx + (supportW / 2));
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
      try {panel.stroked = true; panel.strokeWidth = stroke1px; panel.strokeColor = black;} catch(_eP1) { }
    }
  }

  if(addMeasures && measureOptions) {
    var bounds = {left: left, top: top, right: left + w, bottom: top - h};
    _srh_clearTaggedLightboxMeasures(doc, lightboxId);
    _srh_addLightboxMeasures(doc, bounds, supportCenters, measureOptions, lightboxId);
    if(updateMeasuresLive) {
      _srh_lightboxLiveEnabled[lightboxId] = true;
      _srh_lightboxLiveSignatures[lightboxId] = _srh_buildLightboxSignature(bounds, supportCenters);
    }
  }

  return 'Lightbox created: ' + wMm + ' x ' + hMm + ' mm, depth ' + dMm + ' mm, type ' + type + ', supports ' + supports + '.';
}

function signarama_helper_createLightboxWithLedPanel(jsonStr) {
  var opts = {};
  try {opts = JSON.parse(String(jsonStr));} catch(e) {opts = {};}
  var res = signarama_helper_createLightbox(jsonStr);

  // Reuse the same width/height/depth to create the LED panel and layout.
  var payload = {
    ledWatt: Number(opts.ledWatt || 0),
    ledCode: String(opts.ledCode || ''),
    ledVoltage: Number(opts.ledVoltage || 0),
    ledWidthMm: Number(opts.ledWidthMm || 66),
    ledHeightMm: Number(opts.ledHeightMm || 13),
    allowanceWmm: Number(opts.allowanceWmm || 0),
    allowanceHmm: Number(opts.allowanceHmm || 0),
    maxLedsInSeries: Number(opts.maxLedsInSeries || 50),
    flipLed: !!opts.flipLed,
    layoutWidthMm: Number(opts.widthMm || 0),
    layoutHeightMm: Number(opts.heightMm || 0),
    depthMm: Number(opts.depthMm || 0),
    boxWidthMm: Number(opts.widthMm || 0),
    boxHeightMm: Number(opts.heightMm || 0),
    forceBounds: true,
    ignoreSelection: true
  };

  // If ledOffset is provided, align the LED layout to the lightbox LED panel bounds.
  var wMm = Number(opts.widthMm || 0);
  var hMm = Number(opts.heightMm || 0);
  var ledOffsetMm = Number(opts.ledOffsetMm || 0);
  if(wMm > 0 && hMm > 0 && ledOffsetMm > 0) {
    var ab = app.activeDocument.artboards[app.activeDocument.artboards.getActiveArtboardIndex()].artboardRect;
    var centerX = (ab[0] + ab[2]) / 2;
    var centerY = (ab[1] + ab[3]) / 2;
    var wPt = _srh_mm2ptDoc(wMm);
    var hPt = _srh_mm2ptDoc(hMm);
    var ledOffsetPt = _srh_mm2ptDoc(ledOffsetMm);
    var bounds = {
      left: centerX - (wPt / 2) + ledOffsetPt,
      top: centerY + (hPt / 2) - ledOffsetPt,
      right: centerX + (wPt / 2) - ledOffsetPt,
      bottom: centerY - (hPt / 2) + ledOffsetPt
    };
    payload.boundsOverridePt = bounds;
  }

  signarama_helper_drawLedLayout(JSON.stringify(payload));
  return res + ' + LED panel/layout.';
}

function signarama_helper_updateLightboxMeasures(jsonStr) {
  if(!app.documents.length) return 'No open document.';
  var opts = {};
  try {opts = JSON.parse(String(jsonStr));} catch(e) {opts = {};}

  var measureOptions = opts.measureOptions || null;
  var force = !!opts.force;
  if(!measureOptions) return 'No measure options.';

  var doc = app.activeDocument;
  var lightboxes = _srh_findAllLightboxData(doc);
  if(!lightboxes || !lightboxes.length) return 'No live lightbox found.';

  var changed = 0;
  var supportsTotal = 0;
  for(var i = 0; i < lightboxes.length; i++) {
    var lb = lightboxes[i];
    if(!lb || !lb.id) continue;
    if(force) _srh_lightboxLiveEnabled[lb.id] = true;
    if(!_srh_lightboxLiveEnabled[lb.id]) continue;
    if(!force && _srh_lightboxLiveSignatures[lb.id] === lb.signature) continue;

    _srh_clearTaggedLightboxMeasures(doc, lb.id);
    _srh_addLightboxMeasures(doc, lb.bounds, lb.supportCenters, measureOptions, lb.id);
    _srh_lightboxLiveSignatures[lb.id] = lb.signature;
    changed++;
    supportsTotal += (lb.supportCenters ? lb.supportCenters.length : 0);
  }
  if(changed < 1) return 'No live measure changes.';
  return 'Updated lightbox measures live (' + changed + ' lightboxes, ' + supportsTotal + ' supports).';
}

function signarama_helper_drawLedLayout(jsonStr) {
  if(!app.documents.length) return 'No open document.';
  var doc = app.activeDocument;
  var opts = {};
  try {opts = JSON.parse(String(jsonStr));} catch(e) {opts = {};}

  var ledWatt = Number(opts.ledWatt || 0);
  var ledCode = String(opts.ledCode || '');
  var ledVoltage = Number(opts.ledVoltage || 0);
  var ledWidthMm = Number(opts.ledWidthMm || 0);
  var ledHeightMm = Number(opts.ledHeightMm || 0);
  var allowanceWmm = Number(opts.allowanceWmm || 0);
  var allowanceHmm = Number(opts.allowanceHmm || 0);
  var maxLedsInSeries = parseInt(opts.maxLedsInSeries, 10);
  if(!maxLedsInSeries || maxLedsInSeries < 1) maxLedsInSeries = 50;
  var flipLed = !!opts.flipLed;
  var layoutWidthMm = Number(opts.layoutWidthMm || 0);
  var layoutHeightMm = Number(opts.layoutHeightMm || 0);
  var forceBounds = !!opts.forceBounds;
  var ignoreSelection = !!opts.ignoreSelection;
  var boundsOverridePt = opts.boundsOverridePt || null;
  var depthMm = Number(opts.depthMm || 0);
  if(!(depthMm > 0)) depthMm = 150;

  var panelInsetWmm = 30;
  var panelInsetHmm = 30;

  if(!(ledWidthMm > 0) || !(ledHeightMm > 0)) return 'LED width/height must be > 0.';
  if(flipLed) {
    var tmp = ledWidthMm;
    ledWidthMm = ledHeightMm;
    ledHeightMm = tmp;
  }

  function _getItemBounds(it) {
    var b = null;
    try {b = it.visibleBounds;} catch(_e0) { }
    if(!b || b.length !== 4) {try {b = it.geometricBounds;} catch(_e1) { } }
    if(!b || b.length !== 4) return null;
    return {left: b[0], top: b[1], right: b[2], bottom: b[3]};
  }

  function _getSelectionBounds(doc) {
    if(!doc.selection || doc.selection.length === 0) return [];
    var boundsList = [];
    for(var i = 0; i < doc.selection.length; i++) {
      var it = doc.selection[i];
      if(!it) continue;
      var b = _getItemBounds(it);
      if(!b) continue;
      boundsList.push(b);
    }
    return boundsList;
  }

  function _getArtboardBounds(doc) {
    var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect; // [L,T,R,B]
    return {left: ab[0], top: ab[1], right: ab[2], bottom: ab[3]};
  }

  function _getTargetBounds(doc) {
    if(boundsOverridePt && boundsOverridePt.left != null) {
      return [boundsOverridePt];
    }
    if(forceBounds && (layoutWidthMm > 0) && (layoutHeightMm > 0)) {
      var ab = _getArtboardBounds(doc);
      var centerX = (ab.left + ab.right) / 2;
      var centerY = (ab.top + ab.bottom) / 2;
      var wPt = _srh_mm2ptDoc(layoutWidthMm);
      var hPt = _srh_mm2ptDoc(layoutHeightMm);
      return {
        left: centerX - (wPt / 2),
        top: centerY + (hPt / 2),
        right: centerX + (wPt / 2),
        bottom: centerY - (hPt / 2)
      };
    }
    var selBounds = ignoreSelection ? [] : _getSelectionBounds(doc);
    if(selBounds && selBounds.length) return selBounds;
    if((layoutWidthMm > 0) && (layoutHeightMm > 0)) {
      var ab2 = _getArtboardBounds(doc);
      var cx = (ab2.left + ab2.right) / 2;
      var cy = (ab2.top + ab2.bottom) / 2;
      var wPt2 = _srh_mm2ptDoc(layoutWidthMm);
      var hPt2 = _srh_mm2ptDoc(layoutHeightMm);
      return [{
        left: cx - (wPt2 / 2),
        top: cy + (hPt2 / 2),
        right: cx + (wPt2 / 2),
        bottom: cy - (hPt2 / 2)
      }];
    }
    return [_getArtboardBounds(doc)];
  }

  var boundsList = _getTargetBounds(doc);
  if(!boundsList || !boundsList.length) return 'No bounds found.';

  var depthPt = _srh_mm2ptDoc(depthMm);
  var ledWidthPt = _srh_mm2ptDoc(ledWidthMm);
  var ledHeightPt = _srh_mm2ptDoc(ledHeightMm);
  var allowanceWPt = _srh_mm2ptDoc(allowanceWmm);
  var allowanceHPt = _srh_mm2ptDoc(allowanceHmm);
  var panelInsetWPt = _srh_mm2ptDoc(panelInsetWmm);
  var panelInsetHPt = _srh_mm2ptDoc(panelInsetHmm);

  var panelLayer = _srh_getOrCreateLayer(doc, 'acm led panel');
  var ledLayer = _srh_getOrCreateLayer(doc, 'LEDs');
  var penLayer = _srh_getOrCreateLayer(doc, 'LED Pen');
  var penGroupLayer = _srh_getOrCreateLayer(doc, 'led pen group and text');

  var black = new RGBColor();
  black.red = 0; black.green = 0; black.blue = 0;
  var red = new RGBColor();
  red.red = 255; red.green = 0; red.blue = 0;
  var stroke1px = _srh_pxStrokeDoc(1);

  var totalRows = 0;
  var totalCols = 0;
  var totalLayouts = 0;
  for(var b = 0; b < boundsList.length; b++) {
    var bounds = boundsList[b];
    var boxWidthPt = bounds.right - bounds.left;
    var boxHeightPt = bounds.top - bounds.bottom;
    if(!(boxWidthPt > 0) || !(boxHeightPt > 0)) continue;

    var distanceToOutsideBox = depthPt / 2;
    var rows = Math.ceil(((boxHeightPt - 2 * distanceToOutsideBox) / (depthPt + allowanceHPt)) + 1);
    var cols = Math.ceil((((boxWidthPt - 2 * distanceToOutsideBox) - ledWidthPt) / (depthPt + allowanceWPt)) + 1);
    if(rows < 1) rows = 1;
    if(cols < 1) cols = 1;

    var centresRows = (boxHeightPt - 2 * distanceToOutsideBox + ((rows - 1) * allowanceHPt)) / (rows > 1 ? (rows - 1) : 1);
    var centresCols = (boxWidthPt - 2 * distanceToOutsideBox - ledWidthPt + ((cols - 1) * allowanceWPt)) / (cols > 1 ? (cols - 1) : 1);

    var centresInFromPanelCol = 0.5 * ((boxWidthPt - (cols - 1) * centresCols) - 2 * panelInsetWPt);
    var centresInFromPanelRow = 0.5 * ((boxHeightPt - (rows - 1) * centresRows) - (2 * panelInsetHPt));

    var startX = bounds.left + panelInsetWPt + centresInFromPanelCol;
    var startY = bounds.top - panelInsetHPt - centresInFromPanelRow;

    var panelRect = panelLayer.pathItems.rectangle(bounds.top, bounds.left, boxWidthPt, boxHeightPt);
    try {panelRect.filled = false;} catch(_ePF) { }
    try {panelRect.stroked = true; panelRect.strokeWidth = stroke1px; panelRect.strokeColor = black;} catch(_ePS) { }

    // Column-major ordering (consider LEDs in columns)
    var ledRectsByColumn = [];
    for(var c = 0; c < cols; c++) {
      ledRectsByColumn[c] = [];
      for(var r = 0; r < rows; r++) {
        var cx = startX + c * centresCols;
        var cy = startY - r * centresRows;

        var left = cx - ledWidthPt / 2;
        var top = cy + ledHeightPt / 2;

        var rect = ledLayer.pathItems.rectangle(top, left, ledWidthPt, ledHeightPt);
        try {rect.filled = false;} catch(_eF) { }
        try {rect.stroked = true; rect.strokeWidth = stroke1px; rect.strokeColor = black;} catch(_eS) { }

        var line = penLayer.pathItems.add();
        if(flipLed) {
          line.setEntirePath([[cx, cy - ledHeightPt / 2], [cx, cy + ledHeightPt / 2]]);
        } else {
          line.setEntirePath([[cx - ledWidthPt / 2, cy], [cx + ledWidthPt / 2, cy]]);
        }
        try {line.filled = false;} catch(_eLF) { }
        try {line.stroked = true; line.strokeWidth = stroke1px; line.strokeColor = red;} catch(_eLS) { }

        ledRectsByColumn[c].push({left: left, top: top, right: left + ledWidthPt, bottom: top - ledHeightPt});
      }
    }

    // Group LEDs by columns, max in series, no overlaps
    var colIndex = 0;
    while(colIndex < cols) {
      var colsInGroup = 1;
      if(rows > 0) {
        colsInGroup = Math.floor(maxLedsInSeries / rows);
        if(colsInGroup < 1) colsInGroup = 1;
        if(colsInGroup > (cols - colIndex)) colsInGroup = cols - colIndex;
        // If too many LEDs in group, reduce by one column
        while(colsInGroup > 1 && (colsInGroup * rows) > maxLedsInSeries) {
          colsInGroup -= 1;
        }
      }

      var gLeft = null, gTop = null, gRight = null, gBottom = null;
      var count = 0;
      for(var gc = 0; gc < colsInGroup; gc++) {
        var colRects = ledRectsByColumn[colIndex + gc];
        for(var gr = 0; gr < colRects.length; gr++) {
          var lr = colRects[gr];
          if(gLeft === null) {
            gLeft = lr.left; gTop = lr.top; gRight = lr.right; gBottom = lr.bottom;
          } else {
            if(lr.left < gLeft) gLeft = lr.left;
            if(lr.top > gTop) gTop = lr.top;
            if(lr.right > gRight) gRight = lr.right;
            if(lr.bottom < gBottom) gBottom = lr.bottom;
          }
          count++;
        }
      }

      if(gLeft !== null) {
        var gWidth = gRight - gLeft;
        var gHeight = gTop - gBottom;
        var grpRect = penGroupLayer.pathItems.rectangle(gTop, gLeft, gWidth, gHeight);
        try {grpRect.filled = false;} catch(_eGF) { }
        try {grpRect.stroked = true; grpRect.strokeWidth = stroke1px; grpRect.strokeColor = black;} catch(_eGS) { }

        // Label at bottom center with LED count
        try {
          var label = penGroupLayer.textFrames.pointText([gLeft + (gWidth / 2), gBottom - _srh_mm2ptDoc(5)]);
          label.contents = count + " LEDs";
          try {label.textRange.paragraphAttributes.justification = Justification.CENTER;} catch(_eJust) { }
          try {label.textRange.characterAttributes.size = _srh_ptDoc(20);} catch(_eSize) { }
          try {label.textRange.characterAttributes.fillColor = black;} catch(_eCol) { }
        } catch(_eLbl) { }
      }

      colIndex += colsInGroup;
    }

    // Bottom-center lightbox LED spec label.
    try {
      var ledBits = [];
      if(ledWatt) ledBits.push(ledWatt + 'W');
      if(ledCode && ledCode.length) ledBits.push(ledCode);
      if(ledVoltage) ledBits.push(ledVoltage + 'V');
      if(ledBits.length) {
        var spec = ledBits.join(' | ');
        var specLabel = penGroupLayer.textFrames.pointText([(bounds.left + bounds.right) * 0.5, bounds.bottom - _srh_mm2ptDoc(3)]);
        specLabel.contents = spec;
        try {specLabel.textRange.paragraphAttributes.justification = Justification.CENTER;} catch(_eSpecJust) { }
        try {specLabel.textRange.characterAttributes.size = _srh_ptDoc(12);} catch(_eSpecSize) { }
        try {specLabel.textRange.characterAttributes.fillColor = black;} catch(_eSpecCol) { }
      }
    } catch(_eSpecLbl) { }

    totalRows += rows;
    totalCols += cols;
    totalLayouts++;
  }

  _srh_bringLayerToFront(penLayer);
  _srh_bringLayerToFront(ledLayer);
  _srh_bringLayerToFront(penGroupLayer);

  return 'LED layout drawn. Layouts: ' + totalLayouts + ', Rows: ' + totalRows + ', Columns: ' + totalCols + ', Watt: ' + ledWatt + 'W.';
}

function _srh_addLightboxMeasures(doc, bounds, supportCenters, opts, lightboxId) {
  if(!doc || !bounds || !opts) return;

  var scaleFactor = _srh_getScaleFactor();
  var offsetPt = _dim_mm2ptDoc(opts.offsetMm || 10, scaleFactor);
  var ticLenPt = _dim_mm2ptDoc(opts.ticLenMm || 2, scaleFactor);
  var textPt = _dim_ptDoc(opts.textPt || 10, scaleFactor);
  var strokePt = _dim_ptDoc(opts.strokePt || 1, scaleFactor);
  var decimals = (opts.decimals | 0);
  var textOffsetPt = _dim_mm2ptDoc(opts.labelGapMm || 0, scaleFactor);
  var includeArrowhead = !!opts.includeArrowhead;
  var arrowheadSizePt = _dim_ptDoc(opts.arrowheadSizePt || 0, scaleFactor);
  var scaleAppearance = opts.scaleAppearance || 1;

  var lineColor = null;
  try {
    lineColor = _dim_hexToRGB(opts.lineColor);
    if(!lineColor) lineColor = _dim_parseHexColorToRGBColor(opts.lineColor);
  } catch(_eLc2) {lineColor = null;}
  if(!lineColor) lineColor = _dim_hexToRGB('#000000');
  var textColor = opts.textColor;

  var dOpts = {
    offsetPt: offsetPt * scaleAppearance,
    ticLenPt: ticLenPt * scaleAppearance,
    textPt: textPt * scaleAppearance,
    strokePt: strokePt * scaleAppearance,
    decimals: decimals,
    scaleFactor: scaleFactor,
    textOffsetPt: textOffsetPt * scaleAppearance,
    textColor: textColor,
    lineColor: lineColor,
    includeArrowhead: includeArrowhead,
    arrowheadSizePt: arrowheadSizePt * scaleAppearance
  };

  var lyr = _dim_ensureLayer('Dimensions');

  // Top width measure
  var gTop = _dim_drawHorizontalDim(lyr, bounds.left, bounds.right, bounds.top + dOpts.offsetPt, dOpts.ticLenPt, dOpts.textPt, dOpts.strokePt, dOpts.decimals, 'TOP', dOpts.textOffsetPt, dOpts.scaleFactor, dOpts.textColor, dOpts.lineColor, dOpts.includeArrowhead, dOpts.arrowheadSizePt);
  _srh_tagLightboxMeasureGroup(gTop, lightboxId);
  // Left height measure
  var gLeft = _dim_drawVerticalDim(lyr, bounds.top, bounds.bottom, bounds.left - dOpts.offsetPt, dOpts.ticLenPt, dOpts.textPt, dOpts.strokePt, dOpts.decimals, 90, dOpts.textOffsetPt, 'LEFT', dOpts.scaleFactor, dOpts.textColor, dOpts.lineColor, dOpts.includeArrowhead, dOpts.arrowheadSizePt);
  _srh_tagLightboxMeasureGroup(gLeft, lightboxId);

  // Bottom measures from left edge to each support center
  if(supportCenters && supportCenters.length) {
    var sortedCenters = supportCenters.slice(0);
    sortedCenters.sort(function(a, b) {return a - b;});
    var supportMeasureYOffsetStep = Math.max(dOpts.ticLenPt, dOpts.strokePt * 2);
    for(var i = 0; i < sortedCenters.length; i++) {
      var yLine = bounds.bottom - dOpts.offsetPt - (i * supportMeasureYOffsetStep);
      var gBottom = _dim_drawHorizontalDim(lyr, bounds.left, sortedCenters[i], yLine, dOpts.ticLenPt, dOpts.textPt, dOpts.strokePt, dOpts.decimals, 'BOTTOM', dOpts.textOffsetPt, dOpts.scaleFactor, dOpts.textColor, dOpts.lineColor, dOpts.includeArrowhead, dOpts.arrowheadSizePt);
      _srh_tagLightboxMeasureGroup(gBottom, lightboxId);
    }
  }
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

