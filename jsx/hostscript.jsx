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
function _srh_pt2mm(pt) {return pt / 2.834645669291339;}
function _srh_clamp(n, a, b) {return n < a ? a : (n > b ? b : n);}

function _srh_getDocScaleFactor(doc) {
  var sf = 1.0;
  try {
    sf = (doc && doc.scaleFactor) ? Number(doc.scaleFactor) : 1.0;
  } catch(_eSf) {sf = 1.0;}
  if(!(sf > 0)) sf = 1.0;
  return sf;
}

function _srh_mm2docpt(mm, doc) {
  return _srh_mm2pt(mm) / _srh_getDocScaleFactor(doc || app.activeDocument);
}

function _srh_getAppearanceScale(doc) {
  return 1.0 / _srh_getDocScaleFactor(doc || app.activeDocument);
}

var _srh_lightboxMeasureWatch = {
  enabled: false,
  lastDocName: '',
  lastSupportKey: '',
  wasGrabbed: false,
  measureOptions: null
};

function _srh_supportWatchNum(n) {
  return Math.round(Number(n || 0) * 1000) / 1000;
}

function _srh_getBoundsFromItem(it) {
  var b = null;
  try {b = it.visibleBounds;} catch(_e0) { }
  if(!b || b.length !== 4) {
    try {b = it.geometricBounds;} catch(_e1) { }
  }
  if(!b || b.length !== 4) return null;
  return {left: b[0], top: b[1], right: b[2], bottom: b[3]};
}

function _srh_getLightboxFrameBounds(doc) {
  if(!doc) return null;
  var layer = null;
  try {layer = doc.layers.getByName('lightbox frame');} catch(_eL) {layer = null;}
  if(!layer) return null;

  var best = null;
  var bestArea = 0;
  for(var i = 0; i < layer.pageItems.length; i++) {
    var b = _srh_getBoundsFromItem(layer.pageItems[i]);
    if(!b) continue;
    var w = b.right - b.left;
    var h = b.top - b.bottom;
    var area = w * h;
    if(!(w > 0) || !(h > 0)) continue;
    if(area > bestArea) {bestArea = area; best = b;}
  }
  return best;
}

function _srh_getSupportCentersForFrame(doc, frameBounds) {
  var centers = [];
  if(!doc || !frameBounds) return centers;

  var layer = null;
  try {layer = doc.layers.getByName('lightbox frame');} catch(_eL) {layer = null;}
  if(!layer) return centers;

  var frameW = frameBounds.right - frameBounds.left;
  var frameH = frameBounds.top - frameBounds.bottom;
  if(!(frameW > 0) || !(frameH > 0)) return centers;

  for(var li = 0; li < layer.pathItems.length; li++) {
    var b = _srh_getBoundsFromItem(layer.pathItems[li]);
    if(!b) continue;
    var w = b.right - b.left;
    var h = b.top - b.bottom;
    if(!(w > 0) || !(h > 0)) continue;
    var isSupportLike = (h >= (frameH * 0.9)) && (w < (frameW * 0.5));
    var inFrame = (b.left >= frameBounds.left - 0.01) && (b.right <= frameBounds.right + 0.01);
    if(!isSupportLike || !inFrame) continue;
    centers.push((b.left + b.right) / 2);
  }

  centers.sort(function(a, b) {return a - b;});
  return centers;
}

function _srh_getSelectedSupportCenters(doc, frameBounds) {
  var centers = [];
  if(!doc || !frameBounds || !doc.selection || !doc.selection.length) return centers;

  var frameW = frameBounds.right - frameBounds.left;
  var frameH = frameBounds.top - frameBounds.bottom;
  for(var i = 0; i < doc.selection.length; i++) {
    var b = _srh_getBoundsFromItem(doc.selection[i]);
    if(!b) continue;
    var w = b.right - b.left;
    var h = b.top - b.bottom;
    if(!(w > 0) || !(h > 0)) continue;
    var isSupportLike = (h >= (frameH * 0.9)) && (w < (frameW * 0.5));
    var inFrame = (b.left >= frameBounds.left - 0.01) && (b.right <= frameBounds.right + 0.01);
    if(!isSupportLike || !inFrame) continue;
    centers.push((b.left + b.right) / 2);
  }

  centers.sort(function(a, b) {return a - b;});
  return centers;
}

function _srh_makeSupportKey(centers) {
  if(!centers || !centers.length) return '';
  var out = [];
  for(var i = 0; i < centers.length; i++) out.push(_srh_supportWatchNum(centers[i]));
  return out.join('|');
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

  // Convert all text types to outlines (within group), then expand appearance if needed
  try {
    var tfs = grp.textFrames;
    for(var t = tfs.length - 1; t >= 0; t--) {
      try {
        var tf = tfs[t];
        tf.createOutline();
      } catch(_e1) { }
    }
  } catch(_e2) { }

  // Ensure any remaining text types are converted to outlines/expanded
  try {
    doc.selection = null;
    grp.selected = true;
    app.executeMenuCommand("Expand3");
  } catch(_eExp3) { }

  // Outline strokes after text has been converted (all paths)
  try {
    doc.selection = null;
    grp.selected = true;
    app.executeMenuCommand("OffsetPath v22");
  } catch(_e3) { }

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
          if(!cb || cb.length !== 4) {try {cb = child.geometricBounds;} catch(_eC1) {cb = null;} }
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
  } catch(_eLc2) {lineColor = null;}
  if(!lineColor) lineColor = _dim_hexToRGB('#000000');
  var includeArrowhead = !!opts.includeArrowhead;
  var arrowheadSizePt = opts.arrowheadSizePt || 0;
  var scaleAppearance = (opts.scaleAppearance || 1) * _srh_getAppearanceScale(doc);

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
  var scaleAppearance = (opts.scaleAppearance || 1) * _srh_getAppearanceScale(doc);

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
  var addMeasures = !!opts.addMeasures;
  var measureOptions = opts.measureOptions || null;

  if(!(wMm > 0) || !(hMm > 0)) return 'Width and height must be > 0.';

  var w = _srh_mm2docpt(wMm, doc);
  var h = _srh_mm2docpt(hMm, doc);
  var ledOffset = _srh_mm2docpt(ledOffsetMm, doc);
  var supportW = _srh_mm2docpt(25, doc);
  var appearanceScale = _srh_getAppearanceScale(doc);

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
  var inset = _srh_mm2docpt(25, doc);
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
        it.strokeWidth = 1 * appearanceScale;
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
  var supportItems = [];
  if(supports > 0 && w > supportW) {
    var gap = (w - (supports * supportW)) / (supports + 1);
    if(gap < 0) gap = 0;
    for(var i = 0; i < supports; i++) {
      var sx = left + gap * (i + 1) + supportW * i;
      var s = frameLayer.pathItems.rectangle(top, sx, supportW, h);
      try {s.filled = true; if(frameFill) s.fillColor = frameFill; s.opacity = 50;} catch(_eS0) { }
      try {s.stroked = true; s.strokeWidth = 1 * appearanceScale; s.strokeColor = black;} catch(_eS1) { }
      supportItems.push(s);
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
      try {panel.stroked = true; panel.strokeWidth = 1 * appearanceScale; panel.strokeColor = black;} catch(_eP1) { }
    }
  }

  if(addMeasures && measureOptions) {
    var bounds = {left: left, top: top, right: left + w, bottom: top - h};
    _srh_addLightboxMeasures(doc, bounds, supportItems, measureOptions);

    _srh_lightboxMeasureWatch.enabled = true;
    _srh_lightboxMeasureWatch.lastDocName = String(doc.name || '');
    _srh_lightboxMeasureWatch.lastSupportKey = _srh_makeSupportKey(_srh_getSupportCentersForFrame(doc, bounds));
    _srh_lightboxMeasureWatch.wasGrabbed = false;
    _srh_lightboxMeasureWatch.measureOptions = measureOptions;
  } else {
    _srh_lightboxMeasureWatch.enabled = false;
    _srh_lightboxMeasureWatch.lastSupportKey = '';
    _srh_lightboxMeasureWatch.wasGrabbed = false;
    _srh_lightboxMeasureWatch.measureOptions = null;
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
    var wPt = _srh_mm2docpt(wMm, app.activeDocument);
    var hPt = _srh_mm2docpt(hMm, app.activeDocument);
    var ledOffsetPt = _srh_mm2docpt(ledOffsetMm, app.activeDocument);
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

function signarama_helper_drawLedLayout(jsonStr) {
  if(!app.documents.length) return 'No open document.';
  var doc = app.activeDocument;
  var opts = {};
  try {opts = JSON.parse(String(jsonStr));} catch(e) {opts = {};}

  var ledWatt = Number(opts.ledWatt || 0);
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
  var appearanceScale = _srh_getAppearanceScale(doc);

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
      var wPt = _srh_mm2docpt(layoutWidthMm, doc);
      var hPt = _srh_mm2docpt(layoutHeightMm, doc);
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
      var wPt2 = _srh_mm2docpt(layoutWidthMm, doc);
      var hPt2 = _srh_mm2docpt(layoutHeightMm, doc);
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

  var depthPt = _srh_mm2docpt(depthMm, doc);
  var ledWidthPt = _srh_mm2docpt(ledWidthMm, doc);
  var ledHeightPt = _srh_mm2docpt(ledHeightMm, doc);
  var allowanceWPt = _srh_mm2docpt(allowanceWmm, doc);
  var allowanceHPt = _srh_mm2docpt(allowanceHmm, doc);
  var panelInsetWPt = _srh_mm2docpt(panelInsetWmm, doc);
  var panelInsetHPt = _srh_mm2docpt(panelInsetHmm, doc);

  var panelLayer = _srh_getOrCreateLayer(doc, 'acm led panel');
  var ledLayer = _srh_getOrCreateLayer(doc, 'LEDs');
  var penLayer = _srh_getOrCreateLayer(doc, 'LED Pen');
  var penGroupLayer = _srh_getOrCreateLayer(doc, 'led pen group and text');

  var black = new RGBColor();
  black.red = 0; black.green = 0; black.blue = 0;
  var red = new RGBColor();
  red.red = 255; red.green = 0; red.blue = 0;

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
    try {panelRect.stroked = true; panelRect.strokeWidth = 1 * appearanceScale; panelRect.strokeColor = black;} catch(_ePS) { }

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
        try {rect.stroked = true; rect.strokeWidth = 1 * appearanceScale; rect.strokeColor = black;} catch(_eS) { }

        var line = penLayer.pathItems.add();
        if(flipLed) {
          line.setEntirePath([[cx, cy - ledHeightPt / 2], [cx, cy + ledHeightPt / 2]]);
        } else {
          line.setEntirePath([[cx - ledWidthPt / 2, cy], [cx + ledWidthPt / 2, cy]]);
        }
        try {line.filled = false;} catch(_eLF) { }
        try {line.stroked = true; line.strokeWidth = 1 * appearanceScale; line.strokeColor = red;} catch(_eLS) { }

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
        try {grpRect.stroked = true; grpRect.strokeWidth = 1 * appearanceScale; grpRect.strokeColor = black;} catch(_eGS) { }

        // Label at bottom center with LED count
        try {
          var label = penGroupLayer.textFrames.pointText([gLeft + (gWidth / 2), gBottom - _srh_mm2docpt(5, doc)]);
          label.contents = count + " LEDs";
          try {label.textRange.paragraphAttributes.justification = Justification.CENTER;} catch(_eJust) { }
          try {label.textRange.characterAttributes.size = 10 * appearanceScale;} catch(_eSize) { }
          try {label.textRange.characterAttributes.fillColor = black;} catch(_eCol) { }
        } catch(_eLbl) { }
      }

      colIndex += colsInGroup;
    }

    totalRows += rows;
    totalCols += cols;
    totalLayouts++;
  }

  _srh_bringLayerToFront(penLayer);
  _srh_bringLayerToFront(ledLayer);
  _srh_bringLayerToFront(penGroupLayer);

  return 'LED layout drawn. Layouts: ' + totalLayouts + ', Rows: ' + totalRows + ', Columns: ' + totalCols + ', Watt: ' + ledWatt + 'W.';
}

function _srh_addLightboxMeasures(doc, bounds, supportsData, opts) {
  if(!doc || !bounds || !opts) return;

  var offsetPt = _dim_mm2pt(opts.offsetMm || 10);
  var ticLenPt = _dim_mm2pt(opts.ticLenMm || 2);
  var textPt = opts.textPt || 10;
  var strokePt = opts.strokePt || 1;
  var decimals = (opts.decimals | 0);
  var textOffsetPt = _dim_mm2pt(opts.labelGapMm || 0);
  var includeArrowhead = !!opts.includeArrowhead;
  var arrowheadSizePt = opts.arrowheadSizePt || 0;
  var scaleAppearance = (opts.scaleAppearance || 1) * _srh_getAppearanceScale(doc);

  var lineColor = null;
  try {
    lineColor = _dim_hexToRGB(opts.lineColor);
    if(!lineColor) lineColor = _dim_parseHexColorToRGBColor(opts.lineColor);
  } catch(_eLc2) {lineColor = null;}
  if(!lineColor) lineColor = _dim_hexToRGB('#000000');
  var textColor = opts.textColor;

  var scaleFactor = 1.0;
  try {
    var sfDoc = app.activeDocument.scaleFactor;
    if(sfDoc && sfDoc > 0) scaleFactor = sfDoc;
  } catch(e) {scaleFactor = 1.0;}

  var dOpts = {
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

  var lyr = _dim_ensureLayer('Dimensions');

  // Top width measure
  _dim_drawHorizontalDim(lyr, bounds.left, bounds.right, bounds.top + dOpts.offsetPt, dOpts.ticLenPt, dOpts.textPt, dOpts.strokePt, dOpts.decimals, 'TOP', dOpts.textOffsetPt, dOpts.scaleFactor, dOpts.textColor, dOpts.lineColor, dOpts.includeArrowhead, dOpts.arrowheadSizePt);
  // Left height measure
  _dim_drawVerticalDim(lyr, bounds.top, bounds.bottom, bounds.left - dOpts.offsetPt, dOpts.ticLenPt, dOpts.textPt, dOpts.strokePt, dOpts.decimals, 90, dOpts.textOffsetPt, 'LEFT', dOpts.scaleFactor, dOpts.textColor, dOpts.lineColor, dOpts.includeArrowhead, dOpts.arrowheadSizePt);

  function _srh_getSupportCenterX(supportRef) {
    if(supportRef == null) return null;
    if(typeof supportRef === 'number') return supportRef;

    var b = null;
    try {b = supportRef.visibleBounds;} catch(_eSb0) { }
    if(!b || b.length !== 4) {
      try {b = supportRef.geometricBounds;} catch(_eSb1) { }
    }
    if(!b || b.length !== 4) return null;
    return (b[0] + b[2]) / 2;
  }

  function _srh_collectSupportCentersFromLayer(doc, frameBounds) {
    var centers = [];
    if(!doc || !frameBounds) return centers;

    var layer = null;
    try {layer = doc.layers.getByName('lightbox frame');} catch(_eL) {layer = null;}
    if(!layer) return centers;

    var frameW = frameBounds.right - frameBounds.left;
    var frameH = frameBounds.top - frameBounds.bottom;
    if(!(frameW > 0) || !(frameH > 0)) return centers;

    for(var li = 0; li < layer.pathItems.length; li++) {
      var it = layer.pathItems[li];
      var b = null;
      try {b = it.visibleBounds;} catch(_eB0) { }
      if(!b || b.length !== 4) {
        try {b = it.geometricBounds;} catch(_eB1) { }
      }
      if(!b || b.length !== 4) continue;

      var l = b[0], t = b[1], r = b[2], bt = b[3];
      var w = r - l;
      var h = t - bt;
      if(!(w > 0) || !(h > 0)) continue;

      // Supports are tall rectangles with near frame height but narrower than frame width.
      var isSupportLike = (h >= (frameH * 0.9)) && (w < (frameW * 0.5));
      // Keep only supports inside the frame extents.
      var inFrame = (l >= frameBounds.left - 0.01) && (r <= frameBounds.right + 0.01);
      if(!isSupportLike || !inFrame) continue;

      centers.push((l + r) / 2);
    }

    centers.sort(function(a, b) {return a - b;});
    return centers;
  }

  var supportCenters = [];

  // Resolve support centers from provided data first.
  if(supportsData && supportsData.length) {
    for(var i = 0; i < supportsData.length; i++) {
      var centerX = _srh_getSupportCenterX(supportsData[i]);
      if(centerX == null) continue;
      supportCenters.push(centerX);
    }
  }

  // Fallback: detect support objects from the lightbox layer (useful after supports were manually moved).
  if(!supportCenters.length) {
    supportCenters = _srh_collectSupportCentersFromLayer(doc, bounds);
  }

  // Bottom measures from left edge to each support center.
  for(var ci = 0; ci < supportCenters.length; ci++) {
    _dim_drawHorizontalDim(lyr, bounds.left, supportCenters[ci], bounds.bottom - dOpts.offsetPt, dOpts.ticLenPt, dOpts.textPt, dOpts.strokePt, dOpts.decimals, 'BOTTOM', dOpts.textOffsetPt, dOpts.scaleFactor, dOpts.textColor, dOpts.lineColor, dOpts.includeArrowhead, dOpts.arrowheadSizePt);
  }
}



function signarama_helper_lightboxSupportWatchTick() {
  if(!_srh_lightboxMeasureWatch.enabled) return 'ontick disabled';
  if(!app.documents.length) return 'ontick no-document';

  var doc = app.activeDocument;
  if(!doc) return 'ontick no-document';
  var docName = String(doc.name || '');

  if(_srh_lightboxMeasureWatch.lastDocName && _srh_lightboxMeasureWatch.lastDocName !== docName) {
    _srh_lightboxMeasureWatch.lastDocName = docName;
    _srh_lightboxMeasureWatch.lastSupportKey = '';
    _srh_lightboxMeasureWatch.wasGrabbed = false;
  }

  var frameBounds = _srh_getLightboxFrameBounds(doc);
  if(!frameBounds) return 'ontick no-lightbox-frame';

  var selectedCenters = _srh_getSelectedSupportCenters(doc, frameBounds);
  var isGrabbed = selectedCenters.length > 0;
  var events = ['ontick'];

  if(isGrabbed && !_srh_lightboxMeasureWatch.wasGrabbed) events.push('ongrab');
  if(!isGrabbed && _srh_lightboxMeasureWatch.wasGrabbed) events.push('ondrop');
  _srh_lightboxMeasureWatch.wasGrabbed = isGrabbed;

  var supportCenters = _srh_getSupportCentersForFrame(doc, frameBounds);
  var key = _srh_makeSupportKey(supportCenters);

  if(key && key !== _srh_lightboxMeasureWatch.lastSupportKey) {
    events.push('onmove');
    _srh_lightboxMeasureWatch.lastSupportKey = key;

    if(_srh_lightboxMeasureWatch.measureOptions) {
      try {
        var dimLyr = doc.layers.getByName('Dimensions');
        dimLyr.remove();
      } catch(_eClr) { }
      _srh_addLightboxMeasures(doc, frameBounds, supportCenters, _srh_lightboxMeasureWatch.measureOptions);
      events.push('remeasure');
    } else {
      events.push('remeasure-skipped-no-options');
    }
    if(!b || b.length !== 4) return null;
    return (b[0] + b[2]) / 2;
  }

  function _srh_collectSupportCentersFromLayer(doc, frameBounds) {
    var centers = [];
    if(!doc || !frameBounds) return centers;

    var layer = null;
    try {layer = doc.layers.getByName('lightbox frame');} catch(_eL) {layer = null;}
    if(!layer) return centers;

    var frameW = frameBounds.right - frameBounds.left;
    var frameH = frameBounds.top - frameBounds.bottom;
    if(!(frameW > 0) || !(frameH > 0)) return centers;

    for(var li = 0; li < layer.pathItems.length; li++) {
      var it = layer.pathItems[li];
      var b = null;
      try {b = it.visibleBounds;} catch(_eB0) { }
      if(!b || b.length !== 4) {
        try {b = it.geometricBounds;} catch(_eB1) { }
      }
      if(!b || b.length !== 4) continue;

      var l = b[0], t = b[1], r = b[2], bt = b[3];
      var w = r - l;
      var h = t - bt;
      if(!(w > 0) || !(h > 0)) continue;

      // Supports are tall rectangles with near frame height but narrower than frame width.
      var isSupportLike = (h >= (frameH * 0.9)) && (w < (frameW * 0.5));
      // Keep only supports inside the frame extents.
      var inFrame = (l >= frameBounds.left - 0.01) && (r <= frameBounds.right + 0.01);
      if(!isSupportLike || !inFrame) continue;

      centers.push((l + r) / 2);
    }

    centers.sort(function(a, b) {return a - b;});
    return centers;
  }

  var supportCenters = [];

  // Resolve support centers from provided data first.
  if(supportsData && supportsData.length) {
    for(var i = 0; i < supportsData.length; i++) {
      var centerX = _srh_getSupportCenterX(supportsData[i]);
      if(centerX == null) continue;
      supportCenters.push(centerX);
    }
  }

  // Fallback: detect support objects from the lightbox layer (useful after supports were manually moved).
  if(!supportCenters.length) {
    supportCenters = _srh_collectSupportCentersFromLayer(doc, bounds);
  }

  // Bottom measures from left edge to each support center.
  for(var ci = 0; ci < supportCenters.length; ci++) {
    _dim_drawHorizontalDim(lyr, bounds.left, supportCenters[ci], bounds.bottom - dOpts.offsetPt, dOpts.ticLenPt, dOpts.textPt, dOpts.strokePt, dOpts.decimals, 'BOTTOM', dOpts.textOffsetPt, dOpts.scaleFactor, dOpts.textColor, dOpts.lineColor, dOpts.includeArrowhead, dOpts.arrowheadSizePt);
  }

  events.push('supports=' + supportCenters.length);
  return events.join(' ');
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
