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
function signarama_helper_fitArtboardToArtwork(marginMm) {
  if(!app.documents.length) return 'No open document.';

  var doc = app.activeDocument;

  var b = _srh_getDocumentArtworkBounds(doc);
  if(!b) return 'No eligible artwork found (all items locked/hidden/guides).';

  var marginPt = _srh_mm2ptDoc(Number(marginMm) || 0);
  if(!(marginPt >= 0)) marginPt = 0;

  var fitted = {
    left: b.left - marginPt,
    top: b.top + marginPt,
    right: b.right + marginPt,
    bottom: b.bottom - marginPt
  };

  var idx = doc.artboards.getActiveArtboardIndex();
  doc.artboards[idx].artboardRect = [fitted.left, fitted.top, fitted.right, fitted.bottom];

  return 'Artboard fitted to artwork bounds (margin ' + (Math.round((Number(marginMm) || 0) * 1000) / 1000) + 'mm): ' + _srh_fmtBounds(fitted);
}

/**
 * Debug helper: create a rectangle at active artboard center and apply
 * a brand-new linear gradient intended to use stop rampPoints 10/90.
 */
function signarama_helper_debugCreateGradientRect1090() {
  if(!app.documents.length) return 'No open document.';
  var doc = app.activeDocument;
  var errs = [];
  function _err(tag, e) {
    var t = '';
    try {t = String(e);} catch(_eDbgEr0) {t = 'unknown';}
    errs.push(tag + ': ' + t);
  }

  try {app.beginUndoGroup('SRH Debug Gradient 10/90');} catch(_eDbgUg0) { }
  try {
    var abIdx = 0;
    var ab = null;
    try {abIdx = doc.artboards.getActiveArtboardIndex();} catch(_eDbgGr0) {abIdx = 0; _err('activeArtboardIndex', _eDbgGr0);}
    try {ab = doc.artboards[abIdx].artboardRect;} catch(_eDbgGr1) {ab = null; _err('artboardRect', _eDbgGr1);}
    if(!ab || ab.length !== 4) return 'Could not read active artboard. errors=' + errs.join(' | ');

    var left = Number(ab[0]);
    var top = Number(ab[1]);
    var right = Number(ab[2]);
    var bottom = Number(ab[3]);
    var cx = (left + right) / 2;
    var cy = (top + bottom) / 2;

    var w = _srh_mm2ptDoc(120);
    var h = _srh_mm2ptDoc(60);
    if(!(w > 0)) w = 340;
    if(!(h > 0)) h = 170;
    var rectLeft = cx - (w / 2);
    var rectTop = cy + (h / 2);

    var dbgLayer = null;
    try {dbgLayer = doc.layers.getByName('SRH_DEBUG');} catch(_eDbgLay0) {dbgLayer = null;}
    if(!dbgLayer) {
      try {
        dbgLayer = doc.layers.add();
        dbgLayer.name = 'SRH_DEBUG';
      } catch(_eDbgLay1) {dbgLayer = null; _err('createLayer', _eDbgLay1);}
    }
    if(dbgLayer) {
      try {dbgLayer.visible = true;} catch(_eDbgLay2) {_err('layerVisible', _eDbgLay2);}
      try {dbgLayer.locked = false;} catch(_eDbgLay3) {_err('layerUnlocked', _eDbgLay3);}
      try {doc.activeLayer = dbgLayer;} catch(_eDbgLay5) {_err('setActiveLayer', _eDbgLay5);}
    }

    var rect = null;
    try {
      if(dbgLayer) rect = dbgLayer.pathItems.rectangle(rectTop, rectLeft, w, h);
      else rect = doc.pathItems.rectangle(rectTop, rectLeft, w, h);
    } catch(_eDbgGr2) {rect = null; _err('createRectangle', _eDbgGr2);}
    if(!rect) {
      return 'Failed to create debug rectangle. targetLayer=' + (dbgLayer ? 'SRH_DEBUG' : 'document') +
        ' | rectTop=' + rectTop + ' rectLeft=' + rectLeft + ' w=' + w + ' h=' + h +
        ' | errors=' + errs.join(' | ');
    }

    try {rect.name = 'SRH_DEBUG_GRAD_10_90';} catch(_eDbgGr3) {_err('nameRectangle', _eDbgGr3);}
    try {rect.stroked = false;} catch(_eDbgGr4) {_err('setStrokedFalse', _eDbgGr4);}
    try {rect.filled = true;} catch(_eDbgGr5) {_err('setFilledTrue', _eDbgGr5);}

    var white = new RGBColor();
    white.red = 255; white.green = 255; white.blue = 255;
    var black = new RGBColor();
    black.red = 0; black.green = 0; black.blue = 0;

    var g = null;
    try {g = doc.gradients.add();} catch(_eDbgGr6) {g = null; _err('createGradient', _eDbgGr6);}
    if(!g) return 'Failed to create gradient. errors=' + errs.join(' | ');
    try {g.type = GradientType.LINEAR;} catch(_eDbgGr7) {_err('setGradientType', _eDbgGr7);}
    try {g.name = 'SRH_DEBUG_10_90_' + new Date().getTime();} catch(_eDbgGr8) {_err('setGradientName', _eDbgGr8);}

    try {
      while(g.gradientStops.length > 2) {
        try {g.gradientStops[g.gradientStops.length - 1].remove();} catch(_eDbgGr9) {_err('trimStops', _eDbgGr9); break;}
      }
      while(g.gradientStops.length < 2) g.gradientStops.add();
    } catch(_eDbgGr10) {_err('normalizeStops', _eDbgGr10);}
    try {g.gradientStops[0].color = white;} catch(_eDbgGr11) {_err('setStop0Color', _eDbgGr11);}
    try {g.gradientStops[1].color = black;} catch(_eDbgGr12) {_err('setStop1Color', _eDbgGr12);}
    try {g.gradientStops[0].rampPoint = 10;} catch(_eDbgGr13) {_err('setStop0Ramp10', _eDbgGr13);}
    try {g.gradientStops[1].rampPoint = 90;} catch(_eDbgGr14) {_err('setStop1Ramp90', _eDbgGr14);}

    var gc = null;
    try {gc = new GradientColor();} catch(_eDbgGr15) {gc = null; _err('newGradientColor', _eDbgGr15);}
    if(!gc) return 'Failed to create GradientColor. errors=' + errs.join(' | ');
    try {gc.gradient = g;} catch(_eDbgGr16) {_err('assignGradientToColor', _eDbgGr16);}
    try {gc.angle = 0;} catch(_eDbgGr17) {_err('setAngle', _eDbgGr17);}
    try {gc.origin = [rectLeft, cy];} catch(_eDbgGr18) {_err('setOrigin', _eDbgGr18);}
    try {gc.length = w;} catch(_eDbgGr19) {_err('setLength', _eDbgGr19);}

    try {rect.fillColor = gc;} catch(_eDbgGr20) {
      _err('assignFillColor', _eDbgGr20);
      return 'Failed to assign gradient fill. errors=' + errs.join(' | ');
    }

    try {rect.selected = true;} catch(_eDbgSel0) {_err('selectRect', _eDbgSel0);}
    try {app.redraw();} catch(_eDbgRedraw0) {_err('redraw', _eDbgRedraw0);}

    var readbackStops = [];
    var readbackName = '';
    var readbackCount = 0;
    try {
      var applied = rect.fillColor;
      if(applied && applied.typename === 'GradientColor' && applied.gradient && applied.gradient.gradientStops) {
        try {readbackName = String(applied.gradient.name || '');} catch(_eDbgGr23) {readbackName = '';}
        try {readbackCount = Number(applied.gradient.gradientStops.length || 0);} catch(_eDbgGr24) {readbackCount = 0;}
        for(var i = 0; i < applied.gradient.gradientStops.length; i++) {
          var rp = 0;
          try {rp = Number(applied.gradient.gradientStops[i].rampPoint || 0);} catch(_eDbgGr21) {rp = 0;}
          readbackStops.push(Math.round(rp * 1000) / 1000);
        }
      }
    } catch(_eDbgGr22) {_err('readback', _eDbgGr22);}

    var rb = null;
    try {rb = rect.visibleBounds;} catch(_eDbgRb0) {rb = null; _err('readBounds', _eDbgRb0);}

    return 'Debug rect created: ' + (rect.name || 'unnamed') +
      ' | layer=' + (function() {try {return String((rect.layer && rect.layer.name) || '');} catch(_eDbgLay4) {return '';} })() +
      ' | center=[' + (Math.round(cx * 1000) / 1000) + ', ' + (Math.round(cy * 1000) / 1000) + ']' +
      ' | bounds=' + (rb && rb.length === 4 ? ('[' + (Math.round(rb[0] * 1000) / 1000) + ',' + (Math.round(rb[1] * 1000) / 1000) + ',' + (Math.round(rb[2] * 1000) / 1000) + ',' + (Math.round(rb[3] * 1000) / 1000) + ']') : '[n/a]') +
      ' | requestedStops=[10,90]' +
      ' | readbackStops=[' + readbackStops.join(', ') + ']' +
      ' | readbackStopCount=' + readbackCount +
      ' | gradient=' + readbackName +
      (errs.length ? (' | warnings=' + errs.join(' | ')) : '');
  } finally {
    try {app.endUndoGroup();} catch(_eDbgUg1) { }
  }
}
try {if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_debugCreateGradientRect1090 = signarama_helper_debugCreateGradientRect1090;} catch(_eDbgExp0) { }
try {this.signarama_helper_debugCreateGradientRect1090 = signarama_helper_debugCreateGradientRect1090;} catch(_eDbgExp1) { }

/**
 * Debug helper: for selected items, force gradient stop ramp points into
 * a 10..90 range while preserving stop count/colors.
 */
function signarama_helper_debugSetSelectedGradientStops1090() {
  if(!app.documents.length) return 'No open document.';
  var doc = app.activeDocument;
  if(!doc.selection || !doc.selection.length) return 'No selection.';

  function _isGradientColor(gc) {
    try {return !!(gc && gc.typename === 'GradientColor' && gc.gradient && gc.gradient.gradientStops);} catch(_eSg0) {return false;}
  }
  function _readStops(grad) {
    var out = [];
    if(!grad) return out;
    var stops = null;
    try {stops = grad.gradientStops;} catch(_eSg1) {stops = null;}
    if(!stops || !stops.length) return out;
    for(var i = 0; i < stops.length; i++) {
      out.push({
        color: (function(s) {try {return s.color;} catch(_eSg2) {return null;} })(stops[i]),
        midPoint: (function(s) {try {return Number(s.midPoint || 50);} catch(_eSg3) {return 50;} })(stops[i]),
        opacity: (function(s) {try {return Number(s.opacity || 100);} catch(_eSg4) {return 100;} })(stops[i]),
        rampPoint: (function(s) {try {return Number(s.rampPoint || 0);} catch(_eSg5) {return 0;} })(stops[i])
      });
    }
    return out;
  }
  function _build1090Stops(srcStops) {
    var out = [];
    if(!srcStops || !srcStops.length) return out;
    if(srcStops.length === 1) {
      out.push({color: srcStops[0].color, midPoint: srcStops[0].midPoint, opacity: srcStops[0].opacity, rampPoint: 50});
      return out;
    }
    for(var i = 0; i < srcStops.length; i++) {
      var t = i / (srcStops.length - 1);
      out.push({
        color: srcStops[i].color,
        midPoint: srcStops[i].midPoint,
        opacity: srcStops[i].opacity,
        rampPoint: 10 + (80 * t)
      });
    }
    return out;
  }
  function _rebuildGradientColor(gc) {
    var res = {ok: false, color: null, stopCount: 0, stopsText: ''};
    if(!_isGradientColor(gc)) return res;
    var srcGrad = null;
    try {srcGrad = gc.gradient;} catch(_eSg6) {srcGrad = null;}
    if(!srcGrad) return res;
    var srcStops = _readStops(srcGrad);
    if(!srcStops.length) return res;
    var newStops = _build1090Stops(srcStops);
    if(!newStops.length) return res;

    var newGrad = null;
    try {newGrad = doc.gradients.add();} catch(_eSg7) {newGrad = null;}
    if(!newGrad) return res;
    try {newGrad.type = srcGrad.type;} catch(_eSg8) { }
    try {newGrad.name = 'SRH_SET_10_90_' + new Date().getTime();} catch(_eSg9) { }
    try {
      while(newGrad.gradientStops.length < newStops.length) newGrad.gradientStops.add();
      while(newGrad.gradientStops.length > newStops.length && newGrad.gradientStops.length > 2) {
        try {newGrad.gradientStops[newGrad.gradientStops.length - 1].remove();} catch(_eSg10) {break;}
      }
    } catch(_eSg11) { }
    for(var i = 0; i < newStops.length; i++) {
      var ds = null;
      try {ds = newGrad.gradientStops[i];} catch(_eSg12) {ds = null;}
      if(!ds) continue;
      try {ds.color = newStops[i].color;} catch(_eSg13) { }
      try {ds.midPoint = newStops[i].midPoint;} catch(_eSg14) { }
      try {ds.opacity = newStops[i].opacity;} catch(_eSg15) { }
      try {ds.rampPoint = Number(newStops[i].rampPoint);} catch(_eSg16) { }
    }

    var outGc = null;
    try {outGc = new GradientColor();} catch(_eSg17) {outGc = null;}
    if(!outGc) return res;
    try {outGc.gradient = newGrad;} catch(_eSg18) {return res;}
    try {outGc.angle = gc.angle;} catch(_eSg19) { }
    try {if(gc.origin && gc.origin.length >= 2) outGc.origin = [Number(gc.origin[0]), Number(gc.origin[1])];} catch(_eSg20) { }
    try {outGc.length = gc.length;} catch(_eSg21) { }
    try {outGc.hiliteAngle = gc.hiliteAngle;} catch(_eSg22) { }
    try {outGc.hiliteLength = gc.hiliteLength;} catch(_eSg23) { }
    try {outGc.matrix = gc.matrix;} catch(_eSg24) { }

    var txt = [];
    try {
      var gs = newGrad.gradientStops;
      for(var k = 0; k < gs.length; k++) txt.push(String(Math.round(Number(gs[k].rampPoint || 0) * 1000) / 1000));
    } catch(_eSg25) { }
    res.ok = true;
    res.color = outGc;
    res.stopCount = newStops.length;
    res.stopsText = '[' + txt.join(', ') + ']';
    return res;
  }

  function _setFillGradient(item) {
    var gc = null;
    try {gc = (item.filled ? item.fillColor : null);} catch(_eSg26) {gc = null;}
    if(!_isGradientColor(gc)) return 0;
    var rebuilt = _rebuildGradientColor(gc);
    if(!rebuilt.ok || !rebuilt.color) return 0;
    try {item.filled = true;} catch(_eSg27) { }
    try {item.fillColor = rebuilt.color; return 1;} catch(_eSg28) {return 0;}
  }
  function _setStrokeGradient(item) {
    var gc = null;
    try {gc = (item.stroked ? item.strokeColor : null);} catch(_eSg29) {gc = null;}
    if(!_isGradientColor(gc)) return 0;
    var rebuilt = _rebuildGradientColor(gc);
    if(!rebuilt.ok || !rebuilt.color) return 0;
    try {item.stroked = true;} catch(_eSg30) { }
    try {item.strokeColor = rebuilt.color; return 1;} catch(_eSg31) {return 0;}
  }

  var stack = [];
  for(var s = 0; s < doc.selection.length; s++) stack.push(doc.selection[s]);
  var seen = [];
  var touchedItems = 0;
  var changedGradients = 0;
  while(stack.length) {
    var cur = stack.pop();
    if(!cur) continue;
    var dupe = false;
    for(var si = 0; si < seen.length; si++) {if(seen[si] === cur) {dupe = true; break;} }
    if(dupe) continue;
    seen.push(cur);

    var changed = 0;
    changed += _setFillGradient(cur);
    changed += _setStrokeGradient(cur);
    if(changed > 0) {
      touchedItems++;
      changedGradients += changed;
    }
    try {
      if(cur.typename === 'CompoundPathItem' && cur.pathItems && cur.pathItems.length) {
        for(var cp = 0; cp < cur.pathItems.length; cp++) stack.push(cur.pathItems[cp]);
      } else if(cur.pageItems && cur.pageItems.length) {
        for(var i = 0; i < cur.pageItems.length; i++) stack.push(cur.pageItems[i]);
      }
    } catch(_eSg32) { }
  }
  try {app.redraw();} catch(_eSg33) { }
  return 'Set gradient stops to 10/90. Items touched: ' + touchedItems + ', gradients updated: ' + changedGradients + '.';
}
try {if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_debugSetSelectedGradientStops1090 = signarama_helper_debugSetSelectedGradientStops1090;} catch(_eDbgExp2) { }
try {this.signarama_helper_debugSetSelectedGradientStops1090 = signarama_helper_debugSetSelectedGradientStops1090;} catch(_eDbgExp3) { }

function _srh_fmtBounds(b) {
  function r(n) {return Math.round(n * 1000) / 1000;}
  return '[L ' + r(b.left) + ', T ' + r(b.top) + ', R ' + r(b.right) + ', B ' + r(b.bottom) + ']';
}

function _srh_hasClippingAncestor(item) {
  var p = null;
  try {p = item.parent;} catch(_ePa0) {p = null;}
  while(p) {
    try {
      if(p.typename === "GroupItem" && p.clipped) return true;
    } catch(_ePa1) { }
    try {p = p.parent;} catch(_ePa2) {p = null;}
  }
  return false;
}

function _srh_hasHiddenOrLockedAncestor(item) {
  var p = null;
  try {p = item.parent;} catch(_ePaL0) {p = null;}
  while(p) {
    try {if(p.hidden) return true;} catch(_ePaL1) { }
    try {if(p.locked) return true;} catch(_ePaL2) { }
    try {
      if(p.layer && (p.layer.locked || !p.layer.visible)) return true;
    } catch(_ePaL3) { }
    try {p = p.parent;} catch(_ePaL4) {p = null;}
  }
  return false;
}

function _srh_getClippingPathBounds(groupItem) {
  if(!groupItem) return null;
  try {
    if(!(groupItem.typename === "GroupItem" && groupItem.clipped)) return null;
  } catch(_eGp0) {return null;}
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

function _srh_sendLayerToBack(layer) {
  try {layer.zOrder(ZOrderMethod.SENDTOBACK);} catch(e) { }
}

function _srh_setBleedLayerOrder(cutLayer, originalLayer, bleedLayer) {
  if(bleedLayer) _srh_sendLayerToBack(bleedLayer);
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
    try {p = it.parent;} catch(_eP0) {p = null;}
    while(p) {
      var isClippedParent = false;
      try {isClippedParent = (p.typename === "GroupItem" && p.clipped);} catch(_eP1) {isClippedParent = false;}
      if(isClippedParent) {
        it = p;
        break;
      }
      try {p = p.parent;} catch(_eP2) {p = null;}
    }

    var isClippedGroup = false;
    try {isClippedGroup = (it.typename === "GroupItem" && it.clipped);} catch(_eP3) {isClippedGroup = false;}
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
 * Duplicates selection, then scales to fit 297×210mm (A4 landscape)
 * by either dimension (scales down OR up as needed).
 * Keeps artwork live (no outlining). Workflow: scale geometry first, then scale strokes.
 */
function signarama_helper_duplicateOutlineScaleA4(jsonStr) {
  if(!app.documents.length) return 'No open document.';
  var doc = app.activeDocument;
  if(!doc.selection || doc.selection.length === 0) return 'No selection. Select one or more items.';
  var args = {};
  try {args = jsonStr ? JSON.parse(String(jsonStr)) : {};} catch(_eArgs) {args = {};}
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

  // Scale to fit within 297x210mm.
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

  if(Math.abs(factor - 1) > 0.001) {
    var pct = factor * 100;
    try {grp.resize(pct, pct, true, true, true, true, 100, Transformation.CENTER);} catch(_e6) { }
    var strokeCount = _srh_scaleAllStrokesInContainer(grp, factor);
    if(shouldRasterize) {
      try {
        // Rasterize each selected item using visible bounds.
        doc.selection = null;
        var gi = null;
        try {gi = grp.pageItems;} catch(_eSelA) {gi = null;}
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
          try {bounds = item.visibleBounds;} catch(_eRb0) {bounds = null;}
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

  return 'Duplicate created. No scaling needed (already at 297×210mm fit).';
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
    try {return String(item.typename || 'unknown');} catch(_eDbg1) {return 'unknown';}
  }
  function _dbgName(item) {
    if(!item) return '';
    try {return String(item.name || '');} catch(_eDbg2) {return '';}
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
    } catch(_eGi0) {return null;}
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
      try {isClippedGroup = (p.typename === "GroupItem" && p.clipped);} catch(_eCg0) {isClippedGroup = false;}
      if(isClippedGroup) return p;
      try {p = p.parent;} catch(_eCg1) {p = null;}
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
    } catch(_eSpp0) {return false;}
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
      if(processedItems[si] === sourceItem) {seen = true; break;}
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
    try {prevStrokeW = target.strokeWidth;} catch(_eSw0) {prevStrokeW = null;}

    var resized = false;
    var targetType = "";
    try {targetType = target.typename || "";} catch(_eTy0) {targetType = "";}
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
      } catch(_eGb0) {resized = false;}
      if(!resized) {
        _dbg('sel[' + i + '] geometricBounds set FAILED: ' + _eGb0);
        try {
          target.visibleBounds = desiredBounds;
          resized = true;
          _dbg('sel[' + i + '] visibleBounds set OK');
        } catch(_eGb1) {resized = false;}
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
        } catch(_e2b) {resized = false;}
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
  try {dbgText = _dbgLines.join('\n');} catch(_eDbgJoin) {dbgText = '';}
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
 * Args: JSON string {"offsetMm":number, "createCutline":boolean, "outlineText":boolean, "outlineStroke":boolean, "autoWeld":boolean, "autoCloseOpenPaths":boolean}
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
  var autoCloseOpenPaths = (typeof args.autoCloseOpenPaths === 'undefined') ? true : !!args.autoCloseOpenPaths;

  var offsetPt = _srh_mm2ptDoc(offsetMm);
  if(!(offsetPt > 0)) {return "Offset amount must be > 0 mm.";}
  var forceBleedSolidRedTest = false;
  var forceBleedStopTest2080 = false;
  var forceSolidPrimeBeforeGradientAssign = false;
  var pathBleedDebugVerbose = true;
  var drawGradientStopDebugLines = false;
  var _pathBleedLogLines = [];
  function _pathBleedWithDbg(msg) {
    if(!_pathBleedLogLines || !_pathBleedLogLines.length) return msg;
    return msg + '\n' + _pathBleedLogLines.join('\n');
  }

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
            _expandSelection();
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
    var targets = [];
    var stack = [];
    try {stack.push(container);} catch(_eOs0) { }
    while(stack.length) {
      var cur = stack.pop();
      if(!cur) continue;
      var tn = '';
      try {tn = String(cur.typename || '');} catch(_eOs1) {tn = '';}
      if(tn === 'PathItem') {
        var hasStroke = false;
        try {hasStroke = !!cur.stroked && Number(cur.strokeWidth || 0) > 0;} catch(_eOs2) {hasStroke = false;}
        if(hasStroke) targets.push(cur);
        continue;
      }
      if(tn === 'CompoundPathItem') {
        var compoundHasStroke = false;
        try {
          if(cur.pathItems && cur.pathItems.length) {
            for(var cpi = 0; cpi < cur.pathItems.length; cpi++) {
              var cp = cur.pathItems[cpi];
              if(cp && cp.stroked && Number(cp.strokeWidth || 0) > 0) {
                compoundHasStroke = true;
                break;
              }
            }
          }
        } catch(_eOs3) {compoundHasStroke = false;}
        if(compoundHasStroke) {
          targets.push(cur);
          continue;
        }
      }
      try {
        if(cur.pageItems && cur.pageItems.length) {
          for(var i = 0; i < cur.pageItems.length; i++) stack.push(cur.pageItems[i]);
        }
      } catch(_eOs4) { }
    }
    _pathBleedDebug('outline-stroke | strategy=expand-stroked-targets | count=' + targets.length);
    if(!targets.length) return;
    _runOnSelection(targets, function() {
      try {
        app.executeMenuCommand('Outline Stroke');
        _pathBleedDebug('outline-stroke | command=Outline Stroke | result=OK');
      } catch(_eOsCmd1) {
        _pathBleedDebug('outline-stroke | command=Outline Stroke | result=FAILED');
      }
      _expandSelection();
    });
  }

  function _getMaxStrokeWidthInContainer(container) {
    if(!container) return 0;
    var maxStroke = 0;
    var stack = [];
    try {stack.push(container);} catch(_eGms0) { }
    while(stack.length) {
      var cur = stack.pop();
      if(!cur) continue;
      var tn = '';
      try {tn = String(cur.typename || '');} catch(_eGms1) {tn = '';}
      if(tn === 'PathItem') {
        try {
          if(cur.stroked) {
            var sw = Number(cur.strokeWidth || 0);
            if(sw > maxStroke) maxStroke = sw;
          }
        } catch(_eGms2) { }
      }
      try {
        if(cur.pageItems && cur.pageItems.length) {
          for(var i = 0; i < cur.pageItems.length; i++) stack.push(cur.pageItems[i]);
        }
      } catch(_eGms3) { }
    }
    return maxStroke;
  }

  function _collectNativePathRoots(container) {
    var out = [];
    if(!container) return out;
    var stack = [];
    try {stack.push(container);} catch(_eCnpr0) { }
    while(stack.length) {
      var cur = stack.pop();
      if(!cur) continue;
      var tn = '';
      try {tn = String(cur.typename || '');} catch(_eCnpr1) {tn = ''; }
      if(tn === 'CompoundPathItem') {
        out.push(cur);
        continue;
      }
      if(tn === 'PathItem') {
        var hasCompound = false;
        try {
          var p = cur.parent;
          var guard = 0;
          while(p && guard < 100) {
            guard++;
            if(String(p.typename || '') === 'CompoundPathItem') {hasCompound = true; break;}
            p = p.parent;
          }
        } catch(_eCnpr2) {hasCompound = false; }
        if(!hasCompound) out.push(cur);
      }
      try {
        if(cur.pageItems && cur.pageItems.length) {
          for(var i = 0; i < cur.pageItems.length; i++) stack.push(cur.pageItems[i]);
        }
      } catch(_eCnpr3) { }
    }
    return out;
  }

  function _getOffsetJoinInfoFromItem(item) {
    var info = {joinType: 2, joinName: 'miter', miterLimit: 180};
    if(!item) return info;
    try {
      var sj = item.strokeJoin;
      var sjText = String(sj || '').toLowerCase();
      if(sjText.indexOf('round') >= 0) {
        info.joinType = 1;
        info.joinName = 'round';
      } else if(sjText.indexOf('bevel') >= 0) {
        info.joinType = 0;
        info.joinName = 'bevel';
      } else {
        info.joinType = 2;
        info.joinName = 'miter';
      }
    } catch(_eGoji0) { }
    try {
      var ml = Number(item.strokeMiterLimit || 0);
      if(ml > 0) info.miterLimit = ml;
    } catch(_eGoji1) { }
    return info;
  }

  function _replaceContainerWithItems(container, items, label) {
    if(!container || !items || !items.length) return 0;
    var moved = 0;
    try {
      var existing = [];
      for(var i = 0; i < container.pageItems.length; i++) existing.push(container.pageItems[i]);
      for(var e = 0; e < existing.length; e++) {
        try {existing[e].remove();} catch(_eRci0) { }
      }
    } catch(_eRci1) { }
    for(var j = 0; j < items.length; j++) {
      var it = items[j];
      if(!it || it === container) continue;
      try {
        it.move(container, ElementPlacement.PLACEATEND);
        moved++;
      } catch(_eRci2) { }
    }
    _pathBleedDebug('replace container ' + (label || '') + ' | moved=' + moved);
    return moved;
  }

  function _closeOpenPathsInContainer(container, label) {
    if(!container) return 0;
    var changed = 0;
    var stack = [];
    try {stack.push(container);} catch(_eCp0) { }
    while(stack.length) {
      var cur = stack.pop();
      if(!cur) continue;
      var tn = '';
      try {tn = String(cur.typename || '');} catch(_eCp1) {tn = '';}
      if(tn === 'PathItem') {
        var isClosed = true;
        try {isClosed = !!cur.closed;} catch(_eCp2) {isClosed = true;}
        if(!isClosed) {
          try {cur.closed = true; changed++;} catch(_eCp3) { }
        }
      }
      try {
        if(cur.pageItems && cur.pageItems.length) {
          for(var i = 0; i < cur.pageItems.length; i++) stack.push(cur.pageItems[i]);
        }
      } catch(_eCp4) { }
    }
    if(changed > 0) _pathBleedDebug('auto-close open paths | ' + (label || 'container') + ' | closed=' + changed);
    return changed;
  }

  // Preserve inner subpaths for compound paths (do not delete holes).
  // Retained as a no-op for compatibility with prior call sites/log parsing.
  function _removeNegativeContoursInCompounds(container) {
    return 0;
  }

  function _applyOffsetToContainer(container, ofst, opts) {
    if(!container) return false;
    opts = opts || {};
    var offsetJoinType = (typeof opts.joinType === 'number') ? opts.joinType : 0;
    var offsetMiterLimit = (typeof opts.miterLimit === 'number' && opts.miterLimit > 0) ? opts.miterLimit : 180;
    _pathBleedDebug('offset START | ofst=' + _fmtNum3(ofst) + ' | container=' + _itemRef(container));
    function _safeBounds(it) {
      var b0 = null;
      try {b0 = it.visibleBounds;} catch(_eOb0) {b0 = null;}
      if(!b0 || b0.length !== 4) {
        try {b0 = it.geometricBounds;} catch(_eOb1) {b0 = null;}
      }
      return (b0 && b0.length === 4) ? [Number(b0[0]), Number(b0[1]), Number(b0[2]), Number(b0[3])] : null;
    }
    function _boundsChanged(a, b) {
      if(!a || !b || a.length !== 4 || b.length !== 4) return false;
      var eps = 0.25;
      for(var iB = 0; iB < 4; iB++) {
        if(Math.abs(Number(a[iB]) - Number(b[iB])) > eps) return true;
      }
      return false;
    }
    function _expectedOffsetBounds(b, d) {
      if(!b || b.length !== 4) return null;
      var dd = Number(d || 0);
      return [Number(b[0]) - dd, Number(b[1]) + dd, Number(b[2]) + dd, Number(b[3]) - dd];
    }
    function _boundsDelta(a, b) {
      if(!a || !b || a.length !== 4 || b.length !== 4) return '[n/a]';
      return '[' +
        _fmtNum3(Number(b[0]) - Number(a[0])) + ',' +
        _fmtNum3(Number(b[1]) - Number(a[1])) + ',' +
        _fmtNum3(Number(b[2]) - Number(a[2])) + ',' +
        _fmtNum3(Number(b[3]) - Number(a[3])) + ']';
    }
    function _boundsErr(expected, actual) {
      if(!expected || !actual || expected.length !== 4 || actual.length !== 4) return '[n/a]';
      return '[' +
        _fmtNum3(Number(actual[0]) - Number(expected[0])) + ',' +
        _fmtNum3(Number(actual[1]) - Number(expected[1])) + ',' +
        _fmtNum3(Number(actual[2]) - Number(expected[2])) + ',' +
        _fmtNum3(Number(actual[3]) - Number(expected[3])) + ']';
    }
    function _hasCompoundAncestorLocal(it) {
      var p = null;
      var guard = 0;
      try {p = it.parent;} catch(_eHcaL0) {p = null;}
      while(p && guard < 100) {
        guard++;
        try {if(String(p.typename || '') === 'CompoundPathItem') return true;} catch(_eHcaL1) { }
        try {p = p.parent;} catch(_eHcaL2) {p = null;}
      }
      return false;
    }
    function _collectOffsetTargets(root) {
      var out = [];
      if(!root) return out;
      var st = [];
      try {st.push(root);} catch(_eCot0) { }
      while(st.length) {
        var cur = st.pop();
        if(!cur) continue;
        var tn = '';
        try {tn = String(cur.typename || '');} catch(_eCot1) {tn = '';}
        if(tn === 'CompoundPathItem') {
          out.push(cur);
          continue;
        }
        if(tn === 'PathItem' && !_hasCompoundAncestorLocal(cur)) {
          out.push(cur);
        }
        _pushGradientChildren(cur, st);
      }
      return out;
    }
    function _hasNonNativeDescendant(root) {
      if(!root) return false;
      var stN = [];
      try {stN.push(root);} catch(_eNnD0) { }
      while(stN.length) {
        var itN = stN.pop();
        if(!itN) continue;
        try {if(String(itN.typename || '') === 'NonNativeItem') return true;} catch(_eNnD1) { }
        _pushGradientChildren(itN, stN);
      }
      return false;
    }
    function _logDescendantsBounds(label, root) {
      if(!root) return;
      var st2 = [];
      var idx = 0;
      try {st2.push(root);} catch(_eObL0) { }
      while(st2.length && idx < 60) {
        var it2 = st2.pop();
        if(!it2) continue;
        idx++;
        var t2 = '';
        try {t2 = String(it2.typename || '');} catch(_eObL1) {t2 = '';}
        if(t2 === 'GroupItem' || t2 === 'CompoundPathItem' || t2 === 'PathItem' || t2 === 'NonNativeItem') {
          _pathBleedDebug('offset ' + label + ' | ' + _itemRef(it2));
        }
        try {
          if(it2.pageItems && it2.pageItems.length) {
            for(var pp = 0; pp < it2.pageItems.length; pp++) st2.push(it2.pageItems[pp]);
          }
        } catch(_eObL2) { }
      }
    }
    var beforeBounds = _safeBounds(container);
    var expectedBounds = _expectedOffsetBounds(beforeBounds, ofst);
    var nativeTargets = _collectOffsetTargets(container);
    var hasNonNative = _hasNonNativeDescendant(container);
    var skipExpandForNonNative = hasNonNative;
    _pathBleedDebug(
      'offset mode | strategy=live-effect | targets=' + nativeTargets.length +
      ' | hasNonNative=' + (hasNonNative ? 'yes' : 'no') +
      ' | skipExpand=' + (skipExpandForNonNative ? 'yes' : 'no')
    );
    _pathBleedDebug(
      'offset expect | before=' + _fmtBounds4(beforeBounds) +
      ' | expected=' + _fmtBounds4(expectedBounds) +
      ' | ofst=' + _fmtNum3(ofst)
    );
    _logDescendantsBounds('desc-before', container);

    var fx1 = '<LiveEffect name="Adobe Offset Path"><Dict data="R ofst ' + Number(ofst) + ' I jntp ' + Number(offsetJoinType) + ' R mlim ' + Number(offsetMiterLimit) + '"/></LiveEffect>';
    var fx2 = '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim ' + Number(offsetMiterLimit) + ' R ofst ' + Number(ofst) + ' I jntp ' + Number(offsetJoinType) + '"/></LiveEffect>';
    var applied = false;
    try {
      container.applyEffect(fx1);
      applied = true;
      _pathBleedDebug('offset applyEffect fx1 on container OK');
    } catch(_eFxA) {
      try {
        container.applyEffect(fx2);
        applied = true;
        _pathBleedDebug('offset applyEffect fx2 on container OK');
      } catch(_eFxB) { }
    }
    if(applied) {
      if(skipExpandForNonNative) {
        _pathBleedDebug('offset expand selection SKIP | reason=non-native-content');
      } else {
        _runOnSelection(container, function() {_expandSelection();});
        _pathBleedDebug('offset expand selection OK | target=container');
      }
    }

    var afterContainerBounds = _safeBounds(container);
    var changedByContainer = _boundsChanged(beforeBounds, afterContainerBounds);
    var changedAny = false;
    if(changedByContainer) changedAny = true;
    _pathBleedDebug(
      'offset container-result | before=' + _fmtBounds4(beforeBounds) +
      ' | after=' + _fmtBounds4(afterContainerBounds) +
      ' | delta=' + _boundsDelta(beforeBounds, afterContainerBounds) +
      ' | errVsExpected=' + _boundsErr(expectedBounds, afterContainerBounds) +
      ' | changed=' + (changedByContainer ? 'yes' : 'no')
    );
    _logDescendantsBounds('desc-after-container', container);

    var forcedApplied = 0;
    if(!changedByContainer && nativeTargets.length) {
      for(var nt = 0; nt < nativeTargets.length; nt++) {
        var t = nativeTargets[nt];
        if(!t) continue;
        var tb = _safeBounds(t);
        var used = '';
        try {
          t.applyEffect(fx1);
          used = 'fx1';
          forcedApplied++;
        } catch(_eFxT0) {
          try {
            t.applyEffect(fx2);
            used = 'fx2';
            forcedApplied++;
          } catch(_eFxT1) { }
        }
        var ta = _safeBounds(t);
        var itemChanged = _boundsChanged(tb, ta);
        if(itemChanged) changedAny = true;
        _pathBleedDebug(
          'offset target-result | idx=' + nt +
          ' | item=' + _itemRef(t) +
          ' | effect=' + (used || 'none') +
          ' | before=' + _fmtBounds4(tb) +
          ' | after=' + _fmtBounds4(ta) +
          ' | delta=' + _boundsDelta(tb, ta) +
          ' | changed=' + (itemChanged ? 'yes' : 'no')
        );
      }
      if(forcedApplied > 0) {
        if(skipExpandForNonNative) {
          _pathBleedDebug('offset forced expand SKIP | reason=non-native-content');
        } else {
          _runOnSelection(nativeTargets, function() {_expandSelection();});
          _pathBleedDebug('offset forced expand OK | target=native-targets');
        }
      }
    }

    var afterBounds = _safeBounds(container);
    if(_boundsChanged(beforeBounds, afterBounds)) changedAny = true;
    _pathBleedDebug(
      'offset summary | before=' + _fmtBounds4(beforeBounds) +
      ' | after=' + _fmtBounds4(afterBounds) +
      ' | delta=' + _boundsDelta(beforeBounds, afterBounds) +
      ' | errVsExpected=' + _boundsErr(expectedBounds, afterBounds) +
      ' | applied=' + (applied ? 'yes' : 'no') +
      ' | forcedApplied=' + forcedApplied +
      ' | geometryChanged=' + (changedAny ? 'yes' : 'no')
    );
    _logDescendantsBounds('desc-after-offset', container);
    _pathBleedDebug('offset END | applied=' + ((applied || forcedApplied > 0) ? 'yes' : 'no') + ' | geometryChanged=' + (changedAny ? 'yes' : 'no'));
    return {applied: (applied || forcedApplied > 0), changed: changedAny};
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

  function _isLinearGradient(gc) {
    if(!_isGradientColor(gc)) return false;
    try {
      var gt = gc.gradient && gc.gradient.type;
      return String(gt) === String(GradientType.LINEAR) || String(gt).toLowerCase().indexOf('linear') >= 0;
    } catch(_eGl0) {return false;}
  }

  function _isRadialGradient(gc) {
    if(!_isGradientColor(gc)) return false;
    try {
      var gt = gc.gradient && gc.gradient.type;
      return String(gt) === String(GradientType.RADIAL) || String(gt).toLowerCase().indexOf('radial') >= 0;
    } catch(_eGl1) {return false;}
  }

  function _isFreeformGradient(gc) {
    if(!gc) return false;
    try {
      var tn = String(gc.typename || '').toLowerCase();
      if(tn.indexOf('freeform') >= 0) return true;
    } catch(_eGl2) { }
    try {
      if(gc.gradient && typeof gc.gradient.type !== 'undefined') {
        var gt = String(gc.gradient.type).toLowerCase();
        if(gt.indexOf('freeform') >= 0) return true;
      }
    } catch(_eGl3) { }
    return false;
  }

  function _isGradientLikeColor(gc) {
    return _isGradientColor(gc) || _isFreeformGradient(gc);
  }

  function _getCompensatableGradientMode(gc) {
    if(_isLinearGradient(gc)) return 'linear';
    if(_isRadialGradient(gc)) return 'radial';
    if(_isFreeformGradient(gc)) return 'freeform'; // Freeform follows radial compensation path.
    return '';
  }

  function _pathBleedDebug(msg) {
    if(!pathBleedDebugVerbose) return;
    try {$.writeln('[SRH][pathBleed] ' + msg);} catch(_eDbgPb0) { }
    try {_pathBleedLogLines.push('[SRH][pathBleed] ' + msg);} catch(_eDbgPb1) { }
  }
  var _gradientCloneCache = {};

  function _isGradientColor(gc) {
    return !!(gc && gc.typename === 'GradientColor');
  }

  function _fmtNum3(v) {
    var n = Number(v);
    if(!isFinite(n)) return 'NaN';
    return String(Math.round(n * 1000) / 1000);
  }

  function _fmtPoint(pt) {
    if(!pt || pt.length < 2) return '[n/a]';
    return '[' + _fmtNum3(pt[0]) + ', ' + _fmtNum3(pt[1]) + ']';
  }

  function _fmtBounds4(b) {
    if(!b || b.length !== 4) return '[n/a]';
    return '[' + _fmtNum3(b[0]) + ',' + _fmtNum3(b[1]) + ',' + _fmtNum3(b[2]) + ',' + _fmtNum3(b[3]) + ']';
  }

  function _itemRef(it) {
    if(!it) return 'null-item';
    var tn = 'Item';
    var nm = '';
    var lyr = '';
    var ptn = '';
    var pnm = '';
    var vb = null;
    try {tn = String(it.typename || 'Item');} catch(_eIr0) { }
    try {nm = String(it.name || '');} catch(_eIr1) {nm = '';}
    try {if(it.layer) lyr = String(it.layer.name || '');} catch(_eIr2) {lyr = '';}
    try {if(it.parent) ptn = String(it.parent.typename || '');} catch(_eIr3) {ptn = '';}
    try {if(it.parent) pnm = String(it.parent.name || '');} catch(_eIr4) {pnm = '';}
    try {vb = it.visibleBounds;} catch(_eIr5) {vb = null;}
    return tn +
      (nm ? (' "' + nm + '"') : '') +
      (lyr ? (' @layer=' + lyr) : '') +
      (ptn ? (' @parent=' + ptn + (pnm ? (':' + pnm) : '')) : '') +
      ' @vb=' + _fmtBounds4(vb);
  }

  function _getGradientItemLabel(it) {
    if(!it) return 'UnknownItem';
    var tn = 'Item';
    var nm = '';
    try {tn = String(it.typename || 'Item');} catch(_eLbl0) { }
    try {nm = String(it.name || '');} catch(_eLbl1) {nm = '';}
    if(nm) return tn + ' "' + nm + '"';
    return tn;
  }

  function _hasAncestorNamed(item, ancestorName) {
    if(!item || !ancestorName) return false;
    var p = null;
    try {p = item.parent;} catch(_eHan0) {p = null;}
    var guard = 0;
    while(p && guard < 200) {
      guard++;
      try {
        if(String(p.name || '') === ancestorName) return true;
      } catch(_eHan1) { }
      try {p = p.parent;} catch(_eHan2) {p = null;}
    }
    return false;
  }

  function _itemHasFreeformLike(item) {
    if(!item) return false;
    var stack = [];
    try {stack.push(item);} catch(_eFfLike0) { }
    while(stack.length) {
      var cur = stack.pop();
      if(!cur) continue;
      try {
        if(String(cur.typename || '') === 'NonNativeItem') return true;
      } catch(_eFfLike1) { }
      try {
        var gf = _getGradientColorForKind(cur, 'fill');
        if(_isFreeformGradient(gf)) return true;
      } catch(_eFfLike2) { }
      try {
        var gs = _getGradientColorForKind(cur, 'stroke');
        if(_isFreeformGradient(gs)) return true;
      } catch(_eFfLike3) { }
      _pushGradientChildren(cur, stack);
    }
    return false;
  }

  function _selectionHasFreeformLike(selItems) {
    if(!selItems || !selItems.length) return false;
    for(var i = 0; i < selItems.length; i++) {
      if(_itemHasFreeformLike(selItems[i])) return true;
    }
    return false;
  }

  function _countNonNativeItems(root) {
    if(!root) return 0;
    var count = 0;
    var stack = [];
    try {stack.push(root);} catch(_eCntNn0) { }
    while(stack.length) {
      var cur = stack.pop();
      if(!cur) continue;
      try {if(String(cur.typename || '') === 'NonNativeItem') count++;} catch(_eCntNn1) { }
      _pushGradientChildren(cur, stack);
    }
    return count;
  }

  function _scanForNonNative(root, label) {
    var count = 0;
    if(!root) return count;
    var stack = [];
    try {stack.push(root);} catch(_eScanNn0) { }
    while(stack.length) {
      var cur = stack.pop();
      if(!cur) continue;
      var tn = '';
      try {tn = String(cur.typename || '');} catch(_eScanNn1) {tn = '';}
      if(tn === 'NonNativeItem') {
        count++;
        _pathBleedDebug('non-native scan ' + (label || '') + ' | found | ' + _itemRef(cur));
      }
      _pushGradientChildren(cur, stack);
    }
    _pathBleedDebug('non-native scan ' + (label || '') + ' | count=' + count);
    return count;
  }

  function _describeGradientType(gc) {
    if(!gc) return 'none';
    var colorTn = '';
    var gradTn = '';
    var gradType = '';
    try {colorTn = String(gc.typename || '');} catch(_eDgt0) {colorTn = '';}
    try {if(gc.gradient) gradTn = String(gc.gradient.typename || '');} catch(_eDgt1) {gradTn = '';}
    try {if(gc.gradient && typeof gc.gradient.type !== 'undefined') gradType = String(gc.gradient.type);} catch(_eDgt2) {gradType = '';}
    var mode = 'unknown';
    if(_isLinearGradient(gc)) mode = 'linear';
    else if(_isRadialGradient(gc)) mode = 'radial';
    else if(_isFreeformGradient(gc)) mode = 'freeform';
    return 'mode=' + mode +
      ' colorType=' + (colorTn || 'n/a') +
      ' gradType=' + (gradType || 'n/a') +
      ' gradObjType=' + (gradTn || 'n/a');
  }

  function _logSelectionGradientKinds(selItems, label) {
    if(!pathBleedDebugVerbose) return;
    var stack = [];
    var seen = 0;
    if(selItems && selItems.length) {
      for(var i = 0; i < selItems.length; i++) stack.push(selItems[i]);
    }
    _pathBleedDebug('selection-gradient-kinds START ' + (label || 'selection') + ' | roots=' + (selItems ? selItems.length : 0));
    while(stack.length) {
      var cur = stack.pop();
      if(!cur) continue;
      seen++;
      var tn = '';
      try {tn = String(cur.typename || '');} catch(_eLsg0) {tn = '';}
      var fillDesc = 'none';
      var strokeDesc = 'none';
      try {fillDesc = _describeGradientType(_getGradientColorForKind(cur, 'fill'));} catch(_eLsg1) {fillDesc = 'error';}
      try {strokeDesc = _describeGradientType(_getGradientColorForKind(cur, 'stroke'));} catch(_eLsg2) {strokeDesc = 'error';}
      _pathBleedDebug('selection-gradient-kinds item | ' + _itemRef(cur) + ' | fill{' + fillDesc + '} | stroke{' + strokeDesc + '}');
      if(tn === 'NonNativeItem') {
        _pathBleedDebug('selection-gradient-kinds item | NonNativeItem detected');
      }
      _pushGradientChildren(cur, stack);
    }
    _pathBleedDebug('selection-gradient-kinds END ' + (label || 'selection') + ' | scanned=' + seen);
  }

  function _logGradientStateForItem(it, label) {
    if(!it) return;
    function _logOne(kind) {
      var gc = _getGradientColorForKind(it, kind);
      if(!_isGradientColor(gc)) return;
      var angle = 0;
      var len = 0;
      var origin = null;
      try {angle = Number(gc.angle || 0);} catch(_eLgs0) {angle = 0;}
      try {len = Number(gc.length || 0);} catch(_eLgs1) {len = 0;}
      try {if(gc.origin && gc.origin.length >= 2) origin = [Number(gc.origin[0]), Number(gc.origin[1])];} catch(_eLgs2) {origin = null;}
      _pathBleedDebug(
        'state ' + (label || 'item') +
        ' | ' + kind +
        ' | ' + _itemRef(it) +
        ' | gradient=' + (_getGradientName(gc) || '(unnamed)') +
        ' | stopCount=' + _getGradientStopCount(gc) +
        ' | stops=' + _formatStops(_readGradientStopRampPoints(gc)) +
        ' | angle=' + _fmtNum3(angle) +
        ' | origin=' + _fmtPoint(origin) +
        ' | length=' + _fmtNum3(len)
      );
    }
    _logOne('fill');
    _logOne('stroke');
  }

  function _logContainerState(container, label) {
    if(!pathBleedDebugVerbose || !container) return;
    var stack = [];
    var total = 0;
    var gradCount = 0;
    try {stack.push(container);} catch(_eLcs0) { }
    _pathBleedDebug('container-scan START ' + (label || '') + ' | root=' + _itemRef(container));
    while(stack.length) {
      var cur = stack.pop();
      if(!cur) continue;
      total++;
      var hadGrad = false;
      var gf = _getGradientColorForKind(cur, 'fill');
      var gs = _getGradientColorForKind(cur, 'stroke');
      if(_isGradientColor(gf) || _isGradientColor(gs)) {
        hadGrad = true;
        gradCount++;
        _logGradientStateForItem(cur, label || 'scan');
      } else {
        _pathBleedDebug('state ' + (label || 'scan') + ' | no-gradient | ' + _itemRef(cur));
      }
      _pushGradientChildren(cur, stack);
    }
    _pathBleedDebug('container-scan END ' + (label || '') + ' | totalItems=' + total + ' | gradientItems=' + gradCount);
  }

  function _auditStructure(container, label) {
    if(!container) return;
    var stack = [];
    var total = 0;
    var typeCounts = {};
    var snapshotNamed = 0;
    var nonNative = 0;
    try {stack.push(container);} catch(_eAud0) { }
    while(stack.length) {
      var cur = stack.pop();
      if(!cur) continue;
      total++;
      var tn = '';
      var nm = '';
      try {tn = String(cur.typename || 'Unknown');} catch(_eAud1) {tn = 'Unknown';}
      try {nm = String(cur.name || '');} catch(_eAud2) {nm = '';}
      try {typeCounts[tn] = (typeCounts[tn] || 0) + 1;} catch(_eAud3) { }
      if(tn === 'NonNativeItem') nonNative++;
      if(nm && nm.indexOf('SRH_ORIGINAL_SNAPSHOT') >= 0) snapshotNamed++;
      _pushGradientChildren(cur, stack);
    }
    var parts = [];
    for(var k in typeCounts) {
      if(typeCounts.hasOwnProperty(k)) parts.push(k + '=' + typeCounts[k]);
    }
    _pathBleedDebug(
      'structure audit ' + (label || '') +
      ' | total=' + total +
      ' | nonNative=' + nonNative +
      ' | snapshotNamed=' + snapshotNamed +
      ' | types{' + parts.join(', ') + '}'
    );

    // Show snapshot nesting chain details (why multiple SRH_ORIGINAL_SNAPSHOT groups exist).
    var sst = [];
    var sidx = 0;
    try {sst.push(container);} catch(_eAudS0) { }
    while(sst.length && sidx < 30) {
      var si = sst.pop();
      if(!si) continue;
      var snm = '';
      try {snm = String(si.name || '');} catch(_eAudS1) {snm = '';}
      if(snm && snm.indexOf('SRH_ORIGINAL_SNAPSHOT') >= 0) {
        _pathBleedDebug('structure snapshot ' + (label || '') + ' | ' + _itemRef(si));
      }
      sidx++;
      _pushGradientChildren(si, sst);
    }
  }

  function _collectPathMetrics(container) {
    var out = [];
    if(!container) return out;
    var stack = [];
    var idx = 0;
    try {stack.push(container);} catch(_ePm0) { }
    while(stack.length) {
      var cur = stack.pop();
      if(!cur) continue;
      var tn = '';
      try {tn = String(cur.typename || '');} catch(_ePm1) {tn = '';}
      if(tn === 'PathItem' || tn === 'CompoundPathItem') {
        var b = null;
        try {b = cur.visibleBounds;} catch(_ePm2) {b = null;}
        if(!b || b.length !== 4) {try {b = cur.geometricBounds;} catch(_ePm3) {b = null;} }
        if(b && b.length === 4) {
          var w = Number(b[2]) - Number(b[0]);
          var h = Number(b[1]) - Number(b[3]);
          out.push({
            key: idx + ':' + tn,
            ref: _itemRef(cur),
            b: [Number(b[0]), Number(b[1]), Number(b[2]), Number(b[3])],
            w: w,
            h: h
          });
          idx++;
        }
      }
      _pushGradientChildren(cur, stack);
    }
    return out;
  }

  function _logPathMetrics(label, rows) {
    rows = rows || [];
    _pathBleedDebug('path-metrics ' + (label || '') + ' | count=' + rows.length);
    for(var i = 0; i < rows.length && i < 40; i++) {
      var r = rows[i];
      _pathBleedDebug(
        'path-metrics ' + (label || '') +
        ' | ' + r.key +
        ' | w=' + _fmtNum3(r.w) +
        ' | h=' + _fmtNum3(r.h) +
        ' | b=' + _fmtBounds4(r.b) +
        ' | ' + r.ref
      );
    }
  }

  function _logPathMetricsDelta(beforeRows, afterRows, label) {
    beforeRows = beforeRows || [];
    afterRows = afterRows || [];
    var maxN = (beforeRows.length > afterRows.length) ? beforeRows.length : afterRows.length;
    _pathBleedDebug('path-metrics delta ' + (label || '') + ' | before=' + beforeRows.length + ' | after=' + afterRows.length);
    for(var i = 0; i < maxN && i < 40; i++) {
      var b = (i < beforeRows.length) ? beforeRows[i] : null;
      var a = (i < afterRows.length) ? afterRows[i] : null;
      if(!b || !a) {
        _pathBleedDebug('path-metrics delta ' + (label || '') + ' | idx=' + i + ' | paired=no');
        continue;
      }
      _pathBleedDebug(
        'path-metrics delta ' + (label || '') +
        ' | idx=' + i +
        ' | dw=' + _fmtNum3(a.w - b.w) +
        ' | dh=' + _fmtNum3(a.h - b.h) +
        ' | before=' + _fmtBounds4(b.b) +
        ' | after=' + _fmtBounds4(a.b)
      );
    }
  }

  function _pruneCutlineContainer(container) {
    if(!container) return 0;
    var removed = 0;

    function _walk(it) {
      if(!it) return;
      var tn = '';
      try {tn = String(it.typename || '');} catch(_ePr0) {tn = '';}
      if(tn === 'GroupItem') {
        var kids = [];
        try {
          if(it.pageItems && it.pageItems.length) {
            for(var i = 0; i < it.pageItems.length; i++) kids.push(it.pageItems[i]);
          }
        } catch(_ePr1) { }
        for(var k = 0; k < kids.length; k++) _walk(kids[k]);
        var leftCount = 0;
        try {leftCount = it.pageItems ? Number(it.pageItems.length || 0) : 0;} catch(_ePr2) {leftCount = 0;}
        if(leftCount === 0) {
          try {it.remove(); removed++;} catch(_ePr3) { }
        }
        return;
      }
      if(tn === 'PathItem' || tn === 'CompoundPathItem') {
        return;
      }
      try {
        _pathBleedDebug('cutline prune remove | ' + _itemRef(it));
        it.remove();
        removed++;
      } catch(_ePr4) { }
    }

    _walk(container);
    _pathBleedDebug('cutline prune summary | removed=' + removed);
    return removed;
  }

  // Rebuild a container so it contains only duplicated PathItems at its root.
  function _extractPathOnlySources(container, label) {
    if(!container) return 0;
    var layerRef = null;
    try {layerRef = container.layer;} catch(_eExtP0) {layerRef = null;}
    if(!layerRef) return 0;
    var tmp = null;
    try {
      tmp = layerRef.groupItems.add();
      tmp.name = 'SRH_PATH_ONLY_TMP__' + String(new Date().getTime());
    } catch(_eExtP1) {tmp = null;}
    if(!tmp) return 0;

    var copied = 0;
    var st = [];
    try {st.push(container);} catch(_eExtP2) { }
    while(st.length) {
      var cur = st.pop();
      if(!cur) continue;
      var tn = '';
      try {tn = String(cur.typename || '');} catch(_eExtP3) {tn = '';}
      if(tn === 'PathItem') {
        try {
          var dup = cur.duplicate(tmp, ElementPlacement.PLACEATEND);
          try {dup.clipping = false;} catch(_eExtP4) { }
          copied++;
        } catch(_eExtP5) { }
      }
      _pushGradientChildren(cur, st);
    }

    var removed = 0;
    try {
      var kids = [];
      for(var i = 0; i < container.pageItems.length; i++) kids.push(container.pageItems[i]);
      for(var k = 0; k < kids.length; k++) {
        try {kids[k].remove(); removed++;} catch(_eExtP6) { }
      }
    } catch(_eExtP7) { }

    var moved = 0;
    try {
      var tmpKids = [];
      for(var t = 0; t < tmp.pageItems.length; t++) tmpKids.push(tmp.pageItems[t]);
      for(var m = 0; m < tmpKids.length; m++) {
        try {tmpKids[m].move(container, ElementPlacement.PLACEATEND); moved++;} catch(_eExtP8) { }
      }
    } catch(_eExtP9) { }
    try {tmp.remove();} catch(_eExtP10) { }

    _pathBleedDebug(
      'path-only extract ' + (label || '') +
      ' | copied=' + copied +
      ' | removed=' + removed +
      ' | moved=' + moved
    );
    return moved;
  }

  function _extractNativePathSources(container, label) {
    if(!container) return 0;
    var layerRef = null;
    try {layerRef = container.layer;} catch(_eExt0) {layerRef = null;}
    if(!layerRef) return 0;
    var tmp = null;
    try {
      tmp = layerRef.groupItems.add();
      tmp.name = 'SRH_NATIVE_TMP__' + String(new Date().getTime());
    } catch(_eExt1) {tmp = null;}
    if(!tmp) return 0;

    function _hasCompoundAncestor(it) {
      var p = null;
      var guard = 0;
      try {p = it.parent;} catch(_eExt2) {p = null;}
      while(p && guard < 100) {
        guard++;
        try {if(String(p.typename || '') === 'CompoundPathItem') return true;} catch(_eExt3) { }
        try {p = p.parent;} catch(_eExt4) {p = null;}
      }
      return false;
    }

    var copied = 0;
    var st = [];
    try {st.push(container);} catch(_eExt5) { }
    while(st.length) {
      var cur = st.pop();
      if(!cur) continue;
      var tn = '';
      try {tn = String(cur.typename || '');} catch(_eExt6) {tn = '';}
      if(tn === 'CompoundPathItem') {
        try {cur.duplicate(tmp, ElementPlacement.PLACEATEND); copied++;} catch(_eExt7) { }
      } else if(tn === 'PathItem') {
        if(!_hasCompoundAncestor(cur)) {
          try {cur.duplicate(tmp, ElementPlacement.PLACEATEND); copied++;} catch(_eExt8) { }
        }
      }
      _pushGradientChildren(cur, st);
    }

    var removed = 0;
    try {
      var kids = [];
      for(var i = 0; i < container.pageItems.length; i++) kids.push(container.pageItems[i]);
      for(var k = 0; k < kids.length; k++) {
        try {kids[k].remove(); removed++;} catch(_eExt9) { }
      }
    } catch(_eExt10) { }

    var moved = 0;
    try {
      var tmpKids = [];
      for(var t = 0; t < tmp.pageItems.length; t++) tmpKids.push(tmp.pageItems[t]);
      for(var m = 0; m < tmpKids.length; m++) {
        try {tmpKids[m].move(container, ElementPlacement.PLACEATEND); moved++;} catch(_eExt11) { }
      }
    } catch(_eExt12) { }
    try {tmp.remove();} catch(_eExt13) { }

    _pathBleedDebug(
      'native extract ' + (label || '') +
      ' | copied=' + copied +
      ' | removed=' + removed +
      ' | moved=' + moved
    );
    return moved;
  }

  function _hasCompoundAncestor(it) {
    var p = null;
    var guard = 0;
    try {p = it.parent;} catch(_eHca0) {p = null;}
    while(p && guard < 100) {
      guard++;
      try {if(String(p.typename || '') === 'CompoundPathItem') return true;} catch(_eHca1) { }
      try {p = p.parent;} catch(_eHca2) {p = null;}
    }
    return false;
  }

  function _duplicateNativePathsFromItem(src, targetContainer, label, opts) {
    if(!src || !targetContainer) return 0;
    opts = opts || {};
    var doOutlineText = (typeof opts.outlineText === 'undefined') ? false : !!opts.outlineText;
    var copied = 0;
    var stack = [];
    try {stack.push(src);} catch(_eDupNp0) { }
    while(stack.length) {
      var cur = stack.pop();
      if(!cur) continue;
      var tn = '';
      try {tn = String(cur.typename || '');} catch(_eDupNp1) {tn = '';}
      if(tn === 'CompoundPathItem') {
        try {
          cur.duplicate(targetContainer, ElementPlacement.PLACEATEND);
          copied++;
          _pathBleedDebug('duplicate native ' + (label || '') + ' | compound | ' + _itemRef(cur));
        } catch(_eDupNp2) { }
        continue; // avoid duplicating child pathItems of same compound
      }
      if(tn === 'PathItem') {
        if(!_hasCompoundAncestor(cur)) {
          try {
            cur.duplicate(targetContainer, ElementPlacement.PLACEATEND);
            copied++;
            _pathBleedDebug('duplicate native ' + (label || '') + ' | path | ' + _itemRef(cur));
          } catch(_eDupNp3) { }
        }
      } else if(tn === 'TextFrame') {
        if(doOutlineText) {
          try {
            var dupText = cur.duplicate(targetContainer, ElementPlacement.PLACEATEND);
            var outlinedDup = dupText.createOutline();
            try {dupText.remove();} catch(_eDupNpTxtRm0) { }
            if(outlinedDup) {
              copied += _countPathLikeItems(outlinedDup);
              _pathBleedDebug('duplicate native ' + (label || '') + ' | text->outline | ' + _itemRef(cur));
            }
          } catch(_eDupNpTxt0) {
            _pathBleedDebug('duplicate native ' + (label || '') + ' | text->outline FAILED | ' + _itemRef(cur));
          }
        }
      }
      _pushGradientChildren(cur, stack);
    }
    _pathBleedDebug('duplicate native ' + (label || '') + ' | copied=' + copied);
    return copied;
  }

  function _getGradientColorForKind(item, kind) {
    if(!item) return null;
    try {
      if(kind === 'fill') {
        try {
          if(item.filled && _isGradientLikeColor(item.fillColor)) return item.fillColor;
        } catch(_eGcKind1) { }
        try {
          if(item.typename === "TextFrame" && item.textRange && item.textRange.characterAttributes && _isGradientLikeColor(item.textRange.characterAttributes.fillColor)) {
            return item.textRange.characterAttributes.fillColor;
          }
        } catch(_eGcKind2) { }
        return null;
      }
      try {return (item.stroked && _isGradientLikeColor(item.strokeColor)) ? item.strokeColor : null;} catch(_eGcKind3) {return null;}
    } catch(_eGcKind0) {return null;}
  }

  function _setGradientColorForKind(item, kind, gc) {
    if(!item || !gc) return false;
    try {
      var isGradientAssign = _isGradientColor(gc);
      function _makePrimeRed() {
        var r = new RGBColor();
        r.red = 255; r.green = 0; r.blue = 0;
        return r;
      }
      if(kind === 'fill') {
        try {if(typeof item.filled !== 'undefined') item.filled = true;} catch(_eSetGcFill0) { }
        if(forceSolidPrimeBeforeGradientAssign && isGradientAssign) {
          var primeOkFill = false;
          var redFill = _makePrimeRed();
          try {item.fillColor = redFill; primeOkFill = true;} catch(_eSetGcPrimeF0) { }
          if(!primeOkFill) {
            try {
              if(item.typename === "TextFrame" && item.textRange && item.textRange.characterAttributes) {
                item.textRange.characterAttributes.fillColor = redFill;
                primeOkFill = true;
              }
            } catch(_eSetGcPrimeF1) { }
          }
          if(primeOkFill) {
            try {app.redraw();} catch(_eSetGcPrimeF2) { }
            try {_pathBleedDebug('prime fill red before gradient | item=' + _getGradientItemLabel(item));} catch(_eSetGcPrimeF3) { }
          }
        }
        var setFill = false;
        try {
          item.fillColor = gc;
          setFill = true;
        } catch(_eSetGc1) { }
        if(!setFill) {
          try {
            if(item.typename === "TextFrame" && item.textRange && item.textRange.characterAttributes) {
              item.textRange.characterAttributes.fillColor = gc;
              setFill = true;
            }
          } catch(_eSetGc2) { }
        }
        if(!setFill) return false;
      } else {
        try {if(typeof item.stroked !== 'undefined') item.stroked = true;} catch(_eSetGcStroke0) { }
        if(forceSolidPrimeBeforeGradientAssign && isGradientAssign) {
          var primeOkStroke = false;
          var redStroke = _makePrimeRed();
          try {item.strokeColor = redStroke; primeOkStroke = true;} catch(_eSetGcPrimeS0) { }
          if(primeOkStroke) {
            try {app.redraw();} catch(_eSetGcPrimeS1) { }
            try {_pathBleedDebug('prime stroke red before gradient | item=' + _getGradientItemLabel(item));} catch(_eSetGcPrimeS2) { }
          }
        }
        item.strokeColor = gc;
      }
      return true;
    } catch(_eSetGc0) {return false;}
  }

  function _pushGradientChildren(parent, stack) {
    if(!parent || !stack) return;
    var isCompound = false;
    try {isCompound = (parent.typename === "CompoundPathItem");} catch(_ePgChType0) {isCompound = false;}
    if(isCompound) {
      try {
        if(parent.pathItems && parent.pathItems.length) {
          for(var j = 0; j < parent.pathItems.length; j++) stack.push(parent.pathItems[j]);
        }
      } catch(_ePgCh1) { }
    } else {
      try {
        if(parent.pageItems && parent.pageItems.length) {
          for(var i = 0; i < parent.pageItems.length; i++) stack.push(parent.pageItems[i]);
        }
      } catch(_ePgCh0) { }
      try {
        if(parent.pathItems && parent.pathItems.length && parent.typename === "TextFrame") {
          for(var k = 0; k < parent.pathItems.length; k++) stack.push(parent.pathItems[k]);
        }
      } catch(_ePgCh2) { }
    }
  }

  function _countDocumentGradientFillObjects(docRef) {
    if(!docRef) return 0;
    var count = 0;
    var stack = [];
    try {
      if(docRef.pageItems && docRef.pageItems.length) {
        for(var i = 0; i < docRef.pageItems.length; i++) stack.push(docRef.pageItems[i]);
      }
    } catch(_eCntGf0) { }
    while(stack.length) {
      var cur = stack.pop();
      if(!cur) continue;
      if(_isGradientColor(_getGradientColorForKind(cur, 'fill'))) count++;
      _pushGradientChildren(cur, stack);
    }
    return count;
  }

  function _readGradientStopRampPoints(gc) {
    var out = [];
    if(!gc || !gc.gradient) return out;
    var stops = null;
    try {stops = gc.gradient.gradientStops;} catch(_eGsStops0) {stops = null;}
    if(!stops || !stops.length) return out;
    for(var i = 0; i < stops.length; i++) {
      var rp = 0;
      try {rp = Number(stops[i].rampPoint || 0);} catch(_eGsStops1) {rp = 0;}
      if(!isFinite(rp)) rp = 0;
      out.push(rp);
    }
    return out;
  }

  function _getGradientStopCount(gc) {
    if(!gc || !gc.gradient) return 0;
    try {
      var stops = gc.gradient.gradientStops;
      return (stops && stops.length) ? Number(stops.length) : 0;
    } catch(_eGsc0) {return 0;}
  }

  function _getGradientName(gc) {
    if(!gc || !gc.gradient) return '';
    try {return String(gc.gradient.name || '');} catch(_eGn0) {return '';}
  }

  function _readGradientStopDescriptorsFromGradient(grad) {
    var out = [];
    if(!grad) return out;
    var stops = null;
    try {stops = grad.gradientStops;} catch(_eGsd0) {stops = null;}
    if(!stops || !stops.length) return out;
    for(var i = 0; i < stops.length; i++) {
      var rp = 0;
      try {rp = Number(stops[i].rampPoint || 0);} catch(_eGsd1) {rp = 0;}
      if(!isFinite(rp)) rp = 0;
      out.push({
        rampPoint: rp,
        color: (function(s) {try {return s.color;} catch(_eGsd2) {return null;} })(stops[i]),
        midPoint: (function(s) {try {return Number(s.midPoint || 50);} catch(_eGsd3) {return 50;} })(stops[i]),
        opacity: (function(s) {try {return Number(s.opacity || 100);} catch(_eGsd4) {return 100;} })(stops[i])
      });
    }
    return out;
  }

  function _formatStops(stops) {
    if(!stops || !stops.length) return '[]';
    var parts = [];
    for(var i = 0; i < stops.length; i++) parts.push(_fmtNum3(stops[i]));
    return '[' + parts.join(', ') + ']';
  }

  function _getItemBoundsForGradient(item) {
    if(!item) return null;
    var b = null;
    try {b = item.visibleBounds;} catch(_eGbA0) {b = null;}
    if(!b || b.length !== 4) {
      try {b = item.geometricBounds;} catch(_eGbA1) {b = null;}
    }
    return (b && b.length === 4) ? b : null;
  }

  function _estimateGradientLengthFromItem(item, angleDeg) {
    var b = _getItemBoundsForGradient(item);
    if(!b) return 0;
    var w = Math.abs(Number(b[2]) - Number(b[0]));
    var h = Math.abs(Number(b[1]) - Number(b[3]));
    if(!(w > 0) && !(h > 0)) return 0;
    var rad = Number(angleDeg || 0) * Math.PI / 180.0;
    var c = Math.abs(Math.cos(rad));
    var s = Math.abs(Math.sin(rad));
    var proj = (w * c) + (h * s);
    if(!(proj > 0)) proj = Math.max(w, h);
    return proj;
  }

  function _gradientLinePoints(origin, angleDeg, length) {
    if(!origin || origin.length < 2) return null;
    var ox = Number(origin[0]);
    var oy = Number(origin[1]);
    if(!isFinite(ox) || !isFinite(oy)) return null;
    var len = Number(length || 0);
    if(!isFinite(len)) len = 0;
    var rad = Number(angleDeg || 0) * Math.PI / 180.0;
    return {
      start: [ox, oy],
      end: [ox + (Math.cos(rad) * len), oy + (Math.sin(rad) * len)]
    };
  }

  function _fitGradientLineToItemBounds(item, origin, angleDeg, length) {
    if(!item || !origin || origin.length < 2) return null;
    var b = _getItemBoundsForGradient(item);
    if(!b || b.length !== 4) return null;
    var len = Number(length || 0);
    if(!(len > 0)) return null;

    var ox = Number(origin[0]);
    var oy = Number(origin[1]);
    var rad = Number(angleDeg || 0) * Math.PI / 180.0;
    var ux = Math.cos(rad);
    var uy = Math.sin(rad);
    if(!isFinite(ox) || !isFinite(oy) || !isFinite(ux) || !isFinite(uy)) return null;

    var l = Number(b[0]), t = Number(b[1]), r = Number(b[2]), bt = Number(b[3]);
    if(!isFinite(l) || !isFinite(t) || !isFinite(r) || !isFinite(bt)) return null;

    var corners = [[l, t], [r, t], [r, bt], [l, bt]];
    var minProj = Number.MAX_VALUE;
    var maxProj = -Number.MAX_VALUE;
    for(var i = 0; i < corners.length; i++) {
      var p = corners[i][0] * ux + corners[i][1] * uy;
      if(p < minProj) minProj = p;
      if(p > maxProj) maxProj = p;
    }
    if(!(isFinite(minProj) && isFinite(maxProj) && maxProj > minProj)) return null;

    var originProj = (ox * ux) + (oy * uy);
    var shift = minProj - originProj;
    var fittedOrigin = [ox + (ux * shift), oy + (uy * shift)];
    var fittedLen = maxProj - minProj;
    if(!(fittedLen > 0)) return null;
    return {origin: fittedOrigin, length: fittedLen, shiftAlongAxis: shift};
  }

  function _fmtLinePoints(line) {
    if(!line) return 'start=[n/a], end=[n/a]';
    return 'start=' + _fmtPoint(line.start) + ', end=' + _fmtPoint(line.end);
  }

  function _serializeStops(stops) {
    if(!stops || !stops.length) return '';
    var parts = [];
    for(var i = 0; i < stops.length; i++) {
      var n = Number(stops[i]);
      if(!isFinite(n)) n = 0;
      parts.push(String(n.toFixed(8)));
    }
    return parts.join(',');
  }

  // Stop remap in document points:
  // p' = (((p/100)*Lpt + dpt) / (Lpt + 2*dpt)) * 100
  function _computeInwardRampPoints(stops, baseLengthPt, shiftPtLocal, dbgLabel, mode) {
    if(!stops || !stops.length) return [];
    var oldLen = Number(baseLengthPt || 0);
    var shift = Number(shiftPtLocal || 0);
    var gradMode = String(mode || 'linear');
    if(!(oldLen > 0) || !(shift > 0)) return stops.slice(0);

    var newLen = (gradMode === 'radial')
      ? (oldLen + shift)
      : (oldLen + (shift * 2));
    if(!(newLen > 0)) return stops.slice(0);
    if(pathBleedDebugVerbose && dbgLabel) {
      _pathBleedDebug(
        'formula ' + dbgLabel +
        ' | mode=' + gradMode +
        ' | oldLenPt=' + _fmtNum3(oldLen) +
        ' | shiftPt=' + _fmtNum3(shift) +
        ' | newLenPt=' + _fmtNum3(newLen)
      );
    }

    var out = [];
    for(var i = 0; i < stops.length; i++) {
      var rp = Number(stops[i]);
      if(!isFinite(rp)) rp = 0;
      var t = rp / 100.0;
      var mapped = 0;
      if(gradMode === 'radial') {
        // Radial compensation keeps center fixed and expands radius by +shift.
        // So existing stop radii are scaled by oldRadius/newRadius.
        mapped = (t * oldLen / newLen) * 100.0;
      } else {
        mapped = ((t * oldLen) + shift) / newLen * 100.0;
      }
      if(!isFinite(mapped)) mapped = rp;
      if(mapped < 0) mapped = 0;
      if(mapped > 100) mapped = 100;
      out.push(mapped);
      if(pathBleedDebugVerbose && dbgLabel) {
        _pathBleedDebug(
          'formula ' + dbgLabel +
          ' | stop[' + i + ']' +
          ' | p=' + _fmtNum3(rp) +
          ' | t=' + _fmtNum3(t) +
          ' | mapped=' + _fmtNum3(mapped)
        );
      }
    }

    for(var f = 1; f < out.length; f++) {
      if(out[f] <= out[f - 1]) out[f] = Math.min(100, out[f - 1] + 0.01);
    }
    for(var b = out.length - 2; b >= 0; b--) {
      if(out[b] >= out[b + 1]) out[b] = Math.max(0, out[b + 1] - 0.01);
    }
    if(pathBleedDebugVerbose && dbgLabel) {
      _pathBleedDebug('formula ' + dbgLabel + ' | result=' + _formatStops(out));
    }
    return out;
  }

  // Piecewise remap of ramp points from one stop axis (oldStops) to another (newStops).
  // Useful when the gradient currently has a different stop count than the baseline snapshot.
  function _mapStopsThroughTransform(sourceStops, oldStops, newStops) {
    if(!sourceStops || !sourceStops.length) return [];
    if(!oldStops || !newStops || oldStops.length < 2 || newStops.length < 2 || oldStops.length !== newStops.length) {
      return sourceStops.slice(0);
    }
    var out = [];
    var n = oldStops.length;
    for(var i = 0; i < sourceStops.length; i++) {
      var p = Number(sourceStops[i]);
      if(!isFinite(p)) p = 0;

      var j = -1;
      if(p <= Number(oldStops[0])) j = 0;
      else if(p >= Number(oldStops[n - 1])) j = n - 2;
      else {
        for(var k = 0; k < n - 1; k++) {
          var a = Number(oldStops[k]);
          var b = Number(oldStops[k + 1]);
          if(!isFinite(a) || !isFinite(b) || b <= a) continue;
          if(p >= a && p <= b) {j = k; break;}
        }
      }
      if(j < 0) j = Math.max(0, n - 2);

      var o0 = Number(oldStops[j]), o1 = Number(oldStops[j + 1]);
      var m0 = Number(newStops[j]), m1 = Number(newStops[j + 1]);
      var t = 0;
      if(isFinite(o0) && isFinite(o1) && o1 > o0) t = (p - o0) / (o1 - o0);
      if(!isFinite(t)) t = 0;
      if(t < 0) t = 0;
      if(t > 1) t = 1;
      var mapped = m0 + ((m1 - m0) * t);
      if(!isFinite(mapped)) mapped = p;
      if(mapped < 0) mapped = 0;
      if(mapped > 100) mapped = 100;
      out.push(mapped);
    }
    for(var f = 1; f < out.length; f++) {
      if(out[f] <= out[f - 1]) out[f] = Math.min(100, out[f - 1] + 0.01);
    }
    for(var b = out.length - 2; b >= 0; b--) {
      if(out[b] >= out[b + 1]) out[b] = Math.max(0, out[b + 1] - 0.01);
    }
    return out;
  }

  // Debug override for bleed testing: force all compensated gradients to
  // a fixed 20/80 stop span so we can verify Illustrator actually applies
  // non-endpoint ramp points on bleed results.
  function _forceBleedTestStops2080(stops, startPct, endPct) {
    var out = [];
    if(!stops || !stops.length) return out;
    if(stops.length === 1) return [50];
    var start = Number(startPct);
    var end = Number(endPct);
    if(!isFinite(start)) start = 20;
    if(!isFinite(end)) end = 80;
    if(start < 0) start = 0;
    if(end > 100) end = 100;
    if(end < start) {var t0 = start; start = end; end = t0;}
    var span = end - start;
    for(var i = 0; i < stops.length; i++) {
      var t = (stops.length <= 1) ? 0 : (i / (stops.length - 1));
      out.push(start + (span * t));
    }
    return out;
  }

  function _buildPanelVisibleStops(srcStops, mappedRampPoints) {
    var out = [];
    if(!srcStops || srcStops.length < 2 || !mappedRampPoints || mappedRampPoints.length !== srcStops.length) {
      for(var i0 = 0; i0 < srcStops.length; i0++) {
        out.push({
          rampPoint: srcStops[i0].rampPoint,
          color: srcStops[i0].color,
          midPoint: srcStops[i0].midPoint,
          opacity: srcStops[i0].opacity
        });
      }
      return out;
    }

    function _cloneDesc(base, rp) {
      return {
        rampPoint: rp,
        color: base.color,
        midPoint: base.midPoint,
        opacity: base.opacity
      };
    }

    var first = srcStops[0];
    var last = srcStops[srcStops.length - 1];
    var leftInner = Number(mappedRampPoints[0]);
    var rightInner = Number(mappedRampPoints[mappedRampPoints.length - 1]);
    if(!isFinite(leftInner)) leftInner = 0;
    if(!isFinite(rightInner)) rightInner = 100;
    if(leftInner < 0) leftInner = 0;
    if(leftInner > 100) leftInner = 100;
    if(rightInner < 0) rightInner = 0;
    if(rightInner > 100) rightInner = 100;

    out.push(_cloneDesc(first, 0));
    if(leftInner > 0.01) out.push(_cloneDesc(first, leftInner));

    for(var i = 1; i < srcStops.length - 1; i++) {
      var rpMid = Number(mappedRampPoints[i]);
      if(!isFinite(rpMid)) rpMid = Number(srcStops[i].rampPoint || 0);
      if(rpMid < 0) rpMid = 0;
      if(rpMid > 100) rpMid = 100;
      out.push(_cloneDesc(srcStops[i], rpMid));
    }

    if(rightInner < 99.99) out.push(_cloneDesc(last, rightInner));
    out.push(_cloneDesc(last, 100));

    for(var f = 1; f < out.length; f++) {
      if(out[f].rampPoint <= out[f - 1].rampPoint) out[f].rampPoint = Math.min(100, out[f - 1].rampPoint + 0.01);
    }
    for(var b = out.length - 2; b >= 0; b--) {
      if(out[b].rampPoint >= out[b + 1].rampPoint) out[b].rampPoint = Math.max(0, out[b + 1].rampPoint - 0.01);
    }
    if(out.length > 0) out[0].rampPoint = 0;
    if(out.length > 1) out[out.length - 1].rampPoint = 100;
    return out;
  }

  // Debug-style stop rebuild: keep original stop count/colors and directly
  // assign mapped ramp points (same approach as the working debug rectangle flow).
  function _buildDirectReapplyStops(srcStops, mappedRampPoints) {
    var out = [];
    if(!srcStops || !srcStops.length) return out;
    for(var i = 0; i < srcStops.length; i++) {
      var rp = Number(srcStops[i].rampPoint || 0);
      if(mappedRampPoints && mappedRampPoints.length === srcStops.length) {
        var m = Number(mappedRampPoints[i]);
        if(isFinite(m)) rp = m;
      }
      if(!isFinite(rp)) rp = 0;
      if(rp < 0) rp = 0;
      if(rp > 100) rp = 100;
      out.push({
        rampPoint: rp,
        color: srcStops[i].color,
        midPoint: srcStops[i].midPoint,
        opacity: srcStops[i].opacity
      });
    }
    for(var f = 1; f < out.length; f++) {
      if(out[f].rampPoint <= out[f - 1].rampPoint) out[f].rampPoint = Math.min(100, out[f - 1].rampPoint + 0.01);
    }
    for(var b = out.length - 2; b >= 0; b--) {
      if(out[b].rampPoint >= out[b + 1].rampPoint) out[b].rampPoint = Math.max(0, out[b + 1].rampPoint - 0.01);
    }
    return out;
  }

  function _cloneGradientForStops(gc, rampPoints) {
    if(!gc || !gc.gradient || !rampPoints || !rampPoints.length) return null;
    var srcGrad = null;
    var srcStops = null;
    try {srcGrad = gc.gradient;} catch(_eCg0) {srcGrad = null;}
    if(!srcGrad) return null;
    srcStops = _readGradientStopDescriptorsFromGradient(srcGrad);
    if(!srcStops || !srcStops.length) return null;
    var uiStops = _buildPanelVisibleStops(srcStops, rampPoints);
    if(!uiStops || !uiStops.length) return null;

    var cloned = null;
    try {cloned = doc.gradients.add();} catch(_eCg2) {cloned = null;}
    if(!cloned) return null;

    try {cloned.type = srcGrad.type;} catch(_eCg3) { }
    try {cloned.name = 'SRH_Bleed_' + String(srcGrad.name || 'Gradient') + '_' + new Date().getTime();} catch(_eCg4) { }

    try {
      while(cloned.gradientStops.length < uiStops.length) cloned.gradientStops.add();
      while(cloned.gradientStops.length > uiStops.length && cloned.gradientStops.length > 2) {
        try {cloned.gradientStops[cloned.gradientStops.length - 1].remove();} catch(_eCg5) {break;}
      }
    } catch(_eCg6) { }

    for(var i = 0; i < uiStops.length; i++) {
      var srcStop = uiStops[i];
      var dstStop = null;
      try {dstStop = cloned.gradientStops[i];} catch(_eCg7) {dstStop = null;}
      if(!dstStop) continue;
      try {dstStop.color = srcStop.color;} catch(_eCg8) { }
      try {dstStop.midPoint = srcStop.midPoint;} catch(_eCg9) { }
      try {dstStop.opacity = srcStop.opacity;} catch(_eCg10) { }
      var rp = Number(srcStop.rampPoint || 0);
      if(rp < 0) rp = 0;
      if(rp > 100) rp = 100;
      try {dstStop.rampPoint = rp;} catch(_eCg12) { }
    }
    return cloned;
  }

  function _rebuildGradientColorWithNewGradient(gc, rampPoints, angleDeg, targetOrigin, targetLength, options) {
    var res = {ok: false, color: null, stopCount: 0, gradientName: ''};
    if(!gc || !gc.gradient || !rampPoints || !rampPoints.length) return res;
    options = options || {};
    var panelVisibleStops = !!options.panelVisibleStops;

    var srcGrad = null;
    try {srcGrad = gc.gradient;} catch(_eRb0) {srcGrad = null;}
    if(!srcGrad) return res;

    var srcStops = _readGradientStopDescriptorsFromGradient(srcGrad);
    if(!srcStops || !srcStops.length) return res;
    var uiStops = panelVisibleStops
      ? _buildPanelVisibleStops(srcStops, rampPoints)
      : _buildDirectReapplyStops(srcStops, rampPoints);
    if(!uiStops || !uiStops.length) return res;

    var newGrad = null;
    try {newGrad = doc.gradients.add();} catch(_eRb1) {newGrad = null;}
    if(!newGrad) return res;

    try {newGrad.type = srcGrad.type;} catch(_eRb2) { }
    try {newGrad.name = 'SRH_Reapply_' + String(srcGrad.name || 'Gradient') + '_' + new Date().getTime();} catch(_eRb3) { }

    try {
      while(newGrad.gradientStops.length < uiStops.length) newGrad.gradientStops.add();
      while(newGrad.gradientStops.length > uiStops.length && newGrad.gradientStops.length > 2) {
        try {newGrad.gradientStops[newGrad.gradientStops.length - 1].remove();} catch(_eRb4) {break;}
      }
    } catch(_eRb5) { }

    for(var i = 0; i < uiStops.length; i++) {
      var srcStop = uiStops[i];
      var dstStop = null;
      try {dstStop = newGrad.gradientStops[i];} catch(_eRb6) {dstStop = null;}
      if(!dstStop) continue;
      try {dstStop.color = srcStop.color;} catch(_eRb7) { }
      try {dstStop.midPoint = srcStop.midPoint;} catch(_eRb8) { }
      try {dstStop.opacity = srcStop.opacity;} catch(_eRb9) { }
      var rp = Number(srcStop.rampPoint || 0);
      if(!isFinite(rp)) rp = 0;
      if(rp < 0) rp = 0;
      if(rp > 100) rp = 100;
      try {dstStop.rampPoint = rp;} catch(_eRb10) { }
    }

    var newGc = null;
    try {newGc = new GradientColor();} catch(_eRb11) {newGc = null;}
    if(!newGc) return res;
    try {newGc.gradient = newGrad;} catch(_eRb12) {return res;}

    try {newGc.angle = Number(angleDeg || 0);} catch(_eRb13) { }
    try {
      var ln = Number(targetLength || 0);
      if(isFinite(ln) && ln > 0) newGc.length = ln;
    } catch(_eRb14) { }
    try {
      if(targetOrigin && targetOrigin.length >= 2) {
        newGc.origin = [Number(targetOrigin[0]), Number(targetOrigin[1])];
      }
    } catch(_eRb15) { }

    res.ok = true;
    res.color = newGc;
    res.stopCount = uiStops.length;
    try {res.gradientName = String(newGrad.name || '');} catch(_eRb16) {res.gradientName = '';}
    return res;
  }

  function _applyRampPointsToGradient(gc, rampPoints) {
    if(!gc || !gc.gradient || !rampPoints || !rampPoints.length) return false;
    var stops = null;
    try {stops = gc.gradient.gradientStops;} catch(_eAg0) {stops = null;}
    if(!stops || !stops.length) return false;
    var n = Math.min(stops.length, rampPoints.length);
    var changed = false;
    for(var i = 0; i < n; i++) {
      var rp = Number(rampPoints[i]);
      if(!isFinite(rp)) continue;
      if(rp < 0) rp = 0;
      if(rp > 100) rp = 100;
      try {
        stops[i].rampPoint = rp;
        changed = true;
      } catch(_eAg1) { }
    }
    return changed;
  }

  function _tryApplyGradientStops(gc, rampPoints) {
    var res = {applied: false, usedClone: false, usedCache: false};
    if(!gc || !rampPoints || !rampPoints.length) return res;
    var cacheKey = '';
    try {
      var g = gc.gradient;
      cacheKey =
        String(g && g.name ? g.name : '') + '|' +
        String(g && g.type ? g.type : '') + '|' +
        _serializeStops(rampPoints);
    } catch(_eTasKey0) {cacheKey = '';}

    if(cacheKey && _gradientCloneCache[cacheKey]) {
      try {
        gc.gradient = _gradientCloneCache[cacheKey];
        res.applied = true;
        res.usedClone = true;
        res.usedCache = true;
        return res;
      } catch(_eTasCache0) { }
    }

    var clone = _cloneGradientForStops(gc, rampPoints);
    if(clone) {
      try {
        gc.gradient = clone;
        if(cacheKey) _gradientCloneCache[cacheKey] = clone;
        res.applied = true;
        res.usedClone = true;
        return res;
      } catch(_eTas0) { }
    }
    res.applied = _applyRampPointsToGradient(gc, rampPoints);
    return res;
  }

  function _applyGradientCompensation(item, kind, rec, shiftPt) {
    if(!item) return false;
    var gc = _getGradientColorForKind(item, kind);
    var gradMode = _getCompensatableGradientMode(gc);
    if(!gradMode) return false;
    if(gradMode === 'freeform') return false;

    var itemLabel = _getGradientItemLabel(item);
    var angleDeg = 0;
    var baseLength = 0;
    var baseOrigin = null;

    if(rec) {
      if(isFinite(rec.angle)) angleDeg = Number(rec.angle);
      if(isFinite(rec.length)) baseLength = Number(rec.length);
      if(rec.origin && rec.origin.length >= 2) baseOrigin = [Number(rec.origin[0]), Number(rec.origin[1])];
    }
    if(!isFinite(angleDeg)) angleDeg = 0;
    if(!(baseLength > 0)) {
      try {baseLength = Number(gc.length || 0);} catch(_eAcs0) {baseLength = 0;}
    }
    if(!(baseLength > 0)) {
      baseLength = _estimateGradientLengthFromItem(item, angleDeg);
    }
    if(!baseOrigin) {
      try {
        if(gc.origin && gc.origin.length >= 2) baseOrigin = [Number(gc.origin[0]), Number(gc.origin[1])];
      } catch(_eAcs1) {baseOrigin = null;}
    }
    if(!baseOrigin) return false;

    var sourceStops = null;
    if(rec && rec.stops && rec.stops.length) sourceStops = rec.stops.slice(0);
    else sourceStops = _readGradientStopRampPoints(gc);
    var mappedStops = _computeInwardRampPoints(sourceStops, baseLength, shiftPt, 'primary:' + _getGradientItemLabel(item), gradMode);
    var forcedStopTestMode = forceBleedStopTest2080;
    if(forcedStopTestMode) mappedStops = _forceBleedTestStops2080(sourceStops, 20, 80);
    var beforeLine = _gradientLinePoints(baseOrigin, angleDeg, baseLength);

    var targetOrigin = [baseOrigin[0], baseOrigin[1]];
    if(gradMode === 'linear') {
      var rad = angleDeg * Math.PI / 180.0;
      var dx = Math.cos(rad) * shiftPt;
      var dy = Math.sin(rad) * shiftPt;
      targetOrigin = [baseOrigin[0] - dx, baseOrigin[1] - dy];
    }
    var targetLength = 0;
    if(baseLength > 0) {
      targetLength = (gradMode === 'radial')
        ? (baseLength + shiftPt)
        : (baseLength + (shiftPt * 2));
    }

    var stopApply = {applied: false, usedClone: false, usedCache: false, mode: 'none', rebuiltStopCount: 0, gradientName: ''};
    var panelVisibleForTwoStop = (sourceStops && sourceStops.length <= 2);
    var rebuilt = _rebuildGradientColorWithNewGradient(gc, mappedStops, angleDeg, targetOrigin, targetLength, {panelVisibleStops: panelVisibleForTwoStop});
    var colorSet = false;
    if(rebuilt.ok && rebuilt.color) {
      colorSet = _setGradientColorForKind(item, kind, rebuilt.color);
      if(colorSet) {
        stopApply.applied = true;
        stopApply.mode = panelVisibleForTwoStop ? 'rebuild-panel-visible' : 'rebuild-debugstyle';
        stopApply.rebuiltStopCount = rebuilt.stopCount;
        stopApply.gradientName = rebuilt.gradientName;
      }
    }

    if(!colorSet) {
      var oldStopApply = _tryApplyGradientStops(gc, mappedStops);
      stopApply.applied = oldStopApply.applied;
      stopApply.usedClone = oldStopApply.usedClone;
      stopApply.usedCache = oldStopApply.usedCache;
      stopApply.mode = oldStopApply.applied ? 'legacy' : 'none';

      try {gc.origin = targetOrigin;} catch(_eAcs2) { }
      if(baseLength > 0) {
        try {gc.length = targetLength;} catch(_eAcs3) { }
      }
      colorSet = _setGradientColorForKind(item, kind, gc);
    }

    var appliedGc = _getGradientColorForKind(item, kind);
    var afterStops = _readGradientStopRampPoints(appliedGc || gc);
    var appliedStopCount = _getGradientStopCount(appliedGc || gc);
    var appliedGradientName = _getGradientName(appliedGc || gc);
    var appliedOrigin = null;
    var appliedLength = 0;
    try {
      if(appliedGc && appliedGc.origin && appliedGc.origin.length >= 2) {
        appliedOrigin = [Number(appliedGc.origin[0]), Number(appliedGc.origin[1])];
      }
    } catch(_eAcs4) {appliedOrigin = null;}
    try {appliedLength = Number((appliedGc ? appliedGc.length : gc.length) || 0);} catch(_eAcs5) {appliedLength = 0;}
    if(!appliedOrigin) {
      try {
        if(gc.origin && gc.origin.length >= 2) appliedOrigin = [Number(gc.origin[0]), Number(gc.origin[1])];
      } catch(_eAcs6) {appliedOrigin = null;}
    }
    if(!isFinite(appliedLength)) appliedLength = 0;
    var afterLine = _gradientLinePoints(appliedOrigin, angleDeg, appliedLength);

    _pathBleedDebug(
      'gradient ' + kind +
      ' | mode=' + gradMode +
      ' | item=' + itemLabel +
      ' | angle=' + _fmtNum3(angleDeg) +
      ' | origin ' + _fmtPoint(baseOrigin) + ' -> ' + _fmtPoint(appliedOrigin) +
      ' | length ' + _fmtNum3(baseLength) + ' -> ' + _fmtNum3(appliedLength) +
      ' | stopPos before=' + _formatStops(sourceStops) + ' mapped=' + _formatStops(mappedStops) + ' after=' + _formatStops(afterStops) +
      ' | gradient=' + (appliedGradientName || '(unnamed)') + ' stopCount=' + appliedStopCount +
      ' | points before{' + _fmtLinePoints(beforeLine) + '} after{' + _fmtLinePoints(afterLine) + '}' +
      ' | forcedStopTest2080=' + (forcedStopTestMode ? 'yes' : 'no') +
      ' | stopApply=' + (
        stopApply.applied
          ? (
            stopApply.mode === 'rebuild-reapply'
              ? ('rebuild-reapply(' + stopApply.rebuiltStopCount + ' stops' + (stopApply.gradientName ? ', ' + stopApply.gradientName : '') + ')')
              : (stopApply.mode === 'rebuild-debugstyle'
                ? ('rebuild-debugstyle(' + stopApply.rebuiltStopCount + ' stops' + (stopApply.gradientName ? ', ' + stopApply.gradientName : '') + ')')
                : (stopApply.mode === 'rebuild-panel-visible'
                  ? ('rebuild-panel-visible(' + stopApply.rebuiltStopCount + ' stops' + (stopApply.gradientName ? ', ' + stopApply.gradientName : '') + ')')
                  : (stopApply.usedClone ? (stopApply.usedCache ? 'clone-cache' : 'clone') : 'direct'))
              )
          )
          : 'none'
      ) +
      ' | colorSet=' + (colorSet ? 'yes' : 'no')
    );
    return colorSet;
  }

  function _captureGradientSnapshot(container, stageLabel) {
    var out = [];
    if(!container) return out;
    var stack = [];
    try {stack.push(container);} catch(_eGs0) { }
    while(stack.length) {
      var cur = stack.pop();
      if(!cur) continue;
      function _captureOne(kind) {
        var gc = _getGradientColorForKind(cur, kind);
        if(!_getCompensatableGradientMode(gc)) return;
        var angle = 0;
        var length = 0;
        var origin = null;
        try {angle = Number(gc.angle || 0);} catch(_eGl2) {angle = 0;}
        try {length = Number(gc.length || 0);} catch(_eGl3) {length = 0;}
        try {
          if(gc.origin && gc.origin.length >= 2) origin = [Number(gc.origin[0]), Number(gc.origin[1])];
        } catch(_eGl4) {origin = null;}
        if(!origin) return;
        var stops = _readGradientStopRampPoints(gc);
        var linePts = _gradientLinePoints(origin, angle, length);
        out.push({item: cur, kind: kind, angle: angle, length: length, origin: origin, stops: stops});
        _pathBleedDebug(
          'gradient snapshot' +
          (stageLabel ? ' (' + stageLabel + ')' : '') +
          ' | ' + kind +
          ' | item=' + _getGradientItemLabel(cur) +
          ' | origin=' + _fmtPoint(origin) +
          ' | angle=' + _fmtNum3(angle) +
          ' | length=' + _fmtNum3(length) +
          ' | stops=' + _formatStops(stops) +
          ' | points{' + _fmtLinePoints(linePts) + '}'
        );
      }
      _captureOne('fill');
      _captureOne('stroke');
      _pushGradientChildren(cur, stack);
    }
    return out;
  }

  function _shiftGradientBySnapshot(snapshot, shiftPt, container) {
    if(!snapshot || !snapshot.length || !(shiftPt > 0)) return 0;
    var adjusted = 0;
    var liveSnapshot = [];
    if(container) liveSnapshot = _captureGradientSnapshot(container, 'post-offset pre-comp');
    var liveUsed = {};

    function _takeLiveRecord(kind) {
      if(!liveSnapshot || !liveSnapshot.length) return null;
      for(var li = 0; li < liveSnapshot.length; li++) {
        if(liveUsed[li]) continue;
        var recLive = liveSnapshot[li];
        if(!recLive || !recLive.item || recLive.kind !== kind) continue;
        liveUsed[li] = true;
        return recLive;
      }
      return null;
    }

    for(var i = 0; i < snapshot.length; i++) {
      var rec = snapshot[i];
      if(!rec || !rec.item || !rec.origin) continue;
      var usedItem = rec.item;
      var gc = _getGradientColorForKind(usedItem, rec.kind);
      if(!_getCompensatableGradientMode(gc)) {
        var liveRec = _takeLiveRecord(rec.kind);
        if(liveRec && liveRec.item) {
          usedItem = liveRec.item;
          _pathBleedDebug('snapshot fallback matched ' + rec.kind + ' gradient to ' + _getGradientItemLabel(usedItem));
        }
      }
      if(_applyGradientCompensation(usedItem, rec.kind, rec, shiftPt)) adjusted++;
    }
    if(container) _captureGradientSnapshot(container, 'post-compensation');
    return adjusted;
  }

  function _shiftGradientInContainer(container, shiftPt) {
    if(!container || !(shiftPt > 0)) return 0;
    var adjusted = 0;
    function _shiftOne(it, kind) {
      if(!it) return;
      if(_applyGradientCompensation(it, kind, null, shiftPt)) adjusted++;
    }
    var stack = [];
    try {stack.push(container);} catch(_eGs8) { }
    while(stack.length) {
      var cur = stack.pop();
      if(!cur) continue;
      _shiftOne(cur, 'fill');
      _shiftOne(cur, 'stroke');
      _pushGradientChildren(cur, stack);
    }
    return adjusted;
  }

  function _forceGradientStopsOnItem(it, kind, startPct, endPct) {
    if(!it) return false;
    var gc = _getGradientColorForKind(it, kind);
    if(!_getCompensatableGradientMode(gc)) return false;

    var angleDeg = 0;
    var length = 0;
    var origin = null;
    try {angleDeg = Number(gc.angle || 0);} catch(_eFg0) {angleDeg = 0;}
    try {length = Number(gc.length || 0);} catch(_eFg1) {length = 0;}
    try {
      if(gc.origin && gc.origin.length >= 2) origin = [Number(gc.origin[0]), Number(gc.origin[1])];
    } catch(_eFg2) {origin = null;}
    if(!(length > 0)) length = _estimateGradientLengthFromItem(it, angleDeg);
    if(!origin) {
      var b = _getItemBoundsForGradient(it);
      if(b && b.length === 4) origin = [Number(b[0]), (Number(b[1]) + Number(b[3])) / 2];
    }
    if(!origin) return false;

    var sourceStops = _readGradientStopRampPoints(gc);
    var forcedStops = _forceBleedTestStops2080(sourceStops, startPct, endPct);
    var rebuilt = _rebuildGradientColorWithNewGradient(gc, forcedStops, angleDeg, origin, length, {panelVisibleStops: true});
    var setOk = false;
    if(rebuilt.ok && rebuilt.color) setOk = _setGradientColorForKind(it, kind, rebuilt.color);
    if(!setOk) {
      var oldApply = _tryApplyGradientStops(gc, forcedStops);
      if(oldApply.applied) setOk = _setGradientColorForKind(it, kind, gc);
    }

    if(setOk) {
      var appliedGc = _getGradientColorForKind(it, kind);
      _pathBleedDebug(
        'force-pass ' + kind +
        ' | item=' + _getGradientItemLabel(it) +
        ' | source=' + _formatStops(sourceStops) +
        ' | forced=' + _formatStops(forcedStops) +
        ' | after=' + _formatStops(_readGradientStopRampPoints(appliedGc || gc)) +
        ' | stopCount=' + _getGradientStopCount(appliedGc || gc)
      );
    }
    return setOk;
  }

  function _forceGradientStopsInContainer(container, startPct, endPct) {
    if(!container) return 0;
    var adjusted = 0;
    var stack = [];
    try {stack.push(container);} catch(_eFg3) { }
    while(stack.length) {
      var cur = stack.pop();
      if(!cur) continue;
      if(_forceGradientStopsOnItem(cur, 'fill', startPct, endPct)) adjusted++;
      if(_forceGradientStopsOnItem(cur, 'stroke', startPct, endPct)) adjusted++;
      _pushGradientChildren(cur, stack);
    }
    if(container) _captureGradientSnapshot(container, 'post-force-pass');
    return adjusted;
  }

  function _applyFreeformCompensationByResize(item, shiftPt, kind) {
    if(!item || !(shiftPt > 0)) return false;
    var b = _getItemBoundsForGradient(item);
    if(!b || b.length !== 4) return false;
    var w = Math.abs(Number(b[2]) - Number(b[0]));
    var h = Math.abs(Number(b[1]) - Number(b[3]));
    if(!(w > 0) || !(h > 0)) return false;

    var oldW = w - (2 * shiftPt);
    var oldH = h - (2 * shiftPt);
    if(!(oldW > 0) || !(oldH > 0)) return false;

    var sx = (oldW / w) * 100.0;
    var sy = (oldH / h) * 100.0;
    if(!isFinite(sx) || !isFinite(sy) || sx <= 0 || sy <= 0) return false;
    function _expandBounds(b0, d) {
      if(!b0 || b0.length !== 4) return null;
      var dd = Number(d || 0);
      return [Number(b0[0]) - dd, Number(b0[1]) + dd, Number(b0[2]) + dd, Number(b0[3]) - dd];
    }
    function _fmtFreeformPointsFromBounds(b0, label) {
      if(!b0 || b0.length !== 4) return label + '=[n/a]';
      var l = Number(b0[0]), t = Number(b0[1]), r = Number(b0[2]), bt = Number(b0[3]);
      var cy = (t + bt) / 2;
      var s = [0, 20, 50, 80, 100];
      var parts = [];
      for(var i = 0; i < s.length; i++) {
        var x = l + ((r - l) * (s[i] / 100));
        parts.push(s[i] + '%:' + _fmtPoint([x, cy]));
      }
      return label + '={' + parts.join(', ') + '}';
    }
    var expectedOuter = _expandBounds(b, shiftPt);
    _pathBleedDebug(
      'freeform compensate prep ' + kind +
      ' | item=' + _itemRef(item) +
      ' | before=' + _fmtBounds4(b) +
      ' | expectedOuter=' + _fmtBounds4(expectedOuter) +
      ' | oldInnerW=' + _fmtNum3(oldW) +
      ' | oldInnerH=' + _fmtNum3(oldH) +
      ' | outerW=' + _fmtNum3(w) +
      ' | outerH=' + _fmtNum3(h) +
      ' | shift=' + _fmtNum3(shiftPt) +
      ' | scale=(' + _fmtNum3(sx) + '%,' + _fmtNum3(sy) + '%)' +
      ' | ' + _fmtFreeformPointsFromBounds(b, 'ptsBefore') +
      ' | ' + _fmtFreeformPointsFromBounds(expectedOuter, 'ptsExpectedOuter')
    );

    function _resizeOne(it, label) {
      var bb = _getItemBoundsForGradient(it);
      try {
        // For freeform payloads, resize the object itself (including geometry/pattern/gradient),
        // while keeping center anchor, so clipping path geometry can remain unchanged.
        it.resize(sx, sy, true, true, true, true, 100, Transformation.CENTER);
        var aa = _getItemBoundsForGradient(it);
        _pathBleedDebug(
          'freeform compensate ' + kind +
          ' | target=' + label +
          ' | item=' + _itemRef(it) +
          ' | before=' + _fmtBounds4(bb) +
          ' | after=' + _fmtBounds4(aa) +
          ' | shift=' + _fmtNum3(shiftPt) +
          ' | scale=(' + _fmtNum3(sx) + '%,' + _fmtNum3(sy) + '%)' +
          ' | ' + _fmtFreeformPointsFromBounds(bb, 'ptsTargetBefore') +
          ' | ' + _fmtFreeformPointsFromBounds(aa, 'ptsTargetAfter')
        );
        return true;
      } catch(_eFfR0) {return false;}
    }

    // Prefer direct NonNativeItem children where freeform data usually lives.
    var adjustedAny = false;
    try {
      if(item.pageItems && item.pageItems.length) {
        for(var i = 0; i < item.pageItems.length; i++) {
          var k = item.pageItems[i];
          if(!k) continue;
          var kt = '';
          try {kt = String(k.typename || '');} catch(_eFfK0) {kt = '';}
          if(kt === 'NonNativeItem') {
            if(_resizeOne(k, 'NonNativeItem')) adjustedAny = true;
          }
        }
      }
    } catch(_eFfR1) { }

    if(adjustedAny) return true;

    try {
      // Fallback when no direct NonNativeItem could be resized.
      var bBefore = _getItemBoundsForGradient(item);
      item.resize(sx, sy, true, true, true, true, 100, Transformation.CENTER);
      var bAfter = _getItemBoundsForGradient(item);
      _pathBleedDebug(
        'freeform compensate ' + kind +
        ' | item=' + _itemRef(item) +
        ' | before=' + _fmtBounds4(bBefore) +
        ' | after=' + _fmtBounds4(bAfter) +
        ' | shift=' + _fmtNum3(shiftPt) +
        ' | scale=(' + _fmtNum3(sx) + '%,' + _fmtNum3(sy) + '%)'
      );
      return true;
    } catch(_eFf0) {
      _pathBleedDebug('freeform compensate ' + kind + ' | FAILED resize | item=' + _itemRef(item));
      return false;
    }
  }

  function _hasDirectNonNativeChild(item) {
    if(!item) return false;
    var kids = null;
    try {kids = item.pageItems;} catch(_eNn0) {kids = null;}
    if(!kids || !kids.length) return false;
    for(var i = 0; i < kids.length; i++) {
      var k = kids[i];
      if(!k) continue;
      try {
        if(String(k.typename || '') === 'NonNativeItem') return true;
      } catch(_eNn1) { }
    }
    return false;
  }

  // Final-target pass: adjust stops on the actual post-offset bleed objects.
  // This handles cases where expand/offset creates a new target path that does
  // not preserve the earlier compensated gradient object reference.
  function _finalizeBleedGradientStopsByFinalItems(container, shiftPt, baselineSnapshot, geometryChanged) {
    if(!container || !(shiftPt > 0)) return 0;
    var adjusted = 0;
    var geomChanged = !!geometryChanged;
    var scanned = 0;
    var skippedOriginalBranch = 0;
    var skippedNoGradient = 0;
    var skippedNonLinear = 0;
    var skippedNoStops = 0;
    var skippedNoOriginOrLength = 0;
    var skippedNoMapped = 0;
    var freeformGroupAdjusted = 0;
    var stack = [];
    try {stack.push(container);} catch(_eFgFin0) { }
    while(stack.length) {
      var cur = stack.pop();
      if(!cur) continue;
      scanned++;
      // Freeform compensation is intentionally disabled to avoid appearance expansion/transforms.
      function _one(kind) {
        // Only target final geometry created for this run.
        if(!_hasAncestorNamed(cur, _pbBleedGroupName)) {
          _pathBleedDebug('final-target skip | kind=' + kind + ' | reason=not-in-this-run-bleed-group | item=' + _itemRef(cur));
          return;
        }
        // Skip the preserved source-copy subtree; only final bleed geometry should be compensated here.
        if(_hasAncestorNamed(cur, _pbOriginalGroupName)) {
          skippedOriginalBranch++;
          _pathBleedDebug('final-target skip | kind=' + kind + ' | reason=inside-original-snapshot | item=' + _itemRef(cur));
          return;
        }
        var gc = _getGradientColorForKind(cur, kind);
        if(!gc) {
          skippedNoGradient++;
          return;
        }
        var gradMode = _getCompensatableGradientMode(gc);
        if(!gradMode) {
          skippedNonLinear++;
          _pathBleedDebug('final-target skip | kind=' + kind + ' | reason=unsupported-gradient | item=' + _itemRef(cur));
          return;
        }
        if(!geomChanged) {
          skippedNoMapped++;
          _pathBleedDebug('final-target skip | kind=' + kind + ' | reason=offset-geometry-unchanged | mode=' + gradMode + ' | item=' + _itemRef(cur));
          return;
        }
        var len = 0;
        var ang = 0;
        var org = null;
        try {len = Number(gc.length || 0);} catch(_eFgFin1) {len = 0;}
        try {ang = Number(gc.angle || 0);} catch(_eFgFin2) {ang = 0;}
        try {if(gc.origin && gc.origin.length >= 2) org = [Number(gc.origin[0]), Number(gc.origin[1])];} catch(_eFgFin3) {org = null;}
        if(!(len > 0)) len = _estimateGradientLengthFromItem(cur, ang);
        if(!(len > 0)) {
          skippedNoOriginOrLength++;
          _pathBleedDebug('final-target skip | kind=' + kind + ' | reason=no-length | item=' + _itemRef(cur));
          return;
        }
        var sourceStops = _readGradientStopRampPoints(gc);
        if(!sourceStops || !sourceStops.length) {
          skippedNoStops++;
          _pathBleedDebug('final-target skip | kind=' + kind + ' | reason=no-stops | item=' + _itemRef(cur));
          return;
        }
        if(!org) {
          var bb = _getItemBoundsForGradient(cur);
          if(bb && bb.length === 4) org = [Number(bb[0]), (Number(bb[1]) + Number(bb[3])) / 2];
        }
        if(!org) {
          skippedNoOriginOrLength++;
          _pathBleedDebug('final-target skip | kind=' + kind + ' | reason=no-origin | item=' + _itemRef(cur));
          return;
        }
        if(gradMode === 'linear') {
          var fitFinal = _fitGradientLineToItemBounds(cur, org, ang, len);
          if(fitFinal && fitFinal.origin && fitFinal.origin.length >= 2 && fitFinal.length > 0) {
            _pathBleedDebug(
              'final-target bounds-fit ' + kind +
              ' | item=' + _itemRef(cur) +
              ' | shiftAlongAxis=' + _fmtNum3(fitFinal.shiftAlongAxis) +
              ' | origin ' + _fmtPoint(org) + ' -> ' + _fmtPoint(fitFinal.origin) +
              ' | length ' + _fmtNum3(len) + ' -> ' + _fmtNum3(fitFinal.length)
            );
            org = fitFinal.origin;
            len = fitFinal.length;
          }
        }

        var origLenPtFromFinal = (gradMode === 'radial')
          ? (len - shiftPt)
          : (len - (2 * shiftPt));
        if(!(origLenPtFromFinal > 0)) origLenPtFromFinal = len;

        var baselineStops = sourceStops;
        var baselineLenPt = origLenPtFromFinal;
        var baselineRec = null;
        if(baselineSnapshot && baselineSnapshot.length) {
          var targetBaseLenPt = (gradMode === 'radial')
            ? (len - shiftPt)
            : (len - (2 * shiftPt));
          if(!(targetBaseLenPt > 0)) targetBaseLenPt = len;
          var bestDiff = Number.MAX_VALUE;
          for(var bi = 0; bi < baselineSnapshot.length; bi++) {
            var rec = baselineSnapshot[bi];
            if(!rec || rec.kind !== kind) continue;
            var rLenPt = Number(rec.length || 0);
            var diff = Math.abs(rLenPt - targetBaseLenPt);
            if(diff < bestDiff) {bestDiff = diff; baselineRec = rec;}
          }
        }
        if(baselineRec && baselineRec.stops && baselineRec.stops.length) baselineStops = baselineRec.stops.slice(0);
        if(baselineRec && isFinite(baselineRec.length) && Number(baselineRec.length) > 0) baselineLenPt = Number(baselineRec.length);

        _pathBleedDebug(
          'final-target baseline ' + kind +
          ' | item=' + _itemRef(cur) +
          ' | finalLenPt=' + _fmtNum3(len) +
          ' | origLenPtFromFinal=' + _fmtNum3(origLenPtFromFinal) +
          ' | chosenBaselineLenPt=' + _fmtNum3(baselineLenPt) +
          ' | baselineStops=' + _formatStops(baselineStops)
        );

        var mappedBaseline = _computeInwardRampPoints(baselineStops, baselineLenPt, shiftPt, 'final-target:' + _getGradientItemLabel(cur), gradMode);
        if(!mappedBaseline || !mappedBaseline.length) {
          skippedNoMapped++;
          _pathBleedDebug('final-target skip | kind=' + kind + ' | reason=empty-mapped-baseline | item=' + _itemRef(cur));
          return;
        }

        var mapped = null;
        var panelVisibleStops = false;
        var branch = '';
        if(baselineStops.length <= 2) {
          // Keep 2-stop behavior stable even when current gradient has panel-expanded duplicate endpoints.
          panelVisibleStops = true;
          branch = (sourceStops.length <= 2) ? 'two-stop-direct' : 'two-stop-expanded-source';
          if(sourceStops.length <= 2) {
            mapped = mappedBaseline.slice(0);
          } else {
            var leftNew = Number(mappedBaseline[0]);
            var rightNew = Number(mappedBaseline[mappedBaseline.length - 1]);
            if(!isFinite(leftNew)) leftNew = 0;
            if(!isFinite(rightNew)) rightNew = 100;
            if(rightNew < leftNew) {var _tmpLR = leftNew; leftNew = rightNew; rightNew = _tmpLR;}

            var leftOld = Number(sourceStops[1]);
            var rightOld = Number(sourceStops[sourceStops.length - 2]);
            if(!isFinite(leftOld)) leftOld = 0;
            if(!isFinite(rightOld)) rightOld = 100;
            if(!(rightOld > leftOld)) {leftOld = 0; rightOld = 100;}

            mapped = [];
            var lastIdx = sourceStops.length - 1;
            for(var si = 0; si < sourceStops.length; si++) {
              if(si === 0) {mapped.push(0); continue;}
              if(si === lastIdx) {mapped.push(100); continue;}
              var rpOld = Number(sourceStops[si]);
              if(!isFinite(rpOld)) rpOld = leftOld;
              var tt = (rpOld - leftOld) / (rightOld - leftOld);
              if(!isFinite(tt)) tt = 0;
              if(tt < 0) tt = 0;
              if(tt > 1) tt = 1;
              mapped.push(leftNew + ((rightNew - leftNew) * tt));
            }
          }
        } else {
          // Multi-stop baseline: preserve full stop structure and remap all stops.
          panelVisibleStops = false;
          branch = (sourceStops.length === baselineStops.length) ? 'multi-stop-direct' : 'multi-stop-piecewise';
          if(sourceStops.length === baselineStops.length) mapped = mappedBaseline.slice(0);
          else mapped = _mapStopsThroughTransform(sourceStops, baselineStops, mappedBaseline);
        }

        if(!mapped || !mapped.length) {
          skippedNoMapped++;
          _pathBleedDebug('final-target skip | kind=' + kind + ' | reason=empty-mapped | branch=' + branch + ' | item=' + _itemRef(cur));
          return;
        }
        _pathBleedDebug(
          'final-target remap ' + kind +
          ' | mode=' + gradMode +
          ' | branch=' + branch +
          ' | panelVisible=' + (panelVisibleStops ? 'yes' : 'no') +
          ' | item=' + _itemRef(cur) +
          ' | sourceStops=' + _formatStops(sourceStops) +
          ' | baselineStops=' + _formatStops(baselineStops) +
          ' | mappedBaseline=' + _formatStops(mappedBaseline) +
          ' | mapped=' + _formatStops(mapped)
        );
        var rebuilt = _rebuildGradientColorWithNewGradient(gc, mapped, ang, org, len, {panelVisibleStops: panelVisibleStops});
        var ok = false;
        if(rebuilt.ok && rebuilt.color) ok = _setGradientColorForKind(cur, kind, rebuilt.color);
        if(ok) {
          adjusted++;
          var agc = _getGradientColorForKind(cur, kind);
          _pathBleedDebug(
            'final-target pass ' + kind +
            ' | item=' + _itemRef(cur) +
            ' | len=' + _fmtNum3(len) +
            ' | mapped=' + _formatStops(mapped) +
            ' | after=' + _formatStops(_readGradientStopRampPoints(agc || gc)) +
            ' | stopCount=' + _getGradientStopCount(agc || gc)
          );
        }
      }
      _one('fill');
      _one('stroke');
      _pushGradientChildren(cur, stack);
    }
    _pathBleedDebug(
      'final-target summary | scanned=' + scanned +
      ' | adjusted=' + adjusted +
      ' | skipOriginalBranch=' + skippedOriginalBranch +
      ' | skipNoGradient=' + skippedNoGradient +
      ' | skipNonLinear=' + skippedNonLinear +
      ' | skipNoStops=' + skippedNoStops +
      ' | skipNoOriginOrLength=' + skippedNoOriginOrLength +
      ' | skipNoMapped=' + skippedNoMapped +
      ' | freeformGroupAdjusted=' + freeformGroupAdjusted
    );
    if(container) _captureGradientSnapshot(container, 'post-final-target-pass');
    return adjusted;
  }

  function _forceSolidRedInContainer(container) {
    if(!container) return 0;
    var changed = 0;
    var stack = [];
    var red = new RGBColor();
    red.red = 255; red.green = 0; red.blue = 0;
    var noColor = new NoColor();
    try {stack.push(container);} catch(_eFs0) { }
    while(stack.length) {
      var cur = stack.pop();
      if(!cur) continue;
      var did = false;
      try {
        if(typeof cur.filled !== 'undefined') {
          cur.filled = true;
          cur.fillColor = red;
          try {if(typeof cur.stroked !== 'undefined') cur.stroked = false;} catch(_eFsStroke0) { }
          try {if(typeof cur.strokeColor !== 'undefined') cur.strokeColor = noColor;} catch(_eFsStroke1) { }
          did = true;
        }
      } catch(_eFs1) { }
      if(!did) {
        try {
          if(cur.typename === "TextFrame" && cur.textRange && cur.textRange.characterAttributes) {
            cur.textRange.characterAttributes.fillColor = red;
            try {cur.textRange.characterAttributes.strokeColor = noColor;} catch(_eFsTxtStroke0) { }
            did = true;
          }
        } catch(_eFs2) { }
      }
      if(did) {
        changed++;
        _pathBleedDebug('force-solid-red | item=' + _getGradientItemLabel(cur));
      }
      _pushGradientChildren(cur, stack);
    }
    try {app.redraw();} catch(_eFs3) { }
    return changed;
  }

  function _drawGradientStopDebugLines(originalSnapshot, finalContainer) {
    var out = {original: 0, bleed: 0, total: 0};

    var debugGroup = null;
    try {
      var gItems = bleedLayer.groupItems;
      for(var gi = gItems.length - 1; gi >= 0; gi--) {
        try {
          if(String(gItems[gi].name || '') === 'SRH_GRAD_DEBUG_LINES') gItems[gi].remove();
        } catch(_eGdl0) { }
      }
      debugGroup = bleedLayer.groupItems.add();
      debugGroup.name = 'SRH_GRAD_DEBUG_LINES';
    } catch(_eGdl1) {debugGroup = null;}
    if(!debugGroup) return out;

    var abRect = null;
    try {
      var ai = doc.artboards.getActiveArtboardIndex();
      abRect = doc.artboards[ai].artboardRect;
    } catch(_eGdl2) {abRect = null;}
    if(!abRect || abRect.length !== 4) return out;
    var yTop = Number(abRect[1]);
    var yBottom = Number(abRect[3]);
    if(!isFinite(yTop) || !isFinite(yBottom)) return out;

    var cOrigEnd = new RGBColor(); cOrigEnd.red = 0; cOrigEnd.green = 180; cOrigEnd.blue = 0;
    var cOrigMid = new RGBColor(); cOrigMid.red = 120; cOrigMid.green = 220; cOrigMid.blue = 120;
    var cFinalEnd = new RGBColor(); cFinalEnd.red = 255; cFinalEnd.green = 0; cFinalEnd.blue = 0;
    var cFinalMid = new RGBColor(); cFinalMid.red = 0; cFinalMid.green = 170; cFinalMid.blue = 255;

    function _addLine(x, isEndpoint, isOriginal) {
      try {
        var p = debugGroup.pathItems.add();
        p.setEntirePath([[x, yTop], [x, yBottom]]);
        p.stroked = true;
        p.filled = false;
        p.strokeWidth = _srh_pxStrokeDoc(isOriginal ? 1.6 : 1);
        p.strokeColor = isOriginal
          ? (isEndpoint ? cOrigEnd : cOrigMid)
          : (isEndpoint ? cFinalEnd : cFinalMid);
        return true;
      } catch(_eGdl3) {return false;}
    }

    function _drawFromLine(origin, angle, len, stops, isOriginal) {
      if(!origin || origin.length < 2 || !(len > 0) || !stops || !stops.length) return 0;
      var rad = Number(angle || 0) * Math.PI / 180.0;
      var dx = Math.cos(rad) * len;
      var count = 0;
      for(var i = 0; i < stops.length; i++) {
        var rp = Number(stops[i]);
        if(!isFinite(rp)) continue;
        var x = Number(origin[0]) + (dx * (rp / 100.0));
        if(!isFinite(x)) continue;
        if(_addLine(x, (i === 0 || i === stops.length - 1), isOriginal)) count++;
      }
      return count;
    }

    // Draw final bleed lines first, then original lines so originals remain visible on top when overlapping.
    if(finalContainer) {
      var stack = [];
      try {stack.push(finalContainer);} catch(_eGdl4) { }
      while(stack.length) {
        var cur = stack.pop();
        if(!cur) continue;
        function _drawFinalKind(kind) {
          var gc = _getGradientColorForKind(cur, kind);
          if(!_getCompensatableGradientMode(gc)) return;
          var stops = _readGradientStopRampPoints(gc);
          if(!stops || !stops.length) return;
          var angle = 0, len = 0, origin = null;
          try {angle = Number(gc.angle || 0);} catch(_eGdl5) {angle = 0;}
          try {len = Number(gc.length || 0);} catch(_eGdl6) {len = 0;}
          try {if(gc.origin && gc.origin.length >= 2) origin = [Number(gc.origin[0]), Number(gc.origin[1])];} catch(_eGdl7) {origin = null;}
          if(!(len > 0)) len = _estimateGradientLengthFromItem(cur, angle);
          if(!origin || !(len > 0)) return;
          var fit = _fitGradientLineToItemBounds(cur, origin, angle, len);
          if(fit && fit.origin && fit.origin.length >= 2 && fit.length > 0) {
            origin = fit.origin;
            len = fit.length;
          }
          out.bleed += _drawFromLine(origin, angle, len, stops, false);
        }
        _drawFinalKind('fill');
        _drawFinalKind('stroke');
        _pushGradientChildren(cur, stack);
      }
    }

    if(originalSnapshot && originalSnapshot.length) {
      for(var os = 0; os < originalSnapshot.length; os++) {
        var rec = originalSnapshot[os];
        if(!rec || !rec.origin || !rec.stops || !rec.stops.length) continue;
        var o = [Number(rec.origin[0]), Number(rec.origin[1])];
        var a = Number(rec.angle || 0);
        var l = Number(rec.length || 0);
        var fitO = _fitGradientLineToItemBounds(rec.item, o, a, l);
        if(fitO && fitO.origin && fitO.origin.length >= 2 && fitO.length > 0) {
          o = fitO.origin;
          l = fitO.length;
        }
        out.original += _drawFromLine(o, a, l, rec.stops, true);
      }
    }

    out.total = out.original + out.bleed;
    return out;
  }

  // Back-compat shims for older cached JSX paths that may still call the
  // previous snapshot-based gradient helpers.
  function _collectGradientItems(container) {
    return _captureGradientSnapshot(container);
  }

  function _compensateGradientScale(snapshot) {
    return _shiftGradientBySnapshot(snapshot, offsetPt);
  }

  function _uniteContainerPathfinderAdd(container) {
    if(!container) return false;
    return _runOnSelection(container, function() {
      try {app.executeMenuCommand('group');} catch(_eUpaG0) { }
      try {app.executeMenuCommand('Live Pathfinder Add');} catch(_eUpa0) { }
      try {app.executeMenuCommand('expandStyle');} catch(_eUpa1) { }
      try {app.executeMenuCommand('ungroup');} catch(_eUpa2) { }
    });
  }

  function _uniteContainer(container) {
    if(!container) return false;
    function _collectUniteTargets(root) {
      var out = [];
      var seen = [];
      function _alreadySeen(it) {
        for(var si = 0; si < seen.length; si++) {
          if(seen[si] === it) return true;
        }
        return false;
      }
      function _pushIfValid(it) {
        if(!it || _alreadySeen(it)) return;
        var tn = '';
        try {tn = String(it.typename || '');} catch(_eUT0) {tn = '';}
        if(tn !== 'PathItem') return;
        try {
          if(it.parent && String(it.parent.typename || '') === 'CompoundPathItem') return;
        } catch(_eUT1) { }
        try {if(it.guides) return;} catch(_eUT2) { }
        try {if(it.clipping) return;} catch(_eUT3) { }
        seen.push(it);
        out.push(it);
      }
      try {
        var items = root.pageItems;
        for(var i = 0; i < items.length; i++) _pushIfValid(items[i]);
      } catch(_eUT4) {
        _pushIfValid(root);
      }
      return out;
    }
    var targets = _collectUniteTargets(container);
    if(!targets || targets.length < 2) return false;
    var ok = _runOnSelection(targets, function() {
      try {app.executeMenuCommand('group');} catch(_eUGrp) { }
      try {app.executeMenuCommand('Live Pathfinder Add');} catch(_eU0) { }
      try {app.executeMenuCommand('expandStyle');} catch(_eUEx0) { }
      try {app.executeMenuCommand('ungroup');} catch(_eUUng0) { }
    });
    return !!ok;
  }

  var preSel = [];
  for(var ps = 0; ps < doc.selection.length; ps++) preSel.push(doc.selection[ps]);
  _logSelectionGradientKinds(preSel, 'pre-guard-selection');
  var freeformFallbackMode = _selectionHasFreeformLike(preSel);
  if(freeformFallbackMode) {
    _pathBleedDebug('freeform detect | selection contains NonNative/freeform | mode=offset-only');
  }

  var bleedLayer = _srh_getOrCreateLayer(doc, "bleed");
  var originalLayer = _srh_getOrCreateLayer(doc, "original");
  var cutLayer = createCutline ? _srh_getOrCreateLayer(doc, "cutline") : _srh_getLayerByName(doc, "cutline");
  var _pbRunId = String(new Date().getTime());
  var _pbBleedGroupName = 'SRH_BLEED_WORK__' + _pbRunId;
  var _pbOriginalGroupName = 'SRH_ORIGINAL_SNAPSHOT__' + _pbRunId;
  var _pbCutGroupName = 'SRH_CUTLINE_WORK__' + _pbRunId;
  _srh_setBleedLayerOrder(cutLayer, originalLayer, bleedLayer);
  try {originalLayer.visible = true; originalLayer.locked = false;} catch(_eOlShow0) { }
  try {if(cutLayer) {cutLayer.visible = true; cutLayer.locked = false;} } catch(_eCutShow0) { }

  // Snapshot selection because we'll move items to layers.
  var sel = [];
  for(var s = 0; s < doc.selection.length; s++) {sel.push(doc.selection[s]);}
  _pathBleedDebug('selection snapshot count=' + sel.length);
  for(var ss = 0; ss < sel.length; ss++) _pathBleedDebug('selection[' + ss + '] ' + _itemRef(sel[ss]));

  function _scanBleedSelectionEligibility(items) {
    var out = {nativeCount: 0, textFrameCount: 0};
    var stack = [];
    for(var i = 0; i < items.length; i++) {
      try {stack.push(items[i]);} catch(_eSbse0) { }
    }
    while(stack.length) {
      var cur = stack.pop();
      if(!cur) continue;
      var tn = '';
      try {tn = String(cur.typename || '');} catch(_eSbse1) {tn = '';}
      if(tn === 'TextFrame') {
        out.textFrameCount++;
        if(outlineText) out.nativeCount++;
      } else if(tn === 'CompoundPathItem') {
        out.nativeCount++;
        continue;
      } else if(tn === 'PathItem') {
        if(!_hasCompoundAncestor(cur)) out.nativeCount++;
      }
      try {
        if(cur.pageItems && cur.pageItems.length) {
          for(var j = 0; j < cur.pageItems.length; j++) stack.push(cur.pageItems[j]);
        }
      } catch(_eSbse2) { }
    }
    return out;
  }

  var _selectionEligibility = _scanBleedSelectionEligibility(sel);
  if(!_selectionEligibility.nativeCount) {
    try {doc.selection = null;} catch(_eSelElig0) { }
    try {if(originalLayer) originalLayer.visible = false;} catch(_eOlHideElig0) { }
    if(_selectionEligibility.textFrameCount > 0) {
      return _pathBleedWithDbg('Text must be outlined before using Bleed + Cutline.');
    }
    return _pathBleedWithDbg("No eligible items to offset. Doc gradient-fill objects: " + _countDocumentGradientFillObjects(doc) + ".");
  }

  var bleedGroup = null;
  var originalGroup = null;
  var cutGroup = null;
  try {
    bleedGroup = bleedLayer.groupItems.add();
    bleedGroup.name = _pbBleedGroupName;
  } catch(_eG0) { }
  try {
    originalGroup = originalLayer.groupItems.add();
    originalGroup.name = _pbOriginalGroupName;
  } catch(_eG1) {originalGroup = originalLayer;}
  if(createCutline && cutLayer) {
    try {cutLayer.visible = true; cutLayer.locked = false;} catch(_eCutShow1) { }
    try {
      cutGroup = cutLayer.groupItems.add();
      cutGroup.name = _pbCutGroupName;
    } catch(_eG2) {cutGroup = cutLayer;}
  }

  var originalCount = 0;
  var cutCount = 0;
  var movedCount = 0;
  var originalItems = [];
  var bleedWorkContainers = [];
  var cutWorkContainers = [];
  var docGradientFillCount = _countDocumentGradientFillObjects(doc);
  _pathBleedDebug('doc gradient-fill object count=' + docGradientFillCount);
  _pathBleedDebug('layers: bleed=' + (bleedLayer ? bleedLayer.name : 'n/a') + ', original=' + (originalLayer ? originalLayer.name : 'n/a') + ', cut=' + (cutLayer ? cutLayer.name : 'n/a'));

  // First operation: move selected source artwork to the original layer/group.
  for(var n0 = 0; n0 < sel.length; n0++) {
    var src = sel[n0];
    if(!src || src.locked || src.hidden) continue;
    _pathBleedDebug('move->original candidate[' + n0 + '] ' + _itemRef(src));
    try {
      if(originalGroup) {
        src.move(originalGroup, ElementPlacement.PLACEATEND);
      } else {
        src.move(originalLayer, ElementPlacement.PLACEATBEGINNING);
      }
      originalCount++;
      originalItems.push(src);
      _pathBleedDebug('move->original OK[' + n0 + ']');
    } catch(_eMoveOrig0) { }
  }

  if(!originalCount) {
    try {doc.selection = null;} catch(_eSelOrig0) { }
    try {if(originalLayer) originalLayer.visible = false;} catch(_eOlHide0) { }
    return _pathBleedWithDbg("No eligible items to move into original layer. Doc gradient-fill objects: " + docGradientFillCount + ".");
  }

  // Second pass: duplicate from originals into cutline/bleed work layers.
  for(var n = 0; n < originalItems.length; n++) {
    var item = originalItems[n];
    if(!item) continue;
    _pathBleedDebug('duplicate source[' + n + '] ' + _itemRef(item));
    var bleedTargetContainer = bleedGroup || bleedLayer;
    var cutTargetContainer = cutGroup || cutLayer;

    if(bleedGroup) {
      try {
        bleedTargetContainer = bleedGroup.groupItems.add();
        bleedTargetContainer.name = 'SRH_BLEED_ITEM__' + _pbRunId + '__' + n;
        bleedWorkContainers.push(bleedTargetContainer);
      } catch(_eDupBleedGrp0) {
        bleedTargetContainer = bleedGroup;
      }
    }
    if(createCutline && cutGroup) {
      try {
        cutTargetContainer = cutGroup.groupItems.add();
        cutTargetContainer.name = 'SRH_CUT_ITEM__' + _pbRunId + '__' + n;
        cutWorkContainers.push(cutTargetContainer);
      } catch(_eDupCutGrp0) {
        cutTargetContainer = cutGroup;
      }
    }

    if(createCutline && cutTargetContainer) {
      try {cutCount += _duplicateNativePathsFromItem(item, cutTargetContainer, 'cut', {outlineText: outlineText});} catch(_eDupCut) { }
    }

    if(bleedTargetContainer) {
      try {movedCount += _duplicateNativePathsFromItem(item, bleedTargetContainer, 'bleed', {outlineText: outlineText});} catch(_eDupBleed0) { }
    } else {
      try {movedCount += _duplicateNativePathsFromItem(item, bleedLayer, 'bleed-layer', {outlineText: outlineText});} catch(_eDupBleed1) { }
    }
  }

  function _logCutlineGeometry(container, label) {
    if(!pathBleedDebugVerbose || !container) return;
    var stack = [];
    var idx = 0;
    try {stack.push(container);} catch(_eLcg0) { }
    _pathBleedDebug('cutline-geom START ' + (label || '') + ' | root=' + _itemRef(container));
    while(stack.length && idx < 80) {
      var cur = stack.pop();
      if(!cur) continue;
      idx++;
      var tn = '';
      try {tn = String(cur.typename || '');} catch(_eLcg1) {tn = ''; }
      if(tn === 'PathItem' || tn === 'CompoundPathItem' || tn === 'GroupItem') {
        var vb = null;
        var gb = null;
        var filled = 'n/a';
        var stroked = 'n/a';
        var strokeWidth = 'n/a';
        var strokeColor = 'n/a';
        var fillColor = 'n/a';
        try {vb = cur.visibleBounds;} catch(_eLcg2) {vb = null; }
        try {gb = cur.geometricBounds;} catch(_eLcg3) {gb = null; }
        try {filled = (typeof cur.filled === 'undefined') ? 'n/a' : String(!!cur.filled);} catch(_eLcg4) { }
        try {stroked = (typeof cur.stroked === 'undefined') ? 'n/a' : String(!!cur.stroked);} catch(_eLcg5) { }
        try {strokeWidth = (typeof cur.strokeWidth === 'undefined') ? 'n/a' : _fmtNum3(Number(cur.strokeWidth || 0));} catch(_eLcg6) { }
        try {
          if(cur.strokeColor) strokeColor = String(cur.strokeColor.typename || '');
          if(cur.strokeColor && cur.strokeColor.spot) strokeColor += ':' + String(cur.strokeColor.spot.name || '');
        } catch(_eLcg7) { }
        try {
          if(cur.fillColor) fillColor = String(cur.fillColor.typename || '');
          if(cur.fillColor && cur.fillColor.spot) fillColor += ':' + String(cur.fillColor.spot.name || '');
        } catch(_eLcg8) { }
        _pathBleedDebug(
          'cutline-geom item | idx=' + idx +
          ' | type=' + tn +
          ' | vb=' + _fmtBounds4(vb) +
          ' | gb=' + _fmtBounds4(gb) +
          ' | filled=' + filled +
          ' | stroked=' + stroked +
          ' | strokeWidth=' + strokeWidth +
          ' | strokeColor=' + strokeColor +
          ' | fillColor=' + fillColor +
          ' | ref=' + _itemRef(cur)
        );
      }
      try {
        if(cur.pageItems && cur.pageItems.length) {
          for(var i = 0; i < cur.pageItems.length; i++) stack.push(cur.pageItems[i]);
        }
      } catch(_eLcg9) { }
    }
    _pathBleedDebug('cutline-geom END ' + (label || '') + ' | scanned=' + idx);
  }
  _pathBleedDebug('duplicate summary | originals=' + originalCount + ' | movedToBleed=' + movedCount + ' | cutDup=' + cutCount);
  var bleedContainerRef = bleedGroup || bleedLayer;
  var _bleedNativeCount = 0;
  if(bleedWorkContainers && bleedWorkContainers.length) {
    for(var _bw = 0; _bw < bleedWorkContainers.length; _bw++) {
      _bleedNativeCount += _extractNativePathSources(bleedWorkContainers[_bw], 'bleed-after-duplicate[' + _bw + ']');
    }
  } else {
    _bleedNativeCount = _extractNativePathSources(bleedContainerRef, 'bleed-after-duplicate');
  }
  if(_bleedNativeCount > 0) movedCount = _bleedNativeCount;
  _pathBleedDebug('bleed native summary | pathCount=' + _bleedNativeCount + ' | movedToBleed=' + movedCount);
  _auditStructure(bleedGroup || bleedLayer, 'bleed-after-duplicate');
  if(createCutline && cutGroup) _auditStructure(cutGroup, 'cut-after-duplicate');

  if(!movedCount) {
    try {doc.selection = null;} catch(_eSel0) { }
    try {if(originalLayer) originalLayer.visible = false;} catch(_eOlHide1) { }
    return _pathBleedWithDbg("No eligible items to offset. Doc gradient-fill objects: " + docGradientFillCount + ".");
  }

  // Secondary guard: some freeform payloads are only detectable after duplication into bleed group.
  var _postDupFreeformLike = false;
  var _postDupNonNativeCount = 0;
  try {_postDupFreeformLike = !!(bleedGroup && _itemHasFreeformLike(bleedGroup));} catch(_eFfGuardPd0) {_postDupFreeformLike = false;}
  try {_postDupNonNativeCount = bleedGroup ? _countNonNativeItems(bleedGroup) : 0;} catch(_eFfGuardPd1) {_postDupNonNativeCount = 0;}
  _pathBleedDebug('post-duplicate guard | freeformLike=' + (_postDupFreeformLike ? 'yes' : 'no') + ' | nonNativeCount=' + _postDupNonNativeCount);
  if(bleedGroup && (_postDupFreeformLike || _postDupNonNativeCount > 0)) {
    _logSelectionGradientKinds([bleedGroup], 'post-duplicate-bleed-group');
    _pathBleedDebug('freeform detect | duplicated bleed content contains NonNative/freeform | mode=offset-only');
    freeformFallbackMode = true;
  }

  // Hard isolate processing from originals: never restore commands onto original selection.
  try {doc.selection = null;} catch(_eSelIso0) { }

  // Keep bleed geometry based on the original native path copy.
  // Do not outline/expand bleed work items before offset, otherwise the bleed
  // can become derived from expanded stroke/cutline geometry instead of source paths.
  var closedOpenPathCount = 0;
  if(autoCloseOpenPaths) {
    closedOpenPathCount = _closeOpenPathsInContainer(bleedGroup || bleedLayer, 'bleed-pre-offset');
  }
  _logContainerState(bleedGroup || bleedLayer, 'pre-offset-container');
  var _preOffsetPathMetrics = _collectPathMetrics(bleedGroup || bleedLayer);
  _logPathMetrics('pre-offset', _preOffsetPathMetrics);
  var _preOffsetNonNativeCount = _scanForNonNative(bleedGroup || bleedLayer, 'pre-offset');
  if(_preOffsetNonNativeCount > 0) {
    _pathBleedDebug('freeform detect | pre-offset container contains NonNativeItem | mode=offset-only');
    freeformFallbackMode = true;
  }
  var gradientSnapshot = _captureGradientSnapshot(bleedGroup || bleedLayer, 'pre-offset');
  var _bleedOffsetTargets = _collectNativePathRoots(bleedGroup || bleedLayer);
  var _offsetResult = null;
  if(_bleedOffsetTargets && _bleedOffsetTargets.length) {
    var _bleedApplied = 0;
    var _bleedChanged = 0;
    _pathBleedDebug('bleed offset targets | count=' + _bleedOffsetTargets.length);
    for(var _bti = 0; _bti < _bleedOffsetTargets.length; _bti++) {
      var _bleedTarget = _bleedOffsetTargets[_bti];
      if(!_bleedTarget) continue;
      var _bleedJoinInfo = _getOffsetJoinInfoFromItem(_bleedTarget);
      var _bleedStrokeWidth = 0;
      var _bleedHasStroke = false;
      try {
        _bleedHasStroke = !!_bleedTarget.stroked;
        _bleedStrokeWidth = Number(_bleedTarget.strokeWidth || 0);
      } catch(_eBst0) {
        _bleedHasStroke = false;
        _bleedStrokeWidth = 0;
      }
      _pathBleedDebug(
        'bleed offset target | idx=' + _bti +
        ' | ref=' + _itemRef(_bleedTarget) +
        ' | join=' + _bleedJoinInfo.joinName +
        ' | joinType=' + _bleedJoinInfo.joinType +
        ' | miterLimit=' + _fmtNum3(_bleedJoinInfo.miterLimit) +
        ' | stroked=' + (_bleedHasStroke ? 'yes' : 'no') +
        ' | strokeWidth=' + _fmtNum3(_bleedStrokeWidth)
      );
      var _bleedTargetResult = null;
      if(_bleedHasStroke && _bleedStrokeWidth > 0.01) {
        var _bleedCenterShift = offsetPt / 2;
        _pathBleedDebug(
          'bleed offset target mode | idx=' + _bti +
          ' | mode=stroke-band' +
          ' | centerShift=' + _fmtNum3(_bleedCenterShift) +
          ' | strokeWidthBefore=' + _fmtNum3(_bleedStrokeWidth) +
          ' | strokeWidthAfter=' + _fmtNum3(_bleedStrokeWidth + offsetPt)
        );
        _bleedTargetResult = _applyOffsetToContainer(_bleedTarget, _bleedCenterShift, {
          joinType: _bleedJoinInfo.joinType,
          miterLimit: _bleedJoinInfo.miterLimit
        });
        try {_bleedTarget.strokeWidth = _bleedStrokeWidth + offsetPt;} catch(_eBst1) { }
      } else {
        _pathBleedDebug('bleed offset target mode | idx=' + _bti + ' | mode=plain-offset');
        _bleedTargetResult = _applyOffsetToContainer(_bleedTarget, offsetPt, {
          joinType: _bleedJoinInfo.joinType,
          miterLimit: _bleedJoinInfo.miterLimit
        });
      }
      if(_bleedTargetResult && _bleedTargetResult.applied) _bleedApplied++;
      if(_bleedTargetResult && _bleedTargetResult.changed) _bleedChanged++;
    }
    _offsetResult = {applied: (_bleedApplied > 0), changed: (_bleedChanged > 0)};
    _pathBleedDebug('bleed offset summary | targetsApplied=' + _bleedApplied + ' | targetsChanged=' + _bleedChanged);
  } else {
    _offsetResult = _applyOffsetToContainer(bleedGroup || bleedLayer, offsetPt);
  }
  var removedInnerContours = 0;
  var _postOffsetPathMetrics = _collectPathMetrics(bleedGroup || bleedLayer);
  _logPathMetrics('post-offset', _postOffsetPathMetrics);
  _logPathMetricsDelta(_preOffsetPathMetrics, _postOffsetPathMetrics, 'offset');
  var offsetApplied = false;
  var offsetGeometryChanged = false;
  if(typeof _offsetResult === 'boolean') {
    offsetApplied = _offsetResult;
    offsetGeometryChanged = _offsetResult;
  } else {
    try {offsetApplied = !!_offsetResult.applied;} catch(_eOr0) {offsetApplied = false;}
    try {offsetGeometryChanged = !!_offsetResult.changed;} catch(_eOr1) {offsetGeometryChanged = false;}
  }
  _logContainerState(bleedGroup || bleedLayer, 'after-offset-before-comp');
  var gradientAdjusted = 0;
  if(offsetGeometryChanged) {
    gradientAdjusted = _shiftGradientBySnapshot(gradientSnapshot, offsetPt, bleedGroup || bleedLayer);
    if(gradientAdjusted < 1) gradientAdjusted = _shiftGradientInContainer(bleedGroup || bleedLayer, offsetPt);
  } else {
    _pathBleedDebug('skip compensation | reason=offset-geometry-unchanged');
  }
  var finalTargetAdjusted = 0;
  _pathBleedDebug('gradient compensation count=' + gradientAdjusted);
  _logContainerState(bleedGroup || bleedLayer, 'after-comp');
  var forcedGradientPassCount = 0;
  if(forceBleedStopTest2080) {
    forcedGradientPassCount = _forceGradientStopsInContainer(bleedGroup || bleedLayer, 20, 80);
    _pathBleedDebug('forced gradient 20/80 post-pass count=' + forcedGradientPassCount);
  }
  var forcedSolidRedCount = 0;
  if(forceBleedSolidRedTest) {
    forcedSolidRedCount = _forceSolidRedInContainer(bleedLayer);
    _pathBleedDebug('forced solid-red post-pass count=' + forcedSolidRedCount + ' (container=bleedLayer)');
  }
  _logContainerState(bleedLayer, 'final-bleed-layer');

  var gradientDebugLineCount = 0;
  var gradientDebugLineOriginalCount = 0;
  var gradientDebugLineFinalCount = 0;
  if(drawGradientStopDebugLines) {
    var dbgLines = _drawGradientStopDebugLines(gradientSnapshot, bleedLayer);
    if(dbgLines) {
      gradientDebugLineCount = Number(dbgLines.total || 0);
      gradientDebugLineOriginalCount = Number(dbgLines.original || 0);
      gradientDebugLineFinalCount = Number(dbgLines.bleed || 0);
    }
  } else {
    try {
      var _dbgGroups = bleedLayer.groupItems;
      for(var _dgi = _dbgGroups.length - 1; _dgi >= 0; _dgi--) {
        try {
          if(String(_dbgGroups[_dgi].name || '') === 'SRH_GRAD_DEBUG_LINES') _dbgGroups[_dgi].remove();
        } catch(_eDbgRm1) { }
      }
    } catch(_eDbgRm0) { }
  }

  if(createCutline && cutGroup) {
    var _cutContainers = (cutWorkContainers && cutWorkContainers.length) ? cutWorkContainers : [cutGroup];
    cutCount = 0;
    for(var _cci = 0; _cci < _cutContainers.length; _cci++) {
      var _cutContainer = _cutContainers[_cci];
      if(!_cutContainer) continue;
      var _cutPathCountBefore = _countPathLikeItems(_cutContainer);
      _pathBleedDebug('cutline path summary | container=' + _cci + ' | pathCount=' + _cutPathCountBefore + ' | cutDup=' + cutCount);
      _logCutlineGeometry(_cutContainer, 'before-processing[' + _cci + ']');
      if(outlineText) _outlineTextInContainer(_cutContainer);
      if(outlineStroke) {
        var _cutStrokeWidthPt = _getMaxStrokeWidthInContainer(_cutContainer);
        var _cutStrokeOffsetPt = _cutStrokeWidthPt / 2;
        _pathBleedDebug('cutline outline-stroke fallback | container=' + _cci + ' | strokeWidthPt=' + _fmtNum3(_cutStrokeWidthPt) + ' | offsetPt=' + _fmtNum3(_cutStrokeOffsetPt));
        if(_cutStrokeOffsetPt > 0.01) {
          var _cutTargets = _collectNativePathRoots(_cutContainer);
          var _cutTargetApplied = 0;
          var _cutTargetChanged = 0;
          _pathBleedDebug('cutline outline-stroke targets | container=' + _cci + ' | count=' + _cutTargets.length);
          for(var _cti = 0; _cti < _cutTargets.length; _cti++) {
            var _cutTarget = _cutTargets[_cti];
            if(!_cutTarget) continue;
            var _cutJoinInfo = _getOffsetJoinInfoFromItem(_cutTarget);
            _pathBleedDebug(
              'cutline outline-stroke target | container=' + _cci +
              ' | idx=' + _cti +
              ' | ref=' + _itemRef(_cutTarget) +
              ' | join=' + _cutJoinInfo.joinName +
              ' | joinType=' + _cutJoinInfo.joinType +
              ' | miterLimit=' + _fmtNum3(_cutJoinInfo.miterLimit)
            );
            var _cutOffsetResult = _applyOffsetToContainer(_cutTarget, _cutStrokeOffsetPt, {
              joinType: _cutJoinInfo.joinType,
              miterLimit: _cutJoinInfo.miterLimit
            });
            if(_cutOffsetResult && _cutOffsetResult.applied) _cutTargetApplied++;
            if(_cutOffsetResult && _cutOffsetResult.changed) _cutTargetChanged++;
          }
          _pathBleedDebug(
            'cutline outline-stroke summary | container=' + _cci +
            ' | pathCount=' + _countPathLikeItems(_cutContainer) +
            ' | targetsApplied=' + _cutTargetApplied +
            ' | targetsChanged=' + _cutTargetChanged
          );
          _logCutlineGeometry(_cutContainer, 'after-outline-stroke-offset[' + _cci + ']');
        } else {
          _pathBleedDebug('cutline outline-stroke summary | container=' + _cci + ' | skipped | reason=no-stroke-width');
        }
      }
      _setCutlineStyleOnItem(_cutContainer, {outlineText: outlineText, outlineStroke: false});
      _logCutlineGeometry(_cutContainer, 'after-style-before-prune[' + _cci + ']');
      _pruneCutlineContainer(_cutContainer);
      _auditStructure(_cutContainer, 'cut-after-style[' + _cci + ']');
      _logCutlineGeometry(_cutContainer, 'after-style[' + _cci + ']');
      if(autoWeld) {
        var welded = _uniteContainerPathfinderAdd(_cutContainer);
        if(!welded) welded = _uniteContainer(_cutContainer);
        _setCutlineStyleOnItem(_cutContainer, {outlineText: false, outlineStroke: false});
        _logCutlineGeometry(_cutContainer, 'after-weld-before-prune[' + _cci + ']');
        _pruneCutlineContainer(_cutContainer);
        _auditStructure(_cutContainer, 'cut-after-weld[' + _cci + ']');
        _logCutlineGeometry(_cutContainer, 'after-weld[' + _cci + ']');
      }
      var _cutPathCountAfter = _countPathLikeItems(_cutContainer);
      if(_cutPathCountAfter > 0) cutCount += _cutPathCountAfter;
      _pruneCutlineContainer(_cutContainer);
      _auditStructure(_cutContainer, 'cut-final[' + _cci + ']');
      _logCutlineGeometry(_cutContainer, 'final[' + _cci + ']');
    }
  }

  // Run final-target compensation after offset output is fully materialized on bleed layer.
  try {app.redraw();} catch(_ePbRt0) { }
  _pathBleedDebug('final-target precheck | offsetGeometryChanged=' + (offsetGeometryChanged ? 'yes' : 'no'));
  finalTargetAdjusted = _finalizeBleedGradientStopsByFinalItems(bleedLayer, offsetPt, gradientSnapshot, offsetGeometryChanged);
  _pathBleedDebug('final-target gradient stop pass count=' + finalTargetAdjusted);
  _logContainerState(bleedLayer, 'after-final-target-pass');

  var offsetCount = _countPathLikeItems(bleedGroup || bleedLayer);
  try {doc.selection = null;} catch(_eSel1) { }

  var msg = "Path bleed applied (grouped). Originals: " + originalCount + ", targets: " + movedCount + ", bleed paths: " + offsetCount;
  msg += ", doc gradient-fill objects: " + docGradientFillCount;
  if(gradientAdjusted > 0) msg += ", gradients compensated: " + gradientAdjusted;
  if(finalTargetAdjusted > 0) msg += ", final-target gradients adjusted: " + finalTargetAdjusted;
  if(closedOpenPathCount > 0) msg += ", open paths auto-closed: " + closedOpenPathCount;
  if(forceBleedStopTest2080) msg += ", forced-gradient-20/80 pass: " + forcedGradientPassCount;
  if(forceBleedSolidRedTest) msg += ", forced-solid-red pass: " + forcedSolidRedCount;
  if(createCutline) msg += ", cutlines created: " + cutCount + (autoWeld ? " (auto-weld on)" : " (auto-weld off)");
  if(drawGradientStopDebugLines) msg += ", gradient debug lines: " + gradientDebugLineCount + " (original: " + gradientDebugLineOriginalCount + ", bleed: " + gradientDebugLineFinalCount + ")";
  if(!offsetApplied) msg += ". Note: Offset effect fallback was limited on this selection.";
  _srh_setBleedLayerOrder(cutLayer, originalLayer, bleedLayer);
  try {if(originalLayer) originalLayer.visible = false;} catch(_eOlHide2) { }
  return _pathBleedWithDbg(msg);
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
  } catch(_eCk0) {return null;}
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
  } catch(_eReqAb0) {requestedArtboardIndices = [];}
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
        try {ar = doc.artboards[ai].artboardRect;} catch(_eAbIdx0) {ar = null;}
        if(_rectsEqual(rect, ar)) return ai;
      }
      return -1;
    }
    function _rectFromUnknown(obj) {
      if(!obj) return null;
      var r = null;
      try {r = obj.artboardRect;} catch(_eRu0) {r = null;}
      if(r && r.length === 4) return r;
      try {r = obj.visibleBounds;} catch(_eRu1) {r = null;}
      if(r && r.length === 4) return r;
      try {r = obj.geometricBounds;} catch(_eRu2) {r = null;}
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
      try {selAny = doc.selection;} catch(_eAbSel1) {selAny = null;}
      if(selAny && selAny.length) {
        for(var sA = 0; sA < selAny.length; sA++) {
          var si = selAny[sA];
          if(!si) continue;
          var selIdx = -1;
          try {selIdx = Number(si.artboardIndex);} catch(_eAbSelIdx0) {selIdx = -1;}
          if(!(selIdx >= 0)) {
            try {selIdx = Number(si.index);} catch(_eAbSelIdx1) {selIdx = -1;}
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
      try {isClippedGroup = (p.typename === "GroupItem" && p.clipped);} catch(_eCg0) {isClippedGroup = false;}
      if(isClippedGroup) return p;
      try {p = p.parent;} catch(_eCg1) {p = null;}
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
  try {activeIdx = doc.artboards.getActiveArtboardIndex();} catch(_eAbL0) {activeIdx = 0;}
  for(var i = 0; i < doc.artboards.length; i++) {
    var ab = doc.artboards[i];
    var r = null;
    try {r = _normalizeRect4(ab.artboardRect);} catch(_eAbL1) {r = null;}
    if(!r) continue;
    var w = r[2] - r[0];
    var h = r[1] - r[3];
    var name = '';
    try {name = String(ab.name || '');} catch(_eAbL2) {name = '';}
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
  try {debug.artboardCount = doc.artboards.length;} catch(_eDbg0) {debug.artboardCount = -1;}
  try {debug.activeIndex = doc.artboards.getActiveArtboardIndex();} catch(_eDbg1) {debug.activeIndex = -1;}

  for(var i = 0; i < doc.artboards.length; i++) {
    var row = {index: i, name: '', rectType: 'none', rectRaw: '', rectParsed: null, selectedFlag: null};
    var ab = doc.artboards[i];
    try {row.name = String(ab.name || '');} catch(_eDbg2) {row.name = '';}
    try {row.selectedFlag = !!ab.selected;} catch(_eDbg3) {row.selectedFlag = null;}
    try {
      var r = ab.artboardRect;
      row.rectType = Object.prototype.toString.call(r);
      try {row.rectRaw = String(r);} catch(_eDbg4) {row.rectRaw = '[unprintable]';}
      try {
        if(r && typeof r.length !== 'undefined') {
          row.rectParsed = [Number(r[0]), Number(r[1]), Number(r[2]), Number(r[3])];
        } else if(r && typeof r.left !== 'undefined') {
          row.rectParsed = [Number(r.left), Number(r.top), Number(r.right), Number(r.bottom)];
        }
      } catch(_eDbg5) {row.rectParsed = null;}
    } catch(_eDbg6) { }
    debug.items.push(row);
  }
  return JSON.stringify(debug);
}

function signarama_helper_corebridge_openProofPath(pathText) {
  try {
    var rawPath = String(pathText == null ? '' : pathText).replace(/^\s+|\s+$/g, '');
    if(!rawPath) return 'Error: No proof path supplied.';
    var fileRef = new File(rawPath);
    if(!fileRef.exists) return 'Error: Proof file not found: ' + rawPath;
    app.open(fileRef);
    return 'Proof document opened.';
  } catch(e) {
    return 'Error: ' + e.message;
  }
}

function signarama_helper_corebridge_updatePageNumbers() {
  function _trim(v) {return String(v == null ? '' : v).replace(/^\s+|\s+$/g, '');}
  function _artboardContainsPoint(rect, x, y) {
    var l = Number(rect[0]), t = Number(rect[1]), r = Number(rect[2]), b = Number(rect[3]);
    return x >= Math.min(l, r) && x <= Math.max(l, r) && y >= Math.min(b, t) && y <= Math.max(b, t);
  }
  function _distanceToArtboardCenter(rect, x, y) {
    var l = Number(rect[0]), t = Number(rect[1]), r = Number(rect[2]), b = Number(rect[3]);
    var cx = (l + r) / 2;
    var cy = (t + b) / 2;
    var dx = x - cx, dy = y - cy;
    return (dx * dx) + (dy * dy);
  }
  try {
    if(!app.documents.length) return 'No open document.';
    var doc = app.activeDocument;
    var total = doc.artboards.length;
    if(!total) return 'No artboards.';
    var proofFolderPath = '';
    try {
      if(doc.fullName) {
        var _folderRef = doc.fullName.parent;
        if(_folderRef && _folderRef.fsName) proofFolderPath = String(_folderRef.fsName);
      }
    } catch(_ePfLive) { proofFolderPath = ''; }

    var frames = doc.textFrames;
    var updated = 0;
    var pathUpdated = 0;
    for(var i = 0; i < frames.length; i++) {
      var tf = frames[i];
      var name = '';
      try {name = String(tf.name || '');} catch(_eNm) {name = '';}
      var normName = _trim(name).toLowerCase();
      if(normName === 'file path text') {
        try {
          var curPath = String(tf.contents == null ? '' : tf.contents);
          if(curPath !== proofFolderPath) {
            tf.contents = proofFolderPath;
            pathUpdated++;
          }
        } catch(_eSetPath) { }
      }
      if(normName !== 'page number') continue;

      var gb = tf.geometricBounds; // [L,T,R,B]
      var x = (Number(gb[0]) + Number(gb[2])) / 2;
      var y = (Number(gb[1]) + Number(gb[3])) / 2;

      var idx = -1;
      var bestDist = Number.POSITIVE_INFINITY;
      for(var a = 0; a < total; a++) {
        var ab = doc.artboards[a].artboardRect;
        if(_artboardContainsPoint(ab, x, y)) {
          idx = a;
          break;
        }
        var d = _distanceToArtboardCenter(ab, x, y);
        if(d < bestDist) {
          bestDist = d;
          idx = a;
        }
      }
      if(idx < 0) idx = 0;
      try {
        var nextValue = String(idx + 1) + '/' + String(total);
        var currentValue = String(tf.contents == null ? '' : tf.contents);
        if(currentValue !== nextValue) {
          tf.contents = nextValue;
          updated++;
        }
      } catch(_eSetPg) { }
    }
    return 'Updated page numbers: ' + updated + ' / ' + frames.length + ' text frames scanned. Updated file paths: ' + pathUpdated + '.';
  } catch(e) {
    return 'Error: ' + e.message;
  }
}

var _srhCorebridgeFlashTaskId = null;
var _srhCorebridgeFlashState = null;
var _srhCorebridgeFlashTickCount = 0;
var _srhCorebridgeFlashArrowLayerName = 'SRH Flash Arrows';
function _srh_corebridge_flashDebug(msg) {
  try {$.writeln('[SRH][CorebridgeFlash] ' + String(msg));} catch(_eFlashDbg) { }
}
function _srh_corebridge_makeRgb(r, g, b) {
  var c = new RGBColor();
  c.red = r;
  c.green = g;
  c.blue = b;
  return c;
}
function _srh_corebridge_parseFlashFieldNames(rawText) {
  function _trim(v) {return String(v == null ? '' : v).replace(/^\s+|\s+$/g, '');}
  var out = [];
  var seen = {};
  var lines = String(rawText == null ? '' : rawText).split(/\r?\n/);
  for(var i = 0; i < lines.length; i++) {
    var line = _trim(lines[i]);
    if(!line) continue;
    if(line.indexOf('#') === 0) continue;
    var idx = line.indexOf('->');
    var targetName = idx >= 0 ? _trim(line.substring(idx + 2)) : line;
    if(!targetName) continue;
    var key = targetName.toLowerCase();
    if(seen[key]) continue;
    seen[key] = true;
    out.push(targetName);
  }
  _srh_corebridge_flashDebug('Parsed flash field rows. Input lines=' + lines.length + ' | unique targets=' + out.length + ' | targets=' + out.join(', '));
  return out;
}
function _srh_corebridge_setTextFrameColor(textFrame, colorValue) {
  if(!textFrame || !colorValue) return false;
  try {
    if(textFrame.textRange && textFrame.textRange.characterAttributes) {
      textFrame.textRange.characterAttributes.fillColor = colorValue;
      return true;
    }
  } catch(_eTfClr0) { }
  try {
    if(textFrame.characters && textFrame.characters.length) {
      for(var i = 0; i < textFrame.characters.length; i++) {
        try {textFrame.characters[i].characterAttributes.fillColor = colorValue;} catch(_eTfClr1) { }
      }
      return true;
    }
  } catch(_eTfClr2) { }
  return false;
}
function _srh_corebridge_getArrowLayer(doc, createIfMissing) {
  if(!doc) return null;
  var layer = null;
  try {layer = doc.layers.getByName(_srhCorebridgeFlashArrowLayerName);} catch(_eArrowLayerGet) {layer = null;}
  if(!layer && createIfMissing) {
    try {
      layer = doc.layers.add();
      layer.name = _srhCorebridgeFlashArrowLayerName;
      layer.visible = true;
      layer.printable = false;
      layer.locked = true;
    } catch(_eArrowLayerCreate) {layer = null;}
  }
  return layer;
}
function _srh_corebridge_finalizeArrowLayer(doc) {
  if(!doc) return;
  var layer = _srh_corebridge_getArrowLayer(doc, false);
  if(!layer) return;
  var hasItems = false;
  try {hasItems = !!(layer.pageItems && layer.pageItems.length > 0);} catch(_eArrowHasItems2) {hasItems = true;}
  if(!hasItems) {
    try {layer.locked = false;} catch(_eArrowUnlockDel) { }
    try {layer.remove();} catch(_eArrowDel) { }
    return;
  }
  try {layer.locked = true;} catch(_eArrowLockKeep) { }
}
function _srh_corebridge_removeArrow(entry) {
  if(!entry || !entry.arrowGroup) return;
  try {entry.arrowGroup.remove();} catch(_eArrowRm) { }
  entry.arrowGroup = null;
}
function _srh_corebridge_createArrowForTextFrame(doc, textFrame) {
  if(!doc || !textFrame) return null;
  var layer = _srh_corebridge_getArrowLayer(doc, true);
  if(!layer) return null;
  var group = null;
  try {
    var gb = textFrame.geometricBounds; // [L,T,R,B]
    var left = Number(gb[0]);
    var top = Number(gb[1]);
    var right = Number(gb[2]);
    var bottom = Number(gb[3]);
    if(!isFinite(left) || !isFinite(top) || !isFinite(right) || !isFinite(bottom)) return null;
    var h = Math.max(1, top - bottom);
    var centerY = (top + bottom) / 2;
    var shaftLen = Math.max(_srh_mm2ptDoc(14), h * 1.1);
    var gapFromText = Math.max(_srh_mm2ptDoc(3), h * 0.25);
    var arrowHeadLen = Math.max(_srh_mm2ptDoc(5), h * 0.4);
    var arrowHeadHalfHeight = Math.max(_srh_mm2ptDoc(3), h * 0.3);
    var strokeW = _srh_pxStrokeDoc(3);
    var tipX = right + gapFromText;
    var tailX = tipX + shaftLen;

    var red = _srh_corebridge_makeRgb(255, 0, 0);
    try {layer.locked = false;} catch(_eArrowUnlock) { }

    group = layer.groupItems.add();
    group.name = 'SRH_FlashArrow';

    var shaft = group.pathItems.add();
    shaft.setEntirePath([[tailX, centerY], [tipX, centerY]]);
    shaft.stroked = true;
    shaft.filled = false;
    shaft.strokeWidth = strokeW;
    shaft.strokeColor = red;

    var headA = group.pathItems.add();
    headA.setEntirePath([[tipX, centerY], [tipX + arrowHeadLen, centerY + arrowHeadHalfHeight]]);
    headA.stroked = true;
    headA.filled = false;
    headA.strokeWidth = strokeW;
    headA.strokeColor = red;

    var headB = group.pathItems.add();
    headB.setEntirePath([[tipX, centerY], [tipX + arrowHeadLen, centerY - arrowHeadHalfHeight]]);
    headB.stroked = true;
    headB.filled = false;
    headB.strokeWidth = strokeW;
    headB.strokeColor = red;

    try {layer.locked = true;} catch(_eArrowRelock) { }
    return group;
  } catch(_eArrowCreate) {
    try {if(group) group.remove();} catch(_eArrowCleanup) { }
    try {if(layer) layer.locked = true;} catch(_eArrowRelock2) { }
    return null;
  }
}
function _srh_corebridge_stopFlashing(resetToBlack) {
  if(_srhCorebridgeFlashTaskId != null) {
    _srh_corebridge_flashDebug('Stopping flash task id=' + _srhCorebridgeFlashTaskId + ' (resetToBlack=' + (resetToBlack ? 'yes' : 'no') + ')');
    try {app.cancelTask(_srhCorebridgeFlashTaskId);} catch(_eCancelFlash) { }
    _srhCorebridgeFlashTaskId = null;
  }
  if(resetToBlack && _srhCorebridgeFlashState && _srhCorebridgeFlashState.entries) {
    var black = _srh_corebridge_makeRgb(0, 0, 0);
    for(var i = 0; i < _srhCorebridgeFlashState.entries.length; i++) {
      var entry = _srhCorebridgeFlashState.entries[i];
      try {_srh_corebridge_setTextFrameColor(entry.frame, black);} catch(_eResetClr) { }
      _srh_corebridge_removeArrow(entry);
    }
  }
  if(_srhCorebridgeFlashState && _srhCorebridgeFlashState.doc) {
    _srh_corebridge_finalizeArrowLayer(_srhCorebridgeFlashState.doc);
  }
  _srhCorebridgeFlashState = null;
}
function _srh_corebridge_flashTick() {
  _srhCorebridgeFlashTickCount++;
  if(!_srhCorebridgeFlashState || !_srhCorebridgeFlashState.entries || !_srhCorebridgeFlashState.entries.length) {
    _srh_corebridge_flashDebug('Tick #' + _srhCorebridgeFlashTickCount + ' skipped (no active entries).');
    _srh_corebridge_stopFlashing(false);
    return false;
  }
  var entries = _srhCorebridgeFlashState.entries;
  _srh_corebridge_flashDebug('Tick #' + _srhCorebridgeFlashTickCount + ' starting. entries=' + entries.length + ' | mode=state-check');
  var black = _srh_corebridge_makeRgb(0, 0, 0);
  for(var i = entries.length - 1; i >= 0; i--) {
    var entry = entries[i];
    if(!entry || !entry.frame) {
      entries.splice(i, 1);
      continue;
    }
    var currentValue = '';
    var frameName = '';
    try {frameName = String(entry.frame.name || '');} catch(_eFlashName) {frameName = '';}
    try {currentValue = String(entry.frame.contents == null ? '' : entry.frame.contents);} catch(_eFlashRead) {
      _srh_corebridge_flashDebug('Tick #' + _srhCorebridgeFlashTickCount + ' removing entry index=' + i + ' (failed reading contents).');
      entries.splice(i, 1);
      continue;
    }
    if(currentValue !== entry.baseValue) {
      _srh_corebridge_flashDebug('Tick #' + _srhCorebridgeFlashTickCount + ' resolved field "' + frameName + '". base="' + entry.baseValue + '" current="' + currentValue + '" -> reset BLACK + remove.');
      try {_srh_corebridge_setTextFrameColor(entry.frame, black);} catch(_eFlashDone) { }
      _srh_corebridge_removeArrow(entry);
      entries.splice(i, 1);
      continue;
    }
    var isSelected = false;
    try {isSelected = !!entry.frame.selected;} catch(_eFlashSelected) {isSelected = false;}
    if(isSelected) {
      try {_srh_corebridge_setTextFrameColor(entry.frame, black);} catch(_eFlashDone2) { }
      _srh_corebridge_removeArrow(entry);
      entries.splice(i, 1);
      continue;
    }
  }
  _srhCorebridgeFlashState.isRed = true;
  _srh_corebridge_flashDebug('Tick #' + _srhCorebridgeFlashTickCount + ' complete. remaining=' + entries.length + ' | activeColorNow=RED');
  if(!entries.length) _srh_corebridge_stopFlashing(false);
  return !!(entries && entries.length);
}

function signarama_helper_corebridge_flashTickTask() {
  try {
    var isActive = _srh_corebridge_flashTick();
    var remaining = 0;
    var colorName = 'BLACK';
    if(_srhCorebridgeFlashState && _srhCorebridgeFlashState.entries) remaining = _srhCorebridgeFlashState.entries.length;
    if(_srhCorebridgeFlashState && _srhCorebridgeFlashState.isRed) colorName = 'RED';
    if(isActive && remaining > 0) return 'ACTIVE|' + remaining + '|' + _srhCorebridgeFlashTickCount + '|' + colorName;
    return 'INACTIVE|0|' + _srhCorebridgeFlashTickCount + '|BLACK';
  } catch(_eFlashTaskTick) {
    _srh_corebridge_flashDebug('Tick task exception: ' + (_eFlashTaskTick && _eFlashTaskTick.message ? _eFlashTaskTick.message : _eFlashTaskTick));
    return 'ERROR|' + (_eFlashTaskTick && _eFlashTaskTick.message ? _eFlashTaskTick.message : _eFlashTaskTick);
  }
}
try {this.signarama_helper_corebridge_flashTickTask = signarama_helper_corebridge_flashTickTask;} catch(_eFlashTaskBind) { }
function signarama_helper_corebridge_flashGetState() {
  try {
    var active = (_srhCorebridgeFlashState && _srhCorebridgeFlashState.entries && _srhCorebridgeFlashState.entries.length) ? 1 : 0;
    var remaining = active ? _srhCorebridgeFlashState.entries.length : 0;
    var colorName = (_srhCorebridgeFlashState && _srhCorebridgeFlashState.isRed) ? 'RED' : 'BLACK';
    var taskId = (_srhCorebridgeFlashTaskId == null) ? -1 : _srhCorebridgeFlashTaskId;
    return 'STATE|' + active + '|' + remaining + '|' + _srhCorebridgeFlashTickCount + '|' + colorName + '|' + taskId;
  } catch(_eFlashState) {
    return 'ERROR|' + (_eFlashState && _eFlashState.message ? _eFlashState.message : _eFlashState);
  }
}
try {this.signarama_helper_corebridge_flashGetState = signarama_helper_corebridge_flashGetState;} catch(_eFlashStateBind) { }
function signarama_helper_corebridge_flashTickTaskRunner() {
  try {return signarama_helper_corebridge_flashTickTask();}
  catch(_eFlashTaskRunner) {return 'ERROR|' + (_eFlashTaskRunner && _eFlashTaskRunner.message ? _eFlashTaskRunner.message : _eFlashTaskRunner);}
}
try {this.signarama_helper_corebridge_flashTickTaskRunner = signarama_helper_corebridge_flashTickTaskRunner;} catch(_eFlashTaskRunnerBind) { }

function _srh_corebridge_startFlashing(doc, flashFieldsText) {
  _srh_corebridge_flashDebug('Start flashing requested. Raw text="' + String(flashFieldsText == null ? '' : flashFieldsText) + '"');
  var names = _srh_corebridge_parseFlashFieldNames(flashFieldsText);
  _srh_corebridge_stopFlashing(true);
  _srhCorebridgeFlashTickCount = 0;
  if(!doc || !names.length) return {requested: names.length, found: 0};

  var namesLookup = {};
  for(var n = 0; n < names.length; n++) {
    namesLookup[String(names[n]).toLowerCase()] = true;
  }

  var entries = [];
  var allFrames = null;
  try {allFrames = doc.textFrames;} catch(_eAllFlashTf) {allFrames = null;}
  if(!allFrames || !allFrames.length) return {requested: names.length, found: 0};
  for(var i = 0; i < allFrames.length; i++) {
    var tf = allFrames[i];
    var tfName = '';
    try {tfName = String(tf.name || '');} catch(_eTfNameFlash) {tfName = '';}
    if(!tfName) continue;
    if(!namesLookup[tfName.toLowerCase()]) continue;
    var baseValue = '';
    try {baseValue = String(tf.contents == null ? '' : tf.contents);} catch(_eTfBaseFlash) {baseValue = '';}
    entries.push({frame: tf, baseValue: baseValue});
  }
  if(!entries.length) {
    _srh_corebridge_flashDebug('No matching text frames found for requested flash names.');
    return {requested: names.length, found: 0};
  }

  _srh_corebridge_flashDebug('Starting flashing with ' + entries.length + ' matched text frames.');
  _srhCorebridgeFlashState = {
    doc: doc,
    entries: entries,
    isRed: true
  };
  var red = _srh_corebridge_makeRgb(255, 0, 0);
  for(var e = 0; e < entries.length; e++) {
    try {_srh_corebridge_setTextFrameColor(entries[e].frame, red);} catch(_eFlashInitRed) { }
    entries[e].arrowGroup = _srh_corebridge_createArrowForTextFrame(doc, entries[e].frame);
  }
  _srh_corebridge_flashTick();
  try {
    var tickTaskCode = '((typeof signarama_helper_corebridge_flashTickTaskRunner === "function") ? signarama_helper_corebridge_flashTickTaskRunner() : ((typeof $ !== "undefined" && $.global && typeof $.global.signarama_helper_corebridge_flashTickTaskRunner === "function") ? $.global.signarama_helper_corebridge_flashTickTaskRunner() : ((typeof signarama_helper_corebridge_flashTickTask === "function") ? signarama_helper_corebridge_flashTickTask() : "ERROR|flash tick task missing")))';
    _srhCorebridgeFlashTaskId = app.scheduleTask(tickTaskCode, 300, true);
  } catch(_eScheduleFlash) {
    _srhCorebridgeFlashTaskId = null;
  }
  _srh_corebridge_flashDebug('Flash state primed.');
  return {requested: names.length, found: entries.length};
}

function signarama_helper_corebridge_createProofFromData(pathText, dataJson, mappingText, flashFieldsText) {
  function _trim(v) {
    return String(v == null ? '' : v).replace(/^\s+|\s+$/g, '');
  }
  // Legacy safety guard: some cached CEP/JSX runtimes may still evaluate
  // older return-string variants that referenced linkPlacementRes.
  var linkPlacementRes = '';
  function _todayDdMmYy() {
    var d = new Date();
    function _pad(v) {return (v < 10 ? '0' : '') + String(v);}
    return _pad(d.getDate()) + '/' + _pad(d.getMonth() + 1) + '/' + String(d.getFullYear()).slice(-2);
  }
  function _toTextValue(v) {
    if(v == null) return '';
    if(typeof v === 'string') return v;
    if(typeof v === 'number' || typeof v === 'boolean') return String(v);
    try {return JSON.stringify(v);} catch(_eJsonTxt) {return String(v);}
  }
  function _readByPath(obj, keyPath) {
    if(!obj || !keyPath) return undefined;
    if(obj.hasOwnProperty && obj.hasOwnProperty(keyPath)) return obj[keyPath];
    var parts = String(keyPath).split('.');
    var cur = obj;
    for(var i = 0; i < parts.length; i++) {
      var part = _trim(parts[i]);
      if(!part) continue;
      if(cur == null || typeof cur !== 'object' || !(part in cur)) return undefined;
      cur = cur[part];
    }
    return cur;
  }
  function _parseMappings(rawText) {
    var out = [];
    var lines = String(rawText == null ? '' : rawText).split(/\r?\n/);
    for(var i = 0; i < lines.length; i++) {
      var line = _trim(lines[i]);
      if(!line) continue;
      if(line.indexOf('#') === 0) continue;
      var idx = line.indexOf('->');
      if(idx < 0) continue;
      var sourceKey = _trim(line.substring(0, idx));
      var targetName = _trim(line.substring(idx + 2));
      if(!sourceKey || !targetName) continue;
      out.push({sourceKey: sourceKey, targetName: targetName});
    }
    return out;
  }
  function _looksLikeSvgText(v) {
    var s = String(v == null ? '' : v);
    return s.indexOf('<svg') >= 0 && s.indexOf('</svg>') >= 0;
  }
  function _placeSvgIntoTarget(doc, targetItem, svgText) {
    var tmpFile = null;
    var placed = null;
    try {
      var tmpFolder = Folder.temp;
      if(!tmpFolder || !tmpFolder.exists) return false;
      var unique = (new Date().getTime()) + '_' + Math.floor(Math.random() * 100000);
      tmpFile = new File(tmpFolder.fsName + '/srh-qr-' + unique + '.svg');
      tmpFile.encoding = 'UTF-8';
      if(!tmpFile.open('w')) return false;
      tmpFile.write(String(svgText == null ? '' : svgText));
      tmpFile.close();

      placed = doc.placedItems.add();
      placed.file = tmpFile;

      var tb = targetItem.geometricBounds; // [L,T,R,B]
      var tw = Math.abs(Number(tb[2]) - Number(tb[0]));
      var th = Math.abs(Number(tb[1]) - Number(tb[3]));
      var pb = placed.geometricBounds;
      var pw = Math.abs(Number(pb[2]) - Number(pb[0]));
      var ph = Math.abs(Number(pb[1]) - Number(pb[3]));
      if(pw <= 0 || ph <= 0 || tw <= 0 || th <= 0) return true;

      var sx = (tw / pw) * 100;
      var sy = (th / ph) * 100;
      placed.resize(sx, sy, true, true, true, true, sx, Transformation.TOPLEFT);
      placed.position = [Number(tb[0]), Number(tb[1])];
      return true;
    } catch(_ePlaceSvg) {
      return false;
    } finally {
      try {if(tmpFile && tmpFile.opened) tmpFile.close();} catch(_eCloseTmp) { }
    }
  }
  function _placeSvgOnActiveArtboard(doc, svgText) {
    var tmpFile = null;
    var placed = null;
    try {
      var tmpFolder = Folder.temp;
      if(!tmpFolder || !tmpFolder.exists) return false;
      var unique = (new Date().getTime()) + '_' + Math.floor(Math.random() * 100000);
      tmpFile = new File(tmpFolder.fsName + '/srh-qr-debug-' + unique + '.svg');
      tmpFile.encoding = 'UTF-8';
      if(!tmpFile.open('w')) return false;
      tmpFile.write(String(svgText == null ? '' : svgText));
      tmpFile.close();

      placed = doc.placedItems.add();
      placed.file = tmpFile;

      var activeIdx = 0;
      try {activeIdx = doc.artboards.getActiveArtboardIndex();} catch(_eAbIdx) {activeIdx = 0;}
      var ab = doc.artboards[activeIdx].artboardRect; // [L,T,R,B]
      var left = Number(ab[0]) + _srh_mm2ptDoc(10);
      var top = Number(ab[1]) - _srh_mm2ptDoc(10);

      var targetSize = _srh_mm2ptDoc(35);
      var pb = placed.geometricBounds;
      var pw = Math.abs(Number(pb[2]) - Number(pb[0]));
      var ph = Math.abs(Number(pb[1]) - Number(pb[3]));
      if(pw > 0 && ph > 0) {
        var sx = (targetSize / pw) * 100;
        var sy = (targetSize / ph) * 100;
        placed.resize(sx, sy, true, true, true, true, sx, Transformation.TOPLEFT);
      }
      placed.position = [left, top];
      return true;
    } catch(_eDebugPlace) {
      return false;
    } finally {
      try {if(tmpFile && tmpFile.opened) tmpFile.close();} catch(_eCloseTmp2) { }
    }
  }
  function _looksLikePngDataUrl(v) {
    return /^data:image\/png;base64,/i.test(String(v == null ? '' : v));
  }
  function _decodeBase64(base64Text) {
    var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=';
    var str = String(base64Text == null ? '' : base64Text).replace(/[^A-Za-z0-9\+\/\=]/g, '');
    var out = '';
    var i = 0;
    while(i < str.length) {
      var enc1 = chars.indexOf(str.charAt(i++));
      var enc2 = chars.indexOf(str.charAt(i++));
      var enc3 = chars.indexOf(str.charAt(i++));
      var enc4 = chars.indexOf(str.charAt(i++));
      var chr1 = (enc1 << 2) | (enc2 >> 4);
      var chr2 = ((enc2 & 15) << 4) | (enc3 >> 2);
      var chr3 = ((enc3 & 3) << 6) | enc4;
      out += String.fromCharCode(chr1);
      if(enc3 !== 64) out += String.fromCharCode(chr2);
      if(enc4 !== 64) out += String.fromCharCode(chr3);
    }
    return out;
  }
  function _placePngDataUrlOnActiveArtboard(doc, pngDataUrl) {
    var tmpFile = null;
    var placed = null;
    try {
      var raw = String(pngDataUrl == null ? '' : pngDataUrl);
      var base64Idx = raw.indexOf('base64,');
      if(base64Idx < 0) return false;
      var b64 = raw.substring(base64Idx + 7);
      if(!b64) return false;
      var binary = _decodeBase64(b64);
      if(!binary) return false;

      var tmpFolder = Folder.temp;
      if(!tmpFolder || !tmpFolder.exists) return false;
      var unique = (new Date().getTime()) + '_' + Math.floor(Math.random() * 100000);
      tmpFile = new File(tmpFolder.fsName + '/srh-qr-debug-' + unique + '.png');
      tmpFile.encoding = 'BINARY';
      if(!tmpFile.open('w')) return false;
      tmpFile.write(binary);
      tmpFile.close();

      placed = doc.placedItems.add();
      placed.file = tmpFile;

      var activeIdx = 0;
      try {activeIdx = doc.artboards.getActiveArtboardIndex();} catch(_eAbIdx2) {activeIdx = 0;}
      var ab = doc.artboards[activeIdx].artboardRect; // [L,T,R,B]
      var left = Number(ab[0]) + _srh_mm2ptDoc(10);
      var top = Number(ab[1]) - _srh_mm2ptDoc(10);

      var targetSize = _srh_mm2ptDoc(35);
      var pb = placed.geometricBounds;
      var pw = Math.abs(Number(pb[2]) - Number(pb[0]));
      var ph = Math.abs(Number(pb[1]) - Number(pb[3]));
      if(pw > 0 && ph > 0) {
        var sx = (targetSize / pw) * 100;
        var sy = (targetSize / ph) * 100;
        placed.resize(sx, sy, true, true, true, true, sx, Transformation.TOPLEFT);
      }
      placed.position = [left, top];
      return true;
    } catch(_ePlacePng) {
      return false;
    } finally {
      try {if(tmpFile && tmpFile.opened) tmpFile.close();} catch(_eClosePng) { }
    }
  }
  function _placePngDataUrlIntoTarget(doc, targetItem, pngDataUrl) {
    var tmpFile = null;
    var placed = null;
    try {
      var raw = String(pngDataUrl == null ? '' : pngDataUrl);
      var base64Idx = raw.indexOf('base64,');
      if(base64Idx < 0) return false;
      var b64 = raw.substring(base64Idx + 7);
      if(!b64) return false;
      var binary = _decodeBase64(b64);
      if(!binary) return false;

      var tmpFolder = Folder.temp;
      if(!tmpFolder || !tmpFolder.exists) return false;
      var unique = (new Date().getTime()) + '_' + Math.floor(Math.random() * 100000);
      tmpFile = new File(tmpFolder.fsName + '/srh-qr-target-' + unique + '.png');
      tmpFile.encoding = 'BINARY';
      if(!tmpFile.open('w')) return false;
      tmpFile.write(binary);
      tmpFile.close();

      placed = doc.placedItems.add();
      placed.file = tmpFile;

      var tb = targetItem.geometricBounds; // [L,T,R,B]
      var tLeft = Number(tb[0]);
      var tTop = Number(tb[1]);
      var tRight = Number(tb[2]);
      var tBottom = Number(tb[3]);
      var tw = Math.abs(tRight - tLeft);
      var th = Math.abs(tTop - tBottom);
      if(tw <= 0 || th <= 0) return false;

      var gap = _srh_mm2ptDoc(2); // 2mm inset from target rectangle bounds
      var innerW = Math.max(1, tw - (gap * 2));
      var innerH = Math.max(1, th - (gap * 2));

      var pb = placed.geometricBounds;
      var pw = Math.abs(Number(pb[2]) - Number(pb[0]));
      var ph = Math.abs(Number(pb[1]) - Number(pb[3]));
      if(pw <= 0 || ph <= 0) return false;

      // Fit inside target inner box while preserving aspect ratio.
      var fitScale = Math.min(innerW / pw, innerH / ph);
      var sx = fitScale * 100;
      var sy = fitScale * 100;
      placed.resize(sx, sy, true, true, true, true, sx, Transformation.TOPLEFT);

      var rb = placed.geometricBounds;
      var rw = Math.abs(Number(rb[2]) - Number(rb[0]));
      var rh = Math.abs(Number(rb[1]) - Number(rb[3]));
      var posLeft = tLeft + gap + ((innerW - rw) / 2);
      var posTop = tTop - gap - ((innerH - rh) / 2);
      placed.position = [posLeft, posTop];

      // Keep QR close to target in the same parent stack.
      try {placed.move(targetItem, ElementPlacement.PLACEAFTER);} catch(_eMoveTarget) { }
      return true;
    } catch(_ePlaceTargetPng) {
      return false;
    } finally {
      try {if(tmpFile && tmpFile.opened) tmpFile.close();} catch(_eClosePng2) { }
    }
  }
  function _addGreenTickNextToTextFrame(doc, textFrame) {
    try {
      if(!textFrame) return false;

      var b = textFrame.geometricBounds; // [L,T,R,B]
      var left = Number(b[0]);
      var top = Number(b[1]);
      var right = Number(b[2]);
      var bottom = Number(b[3]);
      if(!isFinite(left) || !isFinite(top) || !isFinite(right) || !isFinite(bottom)) return false;

      var gap = _srh_mm2ptDoc(2);
      var tickW = _srh_mm2ptDoc(3.6);
      var tickH = _srh_mm2ptDoc(3.6);

      var x0 = left - gap - tickW;
      var y0 = top - ((top - bottom) / 2); // vertical center

      // FLIPPED vertically (mirror about horizontal axis)
      var p1 = [x0, y0 - (tickH * 0.05)];
      var p2 = [x0 + (tickW * 0.36), y0 - (tickH * 0.35)];
      var p3 = [x0 + tickW, y0 + (tickH * 0.35)];

      var pathA = doc.activeLayer.pathItems.add();
      pathA.setEntirePath([p1, p2]);
      pathA.stroked = true;
      pathA.filled = false;
      pathA.strokeWidth = _srh_pxStrokeDoc(1.6);

      var pathB = doc.activeLayer.pathItems.add();
      pathB.setEntirePath([p2, p3]);
      pathB.stroked = true;
      pathB.filled = false;
      pathB.strokeWidth = _srh_pxStrokeDoc(1.6);

      var green = new RGBColor();
      green.red = 34;
      green.green = 197;
      green.blue = 94;

      pathA.strokeColor = green;
      pathB.strokeColor = green;

      return true;
    } catch(_eTick) {
      return false;
    }
  }
  function _setTextInContainer(item, value) {
    var updated = 0;
    try {
      if(!item) return 0;
      if(item.typename === 'TextFrame') {
        try {item.contents = value; updated++;} catch(_eSetDirect) { }
        return updated;
      }
      try {
        var tfs = item.textFrames;
        if(tfs && typeof tfs.length !== 'undefined') {
          for(var i = 0; i < tfs.length; i++) {
            try {tfs[i].contents = value; updated++;} catch(_eSetTf) { }
          }
        }
      } catch(_eNoTf) { }
    } catch(_eSetContainer) { }
    return updated;
  }
  function _getPageItemTargetsByName(itemMap, targetName) {
    if(!itemMap) return null;
    var direct = itemMap[targetName];
    if(direct && direct.length) return direct;
    var norm = _trim(targetName).toLowerCase();
    if(!norm) return null;
    for(var key in itemMap) {
      if(!itemMap.hasOwnProperty(key)) continue;
      if(_trim(key).toLowerCase() === norm) return itemMap[key];
    }
    return null;
  }

  try {
    var rawPath = _trim(pathText);
    if(!rawPath) return 'Error: No proof path supplied.';
    var fileRef = new File(rawPath);
    if(!fileRef.exists) return 'Error: Proof file not found: ' + rawPath;
    var proofFolderPath = '';
    try {
      var folderRef = fileRef.parent;
      if(folderRef && folderRef.fsName) proofFolderPath = String(folderRef.fsName);
    } catch(_ePf) {proofFolderPath = '';}

    var parsedData = null;
    try {parsedData = JSON.parse(String(dataJson == null ? '' : dataJson));}
    catch(_eJson) {return 'Error: Invalid data JSON.';}
    var row = null;
    if(parsedData && parsedData.constructor === Array) row = parsedData.length ? parsedData[0] : null;
    else if(parsedData && typeof parsedData === 'object') row = parsedData;
    if(!row || typeof row !== 'object') return 'Error: No row data available.';

    var mappings = _parseMappings(mappingText);
    if(!mappings.length) return 'Error: No valid mappings. Use "source -> target".';

    // Always open a fresh document instance for proof generation.
    // If the source proof file is already open, open a temp copy instead of reusing that document.
    var openFile = fileRef;
    try {
      var tmpFolder = Folder.temp;
      if(tmpFolder && tmpFolder.exists) {
        var srcName = String(fileRef.name || 'proof.ai');
        var dotIdx = srcName.lastIndexOf('.');
        var base = dotIdx > 0 ? srcName.substring(0, dotIdx) : srcName;
        var ext = dotIdx > 0 ? srcName.substring(dotIdx) : '.ai';
        var unique = (new Date().getTime()) + '_' + Math.floor(Math.random() * 100000);
        var tmpPath = tmpFolder.fsName + '/srh-proof-' + base + '-' + unique + ext;
        var tmpFile = new File(tmpPath);
        if(fileRef.copy(tmpFile.fsName)) openFile = tmpFile;
      }
    } catch(_eTmpCopy) { }

    var doc = app.open(openFile);
    if(!doc) return 'Error: Failed to open proof file.';


    var frameMap = {};
    var frameMapNormalized = {};
    var itemMap = {};
    var allFrames = doc.textFrames;
    for(var f = 0; f < allFrames.length; f++) {
      var frame = allFrames[f];
      var frameName = '';
      try {frameName = String(frame.name || '');} catch(_eName) {frameName = '';}
      if(!frameName) continue;
      if(!frameMap[frameName]) frameMap[frameName] = [];
      frameMap[frameName].push(frame);
      var normName = _trim(frameName).toLowerCase();
      if(normName) {
        if(!frameMapNormalized[normName]) frameMapNormalized[normName] = [];
        frameMapNormalized[normName].push(frame);
      }
    }
    function _getTextTargetsByName(targetName) {
      var key = String(targetName == null ? '' : targetName);
      if(frameMap[key] && frameMap[key].length) return frameMap[key];
      var norm = _trim(key).toLowerCase();
      if(norm && frameMapNormalized[norm] && frameMapNormalized[norm].length) return frameMapNormalized[norm];
      return null;
    }
    var allItems = doc.pageItems;
    for(var p = 0; p < allItems.length; p++) {
      var item = allItems[p];
      var itemName = '';
      try {itemName = String(item.name || '');} catch(_eItemName) {itemName = '';}
      if(!itemName) continue;
      if(!itemMap[itemName]) itemMap[itemName] = [];
      itemMap[itemName].push(item);
    }

    // Always keep "File Path Text" up to date on every proof generation.
    var filePathTargets = frameMap['File Path Text'];
    if(filePathTargets && filePathTargets.length) {
      for(var fp = 0; fp < filePathTargets.length; fp++) {
        try {filePathTargets[fp].contents = proofFolderPath;} catch(_eFpSet) { }
      }
    }
    var dateTextTargets = frameMap['Date Text'];
    if(dateTextTargets && dateTextTargets.length) {
      var dateValue = _todayDdMmYy();
      for(var dt = 0; dt < dateTextTargets.length; dt++) {
        try {dateTextTargets[dt].contents = dateValue;} catch(_eDtSet) { }
      }
    }

    var applied = 0;
    var missingSource = 0;
    var missingTarget = 0;
    var appliedImages = 0;
    var forcedAddressApplied = 0;
    var forcedQrPlaced = 0;
    var forcedAddressTicks = 0;
    var forcedDescriptionApplied = 0;
    var forcedMediaApplied = 0;
    var forcedLaminateApplied = 0;
    var forcedSubstrateApplied = 0;
    var forcedPartsApplied = 0;
    var forcedNotesApplied = 0;
    var fallbackContainerTextApplied = 0;
    for(var m = 0; m < mappings.length; m++) {
      var mapRow = mappings[m];
      var sourceValue = _readByPath(row, mapRow.sourceKey);
      if(sourceValue === undefined) {
        missingSource++;
        continue;
      }
      var graphicTargets = itemMap[mapRow.targetName];
      var targets = _getTextTargetsByName(mapRow.targetName);
      if((!targets || !targets.length) && (!graphicTargets || !graphicTargets.length)) {
        missingTarget++;
        continue;
      }

      if(targets && targets.length) {
        var textValue = _toTextValue(sourceValue);
        for(var t = 0; t < targets.length; t++) {
          try {
            targets[t].contents = textValue;
            applied++;
          } catch(_eSetTxt) { }
        }
      }
      // If the named target is a group/container (not direct TextFrame naming), set all descendant text frames.
      if(graphicTargets && graphicTargets.length && (!targets || !targets.length)) {
        var containerTextValue = _toTextValue(sourceValue);
        for(var gt = 0; gt < graphicTargets.length; gt++) {
          fallbackContainerTextApplied += _setTextInContainer(graphicTargets[gt], containerTextValue);
        }
      }

      if(graphicTargets && graphicTargets.length && _looksLikeSvgText(sourceValue)) {
        for(var g = 0; g < graphicTargets.length; g++) {
          if(_placeSvgIntoTarget(doc, graphicTargets[g], sourceValue)) appliedImages++;
        }
      }
    }

    // Fallback (debug): always apply address text to all "Address Text" frames.
    var derivedAddress = _readByPath(row, 'Derived.installAddress');
    var addressFrames = _getTextTargetsByName('Address Text');
    if(addressFrames && addressFrames.length && derivedAddress != null) {
      var forcedAddressValue = _toTextValue(derivedAddress);
      for(var af = 0; af < addressFrames.length; af++) {
        try {
          addressFrames[af].contents = forcedAddressValue;
          forcedAddressApplied++;
          if(_addGreenTickNextToTextFrame(doc, addressFrames[af])) forcedAddressTicks++;
        } catch(_eAddrSet) { }
      }
    }
    // Fallback (debug): always apply description text from derived value.
    var derivedDescription = _readByPath(row, 'Derived.lineItemDescription');
    var descriptionFrames = _getTextTargetsByName('Description Text');
    if((!descriptionFrames || !descriptionFrames.length)) descriptionFrames = _getTextTargetsByName('Description');
    var descriptionContainers = _getPageItemTargetsByName(itemMap, 'Description Text');
    if((!descriptionContainers || !descriptionContainers.length)) descriptionContainers = _getPageItemTargetsByName(itemMap, 'Description');
    if(descriptionFrames && descriptionFrames.length && derivedDescription != null) {
      var descValue = _toTextValue(derivedDescription);
      for(var df = 0; df < descriptionFrames.length; df++) {
        try {
          descriptionFrames[df].contents = descValue;
          forcedDescriptionApplied++;
        } catch(_eDescSet) { }
      }
    }
    if((!descriptionFrames || !descriptionFrames.length) && descriptionContainers && descriptionContainers.length && derivedDescription != null) {
      var descValue2 = _toTextValue(derivedDescription);
      for(var dc = 0; dc < descriptionContainers.length; dc++) {
        forcedDescriptionApplied += _setTextInContainer(descriptionContainers[dc], descValue2);
      }
    }
    // Fallback (debug): always apply media/laminate text from derived values.
    var derivedMedia = _readByPath(row, 'Derived.mediaText');
    var mediaFrames = _getTextTargetsByName('Media Text');
    if(mediaFrames && mediaFrames.length && derivedMedia != null) {
      var mediaValue = _toTextValue(derivedMedia);
      for(var mf = 0; mf < mediaFrames.length; mf++) {
        try {
          mediaFrames[mf].contents = mediaValue;
          forcedMediaApplied++;
        } catch(_eMediaSet) { }
      }
    }
    var derivedLaminate = _readByPath(row, 'Derived.laminateText');
    var laminateFrames = _getTextTargetsByName('Laminate Text');
    if(laminateFrames && laminateFrames.length && derivedLaminate != null) {
      var laminateValue = _toTextValue(derivedLaminate);
      for(var lf = 0; lf < laminateFrames.length; lf++) {
        try {
          laminateFrames[lf].contents = laminateValue;
          forcedLaminateApplied++;
        } catch(_eLamSet) { }
      }
    }
    var derivedSubstrate = _readByPath(row, 'Derived.substrateText');
    var substrateFrames = _getTextTargetsByName('Substrate Text');
    if(substrateFrames && substrateFrames.length && derivedSubstrate != null) {
      var substrateValue = _toTextValue(derivedSubstrate);
      for(var sf = 0; sf < substrateFrames.length; sf++) {
        try {
          substrateFrames[sf].contents = substrateValue;
          forcedSubstrateApplied++;
        } catch(_eSubSet) { }
      }
    }
    var derivedParts = _readByPath(row, 'Derived.partsNumbered');
    var partsFrames = _getTextTargetsByName('Parts');
    if((!partsFrames || !partsFrames.length)) partsFrames = _getTextTargetsByName('Parts Text');
    var partsContainers = _getPageItemTargetsByName(itemMap, 'Parts Text');
    if(partsFrames && partsFrames.length && derivedParts != null) {
      var partsValue = _toTextValue(derivedParts);
      for(var pf = 0; pf < partsFrames.length; pf++) {
        try {
          partsFrames[pf].contents = partsValue;
          forcedPartsApplied++;
        } catch(_ePartsSet) { }
      }
    }
    if((!partsFrames || !partsFrames.length) && partsContainers && partsContainers.length && derivedParts != null) {
      var partsValue2 = _toTextValue(derivedParts);
      for(var pc = 0; pc < partsContainers.length; pc++) {
        forcedPartsApplied += _setTextInContainer(partsContainers[pc], partsValue2);
      }
    }
    var derivedNotes = _readByPath(row, 'Derived.notesAll');
    var notesFrames = _getTextTargetsByName('Notes');
    if((!notesFrames || !notesFrames.length)) notesFrames = _getTextTargetsByName('Notes Text');
    var notesContainers = _getPageItemTargetsByName(itemMap, 'Notes Text');
    if(notesFrames && notesFrames.length && derivedNotes != null) {
      var notesValue = _toTextValue(derivedNotes);
      for(var nf = 0; nf < notesFrames.length; nf++) {
        try {
          notesFrames[nf].contents = notesValue;
          forcedNotesApplied++;
        } catch(_eNotesSet) { }
      }
    }
    if((!notesFrames || !notesFrames.length) && notesContainers && notesContainers.length && derivedNotes != null) {
      var notesValue2 = _toTextValue(derivedNotes);
      for(var nc = 0; nc < notesContainers.length; nc++) {
        forcedNotesApplied += _setTextInContainer(notesContainers[nc], notesValue2);
      }
    }

    // Fallback (debug): place QR inside "Address QR" target with 2mm inset.
    var qrPngDataUrl = _readByPath(row, 'DerivedAssets.addressQrPngDataUrl');
    var qrSvg = _readByPath(row, 'DerivedAssets.addressQrSvg');
    var addressQrTargets = itemMap['Address QR'];
    if(addressQrTargets && addressQrTargets.length) {
      for(var aq = 0; aq < addressQrTargets.length; aq++) {
        if(_looksLikePngDataUrl(qrPngDataUrl)) {
          if(_placePngDataUrlIntoTarget(doc, addressQrTargets[aq], qrPngDataUrl)) forcedQrPlaced++;
        } else if(_looksLikeSvgText(qrSvg)) {
          if(_placeSvgIntoTarget(doc, addressQrTargets[aq], qrSvg)) forcedQrPlaced++;
        }
      }
    } else {
      if(_looksLikePngDataUrl(qrPngDataUrl)) {
        if(_placePngDataUrlOnActiveArtboard(doc, qrPngDataUrl)) forcedQrPlaced++;
      } else if(_looksLikeSvgText(qrSvg)) {
        if(_placeSvgOnActiveArtboard(doc, qrSvg)) forcedQrPlaced++;
      }
    }

    var pageNumberRes = signarama_helper_corebridge_updatePageNumbers();
    var flashRes = _srh_corebridge_startFlashing(doc, flashFieldsText);
    var flashMessage = 'Flash monitoring: ' + flashRes.found + ' text frame' + (flashRes.found === 1 ? '' : 's') +
      ' active from ' + flashRes.requested + ' requested name' + (flashRes.requested === 1 ? '' : 's') + '.';
    return 'Proof created. Applied ' + applied + ' text update' + (applied === 1 ? '' : 's') +
      ' and ' + appliedImages + ' image placement' + (appliedImages === 1 ? '' : 's') +
      '. Forced Address Text updates: ' + forcedAddressApplied +
      '. Forced Address Ticks: ' + forcedAddressTicks +
      '. Forced Description Text updates: ' + forcedDescriptionApplied +
      '. Forced Media Text updates: ' + forcedMediaApplied +
      '. Forced Laminate Text updates: ' + forcedLaminateApplied +
      '. Forced Substrate Text updates: ' + forcedSubstrateApplied +
      '. Forced Parts updates: ' + forcedPartsApplied +
      '. Forced Notes updates: ' + forcedNotesApplied +
      '. Fallback container text updates: ' + fallbackContainerTextApplied +
      '. Forced QR placements: ' + forcedQrPlaced +
      '. Missing source: ' + missingSource + '. Missing target: ' + missingTarget + '. ' + pageNumberRes +
      '. ' + flashMessage +
      '. ' + String(linkPlacementRes || '');
  } catch(e) {
    return 'Error: ' + e.message;
  }
}

function signarama_helper_corebridge_createProofForSelected(pathText, dataJson, mappingText, a4OptionsJson, flashFieldsText) {
  function _trim(v) {
    return String(v == null ? '' : v).replace(/^\s+|\s+$/g, '');
  }
  function _normalizeNameForLookup(value) {
    var s = String(value == null ? '' : value);
    s = s.replace(/\u00a0/g, ' ');
    s = _trim(s).toLowerCase();
    s = s.replace(/\s+/g, ' ');
    return s;
  }
  function _dbg(msg) {
    try {$.writeln('[SRH][CorebridgeProofSelected] ' + msg);} catch(_eDbgCp0) { }
  }
  function _rectToText(rect) {
    if(!rect) return 'null';
    var left = Number(rect.left), top = Number(rect.top), right = Number(rect.right), bottom = Number(rect.bottom);
    var w = Math.max(0, right - left);
    var h = Math.max(0, top - bottom);
    var cx = (left + right) / 2;
    var cy = (top + bottom) / 2;
    return 'L=' + left.toFixed(2) + ', T=' + top.toFixed(2) + ', R=' + right.toFixed(2) + ', B=' + bottom.toFixed(2) +
      ', W=' + w.toFixed(2) + ', H=' + h.toFixed(2) + ', C=(' + cx.toFixed(2) + ', ' + cy.toFixed(2) + ')';
  }
  function _normalizeRect(b) {
    if(!b || b.length !== 4) return null;
    var l = Number(b[0]), t = Number(b[1]), r = Number(b[2]), bt = Number(b[3]);
    if(!isFinite(l) || !isFinite(t) || !isFinite(r) || !isFinite(bt)) return null;
    return {
      left: Math.min(l, r),
      right: Math.max(l, r),
      top: Math.max(t, bt),
      bottom: Math.min(t, bt)
    };
  }
  function _getPlacedItemBounds(item) {
    if(!item) return null;
    var b = null;
    try {b = item.visibleBounds;} catch(_ePb0) {b = null;}
    if(!b || b.length !== 4) {
      try {b = item.geometricBounds;} catch(_ePb1) {b = null;}
    }
    return (b && b.length === 4) ? b : null;
  }
  function _collectNamedItems(doc, nameText) {
    var out = [];
    if(!doc || !nameText) return out;
    var items = null;
    try {items = doc.pageItems;} catch(_eCol0) {items = null;}
    if(!items) return out;
    for(var i = 0; i < items.length; i++) {
      var it = items[i];
      var nm = '';
      try {nm = String(it.name || '');} catch(_eCol1) {nm = '';}
      if(nm === nameText) out.push(it);
    }
    return out;
  }
  function _containsRef(list, item) {
    if(!list || !list.length || !item) return false;
    for(var i = 0; i < list.length; i++) {
      if(list[i] === item) return true;
    }
    return false;
  }
  function _getTargetBounds(item) {
    if(!item) return null;
    if(item.typename === 'PathItem' || item.typename === 'CompoundPathItem') {
      var pb = null;
      try {pb = item.geometricBounds;} catch(_eTbP0) {pb = null;}
      if(!pb || pb.length !== 4) {
        try {pb = item.visibleBounds;} catch(_eTbP1) {pb = null;}
      }
      if(pb && pb.length === 4) return pb;
    }
    var b = null;
    try {b = _srh_getClippingPathBounds(item);} catch(_eTb0) {b = null;}
    if(!b || b.length !== 4) {
      try {b = item.visibleBounds;} catch(_eTb1) {b = null;}
    }
    if(!b || b.length !== 4) {
      try {b = item.geometricBounds;} catch(_eTb2) {b = null;}
    }
    return (b && b.length === 4) ? b : null;
  }
  function _fitItemIntoTarget(item, target) {
    if(!item || !target) return false;
    var tb = _getTargetBounds(target);
    if(!tb || tb.length !== 4) return false;
    function _normalizeRect(b) {
      if(!b || b.length !== 4) return null;
      var l = Number(b[0]), t = Number(b[1]), r = Number(b[2]), bt = Number(b[3]);
      if(!isFinite(l) || !isFinite(t) || !isFinite(r) || !isFinite(bt)) return null;
      return {
        left: Math.min(l, r),
        right: Math.max(l, r),
        top: Math.max(t, bt),
        bottom: Math.min(t, bt)
      };
    }
    function _getItemBounds() {
      var b = null;
      try {b = item.visibleBounds;} catch(_eFb0) {b = null;}
      if(!b || b.length !== 4) {
        try {b = item.geometricBounds;} catch(_eFb1) {b = null;}
      }
      return (b && b.length === 4) ? b : null;
    }
    var tRect = _normalizeRect(tb);
    var iRect = _normalizeRect(_getItemBounds());
    if(!tRect || !iRect) return false;

    var inset = _srh_mm2ptDoc(2);
    var innerLeft = tRect.left + inset;
    var innerTop = tRect.top - inset;
    var innerRight = tRect.right - inset;
    var innerBottom = tRect.bottom + inset;
    var tw = Math.max(0, innerRight - innerLeft);
    var th = Math.max(0, innerTop - innerBottom);
    var iw = Math.max(0, iRect.right - iRect.left);
    var ih = Math.max(0, iRect.top - iRect.bottom);
    if(!(tw > 0 && th > 0 && iw > 0 && ih > 0)) return false;

    var fitScale = Math.min(tw / iw, th / ih);
    var pct = fitScale * 100;
    try {item.resize(pct, pct, true, true, true, true, 100, Transformation.CENTER);} catch(_eFb2) { }

    iRect = _normalizeRect(_getItemBounds());
    if(!iRect) return false;
    iw = Math.max(0, iRect.right - iRect.left);
    ih = Math.max(0, iRect.top - iRect.bottom);

    // Safety clamp: enforce final width/height not larger than target inner bounds.
    for(var clampTry = 0; clampTry < 3; clampTry++) {
      if(iw <= tw && ih <= th) break;
      var down = Math.min(tw / iw, th / ih);
      if(!(down > 0 && down < 1)) break;
      try {item.resize(down * 100, down * 100, true, true, true, true, 100, Transformation.CENTER);} catch(_eFbClamp) {break;}
      iRect = _normalizeRect(_getItemBounds());
      if(!iRect) break;
      iw = Math.max(0, iRect.right - iRect.left);
      ih = Math.max(0, iRect.top - iRect.bottom);
    }

    var targetCx = (innerLeft + innerRight) / 2;
    var targetCy = (innerTop + innerBottom) / 2;
    var itemCx = (iRect.left + iRect.right) / 2;
    var itemCy = (iRect.top + iRect.bottom) / 2;
    var dx = targetCx - itemCx;
    var dy = targetCy - itemCy;
    try {item.translate(dx, dy);} catch(_eFb5) {return false;}

    // Final hard cap using post-translate bounds.
    iRect = _normalizeRect(_getItemBounds());
    if(iRect) {
      iw = Math.max(0, iRect.right - iRect.left);
      ih = Math.max(0, iRect.top - iRect.bottom);
      if(iw > tw || ih > th) {
        var finalDown = Math.min(tw / Math.max(iw, 1e-6), th / Math.max(ih, 1e-6));
        if(finalDown > 0 && finalDown < 1) {
          try {item.resize(finalDown * 100, finalDown * 100, true, true, true, true, 100, Transformation.CENTER);} catch(_eFb6) { }
          iRect = _normalizeRect(_getItemBounds());
          if(iRect) {
            itemCx = (iRect.left + iRect.right) / 2;
            itemCy = (iRect.top + iRect.bottom) / 2;
            try {item.translate(targetCx - itemCx, targetCy - itemCy);} catch(_eFb7) { }
          }
        }
      }
    }
    try {item.move(target, ElementPlacement.PLACEAFTER);} catch(_eFbMove) { }
    return true;
  }

  try {
    if(!app.documents.length) return 'Error: No open source document.';
    var sourceDoc = app.activeDocument;
    if(!sourceDoc.selection || sourceDoc.selection.length === 0) return 'Error: No selection. Select artwork first.';

    var a4Options = {rasterize: false, rasterizeQuality: 'high'};
    try {
      var parsedA4 = a4OptionsJson ? JSON.parse(String(a4OptionsJson)) : null;
      if(parsedA4 && typeof parsedA4 === 'object') {
        a4Options.rasterize = !!parsedA4.rasterize;
        a4Options.rasterizeQuality = _trim(String(parsedA4.rasterizeQuality || 'high')).toLowerCase() || 'high';
      }
    } catch(_eA4Opts) { }
    var a4OptionsSafe = JSON.stringify(a4Options);

    var beforeCopies = _collectNamedItems(sourceDoc, 'SRH_A4_Fit_Copy');
    var a4Res = signarama_helper_duplicateOutlineScaleA4(a4OptionsSafe);
    if(/^Error:/i.test(String(a4Res || '')) || /^No\b/i.test(String(a4Res || ''))) return String(a4Res || 'Error: Failed to scale selected artwork.');

    var afterCopies = _collectNamedItems(sourceDoc, 'SRH_A4_Fit_Copy');
    var scaledCopy = null;
    for(var c = 0; c < afterCopies.length; c++) {
      if(!_containsRef(beforeCopies, afterCopies[c])) {scaledCopy = afterCopies[c]; break;}
    }
    if(!scaledCopy && sourceDoc.selection && sourceDoc.selection.length) {
      try {scaledCopy = sourceDoc.selection[0];} catch(_eSelCopy) {scaledCopy = null;}
    }
    if(!scaledCopy) return 'Error: Could not find scaled proof copy to place.';

    try {sourceDoc.selection = null;} catch(_eSelClr0) { }
    try {scaledCopy.selected = true;} catch(_eSelSet0) { }
    try {app.copy();} catch(_eCopy0) {
      try {app.executeMenuCommand('copy');} catch(_eCopy1) {return 'Error: Failed to copy scaled artwork.';}
    }

    var proofRes = signarama_helper_corebridge_createProofFromData(pathText, dataJson, mappingText, flashFieldsText);
    if(/^Error:/i.test(String(proofRes || ''))) {
      try {scaledCopy.remove();} catch(_eRmCopy0) { }
      return proofRes;
    }
    if(!app.documents.length) {
      try {scaledCopy.remove();} catch(_eRmCopy1) { }
      return 'Error: Proof document was not opened.';
    }
    var proofDoc = app.activeDocument;
    try {proofDoc.selection = null;} catch(_eSelClr1) { }

    var pastedGroup = null;
    var placementMethod = '';
    try {
      app.paste();
      placementMethod = 'paste';
    } catch(_ePaste0) {
      try {
        app.executeMenuCommand('pasteInPlace');
        placementMethod = 'pasteInPlace';
      } catch(_ePaste1) {
        try {
          var duplicateTarget = null;
          try {duplicateTarget = proofDoc.activeLayer;} catch(_eDupLayer0) {duplicateTarget = null;}
          if(!duplicateTarget) {
            try {duplicateTarget = proofDoc.layers[0];} catch(_eDupLayer1) {duplicateTarget = null;}
          }
          if(!duplicateTarget) {
            try {duplicateTarget = proofDoc;} catch(_eDupLayer2) {duplicateTarget = null;}
          }
          if(duplicateTarget) {
            pastedGroup = scaledCopy.duplicate(duplicateTarget, ElementPlacement.PLACEATEND);
            placementMethod = 'duplicate';
          }
        } catch(_eDup0) {pastedGroup = null;}
      }
    }

    if(!pastedGroup) {
      var pastedSel = null;
      try {pastedSel = proofDoc.selection;} catch(_ePasteSel0) {pastedSel = null;}
      if(pastedSel && pastedSel.length) {
        if(pastedSel.length === 1) pastedGroup = pastedSel[0];
        else {
          try {
            pastedGroup = proofDoc.activeLayer.groupItems.add();
            for(var ps = 0; ps < pastedSel.length; ps++) {
              try {pastedSel[ps].move(pastedGroup, ElementPlacement.PLACEATEND);} catch(_eMvPg0) { }
            }
          } catch(_eGroup) {pastedGroup = pastedSel[0];}
        }
      }
    }

    if(!pastedGroup) {
      try {scaledCopy.remove();} catch(_eRmCopy2) { }
      return proofRes + ' Artwork placement failed: unable to transfer artwork to proof document.';
    }

    var target = null;
    var targetPreferred = null;
    var targetFallback = null;
    var nearNames = [];
    var pageItems = null;
    try {pageItems = proofDoc.pageItems;} catch(_ePi0) {pageItems = null;}
    if(pageItems) {
      var targetNeedle = _normalizeNameForLookup('Artwork Placement Area');
      for(var pi = 0; pi < pageItems.length; pi++) {
        var nm = '';
        var itemPi = pageItems[pi];
        try {nm = _trim(String(itemPi.name || ''));} catch(_ePi1) {nm = '';}
        if(!nm) continue;
        var nmNorm = _normalizeNameForLookup(nm);
        var matchesExact = (nmNorm === targetNeedle);
        var matchesLoose = (!matchesExact && (nmNorm.indexOf('artwork placement area') >= 0 || nmNorm.indexOf('artwork placement') >= 0));
        if(!matchesExact && !matchesLoose) {
          if(nmNorm.indexOf('artwork') >= 0 || nmNorm.indexOf('placement') >= 0) {
            if(nearNames.length < 10) nearNames.push(nm);
          }
          continue;
        }
        if(!targetFallback) targetFallback = itemPi;
        var typeName = '';
        try {typeName = String(itemPi.typename || '');} catch(_ePi2) {typeName = '';}
        if(typeName === 'PathItem' || typeName === 'CompoundPathItem') {
          targetPreferred = itemPi;
          break;
        }
      }
    }
    target = targetPreferred || targetFallback;
    var targetBoundsForLog = _normalizeRect(_getTargetBounds(target));
    _dbg('Target "Artwork Placement Area": ' + _rectToText(targetBoundsForLog));
    var placedBoundsBefore = _normalizeRect(_getPlacedItemBounds(pastedGroup));
    _dbg('Placed artwork before fit: ' + _rectToText(placedBoundsBefore));
    if(!target) {
      try {scaledCopy.remove();} catch(_eRmCopy4) { }
      var nearText = nearNames.length ? (' Near matches: ' + nearNames.join(' | ')) : '';
      return proofRes + ' Artwork placement skipped: target "Artwork Placement Area" not found.' + nearText;
    }

    var fitOk = _fitItemIntoTarget(pastedGroup, target);
    var placedBoundsAfter = _normalizeRect(_getPlacedItemBounds(pastedGroup));
    _dbg('Placed artwork after fit: ' + _rectToText(placedBoundsAfter));
    try {scaledCopy.remove();} catch(_eRmCopy5) { }
    if(!fitOk) return proofRes + ' Artwork placement failed: could not fit into target.';
    var methodText = placementMethod ? (' via ' + placementMethod) : '';
    return proofRes + ' Artwork placed into "Artwork Placement Area"' + methodText + '.';
  } catch(e) {
    return 'Error: ' + e.message;
  }
}

this.signarama_helper_transform_makeSize = function(json) {
  return _srh_transform_makeSize_impl(json);
};
this.atlas_transform_makeSize = function(json) {
  return _srh_transform_makeSize_impl(json);
};
try {if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_transform_makeSize = this.signarama_helper_transform_makeSize;} catch(_eTg0) { }
try {if(typeof $ !== 'undefined' && $.global) $.global.atlas_transform_makeSize = this.atlas_transform_makeSize;} catch(_eTg2) { }
try {signarama_helper_transform_makeSize = this.signarama_helper_transform_makeSize;} catch(_eTg1) { }
try {if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_transform_listArtboards = this.signarama_helper_transform_listArtboards;} catch(_eTg3) { }
try {signarama_helper_transform_listArtboards = this.signarama_helper_transform_listArtboards;} catch(_eTg4) { }
try {if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_transform_debugArtboards = this.signarama_helper_transform_debugArtboards;} catch(_eTg5) { }
try {signarama_helper_transform_debugArtboards = this.signarama_helper_transform_debugArtboards;} catch(_eTg6) { }
try {if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_corebridge_openProofPath = this.signarama_helper_corebridge_openProofPath;} catch(_eTg13) { }
try {signarama_helper_corebridge_openProofPath = this.signarama_helper_corebridge_openProofPath;} catch(_eTg14) { }
try {if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_corebridge_createProofFromData = this.signarama_helper_corebridge_createProofFromData;} catch(_eTg15) { }
try {signarama_helper_corebridge_createProofFromData = this.signarama_helper_corebridge_createProofFromData;} catch(_eTg16) { }
try {if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_corebridge_createProofForSelected = this.signarama_helper_corebridge_createProofForSelected;} catch(_eTg16a) { }
try {signarama_helper_corebridge_createProofForSelected = this.signarama_helper_corebridge_createProofForSelected;} catch(_eTg16b) { }
try {if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_corebridge_updatePageNumbers = this.signarama_helper_corebridge_updatePageNumbers;} catch(_eTg17) { }
try {signarama_helper_corebridge_updatePageNumbers = this.signarama_helper_corebridge_updatePageNumbers;} catch(_eTg18) { }
try {if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_corebridge_flashTickTask = this.signarama_helper_corebridge_flashTickTask;} catch(_eTg19) { }
try {signarama_helper_corebridge_flashTickTask = this.signarama_helper_corebridge_flashTickTask;} catch(_eTg20) { }
try {if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_corebridge_flashGetState = this.signarama_helper_corebridge_flashGetState;} catch(_eTg20a) { }
try {signarama_helper_corebridge_flashGetState = this.signarama_helper_corebridge_flashGetState;} catch(_eTg20b) { }
try {if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_corebridge_flashTickTaskRunner = this.signarama_helper_corebridge_flashTickTaskRunner;} catch(_eTg21) { }
try {signarama_helper_corebridge_flashTickTaskRunner = this.signarama_helper_corebridge_flashTickTaskRunner;} catch(_eTg22) { }



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

function _dim_normalizeAngleRad(angleRad) {
  var a = angleRad;
  var twoPi = Math.PI * 2;
  while(a <= -Math.PI) a += twoPi;
  while(a > Math.PI) a -= twoPi;
  return a;
}

function _dim_makeArc(group, cx, cy, radius, startRad, deltaRad, strokePt, lineColor) {
  if(!(radius > 0)) return null;
  var segs = Math.max(8, Math.ceil(Math.abs(deltaRad) / (Math.PI / 18)));
  var pts = [];
  for(var i = 0; i <= segs; i++) {
    var t = i / segs;
    var a = startRad + (deltaRad * t);
    pts.push([cx + Math.cos(a) * radius, cy + Math.sin(a) * radius]);
  }
  var p = group.pathItems.add();
  p.setEntirePath(pts);
  p.closed = false;
  p.stroked = true;
  var sw = (typeof strokePt === 'number' && isFinite(strokePt)) ? strokePt : 1;
  p.strokeWidth = Math.max(sw, 0.1);
  try {if(lineColor) p.strokeColor = lineColor;} catch(_eArcC) { }
  p.filled = false;
  p.opacity = 100;
  p.blendingMode = BlendModes.NORMAL;
  return p;
}

function _dim_collectPathsFromItem(item, out) {
  if(!item || !out) return;
  try {
    if(item.locked || item.hidden) return;
  } catch(_eLk) { }
  try {
    if(item.typename === "PathItem") {
      try {if(item.guides) return;} catch(_eGuides) { }
      try {if(item.clipping) return;} catch(_eClipP) { }
      out.push(item);
      return;
    }
    if(item.typename === "CompoundPathItem" && item.pathItems && item.pathItems.length) {
      for(var cp = 0; cp < item.pathItems.length; cp++) {
        var cpi = item.pathItems[cp];
        if(!cpi) continue;
        try {if(cpi.clipping) continue;} catch(_eCpClip) { }
        out.push(cpi);
      }
      return;
    }
    if(item.typename === "GroupItem" && item.pageItems && item.pageItems.length) {
      for(var gi = 0; gi < item.pageItems.length; gi++) {
        _dim_collectPathsFromItem(item.pageItems[gi], out);
      }
    }
  } catch(_eCollect) { }
}

function _dim_runAngles(opts, mode) {
  opts = opts || {};
  var doc = app.activeDocument;
  if(!doc) return "No document open.";

  var originalSel = _dim_captureSelection();
  if(!originalSel || !originalSel.length) return "Nothing selected.";

  var scaleFactor = _srh_getScaleFactor();
  var textPt = _dim_ptDoc(opts.textPt || 10, scaleFactor);
  var strokePt = _dim_ptDoc(opts.strokePt || 1, scaleFactor);
  var baseRadiusPt = _dim_mm2ptDoc(opts.offsetMm || 10, scaleFactor);
  var textOffsetPt = _dim_mm2ptDoc(opts.labelGapMm || 0, scaleFactor);
  var decimals = opts.decimals | 0;
  var scaleAppearance = opts.scaleAppearance || 1;

  var lineColor = _dim_hexToRGB(opts.lineColor) || _dim_parseHexColorToRGBColor(opts.lineColor) || _dim_hexToRGB('#000000');
  var textColor = opts.textColor;
  var lyr = _dim_ensureLayer('Dimensions');

  var paths = [];
  for(var s = 0; s < originalSel.length; s++) {
    _dim_collectPathsFromItem(originalSel[s], paths);
  }
  if(!paths.length) return 'No measurable paths in selection.';

  var added = 0;
  var pathCount = 0;
  var twoPi = Math.PI * 2;
  var isOuter = String(mode || '').toUpperCase() === 'OUTER';

  function _addAngleAtVertex(prev, cur, next) {
    var v1x = prev[0] - cur[0];
    var v1y = prev[1] - cur[1];
    var v2x = next[0] - cur[0];
    var v2y = next[1] - cur[1];
    var l1 = Math.sqrt(v1x * v1x + v1y * v1y);
    var l2 = Math.sqrt(v2x * v2x + v2y * v2y);
    if(!(l1 > 0) || !(l2 > 0)) return 0;

    var a1 = Math.atan2(v1y, v1x);
    var a2 = Math.atan2(v2y, v2x);
    var innerDelta = _dim_normalizeAngleRad(a2 - a1);
    if(Math.abs(innerDelta) < 0.000001) return 0;
    var outerDelta = (innerDelta > 0) ? (innerDelta - twoPi) : (innerDelta + twoPi);
    var sweep = isOuter ? outerDelta : innerDelta;
    var deg = Math.abs(sweep) * 180 / Math.PI;
    if(!(deg > 0.01)) return 0;

    var radius = Math.max(baseRadiusPt * (isOuter ? 1.35 : 1.0) * scaleAppearance, strokePt * scaleAppearance * 2);
    var g = lyr.groupItems.add();
    _dim_makeArc(g, cur[0], cur[1], radius, a1, sweep, strokePt * scaleAppearance, lineColor);

    var mid = a1 + (sweep * 0.5);
    var labelRadius = radius + (textOffsetPt * scaleAppearance) + (textPt * scaleAppearance * 0.5);
    var label = deg.toFixed(decimals) + String.fromCharCode(176);
    var tx = cur[0] + Math.cos(mid) * labelRadius;
    var ty = cur[1] + Math.sin(mid) * labelRadius;
    var txt = _dim_createAnchoredText(label, "C", tx, ty, {size: textPt * scaleAppearance, textColor: textColor}, lyr);
    if(txt) {
      try {txt.move(g, ElementPlacement.PLACEATEND);} catch(_eMvAng) { }
    }
    return 1;
  }

  for(var p = 0; p < paths.length; p++) {
    var path = paths[p];
    if(!path) continue;
    var pts = null;
    try {pts = path.pathPoints;} catch(_ePts) {pts = null;}
    if(!pts || pts.length < 3) continue;

    pathCount++;
    var n = pts.length;
    var closed = false;
    try {closed = !!path.closed;} catch(_eCl) {closed = false;}

    if(closed) {
      for(var i = 0; i < n; i++) {
        try {
          var prev = pts[(i - 1 + n) % n].anchor;
          var cur = pts[i].anchor;
          var next = pts[(i + 1) % n].anchor;
          added += _addAngleAtVertex(prev, cur, next);
        } catch(_eVA0) { }
      }
    } else {
      for(var j = 1; j < n - 1; j++) {
        try {
          var prev2 = pts[j - 1].anchor;
          var cur2 = pts[j].anchor;
          var next2 = pts[j + 1].anchor;
          added += _addAngleAtVertex(prev2, cur2, next2);
        } catch(_eVA1) { }
      }
    }
  }

  if(!added) return 'No measurable angles in selection.';
  var modeLabel = isOuter ? 'outer' : 'inner';
  return 'Added ' + added + ' ' + modeLabel + ' angle measure' + (added === 1 ? '' : 's') +
    ' on ' + pathCount + ' path' + (pathCount === 1 ? '' : 's') + '.';
}

function _dim_pt2mm2(pt2) {
  var k = 25.4 / 72.0;
  return (pt2 || 0) * k * k;
}

function _dim_dist2(a, b) {
  if(!a || !b) return 0;
  var dx = (a[0] || 0) - (b[0] || 0);
  var dy = (a[1] || 0) - (b[1] || 0);
  return dx * dx + dy * dy;
}

function _dim_cubicPoint(a0, r0, l1, a1, t) {
  var u = 1 - t;
  var uu = u * u;
  var tt = t * t;
  var uuu = uu * u;
  var ttt = tt * t;
  return {
    x: uuu * a0[0] + 3 * uu * t * r0[0] + 3 * u * tt * l1[0] + ttt * a1[0],
    y: uuu * a0[1] + 3 * uu * t * r0[1] + 3 * u * tt * l1[1] + ttt * a1[1]
  };
}

function _dim_estimateCubicLength(a0, r0, l1, a1) {
  var prev = {x: a0[0], y: a0[1]};
  var total = 0;
  var segs = 12;
  for(var i = 1; i <= segs; i++) {
    var t = i / segs;
    var p = _dim_cubicPoint(a0, r0, l1, a1, t);
    var dx = p.x - prev.x;
    var dy = p.y - prev.y;
    total += Math.sqrt(dx * dx + dy * dy);
    prev = p;
  }
  return total;
}

function _dim_pathToPolygonPoints(path, options) {
  if(!path) return [];
  var opts = options || {};
  var stepPt = Number(opts.stepPt);
  if(!(stepPt > 0)) stepPt = _dim_mm2pt(10);

  var pts = null;
  try {pts = path.pathPoints;} catch(_ePts) {pts = null;}
  if(!pts || pts.length < 2) return [];

  var isClosed = false;
  try {isClosed = !!path.closed;} catch(_eCl) {isClosed = false;}
  var segCount = isClosed ? pts.length : (pts.length - 1);
  if(segCount < 1) return [];

  var out = [];
  for(var i = 0; i < segCount; i++) {
    var p0 = pts[i];
    var p1 = pts[(i + 1) % pts.length];
    if(!p0 || !p1) continue;
    var a0 = p0.anchor;
    var r0 = p0.rightDirection;
    var l1 = p1.leftDirection;
    var a1 = p1.anchor;

    var isCurve = (_dim_dist2(a0, r0) > 0.0001) || (_dim_dist2(a1, l1) > 0.0001);
    var segLen = 0;
    if(isCurve) {
      segLen = _dim_estimateCubicLength(a0, r0, l1, a1);
    } else {
      var dx = a1[0] - a0[0];
      var dy = a1[1] - a0[1];
      segLen = Math.sqrt(dx * dx + dy * dy);
    }
    var steps = Math.max(1, Math.ceil(segLen / stepPt));

    for(var s = 0; s < steps; s++) {
      var t = s / steps;
      out.push(_dim_cubicPoint(a0, r0, l1, a1, t));
    }
  }

  if(isClosed && pts.length) {
    try {
      var first = pts[0].anchor;
      out.push({x: first[0], y: first[1]});
    } catch(_eFirst) { }
  } else if(!isClosed && pts.length) {
    try {
      var last = pts[pts.length - 1].anchor;
      out.push({x: last[0], y: last[1]});
    } catch(_eLast) { }
  }
  return out;
}

function _dim_polygonAreaSigned(points) {
  if(!points || points.length < 3) return 0;
  var total = 0;
  for(var i = 0, l = points.length; i < l; i++) {
    var p0 = points[i];
    var p1 = points[(i + 1) % l];
    total += (p0.x * p1.y) - (p1.x * p0.y);
  }
  return total * 0.5;
}

function _dim_drawAreaApproxPolygon(layer, points, strokePt) {
  if(!layer || !points || points.length < 3) return null;
  var path = null;
  try {
    path = layer.pathItems.add();
    var arr = [];
    for(var i = 0; i < points.length; i++) {
      arr.push([points[i].x, points[i].y]);
    }
    path.setEntirePath(arr);
    path.closed = true;
    path.filled = false;
    path.stroked = true;
    path.strokeWidth = Math.max((strokePt > 0 ? strokePt : 1), 0.1);
    try {
      var c = new RGBColor();
      c.red = 255; c.green = 0; c.blue = 0;
      path.strokeColor = c;
    } catch(_eRed) { }
    return path;
  } catch(_eApprox) {
    try {if(path) path.remove();} catch(_eRmApprox) { }
    return null;
  }
}

var _dim_areaDebugLines = null;
function _dim_areaDbg(msg) {
  var line = '[SRH][area] ' + String(msg || '');
  try {$.writeln(line);} catch(_eAreaDbg0) { }
  try {
    if(_dim_areaDebugLines) _dim_areaDebugLines.push(line);
  } catch(_eAreaDbg1) { }
}

function _dim_pointInPolygon(point, polygon) {
  if(!point || !polygon || polygon.length < 3) return false;
  var inside = false;
  for(var i = 0, j = polygon.length - 1; i < polygon.length; j = i++) {
    var xi = polygon[i].x, yi = polygon[i].y;
    var xj = polygon[j].x, yj = polygon[j].y;
    var intersects = ((yi > point.y) !== (yj > point.y)) &&
      (point.x < ((xj - xi) * (point.y - yi) / ((yj - yi) || 0.0000001)) + xi);
    if(intersects) inside = !inside;
  }
  return inside;
}

function _dim_getAreaSamplePoint(path, poly) {
  var b = null;
  try {b = path.geometricBounds;} catch(_eGb) {b = null;}
  if(b && b.length === 4) {
    return {
      x: (Number(b[0]) + Number(b[2])) * 0.5,
      y: (Number(b[1]) + Number(b[3])) * 0.5
    };
  }
  if(poly && poly.length) {
    var mid = poly[Math.floor(poly.length / 2)];
    if(mid) return {x: mid.x, y: mid.y};
  }
  return null;
}

function _dim_collectCompoundAreaPaths(compoundItem, out, stepPt) {
  if(!compoundItem || !compoundItem.pathItems || !compoundItem.pathItems.length || !out) return;
  _dim_areaDbg('collect compound | name=' + String(compoundItem.name || '') + ' | children=' + compoundItem.pathItems.length);
  var entries = [];
  for(var c = 0; c < compoundItem.pathItems.length; c++) {
    var cp = compoundItem.pathItems[c];
    if(!cp) continue;
    try {if(cp.guides || cp.clipping || !cp.closed) continue;} catch(_eCpSkip) { }
    var poly = _dim_pathToPolygonPoints(cp, {stepPt: stepPt});
    if(!poly || poly.length < 3) continue;
    var absArea = Math.abs(_dim_polygonAreaSigned(poly));
    if(!(absArea > 0.000001)) continue;
    entries.push({
      path: cp,
      poly: poly,
      absArea: absArea,
      sample: _dim_getAreaSamplePoint(cp, poly),
      sign: 1,
      depth: 0
    });
  }
  for(var i = 0; i < entries.length; i++) {
    var sample = entries[i].sample;
    if(!sample) continue;
    var depth = 0;
    for(var j = 0; j < entries.length; j++) {
      if(i === j) continue;
      if(entries[j].absArea <= entries[i].absArea) continue;
      if(_dim_pointInPolygon(sample, entries[j].poly)) depth++;
    }
    entries[i].depth = depth;
    entries[i].sign = (depth % 2 === 0) ? 1 : -1;
    _dim_areaDbg(
      'compound child | idx=' + i +
      ' | areaPt2=' + entries[i].absArea +
      ' | depth=' + depth +
      ' | sign=' + entries[i].sign +
      ' | sample=' + (sample ? (sample.x + ',' + sample.y) : 'n/a')
    );
    out.push(entries[i]);
  }
}

function _dim_collectAreaPaths(item, out) {
  if(!item || !out) return;
  var stepPt = arguments.length > 2 ? arguments[2] : null;
  try {
    if(item.locked || item.hidden) return;
  } catch(_eLocked) { }

  try {
    _dim_areaDbg('collect item | type=' + String(item.typename || '') + ' | name=' + String(item.name || ''));
    if(item.typename === "PathItem") {
      var parent = null;
      try {parent = item.parent;} catch(_eParent) {parent = null;}
      try {
        if(parent && parent.typename === "CompoundPathItem") {
          _dim_areaDbg('promote path to compound parent | parentName=' + String(parent.name || ''));
          _dim_collectCompoundAreaPaths(parent, out, stepPt);
          return;
        }
      } catch(_eParentType) { }
    }
    if(item.typename === "PathItem") {
      try {if(item.guides) return;} catch(_eGuides) { }
      try {if(item.clipping) return;} catch(_eClip) { }
      _dim_areaDbg('push simple path');
      out.push({path: item, sign: 1, poly: null});
      return;
    }
    if(item.typename === "CompoundPathItem" && item.pathItems && item.pathItems.length) {
      _dim_collectCompoundAreaPaths(item, out, stepPt);
      return;
    }
    if(item.typename === "GroupItem" && item.pageItems && item.pageItems.length) {
      for(var g = 0; g < item.pageItems.length; g++) {
        _dim_collectAreaPaths(item.pageItems[g], out, stepPt);
      }
    }
  } catch(_eCollect) { }
}

function _dim_runArea(opts) {
  opts = opts || {};
  _dim_areaDebugLines = [];
  var doc = app.activeDocument;
  if(!doc) return "No document open.";

  var originalSel = _dim_captureSelection();
  if(!originalSel || !originalSel.length) return "Nothing selected.";
  _dim_areaDbg('run start | selection=' + originalSel.length);

  var scaleFactor = _srh_getScaleFactor();
  var textPt = _dim_ptDoc(opts.textPt || 10, scaleFactor);
  var decimals = Math.max(3, opts.decimals | 0);
  var textColor = opts.textColor;
  var scaleAppearance = opts.scaleAppearance || 1;
  var areaApproximationStepMm = Number(opts.areaApproximationStep);
  if(!(areaApproximationStepMm > 0)) areaApproximationStepMm = 10;
  var stepPt = _dim_mm2ptDoc(areaApproximationStepMm, scaleFactor);
  var includeStroke = !!opts.measureIncludeStroke;
  var showAreaApproximation = !!opts.showAreaApproximation;
  var approxStrokePt = _dim_ptDoc(opts.strokePt || 1, scaleFactor) * scaleAppearance;

  var lyr = _dim_ensureLayer('Dimensions');
  var objectsProcessed = 0;
  var labelsAdded = 0;

  for(var i = 0; i < originalSel.length; i++) {
    var item = originalSel[i];
    if(!item) continue;
    _dim_areaDbg('selection item | idx=' + i + ' | type=' + String(item.typename || '') + ' | name=' + String(item.name || ''));

    var paths = [];
    _dim_collectAreaPaths(item, paths, stepPt);
    _dim_areaDbg('selection item collected paths=' + paths.length);
    if(!paths.length) continue;

    var totalSignedPt2 = 0;
    for(var p = 0; p < paths.length; p++) {
      var pathEntry = paths[p];
      var path = pathEntry && pathEntry.path ? pathEntry.path : pathEntry;
      if(!path) continue;
      var isClosed = false;
      try {isClosed = !!path.closed;} catch(_eClosedPath) {isClosed = false;}
      if(!isClosed) continue;

      var poly = (pathEntry && pathEntry.poly) ? pathEntry.poly : null;
      var aSigned = 0;
      if(!poly) poly = _dim_pathToPolygonPoints(path, {stepPt: stepPt});
      if(!poly || poly.length < 3) continue;
      aSigned = Math.abs(_dim_polygonAreaSigned(poly));
      if(pathEntry && pathEntry.sign === -1) aSigned = -aSigned;
      var polarityText = 'n/a';
      try {
        polarityText = String(path.polarity);
      } catch(_ePol) { }
      _dim_areaDbg(
        'path sum | idx=' + p +
        ' | type=' + String(path.typename || '') +
        ' | sign=' + (pathEntry && typeof pathEntry.sign !== 'undefined' ? pathEntry.sign : 1) +
        ' | depth=' + (pathEntry && typeof pathEntry.depth !== 'undefined' ? pathEntry.depth : 'n/a') +
        ' | polarity=' + polarityText +
        ' | areaSignedPt2=' + aSigned
      );
      if(showAreaApproximation && poly && poly.length >= 3) {
        _dim_drawAreaApproxPolygon(lyr, poly, approxStrokePt);
      }
      totalSignedPt2 += aSigned;
    }
    _dim_areaDbg('selection item totalSignedPt2=' + totalSignedPt2);

    var areaPt2 = Math.abs(totalSignedPt2);
    if(!(areaPt2 > 0.000001)) continue;

    var areaMm2 = _dim_pt2mm2(areaPt2 * scaleFactor * scaleFactor);
    var areaM2 = areaMm2 / 1000000.0;
    var label = areaM2.toFixed(decimals) + ' m' + String.fromCharCode(178);

    var b = null;
    try {b = _dim_getMetricsFor(item, false, includeStroke);} catch(_eBounds) {b = null;}
    if(!b) continue;
    var cx = b.left + ((b.right - b.left) * 0.5);
    var cy = b.bottom + ((b.top - b.bottom) * 0.5);

    var txt = _dim_createAnchoredText(label, "C", cx, cy, {size: textPt * scaleAppearance, textColor: textColor}, lyr);
    if(txt) {
      labelsAdded++;
      objectsProcessed++;
      _dim_areaDbg('label added | text=' + label);
    }
  }

  var debugText = (_dim_areaDebugLines && _dim_areaDebugLines.length) ? ('\n' + _dim_areaDebugLines.join('\n')) : '';
  _dim_areaDebugLines = null;
  if(!labelsAdded) return 'No measurable closed shapes in selection.' + debugText;
  return 'Added ' + labelsAdded + ' area label' + (labelsAdded === 1 ? '' : 's') +
    ' on ' + objectsProcessed + ' object' + (objectsProcessed === 1 ? '' : 's') + '.' + debugText;
}

function _dim_runLine(opts) {
  opts = opts || {};
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
  var replaceOriginal = !!opts.replaceOriginal;

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
  var replaced = 0;
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
      var off = replaceOriginal ? 0 : lineOff;

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

    if(replaceOriginal) {
      try {
        if(item && (item.typename === "PathItem" || item.typename === "CompoundPathItem")) {
          item.remove();
          replaced++;
        } else if(path && path.typename === "PathItem") {
          path.remove();
          replaced++;
        }
      } catch(_eReplace) { }
    }
  }

  if(!added) return 'No measurable paths in selection.';
  if(replaceOriginal) {
    return 'Added ' + added + ' line measure' + (added === 1 ? '' : 's') +
      ' and replaced ' + replaced + ' source path' + (replaced === 1 ? '' : 's') + '.';
  }
  return 'Added ' + added + ' line measure' + (added === 1 ? '' : 's') + '.';
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
try {if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_panelSettingsSave = this.signarama_helper_panelSettingsSave;} catch(_ePss0) { }
try {signarama_helper_panelSettingsSave = this.signarama_helper_panelSettingsSave;} catch(_ePss1) { }

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
try {if(typeof $ !== 'undefined' && $.global) $.global.signarama_helper_panelSettingsLoad = this.signarama_helper_panelSettingsLoad;} catch(_ePsl0) { }
try {signarama_helper_panelSettingsLoad = this.signarama_helper_panelSettingsLoad;} catch(_ePsl1) { }

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

this.atlas_dimensions_runLineReplace = function(json) {
  try {
    var opts;
    if(typeof json === 'string') {
      try {opts = JSON.parse(json);}
      catch(_) {opts = eval('(' + json + ')');}
    } else {
      opts = json;
    }
    opts = opts || {};
    opts.replaceOriginal = true;
    opts.offsetMm = 0;
    return _dim_runLine(opts);
  } catch(e) {
    return 'Error: ' + e.message;
  }
};

this.atlas_dimensions_runAnglesInner = function(json) {
  try {
    var opts;
    if(typeof json === 'string') {
      try {opts = JSON.parse(json);}
      catch(_) {opts = eval('(' + json + ')');}
    } else {
      opts = json;
    }
    return _dim_runAngles(opts, 'INNER');
  } catch(e) {
    return 'Error: ' + e.message;
  }
};

this.atlas_dimensions_runAnglesOuter = function(json) {
  try {
    var opts;
    if(typeof json === 'string') {
      try {opts = JSON.parse(json);}
      catch(_) {opts = eval('(' + json + ')');}
    } else {
      opts = json;
    }
    return _dim_runAngles(opts, 'OUTER');
  } catch(e) {
    return 'Error: ' + e.message;
  }
};

this.atlas_dimensions_runArea = function(json) {
  try {
    var opts;
    if(typeof json === 'string') {
      try {opts = JSON.parse(json);}
      catch(_) {opts = eval('(' + json + ')');}
    } else {
      opts = json;
    }
    return _dim_runArea(opts);
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

function _srh_letter_collectPlacementRoots(container, out) {
  if(!container || !out) return;
  var items = null;
  try {items = container.pageItems;} catch(_eLcp0) {items = null;}
  if(!items || !items.length) return;
  for(var i = 0; i < items.length; i++) {
    var it = items[i];
    if(!it) continue;
    var tn = '';
    try {tn = String(it.typename || '');} catch(_eLcp1) {tn = '';}
    if(tn === 'PathItem') {
      try {if(it.guides || it.clipping || !it.closed) continue;} catch(_eLcp2) { }
      out.push(it);
      continue;
    }
    if(tn === 'CompoundPathItem') {
      out.push(it);
      continue;
    }
    if(tn === 'GroupItem') _srh_letter_collectPlacementRoots(it, out);
  }
}

var _srh_letterDebugLines = null;
function _srh_letterDbg(msg) {
  var line = '[SRH][letter] ' + String(msg || '');
  try {$.writeln(line);} catch(_eLdbg0) { }
  try {
    if(_srh_letterDebugLines) _srh_letterDebugLines.push(line);
  } catch(_eLdbg1) { }
}

function _srh_letter_collectPlacementPaths(item, out) {
  if(!item || !out) return;
  var tn = '';
  try {tn = String(item.typename || '');} catch(_eLcpp0) {tn = '';}
  if(tn === 'PathItem') {
    try {if(item.guides || item.clipping || !item.closed) return;} catch(_eLcpp1) { }
    out.push(item);
    return;
  }
  if(tn === 'CompoundPathItem') {
    var kids = null;
    try {kids = item.pathItems;} catch(_eLcpp2) {kids = null;}
    if(!kids || !kids.length) return;
    for(var i = 0; i < kids.length; i++) {
      var cp = kids[i];
      if(!cp) continue;
      try {if(cp.guides || cp.clipping || !cp.closed) continue;} catch(_eLcpp3) { }
      out.push(cp);
    }
    return;
  }
  if(tn === 'GroupItem') {
    var items = null;
    try {items = item.pageItems;} catch(_eLcpp4) {items = null;}
    if(!items || !items.length) return;
    for(var j = 0; j < items.length; j++) _srh_letter_collectPlacementPaths(items[j], out);
  }
}

function _srh_letter_pointAndTangentAtDistance(points, distance) {
  if(!points || points.length < 2) return null;
  var remaining = Number(distance || 0);
  if(!(remaining >= 0)) remaining = 0;
  for(var i = 0; i < points.length - 1; i++) {
    var p0 = points[i];
    var p1 = points[i + 1];
    if(!p0 || !p1) continue;
    var dx = p1.x - p0.x;
    var dy = p1.y - p0.y;
    var segLen = Math.sqrt(dx * dx + dy * dy);
    if(!(segLen > 0.0001)) continue;
    if(remaining <= segLen || i === points.length - 2) {
      var t = Math.max(0, Math.min(1, remaining / segLen));
      return {
        x: p0.x + dx * t,
        y: p0.y + dy * t,
        angle: Math.atan2(dy, dx)
      };
    }
    remaining -= segLen;
  }
  return null;
}

function _srh_letter_polylineLength(points) {
  if(!points || points.length < 2) return 0;
  var total = 0;
  for(var i = 0; i < points.length - 1; i++) {
    var p0 = points[i];
    var p1 = points[i + 1];
    if(!p0 || !p1) continue;
    var dx = p1.x - p0.x;
    var dy = p1.y - p0.y;
    total += Math.sqrt(dx * dx + dy * dy);
  }
  return total;
}

function _srh_letter_drawModule(ledLayer, lineLayer, cx, cy, widthPt, heightPt, angleRad, strokePt, black, red) {
  var cosA = Math.cos(angleRad);
  var sinA = Math.sin(angleRad);
  var hx = widthPt * 0.5;
  var hy = heightPt * 0.5;
  var corners = [
    {x: -hx, y: -hy},
    {x: hx, y: -hy},
    {x: hx, y: hy},
    {x: -hx, y: hy}
  ];
  var pathPts = [];
  for(var i = 0; i < corners.length; i++) {
    var pt = corners[i];
    pathPts.push([
      cx + (pt.x * cosA) - (pt.y * sinA),
      cy + (pt.x * sinA) + (pt.y * cosA)
    ]);
  }
  pathPts.push(pathPts[0]);

  var rect = null;
  try {
    rect = ledLayer.pathItems.add();
    rect.setEntirePath(pathPts);
    rect.closed = true;
    rect.filled = false;
    rect.stroked = true;
    rect.strokeWidth = strokePt;
    rect.strokeColor = black;
  } catch(_eLdm0) { }

  try {
    var line = lineLayer.pathItems.add();
    line.setEntirePath([
      [cx - (hx * cosA), cy - (hx * sinA)],
      [cx + (hx * cosA), cy + (hx * sinA)]
    ]);
    line.filled = false;
    line.stroked = true;
    line.strokeWidth = strokePt;
    line.strokeColor = red;
  } catch(_eLdm1) { }

  return rect;
}

function _srh_letterRunOnSelection(doc, items, fn) {
  if(!doc || !items || !fn) return false;
  var list = items.length ? items : [items];
  if(!list.length) return false;
  var prevSel = null;
  try {prevSel = doc.selection;} catch(_eLrs0) {prevSel = null;}
  try {doc.selection = null;} catch(_eLrs1) { }
  try {
    for(var i = 0; i < list.length; i++) {
      try {if(list[i]) list[i].selected = true;} catch(_eLrs2) { }
    }
    fn();
    return true;
  } catch(_eLrs3) {
    return false;
  } finally {
    try {doc.selection = null;} catch(_eLrs4) { }
    try {if(prevSel) doc.selection = prevSel;} catch(_eLrs5) { }
  }
}

function _srh_letterExpandItems(doc, items) {
  if(!doc || !items) return [];
  var list = items.length ? items : [items];
  if(!list.length) return [];
  var prevSel = null;
  try {prevSel = doc.selection;} catch(_eLex0) {prevSel = null;}
  try {doc.selection = null;} catch(_eLex1) { }
  try {
    for(var i = 0; i < list.length; i++) {
      try {if(list[i]) list[i].selected = true;} catch(_eLex2) { }
    }
    _srh_letterExpandSelection();
    var out = [];
    try {
      if(doc.selection && doc.selection.length) {
        for(var j = 0; j < doc.selection.length; j++) {
          try {if(doc.selection[j]) out.push(doc.selection[j]);} catch(_eLex3) { }
        }
      }
    } catch(_eLex4) { }
    return out;
  } catch(_eLex5) {
    return [];
  } finally {
    try {doc.selection = null;} catch(_eLex6) { }
    try {if(prevSel) doc.selection = prevSel;} catch(_eLex7) { }
  }
}

function _srh_letterExpandSelection() {
  try {app.executeMenuCommand('expandStyle');} catch(_eLexp0) { }
  try {app.executeMenuCommand('expand');} catch(_eLexp1) { }
}

function _srh_letter_offsetSourceItem(sourceItem, offsetPt, tempLayer) {
  if(!sourceItem || !tempLayer || !(offsetPt > 0)) return [];
  _srh_letterDbg('offset source | type=' + String(sourceItem.typename || '') + ' | offsetPt=' + offsetPt);
  var working = null;
  var wrapper = null;
  var doc = null;
  try {doc = app.activeDocument;} catch(_eLosDoc) {doc = null;}
  try {
    working = sourceItem.duplicate(tempLayer, ElementPlacement.PLACEATEND);
    _srh_letterDbg('offset duplicate OK | type=' + String(working.typename || ''));
  } catch(_eLos0) {
    _srh_letterDbg('offset duplicate FAILED | ' + String(_eLos0));
    return [];
  }
  try {
    wrapper = tempLayer.groupItems.add();
    working.move(wrapper, ElementPlacement.PLACEATEND);
    _srh_letterDbg('offset wrapper OK');
  } catch(_eLosWrap) {
    _srh_letterDbg('offset wrapper FAILED | ' + String(_eLosWrap));
    try {if(working) working.remove();} catch(_eLosWrapRm) { }
    return [];
  }
  var result = {applied:false, changed:false};
  var expandedItems = [];
  try {
    var beforeBounds = null;
    try {beforeBounds = wrapper.visibleBounds;} catch(_eLosB0) {beforeBounds = null;}
    if(!beforeBounds || beforeBounds.length !== 4) {
      try {beforeBounds = wrapper.geometricBounds;} catch(_eLosB1) {beforeBounds = null;}
    }
    var fx1 = '<LiveEffect name="Adobe Offset Path"><Dict data="R ofst ' + Number(-Math.abs(offsetPt)) + ' I jntp 0 R mlim 180"/></LiveEffect>';
    var fx2 = '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 180 R ofst ' + Number(-Math.abs(offsetPt)) + ' I jntp 0"/></LiveEffect>';
    var applied = false;
    try {wrapper.applyEffect(fx1); applied = true;} catch(_eLosFx0) {
      try {wrapper.applyEffect(fx2); applied = true;} catch(_eLosFx1) { }
    }
    if(applied && doc) {
      expandedItems = _srh_letterExpandItems(doc, wrapper);
      _srh_letterDbg('offset expanded items=' + expandedItems.length);
    }
    var afterTarget = (expandedItems && expandedItems.length) ? expandedItems[0] : wrapper;
    var afterBounds = null;
    try {afterBounds = afterTarget.visibleBounds;} catch(_eLosA0) {afterBounds = null;}
    if(!afterBounds || afterBounds.length !== 4) {
      try {afterBounds = afterTarget.geometricBounds;} catch(_eLosA1) {afterBounds = null;}
    }
    var changed = false;
    if(beforeBounds && afterBounds && beforeBounds.length === 4 && afterBounds.length === 4) {
      for(var bi = 0; bi < 4; bi++) {
        if(Math.abs(Number(beforeBounds[bi]) - Number(afterBounds[bi])) > 0.25) {
          changed = true;
          break;
        }
      }
    }
    result = {applied: applied, changed: changed};
    _srh_letterDbg('offset local apply | applied=' + (applied ? 'yes' : 'no') + ' | changed=' + (changed ? 'yes' : 'no'));
  } catch(_eLos3) {
    _srh_letterDbg('offset apply FAILED | ' + String(_eLos3));
    result = {applied:false, changed:false};
  }
  var changed = false;
  try {changed = !!(result && result.changed);} catch(_eLos4) {changed = false;}
  _srh_letterDbg('offset result | changed=' + (changed ? 'yes' : 'no'));
  if(!changed) {
    try {if(wrapper) wrapper.remove();} catch(_eLos5) { }
    return [];
  }
  var roots = [];
  if(expandedItems && expandedItems.length) {
    for(var ei = 0; ei < expandedItems.length; ei++) {
      var ex = expandedItems[ei];
      if(!ex) continue;
      var exType = '';
      try {exType = String(ex.typename || '');} catch(_eLosExType) {exType = '';}
      if(exType === 'GroupItem') _srh_letter_collectPlacementRoots(ex, roots);
      else if(exType === 'PathItem' || exType === 'CompoundPathItem') roots.push(ex);
    }
  } else {
    _srh_letter_collectPlacementRoots(wrapper, roots);
  }
  if(!roots.length) {
    var workingType = '';
    try {workingType = String(wrapper.typename || '');} catch(_eLosType) {workingType = '';}
    if(workingType === 'PathItem' || workingType === 'CompoundPathItem') roots.push(wrapper);
  }
  _srh_letterDbg('offset roots=' + roots.length);
  if(!roots.length) {
    try {if(wrapper) wrapper.remove();} catch(_eLos6) { }
    return [];
  }
  var out = [];
  for(var i = 0; i < roots.length; i++) {
    var root = roots[i];
    if(!root) continue;
    try {
      var dup = root.duplicate(tempLayer, ElementPlacement.PLACEATEND);
      if(dup) out.push(dup);
    } catch(_eLos7) {
      _srh_letterDbg('offset root duplicate FAILED | ' + String(_eLos7));
    }
  }
  try {if(wrapper) wrapper.remove();} catch(_eLos8) { }
  return out;
}

function signarama_helper_drawLetterLayout(jsonStr) {
  if(!app.documents.length) return 'No open document.';
  _srh_letterDebugLines = [];
  var doc = app.activeDocument;
  if(!doc.selection || !doc.selection.length) return 'No selection. Select one or more path items.';

  var opts = {};
  try {opts = JSON.parse(String(jsonStr));} catch(_eLl0) {opts = {};}

  var ledWatt = Number(opts.ledWatt || 0);
  var ledWidthMm = Number(opts.ledWidthMm || 0);
  var ledHeightMm = Number(opts.ledHeightMm || 0);
  var centerSpacingMm = Number(opts.centerSpacingMm || 0);
  var initialOffsetMm = Number(opts.initialOffsetMm || 0);
  var subsequentOffsetMm = Number(opts.subsequentOffsetMm || 0);
  var rotationDeg = Number(opts.rotationDeg || 0);
  var maxLedsInSeries = parseInt(opts.maxLedsInSeries, 10);
  if(!maxLedsInSeries || maxLedsInSeries < 1) maxLedsInSeries = 50;

  if(!(ledWidthMm > 0) || !(ledHeightMm > 0)) return 'LED width/height must be > 0.';
  if(!(initialOffsetMm > 0)) return 'Initial offset must be > 0.';
  if(!(subsequentOffsetMm > 0)) return 'Subsequent offset must be > 0.';
  if(!(centerSpacingMm > 0)) centerSpacingMm = ledWidthMm;

  var ledWidthPt = _srh_mm2ptDoc(ledWidthMm);
  var ledHeightPt = _srh_mm2ptDoc(ledHeightMm);
  var centerSpacingPt = _srh_mm2ptDoc(centerSpacingMm);
  var initialOffsetPt = _srh_mm2ptDoc(initialOffsetMm);
  var subsequentOffsetPt = _srh_mm2ptDoc(subsequentOffsetMm);
  var rotationRad = rotationDeg * Math.PI / 180.0;
  var stroke1px = _srh_pxStrokeDoc(1);

  var black = new RGBColor();
  black.red = 0; black.green = 0; black.blue = 0;
  var red = new RGBColor();
  red.red = 255; red.green = 0; red.blue = 0;

  var ledLayer = _srh_getOrCreateLayer(doc, 'Letter Layout LEDs');
  var lineLayer = _srh_getOrCreateLayer(doc, 'Letter Layout Center Lines');
  var groupLayer = _srh_getOrCreateLayer(doc, 'Letter Layout Groups');
  var tempLayer = _srh_getOrCreateLayer(doc, '__SRH Letter Layout Temp');
  try {tempLayer.locked = false;} catch(_eLlt0) { }
  try {tempLayer.visible = true;} catch(_eLlt1) { }
  try {
    while(tempLayer.pageItems && tempLayer.pageItems.length) {
      try {tempLayer.pageItems[0].remove();} catch(_eLltClr0) {break;}
    }
  } catch(_eLltClr1) { }
  _srh_letterDbg('temp layer ready | locked=' + (function(){try{return !!tempLayer.locked;}catch(_e){return 'n/a';}}()) + ' | visible=' + (function(){try{return !!tempLayer.visible;}catch(_e){return 'n/a';}}()));

  var currentItems = [];
  for(var s = 0; s < doc.selection.length; s++) {
    try {
      if(doc.selection[s]) currentItems.push(doc.selection[s]);
    } catch(_eLlt2) { }
  }
  _srh_letterDbg('run start | selection=' + currentItems.length);
  if(!currentItems.length) return 'No valid selection items to process.';

  var offsets = [initialOffsetPt];
  var totalPlaced = 0;
  var totalPaths = 0;
  var totalOffsetPaths = 0;
  var currentOffsetPt = initialOffsetPt;
  var loopGuard = 0;

  while(currentItems.length && loopGuard < 500) {
    loopGuard++;
    _srh_letterDbg('offset pass | pass=' + loopGuard + ' | items=' + currentItems.length + ' | offsetPt=' + currentOffsetPt);
    var nextItems = [];
    for(var ci = 0; ci < currentItems.length; ci++) {
      var sourceItem = currentItems[ci];
      if(!sourceItem) continue;
      _srh_letterDbg('source item | idx=' + ci + ' | type=' + String(sourceItem.typename || ''));
      var offsetItems = _srh_letter_offsetSourceItem(sourceItem, currentOffsetPt, tempLayer);
      try {
        if(sourceItem.layer && sourceItem.layer === tempLayer) sourceItem.remove();
      } catch(_eLlt3) { }
      _srh_letterDbg('offset items | idx=' + ci + ' | count=' + (offsetItems ? offsetItems.length : 0));
      if(!offsetItems || !offsetItems.length) continue;

      for(var oi = 0; oi < offsetItems.length; oi++) {
        var offsetItem = offsetItems[oi];
        if(!offsetItem) continue;
        totalOffsetPaths++;
        _srh_letterDbg('offset item | idx=' + oi + ' | type=' + String(offsetItem.typename || ''));

        var placementPaths = [];
        _srh_letter_collectPlacementPaths(offsetItem, placementPaths);
        _srh_letterDbg('placement paths=' + placementPaths.length);
        if(!placementPaths.length) continue;

        for(var pp = 0; pp < placementPaths.length; pp++) {
          var placePath = placementPaths[pp];
          if(!placePath) continue;
          var poly = _dim_pathToPolygonPoints(placePath, {stepPt: Math.max(_srh_mm2ptDoc(2), Math.min(centerSpacingPt / 4, _srh_mm2ptDoc(10)))});
          var pathLen = _srh_letter_polylineLength(poly);
          _srh_letterDbg('path | idx=' + pp + ' | pts=' + (poly ? poly.length : 0) + ' | len=' + pathLen);
          if(!(pathLen > 0.0001)) continue;
          totalPaths++;

          var pathPlaced = 0;
          var groupBounds = null;
          for(var dist = 0; dist <= pathLen + 0.0001; dist += centerSpacingPt) {
            var sample = _srh_letter_pointAndTangentAtDistance(poly, dist);
            if(!sample) continue;
            _srh_letter_drawModule(ledLayer, lineLayer, sample.x, sample.y, ledWidthPt, ledHeightPt, sample.angle + rotationRad, stroke1px, black, red);
            pathPlaced++;
            totalPlaced++;
            if(groupBounds === null) {
              groupBounds = {left: sample.x, top: sample.y, right: sample.x, bottom: sample.y};
            } else {
              if(sample.x < groupBounds.left) groupBounds.left = sample.x;
              if(sample.x > groupBounds.right) groupBounds.right = sample.x;
              if(sample.y > groupBounds.top) groupBounds.top = sample.y;
              if(sample.y < groupBounds.bottom) groupBounds.bottom = sample.y;
            }
          }

          if(groupBounds && pathPlaced > 0) {
            try {
              var label = groupLayer.textFrames.pointText([(groupBounds.left + groupBounds.right) * 0.5, groupBounds.bottom - _srh_mm2ptDoc(5)]);
              label.contents = pathPlaced + ' LEDs';
              try {label.textRange.paragraphAttributes.justification = Justification.CENTER;} catch(_eLlt4) { }
              try {label.textRange.characterAttributes.size = _srh_ptDoc(10);} catch(_eLlt5) { }
              try {label.textRange.characterAttributes.fillColor = black;} catch(_eLlt6) { }
            } catch(_eLlt7) { }
          }
          _srh_letterDbg('path placed=' + pathPlaced);
        }

        nextItems.push(offsetItem);
      }
    }
    currentItems = nextItems;
    currentOffsetPt = subsequentOffsetPt;
  }

  try {tempLayer.locked = false;} catch(_eLlt8a) { }
  try {tempLayer.visible = true;} catch(_eLlt8b) { }
  try {tempLayer.remove();} catch(_eLlt8) { }
  _srh_bringLayerToFront(lineLayer);
  _srh_bringLayerToFront(ledLayer);
  _srh_bringLayerToFront(groupLayer);

  var debugText = (_srh_letterDebugLines && _srh_letterDebugLines.length) ? ('\n' + _srh_letterDebugLines.join('\n')) : '';
  _srh_letterDebugLines = null;
  if(totalPlaced < 1) return 'Letter layout created no LED modules. Check offsets and selected path items.' + debugText;
  return 'Letter layout drawn. Offset paths: ' + totalOffsetPaths + ', Path lines: ' + totalPaths + ', LEDs: ' + totalPlaced + ', Watt: ' + ledWatt + 'W, Max series: ' + maxLedsInSeries + '.' + debugText;
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
