// Round Any Corner

// rounds selected corners of PathItems.
// Especially for the corners at the intersection point of curves,
// this script may work better than "Round Corners" filter (but slower).

// ## How To Use

//  1. Select the anchor(s) or whole path(s) to round.
//  2. Run this script. Adjust the values in the dialog.
//     Then click OK.

// ## Rounding Method
// Basically, the rounding method is compatible with the "Round Corners" filter.
// It is to add two anchors instead of the original anchor, at the points of
// specified line length from each selected corner.  So if there're too many
// anchors on original path, this script can not round nicely.

// ## Radius
// Actually, the specified "radius" is not for a radius of arcs which drawn.
// It is for the line length from each selected corner and is for the base
// to compute the length of handles.  The reason calling it "radius" is
// for compatibility with the "Round Corners" filter.

// ## Max TextBox
// You can change the max value of the slider by the the "max" textbox.
// Input the value and click "apply".
// Inputed number of digits after decimal point is reflected to the slider value.
// If the max value is "10.00", you can set the value like 8.23 with the slider.
// If it is "10", you can set the value as integer.

// This script does not round the corners which already rounded.
// (for example, select a circle and run this script does nothing)

// ### notice
// In the rounding process, the script merges anchors which nearly
// overlapped (when the distance between anchors is less than 0.05 points).
// When two rounded corners share a short segment, their cut-back distances
// are reduced so both corners survive. A very small part of the segment is
// deliberately retained; this prevents Illustrator from collapsing the two
// tangent points into one sharp anchor.

// This script does not work for some part of compound paths.
// When this occurs, please select part of the compound path or release the compound path and
// select them, then run script again.
// I still have not figured out how to get properties from grouped paths inside a compound path.

// If you prefer the slider interface of the previous version,
// uncomment the lines which marked as "uncomment to use slider-version".

// test env: Adobe Illustrator CC (Win/Mac)

// Copyright(c) 2005 Hiroyuki Sato
// https://github.com/shspage
// This script is distributed under the MIT License.
// See the LICENSE file for details.

// 2018.07.20, modified to ignore locked/hidden objects in a selected group

main();
function main(){
    // setting ----------------------------------------------
    
    // -- rr : rounding radius ( unit : point )
    
    // ------------------------------------------------------
    var conf = {
        rr : 5,
        maxSliderValue : "100",
        unit : "pt",
        errmsg : ""
        }
    
    var paths = [];
    getPathItemsInSelection(1, paths); // extract pathItems which pathpoints length is greater than 1
    if(paths.length < 1) return;

    var selectedSpec = getSelectedSpec(paths);
        
    var previewed = false;
    
    var clearPreview = function(){
        if(previewed){
            undo();
            redraw();
            previewed = false;
            applySelectedSpec( paths, selectedSpec );
        }
    }
    
    var drawPreview = function(){
        if( conf.rr > 0){
            roundAnyCorner( paths, conf );
            previewed = true;
        }
    }
    
    var getRoundDigit = function( s ){
        var n = 0;
        if(s.indexOf(".") >= 0){
            n = s.replace(/^[^\.]+\./, "").length;
        }
        return n;
    }
    
    if( conf.rr > conf.maxSliderValue ) conf.rr = conf.maxSliderValue;
    
    var round_digit = getRoundDigit( conf.maxSliderValue );
    
    // show a dialog
    var win = new Window("dialog", "Round Any Corner" );
    win.alignChildren = "fill";

    win.sliderPanel = win.add("panel", undefined, "radius");
    win.sliderPanel.orientation = "column";
    win.sliderPanel.alignChildren = "fill";
    win.sliderPanel.gp1 = win.sliderPanel.add("group");
    /* // uncomment to use slider-version
    win.sliderPanel.gp1.radiusSlider = win.sliderPanel.gp1.add("slider", undefined, conf.rr, 0, conf.maxSliderValue);
    win.sliderPanel.gp1.radiusSlider.size = [180, 20];
    */
    win.sliderPanel.gp1.txtBox = win.sliderPanel.gp1.add("edittext", undefined, conf.rr);
    win.sliderPanel.gp1.txtBox.justify = "right";
    win.sliderPanel.gp1.txtBox.characters = 8;
    /* // uncomment to use slider-version
    win.sliderPanel.gp1.txtBox.characters = 5;
    */
    win.sliderPanel.gp1.txtBox.helpTip = "hit TAB to set the input value temporarily";

    /* // uncomment to use slider-version
    win.sliderPanel.gp2 = win.sliderPanel.add("group");
    win.sliderPanel.gp2.alignment = "right";
    win.sliderPanel.gp2.maxValueCaptionText = win.sliderPanel.gp2.add("statictext", undefined, "max");
    win.sliderPanel.gp2.maxValueTextBox = win.sliderPanel.gp2.add("edittext", undefined, conf.maxSliderValue);
    win.sliderPanel.gp2.maxValueTextBox.characters = 6;
    win.sliderPanel.gp2.maxValueTextBox.justify = "right";
    win.sliderPanel.gp2.applyMaxValueButton = win.sliderPanel.gp2.add("button", undefined, "apply");
    win.sliderPanel.gp2.applyMaxValueButton.size = [60, 24];
    */
    
    win.unitRadioPanel = win.add("panel", undefined, "unit" );
    win.unitRadioPanel.orientation = "row";
    win.sliderPanel.alignChildren = "fill";
    win.unitRadioPanel.ptRb = win.unitRadioPanel.add("radiobutton", undefined, "pt/px");
    win.unitRadioPanel.mmRb = win.unitRadioPanel.add("radiobutton", undefined, "mm");
    win.unitRadioPanel.inchRb = win.unitRadioPanel.add("radiobutton", undefined, "inch");
    
    win.chkGroup = win.add("group");
    win.chkGroup.alignment = "center";
    win.chkGroup.previewChk = win.chkGroup.add("checkbox", undefined, "preview");
    
    win.btnGroup = win.add("group", undefined );
    win.btnGroup.alignment = "center";
    win.btnGroup.okBtn = win.btnGroup.add("button", undefined, "OK");
    win.btnGroup.cancelBtn = win.btnGroup.add("button", undefined, "Cancel");
    
    if( conf.unit == "mm" ){
        win.unitRadioPanel.mmRb.value = true;
    } else if( conf.unit == "inch"){
        win.unitRadioPanel.inchRb.value = true;
    } else {
        win.unitRadioPanel.ptRb.value = true;
        conf.unit = "pt";
    }
    
    var getValues = function(){
        conf.rr = win.sliderPanel.gp1.txtBox.text;
        
        if(win.unitRadioPanel.mmRb.value){
            conf.rr = convertUnit(win.sliderPanel.gp1.txtBox.text, "mm", "pt");
        } else if(win.unitRadioPanel.inchRb.value){
            conf.rr = convertUnit(win.sliderPanel.gp1.txtBox.text, "inch", "pt");
        }
    }
    
    var processPreview = function( is_preview ){
        if( ! is_preview || win.chkGroup.previewChk.value){
            try{
                win.enabled = false;
                getValues();
                clearPreview();
                drawPreview();
                if( is_preview ) redraw();
            } catch(e){
                alert(e);
            } finally{
                win.enabled = true;
            }
        }
    }

    /* // uncomment to use slider-version
    win.sliderPanel.gp2.applyMaxValueButton.onClick = function(){
        try{
            win.enabled = false;
            var mv = win.sliderPanel.gp2.maxValueTextBox.text.replace(/[^0-9\.]/g ,"");
            
            with(win.sliderPanel.gp1.radiusSlider){
                if( mv == "" || isNaN(mv) ){
                    alert("please input a number for the max value");
                    mv = maxvalue;
                } else {
                    round_digit = getRoundDigit( mv );
                }
                
                if(value > mv) value = mv;
                var v = value;
                maxvalue = mv;
                value = v;
            }
            with(win.sliderPanel.gp1){
                txtBox.text = radiusSlider.value.toFixed( round_digit );
            }
            processPreview( true );
        } catch(e){
            alert(e);
        } finally{
            win.enabled = true;
        }
    }
    */
  
    win.unitRadioPanel.ptRb.onClick = function(){
        processPreview( true );
    }
    win.unitRadioPanel.mmRb.onClick = function(){
        processPreview( true );
    }
    win.unitRadioPanel.inchRb.onClick = function(){
        processPreview( true );
    }
    win.chkGroup.previewChk.onClick = function(){
        if( win.chkGroup.previewChk.value ){
            processPreview( true );
        } else {
            if(previewed){
                clearPreview();
                redraw();
            }
        }
    }
    
    win.sliderPanel.gp1.txtBox.onChange = function(){
      var v = parseFloat(this.text);
      
      if(isNaN(v)){
        v = conf.rr;
      } else if(v < 0){
        v = 0;
      /* // uncomment to use slider-version
      } else if(v > win.sliderPanel.gp1.radiusSlider.maxvalue){
        v = win.sliderPanel.gp1.radiusSlider.maxvalue;
      */
      }
       this.text = v;
      /* // uncomment to use slider-version
      win.sliderPanel.gp1.radiusSlider.value = v;
      */
      processPreview( true );
    }

    /* // uncomment to use slider-version
    win.sliderPanel.gp1.radiusSlider.onChanging = function(){
        win.sliderPanel.gp1.txtBox.text = this.value.toFixed( round_digit );
    }
    win.sliderPanel.gp1.radiusSlider.onChange = function(){
        win.sliderPanel.gp1.txtBox.text = this.value.toFixed( round_digit );
        processPreview( true );
    }
    */
  
    win.btnGroup.okBtn.onClick = function(){
        processPreview( false );
        win.close();
    }
    
    win.btnGroup.cancelBtn.onClick = function(){
        try{
            win.enabled = false;
            clearPreview();
        } catch(e){
            alert(e);
        } finally{
            win.enabled = true;
        }
        win.close();
    }
    win.show();
    
    // error messasges aren't implemented for now
    if( conf.errmsg != "") alert( conf.errmsg );
}

function convertUnit(n, fromUnit, toUnit){
    if( fromUnit == "pt" ){
        if( toUnit == "mm"){
            n *= 0.352777778;
        } else if( toUnit == "inch"){
            n /= 72;
        }
    } else if( fromUnit == "mm"){
        if( toUnit == "pt"){
            n *= 2.83464567;
        } else if( toUnit == "inch"){
            n /= 25.4;
        }
    } else if( unit == "inch"){
        if( toUnit == "pt"){
            n *= 72;
        } else if( toUnit == "mm"){
            n *= 25.4;
        }
    }
    return n;
}

function roundAnyCorner( s, conf ){
  var rr = parseFloat(conf.rr);
  if(isNaN(rr) || rr <= 0) return;

  var p, pnts, vertices, rounded, pt;
  var i, j;

  for(j = 0; j < s.length; j++){
    p = s[j].pathPoints;
    if(readjustAnchors(p) < 2) continue;

    vertices = [];
    for(i = 0; i < p.length; i++){
      vertices.push({
        anchor: copyPnt(p[i].anchor),
        rightDirection: copyPnt(p[i].rightDirection),
        leftDirection: copyPnt(p[i].leftDirection),
        pointType: p[i].pointType,
        round: isSelected(p[i])
               && isCorner(p, i)
               && (s[j].closed || (i > 0 && i < p.length - 1))
      });
    }

    rounded = buildRoundedPathData(vertices, s[j].closed, rr);
    if(!rounded.changed) continue;
    pnts = rounded.points;

    for(i = p.length - 1; i > 0; i--) p[i].remove();

    for(i = 0; i < pnts.length; i++){
      pt = i > 0 ? p.add() : p[0];
      pt.anchor = pnts[i][0];
      pt.rightDirection = pnts[i][1];
      pt.leftDirection = pnts[i][2];
      pt.pointType = pnts[i][3];
    }
  }
  activeDocument.selection = s;
}

// ------------------------------------------------
// Build replacement path-point data without touching Illustrator objects.
// This separation keeps the short-segment allocation deterministic and makes
// the geometry testable outside Illustrator.
function buildRoundedPathData(vertices, closed, rr){
  var count = vertices.length;
  var segCount = closed ? count : count - 1;
  var segments = [];
  var effectiveRound = [];
  var pnts = [];
  var hanLen = 4 * (Math.sqrt(2) - 1) / 3;
  var smoothType = (typeof PointType != "undefined") ? PointType.SMOOTH : "SMOOTH";
  var i, nxi, q, fullLen, cornerTrims, t0, t1, sub, prevSeg, nextSeg;
  var v, incoming, outgoing, leftDir, rightDir;

  if(count < 2 || segCount < 1){
    return {points: [], changed: false, segments: []};
  }

  for(i = 0; i < segCount; i++){
    nxi = (i + 1) % count;
    q = [vertices[i].anchor,
         vertices[i].rightDirection,
         vertices[nxi].leftDirection,
         vertices[nxi].anchor];
    fullLen = getT4Len(q, 0);
    segments.push({q: q, fullLen: fullLen});
  }

  // A corner uses one cut-back value on both adjacent segments. Constraints
  // are therefore solved across the whole path, not independently per side.
  // This prevents a close corner from becoming a long, visibly pinched curve.
  cornerTrims = allocateCornerTrims(vertices, closed, rr, segments);

  for(i = 0; i < segCount; i++){
    nxi = (i + 1) % count;
    q = segments[i].q;
    fullLen = segments[i].fullLen;
    segments[i].trimStart = cornerTrims[i];
    segments[i].trimEnd = cornerTrims[nxi];
    t0 = cornerTrims[i] > 0 ? getT4Len(q, cornerTrims[i]) : 0;
    t1 = cornerTrims[nxi] > 0 ? getT4Len(q, -cornerTrims[nxi]) : 1;

    // Arc-length solving can converge to the same parameter on extremely
    // short curves. Retain a tiny parameter interval instead of collapsing it.
    if(t1 <= t0){
      var tm = (t0 + t1) / 2;
      var tp = Math.min(0.0001, 0.25 / Math.max(fullLen, 1));
      t0 = Math.max(0, tm - tp);
      t1 = Math.min(1, tm + tp);
    }

    sub = sliceBezier(q, t0, t1);
    segments[i].t0 = t0;
    segments[i].t1 = t1;
    segments[i].sub = sub;
  }

  for(i = 0; i < count; i++){
    effectiveRound[i] = !!vertices[i].round;
    if(!closed && (i == 0 || i == count - 1)) effectiveRound[i] = false;
    if(effectiveRound[i]){
      prevSeg = segments[(i - 1 + segCount) % segCount];
      nextSeg = segments[i % segCount];
      if(!prevSeg || !nextSeg
         || prevSeg.trimEnd <= 0
         || nextSeg.trimStart <= 0){
        effectiveRound[i] = false;
      }
    }
  }

  for(i = 0; i < count; i++){
    v = vertices[i];
    prevSeg = (!closed && i == 0) ? null
                                  : segments[(i - 1 + segCount) % segCount];
    nextSeg = (!closed && i == count - 1) ? null
                                          : segments[i % segCount];

    if(effectiveRound[i]){
      incoming = copyPnt(prevSeg.sub[3]);
      leftDir = copyPnt(prevSeg.sub[2]);
      rightDir = tangentHandleAtEnd(prevSeg.sub,
                                    prevSeg.trimEnd * hanLen);
      pnts.push([incoming, rightDir, leftDir, smoothType]);

      outgoing = copyPnt(nextSeg.sub[0]);
      leftDir = tangentHandleAtStart(nextSeg.sub,
                                     nextSeg.trimStart * hanLen);
      rightDir = copyPnt(nextSeg.sub[1]);
      pnts.push([outgoing, rightDir, leftDir, smoothType]);
    } else {
      leftDir = prevSeg ? copyPnt(prevSeg.sub[2]) : copyPnt(v.leftDirection);
      rightDir = nextSeg ? copyPnt(nextSeg.sub[1]) : copyPnt(v.rightDirection);
      pnts.push([copyPnt(v.anchor), rightDir, leftDir, v.pointType]);
    }
  }

  for(i = 0; i < effectiveRound.length; i++){
    if(effectiveRound[i]) return {points: pnts, changed: true, segments: segments};
  }
  return {points: pnts, changed: false, segments: segments};
}

// ------------------------------------------------
// Solve a compatible cut-back for every rounded corner. Each segment limits
// the sum of its two endpoint cut-backs. Repeated proportional reduction
// propagates the tightest constraint through clusters of nearby corners.
function allocateCornerTrims(vertices, closed, rr, segments){
  var count = vertices.length;
  var segCount = segments.length;
  var trims = [];
  var maxPasses = count * 4 + 4;
  var i, pass, nxi, requested, usable, scale, changed;

  for(i = 0; i < count; i++){
    trims[i] = vertices[i].round ? rr : 0;
    if(!closed && (i == 0 || i == count - 1)) trims[i] = 0;
  }

  for(pass = 0; pass < maxPasses; pass++){
    changed = false;
    for(i = 0; i < segCount; i++){
      nxi = (i + 1) % count;
      requested = trims[i] + trims[nxi];
      usable = getUsableSegmentLength(segments[i].fullLen);
      if(requested > usable && requested > 0){
        scale = usable / requested;
        trims[i] *= scale;
        trims[nxi] *= scale;
        changed = true;
      }
    }
    if(!changed) break;
  }
  return trims;
}

// ------------------------------------------------
// Share the usable length between rounded endpoints. Keeping a small centre
// gap is the key fix: the old code placed both tangent points at the exact
// midpoint and Illustrator could reduce them to one sharp point.
function allocateSegmentTrims(fullLen, roundStart, roundEnd, rr){
  var trimStart = roundStart ? rr : 0;
  var trimEnd = roundEnd ? rr : 0;
  var requested = trimStart + trimEnd;
  var usable, scale;

  if(fullLen <= 0 || requested <= 0){
    return {start: 0, end: 0};
  }

  usable = getUsableSegmentLength(fullLen);

  if(requested > usable){
    scale = usable / requested;
    trimStart *= scale;
    trimEnd *= scale;
  }
  return {start: trimStart, end: trimEnd};
}

// ------------------------------------------------
function getUsableSegmentLength(fullLen){
  // 0.01 pt is far below production tolerances but large enough to keep two
  // anchors distinct. On microscopic segments, reserve at most 1%.
  var reserve = Math.min(0.01, fullLen * 0.01);
  return Math.max(0, fullLen - reserve);
}

// ------------------------------------------------
function copyPnt(p){
  return [p[0], p[1]];
}

// ------------------------------------------------
function lerpPnt(a, b, t){
  return [a[0] + (b[0] - a[0]) * t,
          a[1] + (b[1] - a[1]) * t];
}

// ------------------------------------------------
function splitBezier(q, t){
  var a = lerpPnt(q[0], q[1], t);
  var b = lerpPnt(q[1], q[2], t);
  var c = lerpPnt(q[2], q[3], t);
  var d = lerpPnt(a, b, t);
  var e = lerpPnt(b, c, t);
  var f = lerpPnt(d, e, t);
  return [[copyPnt(q[0]), a, d, f],
          [f, e, c, copyPnt(q[3])]];
}

// ------------------------------------------------
// Exact cubic sub-curve for the original parameter interval [t0, t1].
function sliceBezier(q, t0, t1){
  var left, rel;
  if(t0 <= 0 && t1 >= 1){
    return [copyPnt(q[0]), copyPnt(q[1]), copyPnt(q[2]), copyPnt(q[3])];
  }
  if(t1 <= 0) return [copyPnt(q[0]), copyPnt(q[0]), copyPnt(q[0]), copyPnt(q[0])];
  left = splitBezier(q, Math.min(1, t1))[0];
  if(t0 <= 0) return left;
  rel = Math.min(1, t0 / t1);
  return splitBezier(left, rel)[1];
}

// ------------------------------------------------
function tangentHandleAtEnd(q, len){
  var anchor = q[3];
  var vx = anchor[0] - q[2][0];
  var vy = anchor[1] - q[2][1];
  var d = Math.sqrt(vx * vx + vy * vy);
  if(d < 0.0000001){
    vx = anchor[0] - q[1][0];
    vy = anchor[1] - q[1][1];
    d = Math.sqrt(vx * vx + vy * vy);
  }
  if(d < 0.0000001) return copyPnt(anchor);
  return [anchor[0] + vx / d * len,
          anchor[1] + vy / d * len];
}

// ------------------------------------------------
function tangentHandleAtStart(q, len){
  var anchor = q[0];
  var vx = q[1][0] - anchor[0];
  var vy = q[1][1] - anchor[1];
  var d = Math.sqrt(vx * vx + vy * vy);
  if(d < 0.0000001){
    vx = q[2][0] - anchor[0];
    vy = q[2][1] - anchor[1];
    d = Math.sqrt(vx * vx + vy * vy);
  }
  if(d < 0.0000001) return copyPnt(anchor);
  return [anchor[0] - vx / d * len,
          anchor[1] - vy / d * len];
}

// ------------------------------------------------
// return [x,y] of the distance "len" and the angle "rad"(in radian)
// from "pt"=[x,y]
function getPnt(pt, rad, len){
  return [pt[0] + Math.cos(rad) * len,
          pt[1] + Math.sin(rad) * len];
}

// ------------------------------------------------
// return the [x, y] coordinate of the handle of the point on the bezier curve
// that corresponds to the parameter "t"
// n=0:leftDir, n=1:rightDir
function defHan(t, q, n){
  return [t * (t * (q[n][0] - 2 * q[n+1][0] + q[n+2][0]) + 2 * (q[n+1][0] - q[n][0])) + q[n][0],
          t * (t * (q[n][1] - 2 * q[n+1][1] + q[n+2][1]) + 2 * (q[n+1][1] - q[n][1])) + q[n][1]];
}

// -----------------------------------------------
// return the [x, y] coordinate on the bezier curve
// that corresponds to the paramter "t"
function bezier(q, t) {
  var u = 1 - t;
  return [u*u*u * q[0][0] + 3*u*t*(u* q[1][0] + t* q[2][0]) + t*t*t * q[3][0],
          u*u*u * q[0][1] + 3*u*t*(u* q[1][1] + t* q[2][1]) + t*t*t * q[3][1]];
}

// ------------------------------------------------
// adjust the length of the handle "dir"
// by the magnification ratio "m",
// returns the modified [x, y] coordinate of the handle
// "anc" is the anchor [x, y]
function adjHan(anc, dir, m){
  return [anc[0] + (dir[0] - anc[0]) * m,
          anc[1] + (dir[1] - anc[1]) * m];
}

// ------------------------------------------------
// return true if the pathPoints "p[idx]" is a corner
function isCorner(p, idx){
  var pnt0 = getAnglePnt(p, idx, -1);
  var pnt1 = getAnglePnt(p, idx,  1);
  if(! pnt0 || ! pnt1) return false;                    // at the end of a open-path
  if(pnt0.length < 1 || pnt1.length<1) return false;   // anchor is overlapping, so cannot determine the angle
  var rad = getRad2(pnt0, p[idx].anchor, pnt1, true);
  if(rad > Math.PI - 0.1) return false;   // set the angle tolerance here
  return true;
}
// ------------------------------------------------
// "p"=pathPoints, "idx1"=index of pathpoint
// "dir" = -1, returns previous point [x,y] to get the angle of tangent at pathpoints[idx1]
// "dir" =  1, returns next ...
function getAnglePnt(p, idx1, dir){
  if(!dir) dir = -1;
  var idx2 = parseIdx(p, idx1 + dir);
  if(idx2 < 0) return null;  // at the end of a open-path
  var p2 = p[idx2];
  with(p[idx1]){
    if(dir<0){
      if(arrEq(leftDirection, anchor)){
        if(arrEq(p2.anchor, anchor)) return [];
        if(arrEq(p2.anchor, p2.rightDirection)
           || arrEq(p2.rightDirection, anchor)) return p2.anchor;
        else return p2.rightDirection;
      } else {
        return leftDirection;
      }
    } else {
      if(arrEq(anchor, rightDirection)){
        if(arrEq(anchor, p2.anchor)) return [];
        if(arrEq(p2.anchor, p2.leftDirection)
           || arrEq(anchor, p2.leftDirection)) return p2.anchor;
        else return p2.leftDirection;
      } else {
        return rightDirection;
      }
    }
  }
}
// --------------------------------------
// if the contents of both arrays are equal, return true (lengthes must be same)
function arrEq(arr1, arr2) {
  for(var i = 0; i < arr1.length; i++){
    if (arr1[i] != arr2[i]) return false;
  }
  return true;
}

// ------------------------------------------------
// return the distance between p1=[x,y] and p2=[x,y]
function dist(p1, p2) {
  return Math.sqrt(Math.pow(p1[0] - p2[0], 2)
                   + Math.pow(p1[1] - p2[1], 2));
}
// ------------------------------------------------
// return the squared distance between p1=[x,y] and p2=[x,y]
function dist2(p1, p2) {
  return Math.pow(p1[0] - p2[0],2)
       + Math.pow(p1[1] - p2[1],2);
}
// --------------------------------------
// return the angle in radian
// of the line drawn from p1=[x,y] from p2
function getRad(p1,p2) {
  return Math.atan2(p2[1] - p1[1],
                    p2[0] - p1[0]);
}

// --------------------------------------
// return the angle between two line segments
// o-p1 and o-p2 ( 0 - Math.PI)
function getRad2(p1, o, p2){
  var v1 = normalize(p1, o);
  var v2 = normalize(p2, o);
  return Math.acos(v1[0] * v2[0] + v1[1] * v2[1]);
}
// ------------------------------------------------
function normalize(p, o){
  var d = dist(p, o);
  return d == 0 ? [0, 0] : [(p[0] - o[0]) / d,
                            (p[1] - o[1]) / d];
}

// ------------------------------------------------
// return the bezier curve parameter "t"
// at the point which the length of the bezier curve segment
// (from the point start drawing) is "len"
// when "len" is 0, return the length of whole this segment.
function getT4Len(q, len){
  var m = [ q[3][0] - q[0][0] + 3 * (q[1][0] - q[2][0]),
            q[0][0] - 2 * q[1][0] + q[2][0],
            q[1][0] - q[0][0] ];
  var n = [ q[3][1] - q[0][1] + 3 * (q[1][1] - q[2][1]),
            q[0][1] - 2 * q[1][1] + q[2][1],
            q[1][1] - q[0][1] ];
  var k = [ m[0] * m[0] + n[0] * n[0],
            4 * (m[0] * m[1] + n[0] * n[1]),
            2 * ((m[0] * m[2] + n[0] * n[2]) + 2 * (m[1] * m[1] + n[1] * n[1])),
            4 * (m[1] * m[2] + n[1] * n[2]),
            m[2] * m[2] + n[2] * n[2] ];
  
   var fullLen = getLength(k, 1);

  if(len == 0){
    return fullLen;
    
  } else if(len < 0){
    len += fullLen;
    if(len < 0) return 0;

  } else if(len > fullLen){
    return 1;
  }
  
  var t, d;
  var t0 = 0;
  var t1 = 1;
  var torelance = 0.001;
  
  for(var h = 1; h < 30; h++){
    t = t0 + (t1 - t0) / 2;
    d = len - getLength(k, t);
    if(Math.abs(d) < torelance) break;
    else if(d < 0) t1 = t;
    else t0 = t;
  }
  return t;
}

// ------------------------------------------------
// return the length of bezier curve segment
// in range of parameter from 0 to "t"
function getLength(k, t){
  var h = t / 128;
  var hh = h * 2;
  var fc = function(t, k){
    return Math.sqrt(t * (t * (t * (t * k[0] + k[1]) + k[2]) + k[3]) + k[4]) || 0 };
  var total = (fc(0, k) - fc(t, k)) / 2;
  for(var i = h; i < t; i += hh) total += 2 * fc(i, k) + fc(i + h, k);
  return total * hh;
}

// ------------------------------------------------
// extract PathItems from the selection which length of PathPoints
// is greater than "n"
function getPathItemsInSelection(n, paths){
  if(documents.length < 1) return;
  
  var s = activeDocument.selection;
  
  if (!(s instanceof Array) || s.length < 1) return;

  extractPaths(s, n, paths);
}

// --------------------------------------
// extract PathItems from "s" (Array of PageItems -- ex. selection),
// and put them into an Array "paths".  If "pp_length_limit" is specified,
// this function extracts PathItems which PathPoints length is greater
// than this number.
function extractPaths(s, pp_length_limit, paths){
  for(var i = 0; i < s.length; i++){
    if(s[i].locked || s[i].hidden){
      continue;
    } else if(s[i].typename == "PathItem"){
      if(pp_length_limit
         && s[i].pathPoints.length <= pp_length_limit){
        continue;
      }
      paths.push(s[i]);
      
    } else if(s[i].typename == "GroupItem"){
      // search for PathItems in GroupItem, recursively
      extractPaths(s[i].pageItems, pp_length_limit, paths);
      
    } else if(s[i].typename == "CompoundPathItem"){
      // searches for pathitems in CompoundPathItem, recursively
      // ( ### Grouped PathItems in CompoundPathItem are ignored ### )
      extractPaths(s[i].pathItems, pp_length_limit , paths);
    }
  }
}

// --------------------------------------
// merge nearly overlapped anchor points 
// return the length of pathpoints after merging
function readjustAnchors(p){
  // Settings ==========================

  // merge the anchor points when the distance between
  // 2 points is within ### square root ### of this value (in point)
  var minDist = 0.0025; 
  
  // ===================================
  if(p.length < 2) return 1;
  var i;

  if(p.parent.closed){
    for(i = p.length - 1; i >= 1; i--){
      if(dist2(p[0].anchor, p[i].anchor) < minDist){
        p[0].leftDirection = p[i].leftDirection;
        p[i].remove();
      } else {
        break;
      }
    }
  }
  
  for(i = p.length - 1; i >= 1; i--){
    if(dist2(p[i].anchor, p[i - 1].anchor) < minDist){
      p[i - 1].rightDirection = p[i].rightDirection;
      p[i].remove();
    }
  }
  
  return p.length;
}
// -----------------------------------------------
// return pathpoint's index. when the argument is out of bounds,
// fixes it if the path is closed (ex. next of last index is 0),
// or return -1 if the path is not closed.
function parseIdx(p, n){ // PathPoints, number for index
  var len = p.length;
  if(p.parent.closed){
    return n >= 0 ? n % len : len - Math.abs(n % len);
  } else {
    return (n < 0 || n > len - 1) ? -1 : n;
  }
}
// -----------------------------------------------
function getDat(p){ // pathPoint
  with(p) return [anchor, rightDirection, leftDirection, pointType];
}
// -----------------------------------------------
function isSelected(p){ // PathPoint
  return p.selected == PathPointSelection.ANCHORPOINT;
}
// -----------------------------------------------
function getSelectedSpec( paths ){
    var specs = [];
    var j, pp, spec;
    for( var i = 0; i < paths.length; i++ ){
        pp = paths[i].pathPoints;
        spec = [];
        for( j = 0; j < pp.length; j++ ){
            spec.push( pp[j].selected );
        }
        specs.push( spec );
    }
    return specs;
}
// -----------------------------------------------
function applySelectedSpec( paths, specs ){
    var j, pp;
    for( var i = 0; i < paths.length; i++ ){
        pp = paths[i].pathPoints;
        for( j = 0; j < pp.length; j++ ){
            pp[j].selected = specs[i][j];
        }
    }
}
