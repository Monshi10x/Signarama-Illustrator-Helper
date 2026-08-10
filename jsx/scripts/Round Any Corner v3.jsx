// CODING MODEL DIRECTIVE: Do not optimize for brevity; prioritize robustness and completeness.
#target illustrator

/*
    Round Any Corner.jsx
    Version 1.0.0

    Production-focused corner filleting for Adobe Illustrator.

    The script deliberately avoids Illustrator Live Corners. It measures each
    eligible line/line corner, solves the requested tangent setback, optionally
    relieves short edges by moving nearby anchors along those edges, and writes
    high-accuracy cubic representations of constant-radius circular arcs.

    Illustrator stores all native path curves as cubic Beziers. Each circular
    arc is therefore split into pieces no larger than 45 degrees. Every piece
    has the mathematically correct endpoints and tangents; splitting keeps the
    maximum radial deviation microscopic while retaining ordinary editable
    Illustrator paths.
*/

(function () {
    "use strict";

    // ---------------------------------------------------------------------
    // Configuration and numeric tolerances
    // ---------------------------------------------------------------------

    var SCRIPT_NAME = "Round Any Corner";
    var SCRIPT_VERSION = "1.0.0";
    var EPS = 1.0e-7;
    var POINT_EPS = 1.0e-4;
    var ANGLE_EPS = 1.0e-5;
    var EDGE_TOLERANCE = 1.0e-4;
    var MAX_RELIEF_PASSES = 40;
    var MAX_SELECTIVE_RETRIES = 12;
    var MAX_ARC_ANGLE = Math.PI / 4.0; // 45 degrees per cubic segment.
    var SELF_INTERSECTION_LIMIT = 450;
    var DEBUG_LAYER_PREFIX = "Round Any Corner Debug ";

    // ---------------------------------------------------------------------
    // Small, side-effect-free vector and scalar helpers
    // ---------------------------------------------------------------------

    function point(x, y) { return { x: x, y: y }; }
    function clonePoint(p) { return point(p.x, p.y); }
    function add(a, b) { return point(a.x + b.x, a.y + b.y); }
    function sub(a, b) { return point(a.x - b.x, a.y - b.y); }
    function mul(a, s) { return point(a.x * s, a.y * s); }
    function dot(a, b) { return a.x * b.x + a.y * b.y; }
    function cross(a, b) { return a.x * b.y - a.y * b.x; }
    function lengthSq(a) { return dot(a, a); }
    function length(a) { return Math.sqrt(lengthSq(a)); }
    function distance(a, b) { return length(sub(a, b)); }
    function perpLeft(a) { return point(-a.y, a.x); }
    function clamp(v, lo, hi) { return Math.max(lo, Math.min(hi, v)); }
    function sign(v) { return v < 0 ? -1 : (v > 0 ? 1 : 0); }

    function normalized(v) {
        var m = length(v);
        return m > EPS ? mul(v, 1.0 / m) : point(0, 0);
    }

    function nearlyEqualPoints(a, b, tolerance) {
        return distance(a, b) <= tolerance;
    }

    function fromArray(a) { return point(Number(a[0]), Number(a[1])); }
    function toArray(p) { return [p.x, p.y]; }

    function signedArea(anchors) {
        var area = 0;
        var i;
        var j;
        for (i = 0; i < anchors.length; i += 1) {
            j = (i + 1) % anchors.length;
            area += anchors[i].x * anchors[j].y - anchors[j].x * anchors[i].y;
        }
        return 0.5 * area;
    }

    function containsReference(items, candidate) {
        var i;
        for (i = 0; i < items.length; i += 1) {
            if (items[i] === candidate) { return true; }
        }
        return false;
    }

    function formatNumber(value, decimals) {
        var p = Math.pow(10, decimals);
        return String(Math.round(value * p) / p);
    }

    // ---------------------------------------------------------------------
    // Unit conversion
    // ---------------------------------------------------------------------

    function unitToPoints(value, unitName) {
        var factors = {
            "Millimetres": 72.0 / 25.4,
            "Centimetres": 72.0 / 2.54,
            "Points": 1.0,
            "Pixels": 1.0,
            "Inches": 72.0
        };
        return value * factors[unitName];
    }

    function pointsToUnit(value, unitName) {
        return value / unitToPoints(1.0, unitName);
    }

    function parsePositiveNumber(textValue, fieldName, allowZero) {
        var cleaned = String(textValue).replace(/,/g, ".").replace(/^\s+|\s+$/g, "");
        var value = Number(cleaned);
        if (!isFinite(value) || (allowZero ? value < 0 : value <= 0)) {
            throw new Error(fieldName + " must be " + (allowZero ? "zero or greater." : "greater than zero."));
        }
        return value;
    }

    // ---------------------------------------------------------------------
    // Selection traversal
    // ---------------------------------------------------------------------

    function itemIsUsable(item) {
        try {
            return !item.locked && !item.hidden;
        } catch (ignore) {
            return true;
        }
    }

    function collectClosedPathsFromItem(item, result, seen) {
        var i;
        var parent;
        if (!item) { return; }

        // Direct-selected anchors can appear as PathPoint objects in some
        // Illustrator selection modes. Promote them to their parent path.
        if (item.typename === "PathPoint") {
            try {
                parent = item.parent;
                collectClosedPathsFromItem(parent, result, seen);
            } catch (ignorePathPoint) {}
            return;
        }

        if (item.typename === "PathItem") {
            if (!containsReference(seen, item)) {
                seen.push(item);
                try {
                    if (item.closed && !item.guides && item.pathPoints.length >= 3 && itemIsUsable(item)) {
                        result.push(item);
                    }
                } catch (ignorePath) {}
            }
            return;
        }

        if (item.typename === "CompoundPathItem") {
            if (!itemIsUsable(item)) { return; }
            for (i = 0; i < item.pathItems.length; i += 1) {
                collectClosedPathsFromItem(item.pathItems[i], result, seen);
            }
            return;
        }

        if (item.typename === "GroupItem") {
            if (!itemIsUsable(item)) { return; }
            // pageItems includes clipping paths and clipped artwork. Traversing
            // it directly retains the clipping-group structure and modifies
            // eligible closed children in place.
            for (i = 0; i < item.pageItems.length; i += 1) {
                collectClosedPathsFromItem(item.pageItems[i], result, seen);
            }
        }
    }

    function collectSelectedClosedPaths(documentRef) {
        var result = [];
        var seen = [];
        var selection = documentRef.selection;
        var i;
        if (!selection || selection.length === 0) { return result; }
        for (i = 0; i < selection.length; i += 1) {
            collectClosedPathsFromItem(selection[i], result, seen);
        }
        return result;
    }

    // ---------------------------------------------------------------------
    // Geometry snapshots and exact restoration for preview/cancel
    // ---------------------------------------------------------------------

    function snapshotPath(pathItem) {
        var records = [];
        var i;
        var pp;
        for (i = 0; i < pathItem.pathPoints.length; i += 1) {
            pp = pathItem.pathPoints[i];
            records.push({
                anchor: fromArray(pp.anchor),
                left: fromArray(pp.leftDirection),
                right: fromArray(pp.rightDirection),
                pointType: pp.pointType
            });
        }
        return {
            item: pathItem,
            points: records,
            closed: pathItem.closed,
            area: signedArea(extractAnchors(records))
        };
    }

    function extractAnchors(records) {
        var anchors = [];
        var i;
        for (i = 0; i < records.length; i += 1) {
            anchors.push(clonePoint(records[i].anchor));
        }
        return anchors;
    }

    function writePathRecords(pathItem, records) {
        var anchors = [];
        var i;
        var pp;
        for (i = 0; i < records.length; i += 1) {
            anchors.push(toArray(records[i].anchor));
        }

        // setEntirePath changes the point count in one native call, which is
        // substantially faster and more reliable than repeatedly adding and
        // removing PathPoint objects during preview.
        pathItem.setEntirePath(anchors);
        pathItem.closed = true;

        for (i = 0; i < records.length; i += 1) {
            pp = pathItem.pathPoints[i];
            pp.leftDirection = toArray(records[i].left);
            pp.rightDirection = toArray(records[i].right);
            pp.pointType = records[i].pointType;
        }
    }

    function restoreSnapshots(snapshots) {
        var i;
        for (i = 0; i < snapshots.length; i += 1) {
            try {
                writePathRecords(snapshots[i].item, snapshots[i].points);
                snapshots[i].item.closed = snapshots[i].closed;
            } catch (restoreError) {
                throw new Error("Could not restore a previewed path. " + restoreError.message);
            }
        }
    }

    // ---------------------------------------------------------------------
    // Corner analysis
    // ---------------------------------------------------------------------

    function segmentIsStraight(records, startIndex, endIndex) {
        return nearlyEqualPoints(records[startIndex].right, records[startIndex].anchor, POINT_EPS) &&
            nearlyEqualPoints(records[endIndex].left, records[endIndex].anchor, POINT_EPS);
    }

    function buildCornerInfo(records, anchors, options, activeMask) {
        var corners = [];
        var area = signedArea(anchors);
        var orientation = sign(area);
        var n = anchors.length;
        var i;
        var prev;
        var next;
        var incoming;
        var outgoing;
        var rayToPrev;
        var rayToNext;
        var cosine;
        var alpha;
        var tangentDistance;
        var turn;
        var isConvex;
        var lineIn;
        var lineOut;
        var eligible;

        for (i = 0; i < n; i += 1) {
            prev = (i - 1 + n) % n;
            next = (i + 1) % n;
            incoming = normalized(sub(anchors[i], anchors[prev]));
            outgoing = normalized(sub(anchors[next], anchors[i]));
            rayToPrev = mul(incoming, -1);
            rayToNext = outgoing;
            cosine = clamp(dot(rayToPrev, rayToNext), -1, 1);
            alpha = Math.acos(cosine);
            turn = cross(incoming, outgoing);
            isConvex = orientation !== 0 && turn * orientation > 0;
            lineIn = segmentIsStraight(records, prev, i);
            lineOut = segmentIsStraight(records, i, next);
            eligible = orientation !== 0 &&
                lineIn && lineOut &&
                length(incoming) > EPS && length(outgoing) > EPS &&
                alpha > ANGLE_EPS && alpha < Math.PI - ANGLE_EPS &&
                Math.abs(turn) > ANGLE_EPS &&
                ((isConvex && options.roundConvex) || (!isConvex && options.roundConcave));

            if (activeMask && activeMask[i] === false) { eligible = false; }

            tangentDistance = eligible ? options.radius / Math.tan(alpha / 2.0) : 0;
            if (!isFinite(tangentDistance) || tangentDistance <= EPS) { eligible = false; }

            corners.push({
                index: i,
                eligible: eligible,
                convex: isConvex,
                alpha: alpha,
                turn: turn,
                tangentDistance: eligible ? tangentDistance : 0,
                radius: eligible ? options.radius : 0,
                scale: 1.0,
                lineIn: lineIn,
                lineOut: lineOut
            });
        }
        return corners;
    }

    function edgeRequirements(anchors, corners) {
        var edges = [];
        var i;
        var j;
        var available;
        var required;
        for (i = 0; i < anchors.length; i += 1) {
            j = (i + 1) % anchors.length;
            available = distance(anchors[i], anchors[j]);
            required = (corners[i].eligible ? corners[i].tangentDistance : 0) +
                (corners[j].eligible ? corners[j].tangentDistance : 0);
            edges.push({
                start: i,
                end: j,
                available: available,
                required: required,
                deficit: Math.max(0, required - available)
            });
        }
        return edges;
    }

    function badEdges(edges) {
        var result = [];
        var i;
        for (i = 0; i < edges.length; i += 1) {
            if (edges[i].deficit > EDGE_TOLERANCE) { result.push(edges[i]); }
        }
        return result;
    }

    // ---------------------------------------------------------------------
    // Adaptive relief engine
    // ---------------------------------------------------------------------

    function maximumStepInsideMovementCap(original, current, direction, cap) {
        var delta = sub(current, original);
        var projection = dot(delta, direction);
        var remainingSquared = cap * cap - lengthSq(delta);
        var discriminant;
        if (remainingSquared < -EDGE_TOLERANCE) { return 0; }
        discriminant = projection * projection + Math.max(0, remainingSquared);
        return Math.max(0, -projection + Math.sqrt(Math.max(0, discriminant)));
    }

    function clonePointArray(points) {
        var result = [];
        var i;
        for (i = 0; i < points.length; i += 1) { result.push(clonePoint(points[i])); }
        return result;
    }

    function expandDeficientEdge(anchors, originals, edge, movementCap) {
        var i = edge.start;
        var j = edge.end;
        var edgeVector = sub(anchors[j], anchors[i]);
        var edgeLength = length(edgeVector);
        var direction;
        var moveI;
        var moveJ;
        var need;
        var share;
        var actualI;
        var actualJ;
        var remaining;

        if (edgeLength <= EPS || edge.deficit <= EDGE_TOLERANCE || movementCap <= EPS) { return 0; }
        direction = mul(edgeVector, 1.0 / edgeLength);
        need = edge.deficit + EDGE_TOLERANCE;

        moveI = maximumStepInsideMovementCap(originals[i], anchors[i], mul(direction, -1), movementCap);
        moveJ = maximumStepInsideMovementCap(originals[j], anchors[j], direction, movementCap);
        share = need / 2.0;
        actualI = Math.min(share, moveI);
        actualJ = Math.min(share, moveJ);
        remaining = need - actualI - actualJ;

        if (remaining > EPS) {
            var extraI = Math.min(remaining, moveI - actualI);
            actualI += extraI;
            remaining -= extraI;
        }
        if (remaining > EPS) {
            var extraJ = Math.min(remaining, moveJ - actualJ);
            actualJ += extraJ;
        }

        anchors[i] = add(anchors[i], mul(direction, -actualI));
        anchors[j] = add(anchors[j], mul(direction, actualJ));
        return actualI + actualJ;
    }

    function runReliefPasses(records, originals, activeMask, options) {
        var anchors = clonePointArray(originals);
        var corners;
        var edges;
        var bad;
        var pass;
        var moved;
        var i;

        for (pass = 0; pass < MAX_RELIEF_PASSES; pass += 1) {
            corners = buildCornerInfo(records, anchors, options, activeMask);
            edges = edgeRequirements(anchors, corners);
            bad = badEdges(edges);
            if (bad.length === 0) {
                return { anchors: anchors, corners: corners, edges: edges, success: true, passes: pass };
            }

            moved = 0;
            // Process the worst deficits first to reduce oscillation where one
            // moved anchor participates in two constrained edges.
            bad.sort(function (a, b) { return b.deficit - a.deficit; });
            for (i = 0; i < bad.length; i += 1) {
                moved += expandDeficientEdge(anchors, originals, bad[i], options.maxMovement);
            }
            if (moved <= EDGE_TOLERANCE) { break; }
        }

        corners = buildCornerInfo(records, anchors, options, activeMask);
        edges = edgeRequirements(anchors, corners);
        return {
            anchors: anchors,
            corners: corners,
            edges: edges,
            success: badEdges(edges).length === 0,
            passes: MAX_RELIEF_PASSES
        };
    }

    function allTrue(count) {
        var result = [];
        var i;
        for (i = 0; i < count; i += 1) { result.push(true); }
        return result;
    }

    function solveExactRelief(snapshot, options) {
        var originals = extractAnchors(snapshot.points);
        var activeMask = allTrue(originals.length);
        var initiallyEligible = buildCornerInfo(snapshot.points, originals, options, activeMask);
        var failed = [];
        var attempt;
        var solution;
        var bad;
        var i;
        var changed;

        for (attempt = 0; attempt < MAX_SELECTIVE_RETRIES; attempt += 1) {
            solution = runReliefPasses(snapshot.points, originals, activeMask, options);
            if (solution.success) { break; }
            bad = badEdges(solution.edges);
            changed = false;

            // Disable only rounded endpoints that still participate in an
            // impossible edge. The solver is then restarted from the original
            // geometry so failed corners cannot leave residual deformation.
            for (i = 0; i < bad.length; i += 1) {
                if (solution.corners[bad[i].start].eligible && activeMask[bad[i].start]) {
                    activeMask[bad[i].start] = false;
                    failed.push(bad[i].start);
                    changed = true;
                }
                if (solution.corners[bad[i].end].eligible && activeMask[bad[i].end]) {
                    activeMask[bad[i].end] = false;
                    failed.push(bad[i].end);
                    changed = true;
                }
            }
            if (!changed) { break; }
        }

        solution = runReliefPasses(snapshot.points, originals, activeMask, options);
        solution.failed = uniqueNumbers(failed);
        solution.initiallyEligible = initiallyEligible;
        return solution;
    }

    function uniqueNumbers(values) {
        var result = [];
        var i;
        var j;
        var found;
        for (i = 0; i < values.length; i += 1) {
            found = false;
            for (j = 0; j < result.length; j += 1) {
                if (result[j] === values[i]) { found = true; break; }
            }
            if (!found) { result.push(values[i]); }
        }
        return result;
    }

    function solveReducedRadii(snapshot, options) {
        var anchors = extractAnchors(snapshot.points);
        var corners = buildCornerInfo(snapshot.points, anchors, options, null);
        var iteration;
        var edges;
        var i;
        var j;
        var required;
        var factor;

        // Non-exact mode keeps anchors fixed and proportionally reduces the
        // radii sharing any over-constrained edge.
        for (iteration = 0; iteration < 20; iteration += 1) {
            edges = edgeRequirements(anchors, corners);
            for (i = 0; i < edges.length; i += 1) {
                if (edges[i].deficit > EDGE_TOLERANCE && edges[i].required > EPS) {
                    factor = clamp((edges[i].available - EDGE_TOLERANCE) / edges[i].required, 0, 1);
                    j = edges[i].end;
                    if (corners[i].eligible) {
                        corners[i].scale *= factor;
                        corners[i].radius *= factor;
                        corners[i].tangentDistance *= factor;
                    }
                    if (corners[j].eligible) {
                        corners[j].scale *= factor;
                        corners[j].radius *= factor;
                        corners[j].tangentDistance *= factor;
                    }
                }
            }
        }

        // Remove vanishingly small results rather than generating coincident
        // points that Illustrator may discard unpredictably.
        for (i = 0; i < corners.length; i += 1) {
            if (corners[i].eligible && corners[i].radius <= POINT_EPS) {
                corners[i].eligible = false;
            }
        }
        required = edgeRequirements(anchors, corners);
        return { anchors: anchors, corners: corners, edges: required, success: true, failed: [] };
    }

    // ---------------------------------------------------------------------
    // Polygon safety checks
    // ---------------------------------------------------------------------

    function orientationValue(a, b, c) { return cross(sub(b, a), sub(c, a)); }

    function onSegment(a, b, p) {
        return p.x >= Math.min(a.x, b.x) - EDGE_TOLERANCE &&
            p.x <= Math.max(a.x, b.x) + EDGE_TOLERANCE &&
            p.y >= Math.min(a.y, b.y) - EDGE_TOLERANCE &&
            p.y <= Math.max(a.y, b.y) + EDGE_TOLERANCE;
    }

    function segmentsIntersect(a, b, c, d) {
        var o1 = orientationValue(a, b, c);
        var o2 = orientationValue(a, b, d);
        var o3 = orientationValue(c, d, a);
        var o4 = orientationValue(c, d, b);
        if (((o1 > EDGE_TOLERANCE && o2 < -EDGE_TOLERANCE) || (o1 < -EDGE_TOLERANCE && o2 > EDGE_TOLERANCE)) &&
                ((o3 > EDGE_TOLERANCE && o4 < -EDGE_TOLERANCE) || (o3 < -EDGE_TOLERANCE && o4 > EDGE_TOLERANCE))) {
            return true;
        }
        if (Math.abs(o1) <= EDGE_TOLERANCE && onSegment(a, b, c)) { return true; }
        if (Math.abs(o2) <= EDGE_TOLERANCE && onSegment(a, b, d)) { return true; }
        if (Math.abs(o3) <= EDGE_TOLERANCE && onSegment(c, d, a)) { return true; }
        if (Math.abs(o4) <= EDGE_TOLERANCE && onSegment(c, d, b)) { return true; }
        return false;
    }

    function polygonHasSelfIntersection(anchors) {
        var n = anchors.length;
        var i;
        var j;
        var i2;
        var j2;
        if (n > SELF_INTERSECTION_LIMIT) { return false; }
        for (i = 0; i < n; i += 1) {
            i2 = (i + 1) % n;
            for (j = i + 1; j < n; j += 1) {
                j2 = (j + 1) % n;
                if (i === j || i2 === j || j2 === i) { continue; }
                if (segmentsIntersect(anchors[i], anchors[i2], anchors[j], anchors[j2])) { return true; }
            }
        }
        return false;
    }

    function anyAnchorMoved(originals, current) {
        var i;
        for (i = 0; i < originals.length; i += 1) {
            if (distance(originals[i], current[i]) > EDGE_TOLERANCE) { return true; }
        }
        return false;
    }

    // ---------------------------------------------------------------------
    // Constant-radius arc construction
    // ---------------------------------------------------------------------

    function makePathRecord(anchor, left, right, pointType) {
        return { anchor: anchor, left: left, right: right, pointType: pointType };
    }

    function smoothPointType() {
        return PointType.SMOOTH;
    }

    function buildArcRecords(vertex, previous, next, corner) {
        var rayToPrev = normalized(sub(previous, vertex));
        var rayToNext = normalized(sub(next, vertex));
        var tangentIn = add(vertex, mul(rayToPrev, corner.tangentDistance));
        var tangentOut = add(vertex, mul(rayToNext, corner.tangentDistance));
        var bisector = normalized(add(rayToPrev, rayToNext));
        var centerDistance = corner.radius / Math.sin(corner.alpha / 2.0);
        var center = add(vertex, mul(bisector, centerDistance));
        var startVector = sub(tangentIn, center);
        var endVector = sub(tangentOut, center);
        var a0 = Math.atan2(startVector.y, startVector.x);
        var a1 = Math.atan2(endVector.y, endVector.x);
        var directionSign = sign(corner.turn);
        var sweep = a1 - a0;
        var count;
        var step;
        var records = [];
        var s;
        var startAngle;
        var endAngle;
        var p0;
        var p1;
        var tangent0;
        var tangent1;
        var handleLength;
        var leftHandle;
        var rightHandle;

        // Normalize the angular difference into the direction of path travel.
        if (directionSign > 0) {
            while (sweep <= 0) { sweep += Math.PI * 2.0; }
            while (sweep > Math.PI) { sweep -= Math.PI * 2.0; }
        } else {
            while (sweep >= 0) { sweep -= Math.PI * 2.0; }
            while (sweep < -Math.PI) { sweep += Math.PI * 2.0; }
        }

        // Subtract a tiny quotient tolerance so a numerically noisy 90-degree
        // sweep (for example 2.0000000000000004 x 45 degrees) does not create
        // an unnecessary third cubic segment.
        count = Math.max(1, Math.ceil(Math.abs(sweep) / MAX_ARC_ANGLE - 1.0e-10));
        step = sweep / count;

        for (s = 0; s < count; s += 1) {
            startAngle = a0 + step * s;
            endAngle = startAngle + step;
            p0 = add(center, point(Math.cos(startAngle) * corner.radius, Math.sin(startAngle) * corner.radius));
            p1 = add(center, point(Math.cos(endAngle) * corner.radius, Math.sin(endAngle) * corner.radius));
            tangent0 = mul(perpLeft(normalized(sub(p0, center))), sign(step));
            tangent1 = mul(perpLeft(normalized(sub(p1, center))), sign(step));
            handleLength = (4.0 / 3.0) * Math.tan(Math.abs(step) / 4.0) * corner.radius;
            rightHandle = add(p0, mul(tangent0, handleLength));
            leftHandle = sub(p1, mul(tangent1, handleLength));

            if (s === 0) {
                records.push(makePathRecord(p0, clonePoint(p0), rightHandle, smoothPointType()));
            } else {
                records[records.length - 1].right = rightHandle;
            }
            records.push(makePathRecord(p1, leftHandle, clonePoint(p1), smoothPointType()));
        }
        return { records: records, center: center, tangentIn: tangentIn, tangentOut: tangentOut };
    }

    function buildOutputRecords(snapshot, solution) {
        var output = [];
        var arcDebug = [];
        var originals = snapshot.points;
        var anchors = solution.anchors;
        var n = anchors.length;
        var i;
        var prev;
        var next;
        var delta;
        var arc;
        var j;

        for (i = 0; i < n; i += 1) {
            prev = (i - 1 + n) % n;
            next = (i + 1) % n;
            if (solution.corners[i].eligible) {
                arc = buildArcRecords(anchors[i], anchors[prev], anchors[next], solution.corners[i]);
                for (j = 0; j < arc.records.length; j += 1) { output.push(arc.records[j]); }
                arcDebug.push({ index: i, center: arc.center, tangentIn: arc.tangentIn, tangentOut: arc.tangentOut });
            } else {
                delta = sub(anchors[i], originals[i].anchor);
                output.push(makePathRecord(
                    clonePoint(anchors[i]),
                    add(originals[i].left, delta),
                    add(originals[i].right, delta),
                    originals[i].pointType
                ));
            }
        }
        return { records: output, arcs: arcDebug };
    }

    // ---------------------------------------------------------------------
    // Debug drawing
    // ---------------------------------------------------------------------

    function rgb(r, g, b) {
        var c = new RGBColor();
        c.red = r; c.green = g; c.blue = b;
        return c;
    }

    function removeDebugLayer(state) {
        if (state.debugLayer) {
            try { state.debugLayer.remove(); } catch (ignore) {}
            state.debugLayer = null;
        }
    }

    function getDebugLayer(documentRef, state) {
        if (!state.debugLayer) {
            state.debugLayer = documentRef.layers.add();
            state.debugLayer.name = DEBUG_LAYER_PREFIX + state.sessionId;
            state.debugLayer.printable = false;
        }
        return state.debugLayer;
    }

    function drawDebugLine(layer, a, b, colour, width) {
        var p = layer.pathItems.add();
        p.setEntirePath([toArray(a), toArray(b)]);
        p.closed = false;
        p.filled = false;
        p.stroked = true;
        p.strokeColor = colour;
        p.strokeWidth = width;
        return p;
    }

    function drawDebugCircle(layer, center, diameter, colour) {
        var circle = layer.pathItems.ellipse(center.y + diameter / 2.0, center.x - diameter / 2.0, diameter, diameter);
        circle.filled = true;
        circle.fillColor = colour;
        circle.stroked = false;
        return circle;
    }

    function drawDebugCross(layer, center, size, colour) {
        drawDebugLine(layer, point(center.x - size, center.y - size), point(center.x + size, center.y + size), colour, 0.75);
        drawDebugLine(layer, point(center.x - size, center.y + size), point(center.x + size, center.y - size), colour, 0.75);
    }

    function drawPathDebug(documentRef, state, snapshot, solution, output, options) {
        var layer = getDebugLayer(documentRef, state);
        var originals = extractAnchors(snapshot.points);
        var size = Math.max(2, Math.min(options.radius * 0.18, 8));
        var i;
        var failed;
        for (i = 0; i < originals.length; i += 1) {
            if (distance(originals[i], solution.anchors[i]) > EDGE_TOLERANCE) {
                drawDebugLine(layer, originals[i], solution.anchors[i], rgb(255, 140, 0), 0.75);
                drawDebugCircle(layer, solution.anchors[i], size, rgb(255, 140, 0));
            }
        }
        for (i = 0; i < output.arcs.length; i += 1) {
            drawDebugCircle(layer, output.arcs[i].center, size, rgb(0, 170, 255));
            drawDebugLine(layer, output.arcs[i].center, output.arcs[i].tangentIn, rgb(0, 170, 255), 0.5);
            drawDebugLine(layer, output.arcs[i].center, output.arcs[i].tangentOut, rgb(0, 170, 255), 0.5);
        }
        failed = solution.failed || [];
        for (i = 0; i < failed.length; i += 1) {
            drawDebugCross(layer, originals[failed[i]], size, rgb(255, 0, 60));
        }
    }

    // ---------------------------------------------------------------------
    // Whole-document application and reporting
    // ---------------------------------------------------------------------

    function emptyReport() {
        return {
            pathsVisited: 0,
            pathsChanged: 0,
            cornersRounded: 0,
            cornersRelieved: 0,
            cornersReduced: 0,
            cornersFailed: 0,
            pathsRejectedForIntersection: 0,
            errors: []
        };
    }

    function countEligible(corners) {
        var count = 0;
        var i;
        for (i = 0; i < corners.length; i += 1) {
            if (corners[i].eligible) { count += 1; }
        }
        return count;
    }

    function applyToSnapshots(documentRef, snapshots, options, state) {
        var report = emptyReport();
        var i;
        var snapshot;
        var solution;
        var output;
        var originals;
        var originalIntersected;
        var movedIntersected;
        var j;

        removeDebugLayer(state);

        for (i = 0; i < snapshots.length; i += 1) {
            snapshot = snapshots[i];
            report.pathsVisited += 1;
            try {
                solution = options.exactRadius ? solveExactRelief(snapshot, options) : solveReducedRadii(snapshot, options);
                originals = extractAnchors(snapshot.points);

                // Relief must never introduce a new coarse polygon crossing.
                // If it does, retain only corners that already fit without any
                // movement. This is conservative and prevents damaged artwork.
                if (options.exactRadius && anyAnchorMoved(originals, solution.anchors)) {
                    originalIntersected = polygonHasSelfIntersection(originals);
                    movedIntersected = polygonHasSelfIntersection(solution.anchors);
                    if (!originalIntersected && movedIntersected) {
                        var noMoveOptions = copyOptions(options);
                        noMoveOptions.maxMovement = 0;
                        solution = solveExactRelief(snapshot, noMoveOptions);
                        report.pathsRejectedForIntersection += 1;
                    }
                }

                output = buildOutputRecords(snapshot, solution);
                if (countEligible(solution.corners) > 0) {
                    writePathRecords(snapshot.item, output.records);
                    report.pathsChanged += 1;
                    report.cornersRounded += countEligible(solution.corners);
                }
                report.cornersFailed += solution.failed ? solution.failed.length : 0;

                for (j = 0; j < solution.corners.length; j += 1) {
                    if (solution.corners[j].eligible && solution.corners[j].scale < 1.0 - EDGE_TOLERANCE) {
                        report.cornersReduced += 1;
                    }
                    if (distance(originals[j], solution.anchors[j]) > EDGE_TOLERANCE) {
                        report.cornersRelieved += 1;
                    }
                }

                if (options.debug) {
                    drawPathDebug(documentRef, state, snapshot, solution, output, options);
                }
            } catch (pathError) {
                report.errors.push("Path " + (i + 1) + ": " + pathError.message);
                try { writePathRecords(snapshot.item, snapshot.points); } catch (ignoreRestore) {}
            }
        }
        return report;
    }

    function copyOptions(options) {
        return {
            radius: options.radius,
            maxMovement: options.maxMovement,
            exactRadius: options.exactRadius,
            roundConvex: options.roundConvex,
            roundConcave: options.roundConcave,
            debug: options.debug,
            unitName: options.unitName
        };
    }

    function reportText(report, options) {
        var unit = options.unitName;
        var lines = [];
        lines.push("Rounded " + report.cornersRounded + " corner" + (report.cornersRounded === 1 ? "" : "s") +
            " across " + report.pathsChanged + " path" + (report.pathsChanged === 1 ? "" : "s") + ".");
        if (report.cornersRelieved > 0) {
            lines.push(report.cornersRelieved + " anchor" + (report.cornersRelieved === 1 ? " was" : "s were") +
                " moved for relief (cap " + formatNumber(pointsToUnit(options.maxMovement, unit), 3) + " " + unit.toLowerCase() + ").");
        }
        if (report.cornersReduced > 0) {
            lines.push(report.cornersReduced + " corner radius" + (report.cornersReduced === 1 ? " was" : "es were") +
                " reduced because Exact radius was off.");
        }
        if (report.cornersFailed > 0) {
            lines.push(report.cornersFailed + " corner" + (report.cornersFailed === 1 ? " was" : "s were") +
                " left unchanged because the exact radius could not fit within the movement cap.");
        }
        if (report.pathsRejectedForIntersection > 0) {
            lines.push(report.pathsRejectedForIntersection + " path" + (report.pathsRejectedForIntersection === 1 ? " used" : "s used") +
                " conservative fallback because relief would have introduced a crossing.");
        }
        if (report.errors.length > 0) {
            lines.push("\nErrors:\n" + report.errors.slice(0, 8).join("\n"));
            if (report.errors.length > 8) { lines.push("...and " + (report.errors.length - 8) + " more."); }
        }
        return lines.join("\n");
    }

    // ---------------------------------------------------------------------
    // ScriptUI
    // ---------------------------------------------------------------------

    function buildDialog(documentRef, snapshots, state) {
        var dialog = new Window("dialog", SCRIPT_NAME + "  " + SCRIPT_VERSION);
        var panel = dialog.add("panel", undefined, "Fillet settings");
        var radiusGroup = panel.add("group");
        var radiusInput;
        var units;
        var typeGroup;
        var convex;
        var concave;
        var exact;
        var movementGroup;
        var movementInput;
        var movementLabel;
        var behaviourPanel;
        var preview;
        var debug;
        var status;
        var buttons;
        var applyButton;
        var cancelButton;
        var applied = false;
        var previewApplied = false;
        var lastReport = null;
        var previousUnit = "Millimetres";

        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.margins = 14;

        radiusGroup.add("statictext", undefined, "Radius:");
        radiusInput = radiusGroup.add("edittext", undefined, "5");
        radiusInput.characters = 10;
        units = radiusGroup.add("dropdownlist", undefined, ["Millimetres", "Centimetres", "Points", "Pixels", "Inches"]);
        units.selection = 0;

        typeGroup = panel.add("group");
        typeGroup.add("statictext", undefined, "Corners:");
        convex = typeGroup.add("checkbox", undefined, "Convex");
        concave = typeGroup.add("checkbox", undefined, "Concave");
        convex.value = true;
        concave.value = true;

        exact = panel.add("checkbox", undefined, "Exact radius (move anchors when required)");
        exact.value = true;

        movementGroup = panel.add("group");
        movementLabel = movementGroup.add("statictext", undefined, "Maximum vertex movement:");
        movementInput = movementGroup.add("edittext", undefined, "25");
        movementInput.characters = 10;

        behaviourPanel = dialog.add("panel", undefined, "Preview and diagnostics");
        behaviourPanel.orientation = "column";
        behaviourPanel.alignChildren = ["left", "top"];
        behaviourPanel.margins = 14;
        preview = behaviourPanel.add("checkbox", undefined, "Preview");
        debug = behaviourPanel.add("checkbox", undefined, "Debug overlays (centres, radii, movement, failures)");

        status = dialog.add("statictext", undefined, snapshots.length + " closed path" + (snapshots.length === 1 ? "" : "s") + " ready.", { multiline: true });
        status.preferredSize.width = 430;
        status.preferredSize.height = 36;

        buttons = dialog.add("group");
        buttons.alignment = "right";
        applyButton = buttons.add("button", undefined, "Apply", { name: "ok" });
        cancelButton = buttons.add("button", undefined, "Cancel", { name: "cancel" });

        function readOptions() {
            var unitName = units.selection.text;
            var radiusValue = parsePositiveNumber(radiusInput.text, "Radius", false);
            var movementValue = parsePositiveNumber(movementInput.text, "Maximum vertex movement", true);
            if (!convex.value && !concave.value) {
                throw new Error("Select Convex, Concave, or both.");
            }
            return {
                radius: unitToPoints(radiusValue, unitName),
                maxMovement: exact.value ? unitToPoints(movementValue, unitName) : 0,
                exactRadius: exact.value,
                roundConvex: convex.value,
                roundConcave: concave.value,
                debug: debug.value,
                unitName: unitName
            };
        }

        function updatePreview() {
            var options;
            if (!preview.value) {
                if (previewApplied) {
                    restoreSnapshots(snapshots);
                    removeDebugLayer(state);
                    previewApplied = false;
                    app.redraw();
                }
                status.text = snapshots.length + " closed path" + (snapshots.length === 1 ? "" : "s") + " ready.";
                return;
            }
            try {
                options = readOptions();
                restoreSnapshots(snapshots);
                lastReport = applyToSnapshots(documentRef, snapshots, options, state);
                previewApplied = true;
                status.text = lastReport.cornersRounded + " corners previewed; " + lastReport.cornersFailed + " exact-radius failures.";
                app.redraw();
            } catch (previewError) {
                try { restoreSnapshots(snapshots); } catch (ignore) {}
                removeDebugLayer(state);
                previewApplied = false;
                status.text = previewError.message;
                app.redraw();
            }
        }

        function updateMovementEnabled() {
            movementInput.enabled = exact.value;
            movementLabel.enabled = exact.value;
            updatePreview();
        }

        function convertDisplayedUnits() {
            var newUnit = units.selection.text;
            var radiusNumber;
            var movementNumber;
            try {
                radiusNumber = parsePositiveNumber(radiusInput.text, "Radius", false);
                movementNumber = parsePositiveNumber(movementInput.text, "Maximum vertex movement", true);
                radiusInput.text = formatNumber(pointsToUnit(unitToPoints(radiusNumber, previousUnit), newUnit), 4);
                movementInput.text = formatNumber(pointsToUnit(unitToPoints(movementNumber, previousUnit), newUnit), 4);
            } catch (ignoreConversion) {}
            previousUnit = newUnit;
            updatePreview();
        }

        radiusInput.onChange = updatePreview;
        movementInput.onChange = updatePreview;
        units.onChange = convertDisplayedUnits;
        convex.onClick = updatePreview;
        concave.onClick = updatePreview;
        exact.onClick = updateMovementEnabled;
        preview.onClick = updatePreview;
        debug.onClick = updatePreview;

        applyButton.onClick = function () {
            var options;
            try {
                options = readOptions();
                restoreSnapshots(snapshots);
                lastReport = applyToSnapshots(documentRef, snapshots, options, state);
                applied = true;
                previewApplied = false;
                if (!options.debug) { removeDebugLayer(state); }
                dialog.close(1);
                alert(reportText(lastReport, options), SCRIPT_NAME);
            } catch (applyError) {
                try { restoreSnapshots(snapshots); } catch (ignoreRestore) {}
                removeDebugLayer(state);
                previewApplied = false;
                alert(applyError.message, SCRIPT_NAME);
            }
        };

        cancelButton.onClick = function () {
            try { restoreSnapshots(snapshots); } catch (ignore) {}
            removeDebugLayer(state);
            previewApplied = false;
            dialog.close(0);
            app.redraw();
        };

        dialog.onClose = function () {
            if (!applied) {
                try { restoreSnapshots(snapshots); } catch (ignore) {}
                removeDebugLayer(state);
                app.redraw();
            }
        };

        dialog.center();
        dialog.show();
    }

    // ---------------------------------------------------------------------
    // Entry point
    // ---------------------------------------------------------------------

    function main() {
        var documentRef;
        var paths;
        var snapshots = [];
        var i;
        var state;

        if (app.documents.length === 0) {
            alert("Open a document and select one or more closed paths.", SCRIPT_NAME);
            return;
        }

        documentRef = app.activeDocument;
        paths = collectSelectedClosedPaths(documentRef);
        if (paths.length === 0) {
            alert("No usable closed paths were found. Open paths, guides, hidden items, and locked items are ignored.", SCRIPT_NAME);
            return;
        }

        for (i = 0; i < paths.length; i += 1) {
            try { snapshots.push(snapshotPath(paths[i])); } catch (ignoreSnapshot) {}
        }
        if (snapshots.length === 0) {
            alert("The selected artwork could not be read.", SCRIPT_NAME);
            return;
        }

        state = {
            debugLayer: null,
            sessionId: String((new Date()).getTime())
        };
        buildDialog(documentRef, snapshots, state);
    }

    try {
        main();
    } catch (fatalError) {
        alert(SCRIPT_NAME + " stopped:\n" + fatalError.message, SCRIPT_NAME);
    }
}());
