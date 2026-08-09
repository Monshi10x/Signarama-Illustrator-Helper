#target illustrator

/*
 * Illustrator-Drain-Points.jsx
 * Computational-geometry drain placement for Adobe Illustrator.
 * ExtendScript / ECMAScript 3 compatible. No external dependencies.
 */

var DrainPoints = (function () {
    var CONFIG = {
        OUTPUT_LAYER_NAME: "Drain Holes",
        DEBUG_LAYER_NAME: "Drain Debug",
        TEMP_LAYER_NAME: "__DrainPoints_Normalize__",

        DRAIN_DIAMETER_MM: 3.0,
        FLATTEN_TOLERANCE_MM: 0.20,
        MAX_FLATTEN_DEPTH: 12,
        COINCIDENT_EPSILON_MM: 0.15,
        HORIZONTAL_EPSILON_MM: 0.03,
        SWEEP_EPSILON_MM: 0.03,

        IGNORE_HIDDEN_OR_LOCKED: true,
        EXPAND_APPEARANCE: true,
        CREATE_NEW_OUTPUT_LAYER: true,
        DEBUG: false,
        DEBUG_LABELS: false,

        MAX_FLATTENED_VERTICES: 250000,
        MAX_CANDIDATES: 20000,
        SHOW_SUMMARY: true
    };

    var PT_PER_MM = 72.0 / 25.4;

    function mm(value) { return value * PT_PER_MM; }
    function abs(value) { return value < 0 ? -value : value; }
    function min(a, b) { return a < b ? a : b; }
    function max(a, b) { return a > b ? a : b; }
    function sq(value) { return value * value; }
    function dist2(a, b) { return sq(a.x - b.x) + sq(a.y - b.y); }
    function point(x, y) { return { x: x, y: y }; }
    function clonePoint(p) { return point(p.x, p.y); }

    function makeSettings(overrides) {
        var result = {};
        var key;
        for (key in CONFIG) {
            if (CONFIG.hasOwnProperty(key)) { result[key] = CONFIG[key]; }
        }
        if (overrides) {
            for (key in overrides) {
                if (overrides.hasOwnProperty(key)) { result[key] = overrides[key]; }
            }
        }
        result.drainDiameter = mm(result.DRAIN_DIAMETER_MM);
        result.flattenTolerance = mm(result.FLATTEN_TOLERANCE_MM);
        result.coincidentEpsilon = mm(result.COINCIDENT_EPSILON_MM);
        result.horizontalEpsilon = mm(result.HORIZONTAL_EPSILON_MM);
        result.sweepEpsilon = mm(result.SWEEP_EPSILON_MM);
        return result;
    }

    function Logger() {
        this.warnings = [];
        this.stats = {
            sourceItems: 0,
            units: 0,
            contours: 0,
            flattenedVertices: 0,
            candidates: 0,
            rejectedCandidates: 0,
            drains: 0
        };
    }
    Logger.prototype.warn = function (message) { this.warnings.push(message); };

    function samePoint(a, b, epsilon) {
        return dist2(a, b) <= epsilon * epsilon;
    }

    function pointLineDistance(p, a, b) {
        var dx = b.x - a.x;
        var dy = b.y - a.y;
        var length2 = dx * dx + dy * dy;
        var cross;
        if (length2 === 0) { return Math.sqrt(dist2(p, a)); }
        cross = dy * p.x - dx * p.y + b.x * a.y - b.y * a.x;
        return abs(cross) / Math.sqrt(length2);
    }

    function midpoint(a, b) {
        return point((a.x + b.x) * 0.5, (a.y + b.y) * 0.5);
    }

    function flattenCubic(p0, p1, p2, p3, tolerance, maxDepth, out, depth) {
        var flatness = max(pointLineDistance(p1, p0, p3), pointLineDistance(p2, p0, p3));
        var p01, p12, p23, p012, p123, p0123;
        if (depth >= maxDepth || flatness <= tolerance) {
            out.push(clonePoint(p3));
            return;
        }
        p01 = midpoint(p0, p1);
        p12 = midpoint(p1, p2);
        p23 = midpoint(p2, p3);
        p012 = midpoint(p01, p12);
        p123 = midpoint(p12, p23);
        p0123 = midpoint(p012, p123);
        flattenCubic(p0, p01, p012, p0123, tolerance, maxDepth, out, depth + 1);
        flattenCubic(p0123, p123, p23, p3, tolerance, maxDepth, out, depth + 1);
    }

    function isZeroHandle(anchor, handle, epsilon) {
        return samePoint(anchor, handle, epsilon);
    }

    function flattenPathItem(pathItem, settings, logger) {
        var pathPoints = pathItem.pathPoints;
        var count = pathPoints.length;
        var result = [];
        var i, next, a, b, p0, p1, p2, p3;
        if (!pathItem.closed || count < 2) { return null; }
        p0 = point(pathPoints[0].anchor[0], pathPoints[0].anchor[1]);
        result.push(p0);
        for (i = 0; i < count; i += 1) {
            next = (i + 1) % count;
            a = pathPoints[i];
            b = pathPoints[next];
            p0 = point(a.anchor[0], a.anchor[1]);
            p1 = point(a.rightDirection[0], a.rightDirection[1]);
            p2 = point(b.leftDirection[0], b.leftDirection[1]);
            p3 = point(b.anchor[0], b.anchor[1]);
            if (isZeroHandle(p0, p1, 0.0001) && isZeroHandle(p3, p2, 0.0001)) {
                result.push(p3);
            } else {
                flattenCubic(p0, p1, p2, p3, settings.flattenTolerance,
                    settings.MAX_FLATTEN_DEPTH, result, 0);
            }
            if (logger.stats.flattenedVertices + result.length > settings.MAX_FLATTENED_VERTICES) {
                throw new Error("Flattened-vertex budget exceeded (" + settings.MAX_FLATTENED_VERTICES + "). Increase FLATTEN_TOLERANCE_MM or the budget.");
            }
        }
        if (result.length > 1 && samePoint(result[0], result[result.length - 1], 0.0001)) {
            result.pop();
        }
        logger.stats.flattenedVertices += result.length;
        return result.length >= 3 ? result : null;
    }

    function signedArea(points) {
        var area = 0;
        var i, j;
        for (i = 0; i < points.length; i += 1) {
            j = (i + 1) % points.length;
            area += points[i].x * points[j].y - points[j].x * points[i].y;
        }
        return area * 0.5;
    }

    function boundsOfContours(contours) {
        var left = Number.POSITIVE_INFINITY;
        var right = Number.NEGATIVE_INFINITY;
        var bottom = Number.POSITIVE_INFINITY;
        var top = Number.NEGATIVE_INFINITY;
        var i, j, p;
        for (i = 0; i < contours.length; i += 1) {
            for (j = 0; j < contours[i].points.length; j += 1) {
                p = contours[i].points[j];
                left = min(left, p.x); right = max(right, p.x);
                bottom = min(bottom, p.y); top = max(top, p.y);
            }
        }
        return { left: left, right: right, bottom: bottom, top: top };
    }

    function readEvenOdd(pathItem) {
        try { return pathItem.evenodd === true; } catch (error) { return false; }
    }

    function unitFromPath(pathItem, settings, logger, sourceName) {
        var points = flattenPathItem(pathItem, settings, logger);
        if (!points) { return null; }
        return {
            name: sourceName || pathItem.name || "Path",
            evenOdd: readEvenOdd(pathItem),
            contours: [{ points: points, area: signedArea(points) }]
        };
    }

    function unitFromCompound(compound, settings, logger) {
        var contours = [];
        var i, points;
        for (i = 0; i < compound.pathItems.length; i += 1) {
            points = flattenPathItem(compound.pathItems[i], settings, logger);
            if (points) { contours.push({ points: points, area: signedArea(points) }); }
        }
        if (!contours.length) { return null; }
        return {
            name: compound.name || "Compound Path",
            evenOdd: readEvenOdd(compound.pathItems[0]),
            contours: contours
        };
    }

    function collectUnitsFromItem(item, units, settings, logger) {
        var i, unit;
        if (!item) { return; }
        try {
            if (settings.IGNORE_HIDDEN_OR_LOCKED && (item.hidden || item.locked)) { return; }
        } catch (error0) {}

        switch (item.typename) {
        case "PathItem":
            if (item.parent && item.parent.typename === "CompoundPathItem") { return; }
            if (!item.closed) { return; }
            unit = unitFromPath(item, settings, logger, item.name);
            if (unit) { units.push(unit); }
            break;
        case "CompoundPathItem":
            unit = unitFromCompound(item, settings, logger);
            if (unit) { units.push(unit); }
            break;
        case "GroupItem":
        case "Layer":
            for (i = 0; i < item.pageItems.length; i += 1) {
                if (item.pageItems[i].parent === item) {
                    collectUnitsFromItem(item.pageItems[i], units, settings, logger);
                }
            }
            break;
        default:
            logger.warn("Ignored unsupported normalized item: " + item.typename);
            break;
        }
    }

    function deselectAll(document) {
        try { app.executeMenuCommand("deselectall"); }
        catch (error) {
            try { document.selection = null; } catch (error2) {}
        }
    }

    function duplicateSelectionToLayer(document, selection, tempLayer, logger) {
        var duplicates = [];
        var i, duplicate;
        deselectAll(document);
        for (i = 0; i < selection.length; i += 1) {
            try {
                duplicate = selection[i].duplicate(tempLayer, ElementPlacement.PLACEATEND);
                duplicate.selected = true;
                duplicates.push(duplicate);
                logger.stats.sourceItems += 1;
            } catch (error) {
                logger.warn("Could not duplicate " + selection[i].typename + ": " + error.message);
            }
        }
        return duplicates;
    }

    function outlineTextIn(container, logger) {
        var texts = [];
        var i;
        try {
            for (i = 0; i < container.textFrames.length; i += 1) { texts.push(container.textFrames[i]); }
        } catch (error0) { return; }
        for (i = texts.length - 1; i >= 0; i -= 1) {
            try { texts[i].createOutline(); }
            catch (error) { logger.warn("Could not outline a text frame: " + error.message); }
        }
    }

    function normalizeSelection(document, selection, settings, logger) {
        var tempLayer = document.layers.add();
        var duplicates;
        tempLayer.name = settings.TEMP_LAYER_NAME;
        duplicates = duplicateSelectionToLayer(document, selection, tempLayer, logger);
        outlineTextIn(tempLayer, logger);

        if (settings.EXPAND_APPEARANCE && duplicates.length) {
            try {
                app.executeMenuCommand("expandStyle");
            } catch (error) {
                logger.warn("Expand Appearance was unavailable; basic paths and outlined text were still processed.");
            }
        }
        return tempLayer;
    }

    function restoreSelection(document, originalSelection) {
        var i;
        deselectAll(document);
        for (i = 0; i < originalSelection.length; i += 1) {
            try { originalSelection[i].selected = true; } catch (error) {}
        }
    }

    function candidatePoint(p, kind, spanLeft, spanRight, contourIndex) {
        return {
            x: p.x,
            y: p.y,
            kind: kind,
            spanLeft: spanLeft,
            spanRight: spanRight,
            contourIndex: contourIndex,
            accepted: false,
            reason: "unclassified"
        };
    }

    function horizontalEdges(points, epsilon) {
        var flags = [];
        var i, j;
        for (i = 0; i < points.length; i += 1) {
            j = (i + 1) % points.length;
            flags[i] = abs(points[j].y - points[i].y) <= epsilon;
        }
        return flags;
    }

    function collectContourCandidates(points, contourIndex, settings) {
        var candidates = [];
        var count = points.length;
        var horizontal = horizontalEdges(points, settings.horizontalEpsilon);
        var visited = [];
        var i, start, end, prevIndex, nextIndex, left, right, y, cursor, guard;
        var prev, current, next;

        for (i = 0; i < count; i += 1) {
            if (!horizontal[i] || visited[i]) { continue; }
            start = i;
            cursor = i;
            guard = 0;
            left = min(points[i].x, points[(i + 1) % count].x);
            right = max(points[i].x, points[(i + 1) % count].x);
            y = (points[i].y + points[(i + 1) % count].y) * 0.5;
            while (horizontal[cursor] && !visited[cursor] && guard < count) {
                visited[cursor] = true;
                left = min(left, min(points[cursor].x, points[(cursor + 1) % count].x));
                right = max(right, max(points[cursor].x, points[(cursor + 1) % count].x));
                end = cursor;
                cursor = (cursor + 1) % count;
                guard += 1;
            }
            prevIndex = (start - 1 + count) % count;
            nextIndex = (end + 2) % count;
            if (points[prevIndex].y > y + settings.horizontalEpsilon &&
                    points[nextIndex].y > y + settings.horizontalEpsilon) {
                candidates.push(candidatePoint(point((left + right) * 0.5, y),
                    "flat", left, right, contourIndex));
            }
        }

        for (i = 0; i < count; i += 1) {
            if (horizontal[i] || horizontal[(i - 1 + count) % count]) { continue; }
            prev = points[(i - 1 + count) % count];
            current = points[i];
            next = points[(i + 1) % count];
            if (current.y < prev.y - settings.horizontalEpsilon &&
                    current.y < next.y - settings.horizontalEpsilon) {
                candidates.push(candidatePoint(current, "vertex", current.x, current.x, contourIndex));
            }
        }
        return candidates;
    }

    function collectCandidates(unit, settings, logger) {
        var result = [];
        var i, local, j;
        for (i = 0; i < unit.contours.length; i += 1) {
            local = collectContourCandidates(unit.contours[i].points, i, settings);
            for (j = 0; j < local.length; j += 1) { result.push(local[j]); }
        }
        logger.stats.candidates += result.length;
        if (logger.stats.candidates > settings.MAX_CANDIDATES) {
            throw new Error("Candidate budget exceeded (" + settings.MAX_CANDIDATES + "). Simplify the artwork or raise MAX_CANDIDATES.");
        }
        return result;
    }

    function sortNumbers(a, b) { return a - b; }
    function sortCrossings(a, b) { return a.x - b.x; }

    function lineCrossings(unit, y, epsilon) {
        var crossings = [];
        var c, pts, i, j, a, b, t, x, delta;
        for (c = 0; c < unit.contours.length; c += 1) {
            pts = unit.contours[c].points;
            for (i = 0; i < pts.length; i += 1) {
                j = (i + 1) % pts.length;
                a = pts[i]; b = pts[j];
                if (abs(a.y - b.y) <= epsilon) { continue; }
                if ((a.y <= y && y < b.y) || (b.y <= y && y < a.y)) {
                    t = (y - a.y) / (b.y - a.y);
                    x = a.x + t * (b.x - a.x);
                    delta = b.y > a.y ? 1 : -1;
                    crossings.push({ x: x, delta: delta });
                }
            }
        }
        crossings.sort(sortCrossings);
        return crossings;
    }

    function intervalsAtY(unit, y, epsilon) {
        var crossings = lineCrossings(unit, y, epsilon * 0.1);
        var intervals = [];
        var winding = 0;
        var parity = 0;
        var previousX = null;
        var i = 0;
        var x, delta, count, inside;
        while (i < crossings.length) {
            x = crossings[i].x;
            inside = unit.evenOdd ? (parity % 2 !== 0) : (winding !== 0);
            if (previousX !== null && inside && x - previousX > epsilon * 0.1) {
                intervals.push({ left: previousX, right: x });
            }
            delta = 0;
            count = 0;
            while (i < crossings.length && abs(crossings[i].x - x) <= epsilon * 0.1) {
                delta += crossings[i].delta;
                count += 1;
                i += 1;
            }
            winding += delta;
            parity += count;
            previousX = x;
        }
        return intervals;
    }

    function intervalOverlaps(a, b, epsilon) {
        return min(a.right, b.right) >= max(a.left, b.left) - epsilon;
    }

    function candidateTouchesInterval(candidate, interval, epsilon) {
        if (candidate.kind === "flat") {
            return min(candidate.spanRight, interval.right) >= max(candidate.spanLeft, interval.left) - epsilon;
        }
        return candidate.x >= interval.left - epsilon && candidate.x <= interval.right + epsilon;
    }

    function uniqueEventYs(unit, epsilon) {
        var values = [];
        var i, j;
        for (i = 0; i < unit.contours.length; i += 1) {
            for (j = 0; j < unit.contours[i].points.length; j += 1) {
                values.push(unit.contours[i].points[j].y);
            }
        }
        values.sort(sortNumbers);
        if (!values.length) { return values; }
        var unique = [values[0]];
        for (i = 1; i < values.length; i += 1) {
            if (values[i] - unique[unique.length - 1] > epsilon) { unique.push(values[i]); }
        }
        return unique;
    }

    function sweepOffsetForY(eventYs, y, settings) {
        var previous = null;
        var next = null;
        var i, gap, offset;
        for (i = 0; i < eventYs.length; i += 1) {
            if (eventYs[i] < y - settings.horizontalEpsilon) { previous = eventYs[i]; }
            if (eventYs[i] > y + settings.horizontalEpsilon) { next = eventYs[i]; break; }
        }
        gap = Number.POSITIVE_INFINITY;
        if (previous !== null) { gap = min(gap, y - previous); }
        if (next !== null) { gap = min(gap, next - y); }
        if (gap === Number.POSITIVE_INFINITY) { gap = settings.sweepEpsilon * 20; }
        offset = min(max(settings.sweepEpsilon, gap * 0.02), max(settings.sweepEpsilon, gap * 0.25));
        if (gap > 0) { offset = min(offset, gap * 0.25); }
        return max(offset, 0.0001);
    }

    function groupCandidatesByY(candidates, epsilon) {
        var sorted = candidates.slice(0);
        var groups = [];
        var i, group;
        sorted.sort(function (a, b) { return a.y - b.y; });
        for (i = 0; i < sorted.length; i += 1) {
            if (!groups.length || sorted[i].y - groups[groups.length - 1].y > epsilon) {
                groups.push({ y: sorted[i].y, candidates: [sorted[i]] });
            } else {
                group = groups[groups.length - 1];
                group.candidates.push(sorted[i]);
                group.y = (group.y * (group.candidates.length - 1) + sorted[i].y) / group.candidates.length;
            }
        }
        return groups;
    }

    function chooseRepresentative(candidates) {
        var i;
        for (i = 0; i < candidates.length; i += 1) {
            if (candidates[i].kind === "flat") { return candidates[i]; }
        }
        return candidates[0];
    }

    function classifyBasinBirths(unit, candidates, settings, debugData) {
        var accepted = [];
        var groups = groupCandidatesByY(candidates, settings.horizontalEpsilon);
        var eventYs = uniqueEventYs(unit, settings.horizontalEpsilon);
        var g, y, offset, below, above, connected, ai, bi, ci, touched, reps, rep;
        var candidate, sameBirth;

        for (g = 0; g < groups.length; g += 1) {
            y = groups[g].y;
            offset = sweepOffsetForY(eventYs, y, settings);
            below = intervalsAtY(unit, y - offset, settings.sweepEpsilon);
            above = intervalsAtY(unit, y + offset, settings.sweepEpsilon);
            connected = [];
            for (ai = 0; ai < above.length; ai += 1) {
                connected[ai] = false;
                for (bi = 0; bi < below.length; bi += 1) {
                    if (intervalOverlaps(above[ai], below[bi], settings.coincidentEpsilon)) {
                        connected[ai] = true;
                        break;
                    }
                }
            }
            reps = [];
            for (ci = 0; ci < groups[g].candidates.length; ci += 1) {
                candidate = groups[g].candidates[ci];
                touched = [];
                for (ai = 0; ai < above.length; ai += 1) {
                    if (candidateTouchesInterval(candidate, above[ai], settings.coincidentEpsilon)) {
                        touched.push(ai);
                    }
                }
                if (!touched.length) {
                    candidate.reason = "no-filled-region-above";
                    continue;
                }
                sameBirth = -1;
                for (ai = 0; ai < touched.length; ai += 1) {
                    if (!connected[touched[ai]]) { sameBirth = touched[ai]; break; }
                }
                if (sameBirth < 0) {
                    candidate.reason = "connected-to-lower-water";
                    continue;
                }
                if (!reps[sameBirth]) { reps[sameBirth] = []; }
                reps[sameBirth].push(candidate);
            }
            for (ai = 0; ai < reps.length; ai += 1) {
                if (!reps[ai] || !reps[ai].length) { continue; }
                rep = chooseRepresentative(reps[ai]);
                rep.accepted = true;
                rep.reason = "new-basin";
                accepted.push(rep);
                for (ci = 0; ci < reps[ai].length; ci += 1) {
                    if (reps[ai][ci] !== rep) { reps[ai][ci].reason = "same-basin"; }
                }
            }
            if (debugData) {
                debugData.sweeps.push({ y: y, offset: offset, below: below, above: above });
            }
        }
        return accepted;
    }

    function mergeCoincident(drains, epsilon) {
        var merged = [];
        var i, j, found, existing;
        for (i = 0; i < drains.length; i += 1) {
            found = false;
            for (j = 0; j < merged.length; j += 1) {
                existing = merged[j];
                if (dist2(existing, drains[i]) <= epsilon * epsilon) {
                    if (drains[i].kind === "flat" && existing.kind !== "flat") {
                        merged[j] = drains[i];
                    }
                    found = true;
                    break;
                }
            }
            if (!found) { merged.push(drains[i]); }
        }
        return merged;
    }

    function analyzeUnit(unit, settings, logger, debugData) {
        var candidates;
        var accepted;
        unit.bounds = boundsOfContours(unit.contours);
        candidates = collectCandidates(unit, settings, logger);
        accepted = classifyBasinBirths(unit, candidates, settings, debugData);
        accepted = mergeCoincident(accepted, settings.coincidentEpsilon);
        logger.stats.rejectedCandidates += candidates.length - accepted.length;
        if (debugData) {
            var i;
            for (i = 0; i < candidates.length; i += 1) { debugData.candidates.push(candidates[i]); }
        }
        return accepted;
    }

    function getOrCreateOutputLayer(document, settings) {
        var layer;
        if (!settings.CREATE_NEW_OUTPUT_LAYER) {
            try { return document.layers.getByName(settings.OUTPUT_LAYER_NAME); }
            catch (error0) {}
        }
        layer = document.layers.add();
        try { layer.name = settings.OUTPUT_LAYER_NAME; }
        catch (error) { layer.name = settings.OUTPUT_LAYER_NAME + " (new)"; }
        return layer;
    }

    function rgb(red, green, blue) {
        var color = new RGBColor();
        color.red = red; color.green = green; color.blue = blue;
        return color;
    }

    function drawCircle(layer, center, diameter, color) {
        var radius = diameter * 0.5;
        var circle = layer.pathItems.ellipse(center.y + radius, center.x - radius, diameter, diameter);
        circle.stroked = false;
        circle.filled = true;
        circle.fillColor = color;
        return circle;
    }

    function drawLine(layer, x1, y1, x2, y2, color, width) {
        var path = layer.pathItems.add();
        path.setEntirePath([[x1, y1], [x2, y2]]);
        path.stroked = true; path.strokeColor = color; path.strokeWidth = width;
        path.filled = false;
        return path;
    }

    function drawDebug(document, debugData, settings) {
        var layer = document.layers.add();
        var i, j, candidate, sweep, interval, label;
        layer.name = settings.DEBUG_LAYER_NAME;
        for (i = 0; i < debugData.sweeps.length; i += 1) {
            sweep = debugData.sweeps[i];
            for (j = 0; j < sweep.below.length; j += 1) {
                interval = sweep.below[j];
                drawLine(layer, interval.left, sweep.y - sweep.offset, interval.right,
                    sweep.y - sweep.offset, rgb(0, 170, 255), 0.5);
            }
            for (j = 0; j < sweep.above.length; j += 1) {
                interval = sweep.above[j];
                drawLine(layer, interval.left, sweep.y + sweep.offset, interval.right,
                    sweep.y + sweep.offset, rgb(255, 0, 200), 0.5);
            }
        }
        for (i = 0; i < debugData.candidates.length; i += 1) {
            candidate = debugData.candidates[i];
            drawCircle(layer, candidate, mm(1.2), candidate.accepted ? rgb(0, 200, 0) : rgb(120, 120, 120));
            if (settings.DEBUG_LABELS) {
                label = layer.textFrames.add();
                label.contents = candidate.reason;
                label.position = [candidate.x + mm(1), candidate.y + mm(1)];
                label.textRange.characterAttributes.size = 6;
            }
        }
        return layer;
    }

    function drawDrains(document, drains, settings) {
        var layer = getOrCreateOutputLayer(document, settings);
        var red = rgb(255, 0, 0);
        var i, circle;
        for (i = 0; i < drains.length; i += 1) {
            circle = drawCircle(layer, drains[i], settings.drainDiameter, red);
            circle.name = "Drain " + (i + 1) + " [" + drains[i].kind + "]";
        }
        return layer;
    }

    function summaryMessage(logger, elapsedMs, outputLayerName) {
        var s = logger.stats;
        var message = "Drain Points complete.\n\n" +
            "Drains: " + s.drains + "\n" +
            "Shape units: " + s.units + "\n" +
            "Contours: " + s.contours + "\n" +
            "Flattened vertices: " + s.flattenedVertices + "\n" +
            "Candidates rejected by connectivity: " + s.rejectedCandidates + "\n" +
            "Output layer: " + outputLayerName + "\n" +
            "Time: " + (elapsedMs / 1000).toFixed(2) + " s";
        if (logger.warnings.length) {
            message += "\n\nWarnings (" + logger.warnings.length + "):\n- " + logger.warnings.slice(0, 8).join("\n- ");
            if (logger.warnings.length > 8) { message += "\n- ..."; }
        }
        return message;
    }

    function run(overrides) {
        if (typeof app === "undefined" || !app.documents || app.documents.length === 0) {
            alert("Open an Illustrator document and select one or more vector shapes.");
            return;
        }
        var document = app.activeDocument;
        var originalSelection = [];
        var settings = makeSettings(overrides);
        var logger = new Logger();
        var tempLayer = null;
        var outputLayer = null;
        var debugData = settings.DEBUG ? { candidates: [], sweeps: [] } : null;
        var units = [];
        var drains = [];
        var i, j, unitDrains, started = new Date().getTime();

        for (i = 0; i < document.selection.length; i += 1) { originalSelection.push(document.selection[i]); }
        if (!originalSelection.length) {
            alert("Select at least one closed vector shape or text object.");
            return;
        }

        try {
            tempLayer = normalizeSelection(document, originalSelection, settings, logger);
            collectUnitsFromItem(tempLayer, units, settings, logger);
            logger.stats.units = units.length;
            for (i = 0; i < units.length; i += 1) { logger.stats.contours += units[i].contours.length; }
            if (!units.length) { throw new Error("No closed filled paths were found after normalization."); }

            for (i = 0; i < units.length; i += 1) {
                unitDrains = analyzeUnit(units[i], settings, logger, debugData);
                for (j = 0; j < unitDrains.length; j += 1) { drains.push(unitDrains[j]); }
            }
            drains = mergeCoincident(drains, settings.coincidentEpsilon);
            logger.stats.drains = drains.length;

            tempLayer.remove();
            tempLayer = null;
            restoreSelection(document, originalSelection);
            outputLayer = drawDrains(document, drains, settings);
            if (settings.DEBUG) { drawDebug(document, debugData, settings); }
            app.redraw();
            if (settings.SHOW_SUMMARY) {
                alert(summaryMessage(logger, new Date().getTime() - started, outputLayer.name));
            }
        } catch (error) {
            try { if (tempLayer) { tempLayer.remove(); } } catch (cleanupError) {}
            restoreSelection(document, originalSelection);
            alert("Drain Points failed:\n" + error.message + (error.line ? "\nLine: " + error.line : ""));
        }
    }

    return {
        CONFIG: CONFIG,
        run: run
    };
}());

DrainPoints.run();
