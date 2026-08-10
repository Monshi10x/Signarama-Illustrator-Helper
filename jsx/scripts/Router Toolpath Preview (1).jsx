#target illustrator

/*
Router Toolpath Preview.jsx

Visualises a CNC router profile cut from selected closed Illustrator paths.
The script uses Illustrator's native Offset Path effect with round joins and a
180 miter limit. Outside cuts use +radius then -radius; inside cuts use -radius
then +radius. The first pass is the tool centreline and the second pass is the
resulting machined edge.

The selected artwork is never modified. Preview artwork is created on a new
layer and is removed if the dialog is cancelled.

Compatible with Adobe Illustrator ExtendScript (ES3 syntax).
*/

(function () {
    var SCRIPT_NAME = "Router Toolpath Preview";
    var VERSION = "1.1.0";
    var EPS = 0.000001;
    var CORNER_ANGLE_RADIANS = Math.PI / 180.0;
    var MAX_FLATTEN_DEPTH = 16;
    var MAX_POINTS_PER_PATH = 12000;

    function pt(x, y) {
        return { x: x, y: y };
    }

    function copyPoint(p) {
        return { x: p.x, y: p.y };
    }

    function add(a, b) {
        return { x: a.x + b.x, y: a.y + b.y };
    }

    function sub(a, b) {
        return { x: a.x - b.x, y: a.y - b.y };
    }

    function mul(a, s) {
        return { x: a.x * s, y: a.y * s };
    }

    function dot(a, b) {
        return a.x * b.x + a.y * b.y;
    }

    function cross(a, b) {
        return a.x * b.y - a.y * b.x;
    }

    function lengthOf(a) {
        return Math.sqrt(a.x * a.x + a.y * a.y);
    }

    function distance(a, b) {
        return lengthOf(sub(a, b));
    }

    function normalize(a) {
        var len = lengthOf(a);
        if (len < EPS) {
            return null;
        }
        return { x: a.x / len, y: a.y / len };
    }

    function leftNormal(unitDirection) {
        return { x: -unitDirection.y, y: unitDirection.x };
    }

    function nearlyEqual(a, b, tolerance) {
        return distance(a, b) <= (tolerance || EPS);
    }

    function clamp(value, minimum, maximum) {
        return Math.max(minimum, Math.min(maximum, value));
    }

    function finiteNumber(value) {
        return typeof value === "number" && isFinite(value);
    }

    function trimString(value) {
        return String(value).replace(/^\s+|\s+$/g, "");
    }

    function parsePositiveNumber(value) {
        var number = Number(trimString(value));
        if (!finiteNumber(number) || number <= 0) {
            return null;
        }
        return number;
    }

    function roundForField(value) {
        var rounded = Math.round(value * 10000) / 10000;
        return String(rounded);
    }

    function unitToPoints(value, unitName) {
        if (unitName === "mm") {
            return value * 72.0 / 25.4;
        }
        if (unitName === "in") {
            return value * 72.0;
        }
        return value;
    }

    function pointsToUnit(value, unitName) {
        if (unitName === "mm") {
            return value * 25.4 / 72.0;
        }
        if (unitName === "in") {
            return value / 72.0;
        }
        return value;
    }

    function defaultUnitForDocument(documentRef) {
        try {
            if (documentRef.rulerUnits === RulerUnits.Inches) {
                return "in";
            }
            if (documentRef.rulerUnits === RulerUnits.Points ||
                    documentRef.rulerUnits === RulerUnits.Picas) {
                return "pt";
            }
        } catch (ignore) {
        }
        return "mm";
    }

    function defaultDiameterForUnit(unitName) {
        if (unitName === "in") {
            return 0.25;
        }
        if (unitName === "pt") {
            return unitToPoints(6.0, "mm");
        }
        return 6.0;
    }

    function defaultToleranceForUnit(unitName) {
        if (unitName === "in") {
            return 0.004;
        }
        if (unitName === "pt") {
            return unitToPoints(0.1, "mm");
        }
        return 0.1;
    }

    function rgb(red, green, blue) {
        var color = new RGBColor();
        color.red = red;
        color.green = green;
        color.blue = blue;
        return color;
    }

    function arrayPoint(value) {
        return { x: Number(value[0]), y: Number(value[1]) };
    }

    function snapshotPath(pathItem, label, compoundKey) {
        var snapshot = {
            label: label,
            compoundKey: compoundKey,
            closed: Boolean(pathItem.closed),
            points: []
        };
        var index;
        for (index = 0; index < pathItem.pathPoints.length; index += 1) {
            var sourcePoint = pathItem.pathPoints[index];
            snapshot.points.push({
                anchor: arrayPoint(sourcePoint.anchor),
                left: arrayPoint(sourcePoint.leftDirection),
                right: arrayPoint(sourcePoint.rightDirection),
                pointType: sourcePoint.pointType
            });
        }
        return snapshot;
    }

    function collectSelectedPaths(documentRef) {
        var result = {
            paths: [],
            originalSelection: [],
            skippedOpen: 0,
            skippedUnsupported: 0,
            skippedGuides: 0
        };
        var seen = [];
        var compoundCounter = 0;
        var pathCounter = 0;

        function alreadySeen(item) {
            var i;
            for (i = 0; i < seen.length; i += 1) {
                if (seen[i] === item) {
                    return true;
                }
            }
            seen.push(item);
            return false;
        }

        function addPath(item, compoundKey) {
            if (alreadySeen(item)) {
                return;
            }
            try {
                if (item.guides) {
                    result.skippedGuides += 1;
                    return;
                }
                if (!item.closed) {
                    result.skippedOpen += 1;
                    return;
                }
                if (item.pathPoints.length < 3) {
                    result.skippedUnsupported += 1;
                    return;
                }
                pathCounter += 1;
                result.paths.push(snapshotPath(
                    item,
                    item.name ? item.name : "Path " + pathCounter,
                    compoundKey
                ));
            } catch (error) {
                result.skippedUnsupported += 1;
            }
        }

        function collectClippingMembers(groupItem) {
            var clippingMembers = [];
            var i;
            for (i = 0; i < groupItem.pageItems.length; i += 1) {
                var child = groupItem.pageItems[i];
                if (child.typename === "PathItem") {
                    try {
                        if (child.clipping) {
                            clippingMembers.push(child);
                        }
                    } catch (ignorePath) {
                    }
                } else if (child.typename === "CompoundPathItem") {
                    var j;
                    for (j = 0; j < child.pathItems.length; j += 1) {
                        try {
                            if (child.pathItems[j].clipping) {
                                clippingMembers.push(child);
                                break;
                            }
                        } catch (ignoreCompound) {
                        }
                    }
                }
            }
            return clippingMembers;
        }

        function visit(item, inheritedCompoundKey) {
            if (!item || !item.typename) {
                result.skippedUnsupported += 1;
                return;
            }
            if (item.typename === "PathItem") {
                addPath(item, inheritedCompoundKey);
                return;
            }
            if (item.typename === "CompoundPathItem") {
                compoundCounter += 1;
                var compoundKey = "compound-" + compoundCounter;
                var compoundIndex;
                for (compoundIndex = 0;
                        compoundIndex < item.pathItems.length;
                        compoundIndex += 1) {
                    addPath(item.pathItems[compoundIndex], compoundKey);
                }
                return;
            }
            if (item.typename === "GroupItem") {
                var clippingMembers = [];
                try {
                    if (item.clipped) {
                        clippingMembers = collectClippingMembers(item);
                    }
                } catch (ignoreClipped) {
                }
                var groupIndex;
                if (clippingMembers.length > 0) {
                    for (groupIndex = 0;
                            groupIndex < clippingMembers.length;
                            groupIndex += 1) {
                        visit(clippingMembers[groupIndex], null);
                    }
                } else {
                    for (groupIndex = 0;
                            groupIndex < item.pageItems.length;
                            groupIndex += 1) {
                        visit(item.pageItems[groupIndex], null);
                    }
                }
                return;
            }
            result.skippedUnsupported += 1;
        }

        var selection = documentRef.selection;
        var selectionIndex;
        for (selectionIndex = 0;
                selection && selectionIndex < selection.length;
                selectionIndex += 1) {
            result.originalSelection.push(selection[selectionIndex]);
            visit(selection[selectionIndex], null);
        }
        return result;
    }

    function pointLineDistance(pointValue, lineStart, lineEnd) {
        var chord = sub(lineEnd, lineStart);
        var chordLength = lengthOf(chord);
        if (chordLength < EPS) {
            return distance(pointValue, lineStart);
        }
        return Math.abs(cross(chord, sub(pointValue, lineStart))) / chordLength;
    }

    function midpoint(a, b) {
        return { x: (a.x + b.x) * 0.5, y: (a.y + b.y) * 0.5 };
    }

    function flattenCubic(p0, p1, p2, p3, tolerance, depth, output, state) {
        if (state.count >= MAX_POINTS_PER_PATH) {
            state.capped = true;
            return;
        }
        var flatness = Math.max(
            pointLineDistance(p1, p0, p3),
            pointLineDistance(p2, p0, p3)
        );
        if (depth >= MAX_FLATTEN_DEPTH || flatness <= tolerance) {
            output.push(copyPoint(p3));
            state.count += 1;
            return;
        }

        var p01 = midpoint(p0, p1);
        var p12 = midpoint(p1, p2);
        var p23 = midpoint(p2, p3);
        var p012 = midpoint(p01, p12);
        var p123 = midpoint(p12, p23);
        var p0123 = midpoint(p012, p123);

        flattenCubic(p0, p01, p012, p0123,
            tolerance, depth + 1, output, state);
        flattenCubic(p0123, p123, p23, p3,
            tolerance, depth + 1, output, state);
    }

    function removeConsecutiveDuplicates(vertices, tolerance) {
        var cleaned = [];
        var index;
        for (index = 0; index < vertices.length; index += 1) {
            var current = vertices[index];
            if (cleaned.length === 0 ||
                    !nearlyEqual(cleaned[cleaned.length - 1], current, tolerance)) {
                cleaned.push(current);
            } else if (current.original) {
                cleaned[cleaned.length - 1].original = true;
            }
        }
        if (cleaned.length > 1 &&
                nearlyEqual(cleaned[0], cleaned[cleaned.length - 1], tolerance)) {
            if (cleaned[cleaned.length - 1].original) {
                cleaned[0].original = true;
            }
            cleaned.pop();
        }
        return cleaned;
    }

    function markGeometricCorners(vertices) {
        var count = vertices.length;
        var index;
        for (index = 0; index < count; index += 1) {
            var current = vertices[index];
            current.corner = false;
            if (!current.original) {
                continue;
            }
            var previous = vertices[(index - 1 + count) % count];
            var next = vertices[(index + 1) % count];
            var incoming = normalize(sub(current, previous));
            var outgoing = normalize(sub(next, current));
            if (!incoming || !outgoing) {
                continue;
            }
            var angle = Math.acos(clamp(dot(incoming, outgoing), -1, 1));
            current.corner = angle >= CORNER_ANGLE_RADIANS;
        }
    }

    function flattenSnapshot(snapshot, tolerance) {
        var source = snapshot.points;
        var vertices = [];
        var state = { count: 0, capped: false };
        if (source.length < 3) {
            return { vertices: [], capped: false };
        }
        vertices.push({
            x: source[0].anchor.x,
            y: source[0].anchor.y,
            original: true
        });

        var segmentCount = snapshot.closed ? source.length : source.length - 1;
        var segmentIndex;
        for (segmentIndex = 0; segmentIndex < segmentCount; segmentIndex += 1) {
            var nextIndex = (segmentIndex + 1) % source.length;
            var segmentPoints = [];
            flattenCubic(
                source[segmentIndex].anchor,
                source[segmentIndex].right,
                source[nextIndex].left,
                source[nextIndex].anchor,
                tolerance,
                0,
                segmentPoints,
                state
            );
            var pointIndex;
            for (pointIndex = 0; pointIndex < segmentPoints.length; pointIndex += 1) {
                if (segmentIndex === segmentCount - 1 &&
                        pointIndex === segmentPoints.length - 1 &&
                        snapshot.closed) {
                    continue;
                }
                vertices.push({
                    x: segmentPoints[pointIndex].x,
                    y: segmentPoints[pointIndex].y,
                    original: pointIndex === segmentPoints.length - 1
                });
            }
        }
        vertices = removeConsecutiveDuplicates(vertices,
            Math.max(EPS, tolerance * 0.001));
        if (vertices.length >= 3) {
            markGeometricCorners(vertices);
        }
        return { vertices: vertices, capped: state.capped };
    }

    function signedArea(vertices) {
        var total = 0;
        var index;
        for (index = 0; index < vertices.length; index += 1) {
            var next = vertices[(index + 1) % vertices.length];
            total += vertices[index].x * next.y - next.x * vertices[index].y;
        }
        return total * 0.5;
    }

    function pointInPolygon(pointValue, vertices) {
        var inside = false;
        var i;
        var j = vertices.length - 1;
        for (i = 0; i < vertices.length; i += 1) {
            var pi = vertices[i];
            var pj = vertices[j];
            var crossesRay = ((pi.y > pointValue.y) !== (pj.y > pointValue.y));
            if (crossesRay) {
                var intersectX = (pj.x - pi.x) *
                    (pointValue.y - pi.y) / (pj.y - pi.y) + pi.x;
                if (pointValue.x < intersectX) {
                    inside = !inside;
                }
            }
            j = i;
        }
        return inside;
    }

    function assignCompoundDepths(flattenedRecords) {
        var index;
        for (index = 0; index < flattenedRecords.length; index += 1) {
            flattenedRecords[index].depth = 0;
            if (!flattenedRecords[index].snapshot.compoundKey ||
                    flattenedRecords[index].vertices.length < 3) {
                continue;
            }
            var testPoint = flattenedRecords[index].vertices[0];
            var otherIndex;
            for (otherIndex = 0;
                    otherIndex < flattenedRecords.length;
                    otherIndex += 1) {
                if (index === otherIndex ||
                        flattenedRecords[otherIndex].snapshot.compoundKey !==
                            flattenedRecords[index].snapshot.compoundKey ||
                        flattenedRecords[otherIndex].vertices.length < 3) {
                    continue;
                }
                if (pointInPolygon(testPoint,
                        flattenedRecords[otherIndex].vertices)) {
                    flattenedRecords[index].depth += 1;
                }
            }
        }
    }

    function lineIntersection(linePointA, directionA, linePointB, directionB) {
        var denominator = cross(directionA, directionB);
        if (Math.abs(denominator) < EPS) {
            return null;
        }
        var parameter = cross(sub(linePointB, linePointA), directionB) /
            denominator;
        return add(linePointA, mul(directionA, parameter));
    }

    function commandTarget(command) {
        if (command.type === "C") {
            return command.p;
        }
        return command.p;
    }

    function appendLine(commands, pointValue) {
        var current = commandTarget(commands[commands.length - 1]);
        if (!nearlyEqual(current, pointValue, EPS)) {
            commands.push({ type: "L", p: copyPoint(pointValue) });
        }
    }

    function normaliseArcDelta(startAngle, endAngle, direction) {
        var delta = endAngle - startAngle;
        if (direction > 0) {
            while (delta <= EPS) {
                delta += Math.PI * 2.0;
            }
        } else {
            while (delta >= -EPS) {
                delta -= Math.PI * 2.0;
            }
        }
        return delta;
    }

    function appendArc(commands, center, startPoint, endPoint, direction) {
        appendLine(commands, startPoint);
        var radius = distance(center, startPoint);
        if (radius < EPS || nearlyEqual(startPoint, endPoint, EPS)) {
            appendLine(commands, endPoint);
            return;
        }
        var startAngle = Math.atan2(startPoint.y - center.y,
            startPoint.x - center.x);
        var endAngle = Math.atan2(endPoint.y - center.y,
            endPoint.x - center.x);
        var delta = normaliseArcDelta(startAngle, endAngle, direction);
        var pieceCount = Math.max(1,
            Math.ceil(Math.abs(delta) / (Math.PI * 0.5)));
        var pieceDelta = delta / pieceCount;
        var pieceIndex;
        for (pieceIndex = 0; pieceIndex < pieceCount; pieceIndex += 1) {
            var angle0 = startAngle + pieceDelta * pieceIndex;
            var angle1 = angle0 + pieceDelta;
            var p0 = pt(
                center.x + radius * Math.cos(angle0),
                center.y + radius * Math.sin(angle0)
            );
            var p1 = pt(
                center.x + radius * Math.cos(angle1),
                center.y + radius * Math.sin(angle1)
            );
            var kappa = 4.0 / 3.0 * Math.tan(pieceDelta * 0.25);
            var control1 = pt(
                p0.x + kappa * radius * -Math.sin(angle0),
                p0.y + kappa * radius * Math.cos(angle0)
            );
            var control2 = pt(
                p1.x - kappa * radius * -Math.sin(angle1),
                p1.y - kappa * radius * Math.cos(angle1)
            );
            commands.push({
                type: "C",
                c1: control1,
                c2: control2,
                p: p1
            });
        }
    }

    function offsetClosedPolyline(vertices, signedOffset) {
        var result = {
            commands: [],
            limitedCenters: [],
            tightCount: 0,
            invalid: false
        };
        var count = vertices.length;
        if (count < 3) {
            result.invalid = true;
            return result;
        }
        if (Math.abs(signedOffset) < EPS) {
            result.commands.push({ type: "M", p: copyPoint(vertices[0]) });
            var directIndex;
            for (directIndex = 1; directIndex < count; directIndex += 1) {
                appendLine(result.commands, vertices[directIndex]);
            }
            appendLine(result.commands, vertices[0]);
            return result;
        }

        var edges = [];
        var edgeIndex;
        for (edgeIndex = 0; edgeIndex < count; edgeIndex += 1) {
            var nextIndex = (edgeIndex + 1) % count;
            var vector = sub(vertices[nextIndex], vertices[edgeIndex]);
            var edgeLength = lengthOf(vector);
            var unit = normalize(vector);
            if (!unit) {
                result.invalid = true;
                return result;
            }
            edges.push({
                unit: unit,
                normal: leftNormal(unit),
                length: edgeLength
            });
        }

        var joins = [];
        var vertexIndex;
        for (vertexIndex = 0; vertexIndex < count; vertexIndex += 1) {
            var previousEdge = edges[(vertexIndex - 1 + count) % count];
            var nextEdge = edges[vertexIndex];
            var vertex = vertices[vertexIndex];
            var turn = cross(previousEdge.unit, nextEdge.unit);
            var directionDot = dot(previousEdge.unit, nextEdge.unit);
            var previousOffsetPoint = add(vertex,
                mul(previousEdge.normal, signedOffset));
            var nextOffsetPoint = add(vertex,
                mul(nextEdge.normal, signedOffset));

            if (Math.abs(turn) < EPS && directionDot < -0.999999) {
                result.invalid = true;
                return result;
            }

            var expandingSharpCorner = vertex.corner &&
                turn * signedOffset < -EPS;
            if (expandingSharpCorner) {
                joins.push({
                    arc: true,
                    start: previousOffsetPoint,
                    end: nextOffsetPoint,
                    center: copyPoint(vertex),
                    direction: turn > 0 ? 1 : -1
                });
                continue;
            }

            var intersection;
            if (Math.abs(turn) < EPS && directionDot > 0) {
                intersection = nextOffsetPoint;
            } else {
                intersection = lineIntersection(
                    previousOffsetPoint,
                    previousEdge.unit,
                    nextOffsetPoint,
                    nextEdge.unit
                );
            }
            if (!intersection || !finiteNumber(intersection.x) ||
                    !finiteNumber(intersection.y)) {
                result.invalid = true;
                return result;
            }

            var limited = vertex.corner && turn * signedOffset > EPS;
            var tight = false;
            if (limited) {
                var tangentDistancePrevious = Math.abs(dot(
                    sub(intersection, vertex), previousEdge.unit));
                var tangentDistanceNext = Math.abs(dot(
                    sub(intersection, vertex), nextEdge.unit));
                tight = tangentDistancePrevious > previousEdge.length + EPS ||
                    tangentDistanceNext > nextEdge.length + EPS;
                if (tight) {
                    result.tightCount += 1;
                }
                result.limitedCenters.push({
                    point: copyPoint(intersection),
                    tight: tight
                });
            }
            joins.push({
                arc: false,
                start: copyPoint(intersection),
                end: copyPoint(intersection)
            });
        }

        result.commands.push({ type: "M", p: copyPoint(joins[0].end) });
        for (vertexIndex = 1; vertexIndex < count; vertexIndex += 1) {
            appendLine(result.commands, joins[vertexIndex].start);
            if (joins[vertexIndex].arc) {
                appendArc(
                    result.commands,
                    joins[vertexIndex].center,
                    joins[vertexIndex].start,
                    joins[vertexIndex].end,
                    joins[vertexIndex].direction
                );
            }
        }
        appendLine(result.commands, joins[0].start);
        if (joins[0].arc) {
            appendArc(
                result.commands,
                joins[0].center,
                joins[0].start,
                joins[0].end,
                joins[0].direction
            );
        } else {
            appendLine(result.commands, joins[0].end);
        }
        return result;
    }

    function createPathFromCommands(container, commands, closed) {
        if (!commands || commands.length < 2 || commands[0].type !== "M") {
            return null;
        }
        var firstPoint = copyPoint(commands[0].p);
        var nodes = [{
            anchor: firstPoint,
            left: copyPoint(firstPoint),
            right: copyPoint(firstPoint)
        }];
        var commandIndex;
        for (commandIndex = 1;
                commandIndex < commands.length;
                commandIndex += 1) {
            var command = commands[commandIndex];
            var target = commandTarget(command);
            var isClosingCommand = closed &&
                commandIndex === commands.length - 1 &&
                nearlyEqual(target, firstPoint, 0.0001);
            var previousNode = nodes[nodes.length - 1];
            if (command.type === "C") {
                previousNode.right = copyPoint(command.c1);
            } else {
                previousNode.right = copyPoint(previousNode.anchor);
            }
            if (isClosingCommand) {
                nodes[0].left = command.type === "C" ?
                    copyPoint(command.c2) : copyPoint(nodes[0].anchor);
            } else if (!nearlyEqual(previousNode.anchor, target, EPS)) {
                nodes.push({
                    anchor: copyPoint(target),
                    left: command.type === "C" ?
                        copyPoint(command.c2) : copyPoint(target),
                    right: copyPoint(target)
                });
            }
        }
        if (nodes.length < 2) {
            return null;
        }

        var path = container.pathItems.add();
        var anchors = [];
        var nodeIndex;
        for (nodeIndex = 0; nodeIndex < nodes.length; nodeIndex += 1) {
            anchors.push([nodes[nodeIndex].anchor.x, nodes[nodeIndex].anchor.y]);
        }
        path.setEntirePath(anchors);
        path.closed = closed;
        for (nodeIndex = 0; nodeIndex < nodes.length; nodeIndex += 1) {
            path.pathPoints[nodeIndex].leftDirection = [
                nodes[nodeIndex].left.x,
                nodes[nodeIndex].left.y
            ];
            path.pathPoints[nodeIndex].rightDirection = [
                nodes[nodeIndex].right.x,
                nodes[nodeIndex].right.y
            ];
            path.pathPoints[nodeIndex].pointType = PointType.CORNER;
        }
        return path;
    }

    function createExactDesiredPath(container, snapshot) {
        var path = container.pathItems.add();
        var anchors = [];
        var index;
        for (index = 0; index < snapshot.points.length; index += 1) {
            anchors.push([
                snapshot.points[index].anchor.x,
                snapshot.points[index].anchor.y
            ]);
        }
        path.setEntirePath(anchors);
        path.closed = snapshot.closed;
        for (index = 0; index < snapshot.points.length; index += 1) {
            path.pathPoints[index].leftDirection = [
                snapshot.points[index].left.x,
                snapshot.points[index].left.y
            ];
            path.pathPoints[index].rightDirection = [
                snapshot.points[index].right.x,
                snapshot.points[index].right.y
            ];
            try {
                path.pathPoints[index].pointType =
                    snapshot.points[index].pointType;
            } catch (ignoreType) {
                path.pathPoints[index].pointType = PointType.CORNER;
            }
        }
        return path;
    }

    function forEachPathInArtwork(item, callback) {
        if (!item || !item.typename) {
            return;
        }
        if (item.typename === "PathItem") {
            callback(item);
            return;
        }
        if (item.typename === "CompoundPathItem") {
            var compoundIndex;
            for (compoundIndex = 0;
                    compoundIndex < item.pathItems.length;
                    compoundIndex += 1) {
                callback(item.pathItems[compoundIndex]);
            }
            return;
        }
        if (item.typename === "GroupItem") {
            var children = [];
            var childIndex;
            for (childIndex = 0;
                    childIndex < item.pageItems.length;
                    childIndex += 1) {
                children.push(item.pageItems[childIndex]);
            }
            for (childIndex = 0;
                    childIndex < children.length;
                    childIndex += 1) {
                forEachPathInArtwork(children[childIndex], callback);
            }
        }
    }

    function prepareNativeOffsetArtwork(item) {
        forEachPathInArtwork(item, function (path) {
            path.filled = true;
            path.fillColor = rgb(0, 0, 0);
            path.stroked = false;
            path.opacity = 100;
        });
        item.opacity = 100;
    }

    function nativeOffsetSequence(depth, cutMode, radius) {
        if (cutMode === 2) {
            return { first: 0, second: 0 };
        }
        var compoundParity = depth % 2 === 0 ? 1 : -1;
        var outsideFirst = radius * compoundParity;
        var first = cutMode === 0 ? outsideFirst : -outsideFirst;
        return { first: first, second: -first };
    }

    function nativeOffsetEffectXML(offsetPoints) {
        var cleanOffset = Math.round(offsetPoints * 100000000) / 100000000;
        return '<LiveEffect name="Adobe Offset Path"><Dict data="' +
            'R mlim 180 R ofst ' + String(cleanOffset) +
            ' I jntp 0 "/></LiveEffect>';
    }

    function applyNativeOffsetEffect(item, offsetPoints) {
        prepareNativeOffsetArtwork(item);
        var effectXML = nativeOffsetEffectXML(offsetPoints);
        try {
            item.applyEffect(effectXML);
            return item;
        } catch (primaryError) {
            /* Some Illustrator releases expose applyEffect only on leaf paths
               after an appearance expands into a group/compound path. */
            var appliedCount = 0;
            if (item.typename === "CompoundPathItem" &&
                    item.pathItems.length > 0) {
                item.pathItems[0].applyEffect(effectXML);
                appliedCount = 1;
            } else if (item.typename === "GroupItem") {
                var childIndex;
                for (childIndex = 0;
                        childIndex < item.pageItems.length;
                        childIndex += 1) {
                    try {
                        item.pageItems[childIndex].applyEffect(effectXML);
                        appliedCount += 1;
                    } catch (ignoreChildEffect) {
                    }
                }
            }
            if (appliedCount === 0) {
                throw primaryError;
            }
        }
        return item;
    }

    function expandNativeAppearance(documentRef, item, container) {
        documentRef.selection = null;
        item.selected = true;
        app.executeMenuCommand("expandStyle");

        var selectedItems = [];
        var selectionIndex;
        for (selectionIndex = 0;
                documentRef.selection &&
                selectionIndex < documentRef.selection.length;
                selectionIndex += 1) {
            selectedItems.push(documentRef.selection[selectionIndex]);
        }
        documentRef.selection = null;

        if (selectedItems.length === 0) {
            throw new Error(
                "Illustrator did not return artwork after expanding Offset Path."
            );
        }
        if (selectedItems.length === 1) {
            var expandedItem = selectedItems[0];
            try {
                if (expandedItem.parent !== container) {
                    expandedItem.move(container, ElementPlacement.PLACEATEND);
                }
            } catch (ignoreMove) {
            }
            return expandedItem;
        }

        var expandedGroup = container.groupItems.add();
        expandedGroup.name = "Expanded native offset";
        for (selectionIndex = selectedItems.length - 1;
                selectionIndex >= 0;
                selectionIndex -= 1) {
            selectedItems[selectionIndex].move(
                expandedGroup,
                ElementPlacement.PLACEATBEGINNING
            );
        }
        return expandedGroup;
    }

    function duplicateArtwork(item, container) {
        return item.duplicate(container, ElementPlacement.PLACEATEND);
    }

    function styleArtwork(item, name, pathStyler) {
        item.name = name;
        item.opacity = 100;
        forEachPathInArtwork(item, pathStyler);
    }

    function styleDesiredPath(path) {
        path.name = "Desired edge";
        path.filled = false;
        path.stroked = true;
        path.strokeColor = rgb(0, 112, 255);
        path.strokeWidth = 0.75;
        path.opacity = 90;
        try {
            path.strokeDashes = [2, 2];
        } catch (ignoreDashes) {
        }
    }

    function styleSweepPath(path, diameterPoints) {
        path.name = "Cutter sweep";
        path.filled = false;
        path.stroked = true;
        path.strokeColor = rgb(230, 35, 65);
        path.strokeWidth = diameterPoints;
        path.opacity = 20;
        try {
            path.strokeCap = StrokeCap.ROUNDENDCAP;
            path.strokeJoin = StrokeJoin.ROUNDENDJOIN;
        } catch (ignoreCaps) {
        }
    }

    function styleResultPath(path) {
        path.name = "Resulting machined edge";
        path.filled = false;
        path.stroked = true;
        path.strokeColor = rgb(0, 155, 90);
        path.strokeWidth = 1.2;
        path.opacity = 100;
        try {
            path.strokeDashes = [];
            path.strokeCap = StrokeCap.ROUNDENDCAP;
            path.strokeJoin = StrokeJoin.ROUNDENDJOIN;
        } catch (ignoreResultStyle) {
        }
    }

    function styleCentrelinePath(path, diameterPoints) {
        path.name = "Router bit centreline";
        path.filled = false;
        path.stroked = true;
        path.strokeColor = rgb(210, 0, 25);
        path.strokeWidth = Math.max(0.6,
            Math.min(1.5, diameterPoints * 0.08));
        path.opacity = 100;
        var dash = Math.max(3, Math.min(12, diameterPoints * 0.75));
        try {
            path.strokeDashes = [dash, dash * 0.65];
            path.strokeCap = StrokeCap.ROUNDENDCAP;
            path.strokeJoin = StrokeJoin.ROUNDENDJOIN;
        } catch (ignoreDashes) {
        }
    }

    function createCutterCircle(container, center, radius, tight) {
        var circle = container.pathItems.ellipse(
            center.y + radius,
            center.x - radius,
            radius * 2,
            radius * 2
        );
        circle.name = tight ?
            "Cutter at overlapping tight corner" :
            "Cutter at limited corner";
        circle.filled = true;
        circle.fillColor = tight ? rgb(255, 145, 0) : rgb(105, 105, 105);
        circle.stroked = true;
        circle.strokeColor = tight ? rgb(180, 80, 0) : rgb(55, 55, 55);
        circle.strokeWidth = 0.6;
        circle.opacity = tight ? 42 : 34;
        return circle;
    }

    function uniqueLayerName(documentRef, baseName) {
        var candidate = baseName;
        var suffix = 2;
        var found;
        while (true) {
            found = null;
            try {
                found = documentRef.layers.getByName(candidate);
            } catch (ignore) {
            }
            if (!found) {
                return candidate;
            }
            candidate = baseName + " " + suffix;
            suffix += 1;
        }
    }

    function restoreSelection(documentRef, selectionItems) {
        try {
            documentRef.selection = null;
        } catch (ignoreClear) {
        }
        var index;
        for (index = 0; index < selectionItems.length; index += 1) {
            try {
                selectionItems[index].selected = true;
            } catch (ignoreSelect) {
            }
        }
    }

    function machiningOffset(vertices, depth, cutMode, radius) {
        if (cutMode === 2) {
            return 0;
        }
        var area = signedArea(vertices);
        if (Math.abs(area) < EPS) {
            return null;
        }
        var orientation = area > 0 ? 1 : -1;
        var compoundParity = depth % 2 === 0 ? 1 : -1;
        var outsideFilledShape = -orientation * radius * compoundParity;
        return cutMode === 0 ? outsideFilledShape : -outsideFilledShape;
    }

    function buildPreview(documentRef, sourceData, settings, outputLayer) {
        var stats = {
            renderedPaths: 0,
            invalidPaths: 0,
            limitedCorners: 0,
            tightCorners: 0,
            cappedPaths: 0,
            nativeErrorMessage: ""
        };
        var flattenedRecords = [];
        var sourceIndex;
        for (sourceIndex = 0;
                sourceIndex < sourceData.paths.length;
                sourceIndex += 1) {
            var flattened = flattenSnapshot(
                sourceData.paths[sourceIndex],
                settings.tolerancePoints
            );
            flattenedRecords.push({
                snapshot: sourceData.paths[sourceIndex],
                vertices: flattened.vertices,
                capped: flattened.capped,
                depth: 0
            });
            if (flattened.capped) {
                stats.cappedPaths += 1;
            }
        }
        assignCompoundDepths(flattenedRecords);

        var root = outputLayer.groupItems.add();
        root.name = "Router toolpath preview";
        var radius = settings.diameterPoints * 0.5;

        for (sourceIndex = 0;
                sourceIndex < flattenedRecords.length;
                sourceIndex += 1) {
            var record = flattenedRecords[sourceIndex];
            if (record.vertices.length < 3) {
                stats.invalidPaths += 1;
                continue;
            }

            /* The custom polyline analysis is retained only for cutter-position
               markers. Native Illustrator Offset Path creates all displayed
               centreline and two-pass result geometry. */
            var route = {
                limitedCenters: [],
                tightCount: 0,
                invalid: true
            };
            var analysisOffset = machiningOffset(
                record.vertices,
                record.depth,
                settings.cutMode,
                radius
            );
            if (analysisOffset !== null && !record.capped) {
                var analysedRoute = offsetClosedPolyline(
                    record.vertices,
                    analysisOffset
                );
                if (!analysedRoute.invalid) {
                    route = analysedRoute;
                }
            }

            var pathGroup = root.groupItems.add();
            pathGroup.name = record.snapshot.label + " - router preview";
            var sweepArtwork = null;
            var desiredPath = null;
            var centrelineArtwork = null;
            var resultArtwork = null;
            var centrelineBase = null;

            try {
                if (settings.showDesired) {
                    desiredPath = createExactDesiredPath(
                        pathGroup,
                        record.snapshot
                    );
                    styleDesiredPath(desiredPath);
                }

                centrelineBase = createExactDesiredPath(
                    pathGroup,
                    record.snapshot
                );

                if (settings.cutMode !== 2) {
                    var sequence = nativeOffsetSequence(
                        record.depth,
                        settings.cutMode,
                        radius
                    );

                    /* First native offset: actual router-bit centreline. */
                    applyNativeOffsetEffect(centrelineBase, sequence.first);

                    /* Duplicate and expand the first pass before applying the
                       reverse offset. Expansion between passes guarantees the
                       requested order instead of relying on Appearance-panel
                       live-effect stacking order. */
                    var resultFirstPassLive = duplicateArtwork(
                        centrelineBase,
                        pathGroup
                    );
                    centrelineBase = expandNativeAppearance(
                        documentRef,
                        centrelineBase,
                        pathGroup
                    );
                    var resultFirstPass = expandNativeAppearance(
                        documentRef,
                        resultFirstPassLive,
                        pathGroup
                    );
                    applyNativeOffsetEffect(
                        resultFirstPass,
                        sequence.second
                    );
                    resultArtwork = expandNativeAppearance(
                        documentRef,
                        resultFirstPass,
                        pathGroup
                    );
                }

                if (settings.showSweep) {
                    sweepArtwork = duplicateArtwork(
                        centrelineBase,
                        pathGroup
                    );
                    styleArtwork(
                        sweepArtwork,
                        "Cutter sweep",
                        function (path) {
                            styleSweepPath(path, settings.diameterPoints);
                        }
                    );
                }
                if (settings.showCentreline) {
                    centrelineArtwork = duplicateArtwork(
                        centrelineBase,
                        pathGroup
                    );
                    styleArtwork(
                        centrelineArtwork,
                        "Router bit centreline",
                        function (path) {
                            styleCentrelinePath(
                                path,
                                settings.diameterPoints
                            );
                        }
                    );
                }
                if (resultArtwork) {
                    if (settings.showResult) {
                        styleArtwork(
                            resultArtwork,
                            "Resulting machined edge",
                            styleResultPath
                        );
                    } else {
                        resultArtwork.remove();
                        resultArtwork = null;
                    }
                }

                centrelineBase.remove();
                centrelineBase = null;

                if (settings.showCutterCircles) {
                    var circleIndex;
                    for (circleIndex = 0;
                            circleIndex < route.limitedCenters.length;
                            circleIndex += 1) {
                        createCutterCircle(
                            pathGroup,
                            route.limitedCenters[circleIndex].point,
                            radius,
                            route.limitedCenters[circleIndex].tight
                        );
                    }
                }

                try {
                    if (sweepArtwork) {
                        sweepArtwork.zOrder(ZOrderMethod.SENDTOBACK);
                    }
                    if (desiredPath) {
                        desiredPath.zOrder(ZOrderMethod.BRINGTOFRONT);
                    }
                    if (resultArtwork) {
                        resultArtwork.zOrder(ZOrderMethod.BRINGTOFRONT);
                    }
                    if (centrelineArtwork) {
                        centrelineArtwork.zOrder(ZOrderMethod.BRINGTOFRONT);
                    }
                } catch (ignoreOrder) {
                }

                stats.renderedPaths += 1;
                stats.limitedCorners += route.limitedCenters.length;
                stats.tightCorners += route.tightCount;
            } catch (nativeError) {
                documentRef.selection = null;
                if (!stats.nativeErrorMessage) {
                    stats.nativeErrorMessage = nativeError.message ?
                        nativeError.message : String(nativeError);
                }
                try {
                    pathGroup.remove();
                } catch (ignoreFailedGroup) {
                }
                stats.invalidPaths += 1;
            }
        }
        restoreSelection(documentRef, sourceData.originalSelection);
        return { root: root, stats: stats };
    }

    function settingsFromControls(controls, unitName) {
        var diameter = parsePositiveNumber(controls.diameter.text);
        var tolerance = parsePositiveNumber(controls.tolerance.text);
        if (diameter === null) {
            throw new Error("Enter a tool diameter greater than zero.");
        }
        if (tolerance === null) {
            throw new Error("Enter a curve tolerance greater than zero.");
        }
        var diameterPoints = unitToPoints(diameter, unitName);
        var tolerancePoints = unitToPoints(tolerance, unitName);
        if (tolerancePoints > diameterPoints * 0.25) {
            throw new Error(
                "Curve tolerance is too large for this cutter. " +
                "Use no more than one quarter of the tool diameter."
            );
        }
        return {
            diameterValue: diameter,
            diameterPoints: diameterPoints,
            toleranceValue: tolerance,
            tolerancePoints: tolerancePoints,
            unitName: unitName,
            cutMode: controls.cutMode.selection.index,
            showDesired: controls.showDesired.value,
            showResult: controls.showResult.value,
            showSweep: controls.showSweep.value,
            showCentreline: controls.showCentreline.value,
            showCutterCircles: controls.showCutterCircles.value
        };
    }

    function cutModeShortName(cutMode) {
        if (cutMode === 0) {
            return "Outside";
        }
        if (cutMode === 1) {
            return "Inside";
        }
        return "On-path";
    }

    function showDialog(documentRef, sourceData) {
        var unitName = defaultUnitForDocument(documentRef);
        var previousUnitName = unitName;
        var outputLayer = null;
        var previewRoot = null;
        var latestStats = null;
        var acceptedSettings = null;
        var busy = false;

        var dialog = new Window("dialog", SCRIPT_NAME + " " + VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.spacing = 10;
        dialog.margins = 16;

        var toolPanel = dialog.add("panel", undefined, "Tool and cut");
        toolPanel.orientation = "column";
        toolPanel.alignChildren = ["fill", "top"];
        toolPanel.margins = 12;

        var diameterRow = toolPanel.add("group");
        diameterRow.add("statictext", undefined, "Tool diameter:");
        var diameterInput = diameterRow.add("edittext", undefined,
            roundForField(defaultDiameterForUnit(unitName)));
        diameterInput.characters = 10;
        var unitDropdown = diameterRow.add("dropdownlist", undefined,
            ["mm", "in", "pt"]);
        unitDropdown.selection = unitName === "in" ? 1 :
            (unitName === "pt" ? 2 : 0);

        var cutRow = toolPanel.add("group");
        cutRow.add("statictext", undefined, "Cut side:");
        var cutModeDropdown = cutRow.add("dropdownlist", undefined, [
            "Outside profile (male part)",
            "Inside profile / pocket",
            "Centre on selected path"
        ]);
        cutModeDropdown.selection = 0;
        cutModeDropdown.preferredSize.width = 220;

        var toleranceRow = toolPanel.add("group");
        toleranceRow.add("statictext", undefined, "Corner-analysis tolerance:");
        var toleranceInput = toleranceRow.add("edittext", undefined,
            roundForField(defaultToleranceForUnit(unitName)));
        toleranceInput.characters = 10;
        var toleranceUnitLabel = toleranceRow.add("statictext", undefined,
            unitName);

        var displayPanel = dialog.add("panel", undefined, "Display");
        displayPanel.orientation = "column";
        displayPanel.alignChildren = ["left", "top"];
        displayPanel.margins = 12;
        var showDesired = displayPanel.add("checkbox", undefined,
            "Show desired edge (blue dashed)");
        showDesired.value = true;
        var showResult = displayPanel.add("checkbox", undefined,
            "Show resulting two-pass cut edge (green)");
        showResult.value = true;
        var showSweep = displayPanel.add("checkbox", undefined,
            "Show cutter sweep at full tool diameter");
        showSweep.value = true;
        var showCentreline = displayPanel.add("checkbox", undefined,
            "Show router-bit centreline (red dashed)");
        showCentreline.value = true;
        var showCutterCircles = displayPanel.add("checkbox", undefined,
            "Show cutter at corners the bit cannot reproduce");
        showCutterCircles.value = true;

        var previewCheckbox = dialog.add("checkbox", undefined,
            "Live preview");
        previewCheckbox.value = true;

        var statusText = dialog.add("statictext", undefined,
            sourceData.paths.length + " closed path(s) ready.",
            { multiline: true });
        statusText.preferredSize = [390, 42];

        var information = dialog.add("statictext", undefined,
            "Green is the native two-pass result. Grey circles mark normal " +
            "cutter-limited corners; orange marks overlap with a nearby edge.",
            { multiline: true });
        information.preferredSize = [390, 48];

        var buttonRow = dialog.add("group");
        buttonRow.alignment = "right";
        var cancelButton = buttonRow.add("button", undefined, "Cancel",
            { name: "cancel" });
        var okButton = buttonRow.add("button", undefined, "Create Toolpath",
            { name: "ok" });

        var controls = {
            diameter: diameterInput,
            tolerance: toleranceInput,
            cutMode: cutModeDropdown,
            showDesired: showDesired,
            showResult: showResult,
            showSweep: showSweep,
            showCentreline: showCentreline,
            showCutterCircles: showCutterCircles
        };

        function removePreviewRoot() {
            if (previewRoot) {
                try {
                    previewRoot.remove();
                } catch (ignoreRemove) {
                }
                previewRoot = null;
            }
        }

        function ensureOutputLayer() {
            if (!outputLayer) {
                outputLayer = documentRef.layers.add();
                outputLayer.name = uniqueLayerName(
                    documentRef,
                    "Router Toolpath Preview"
                );
                outputLayer.locked = false;
                outputLayer.visible = true;
            }
            return outputLayer;
        }

        function updateStatus(stats) {
            var message = stats.renderedPaths + " path(s); " +
                stats.limitedCorners + " cutter-limited corner(s)";
            if (stats.tightCorners > 0) {
                message += "; " + stats.tightCorners + " overlapping tight corner(s)";
            }
            if (stats.invalidPaths > 0) {
                message += "; " + stats.invalidPaths + " invalid path(s) skipped";
            }
            if (stats.nativeErrorMessage) {
                message += ". Offset Path error: " + stats.nativeErrorMessage;
            }
            statusText.text = message + ".";
        }

        function refreshPreview(showErrors) {
            if (busy) {
                return;
            }
            busy = true;
            try {
                if (!previewCheckbox.value) {
                    removePreviewRoot();
                    statusText.text = sourceData.paths.length +
                        " closed path(s) ready.";
                    app.redraw();
                    return;
                }
                var settings = settingsFromControls(controls, unitName);
                removePreviewRoot();
                var built = buildPreview(
                    documentRef,
                    sourceData,
                    settings,
                    ensureOutputLayer()
                );
                previewRoot = built.root;
                latestStats = built.stats;
                updateStatus(latestStats);
                app.redraw();
            } catch (error) {
                removePreviewRoot();
                restoreSelection(documentRef, sourceData.originalSelection);
                statusText.text = error.message;
                if (showErrors) {
                    alert(error.message, SCRIPT_NAME);
                }
            } finally {
                busy = false;
            }
        }

        function currentUnitFromDropdown() {
            return unitDropdown.selection ?
                unitDropdown.selection.text : previousUnitName;
        }

        unitDropdown.onChange = function () {
            var nextUnitName = currentUnitFromDropdown();
            var oldDiameter = parsePositiveNumber(diameterInput.text);
            var oldTolerance = parsePositiveNumber(toleranceInput.text);
            if (oldDiameter !== null) {
                diameterInput.text = roundForField(pointsToUnit(
                    unitToPoints(oldDiameter, previousUnitName),
                    nextUnitName
                ));
            }
            if (oldTolerance !== null) {
                toleranceInput.text = roundForField(pointsToUnit(
                    unitToPoints(oldTolerance, previousUnitName),
                    nextUnitName
                ));
            }
            unitName = nextUnitName;
            previousUnitName = nextUnitName;
            toleranceUnitLabel.text = unitName;
            refreshPreview(false);
        };

        diameterInput.onChange = function () {
            refreshPreview(false);
        };
        toleranceInput.onChange = function () {
            refreshPreview(false);
        };
        cutModeDropdown.onChange = function () {
            refreshPreview(false);
        };
        showDesired.onClick = function () {
            refreshPreview(false);
        };
        showResult.onClick = function () {
            refreshPreview(false);
        };
        showSweep.onClick = function () {
            refreshPreview(false);
        };
        showCentreline.onClick = function () {
            refreshPreview(false);
        };
        showCutterCircles.onClick = function () {
            refreshPreview(false);
        };
        previewCheckbox.onClick = function () {
            refreshPreview(false);
        };

        dialog.onShow = function () {
            refreshPreview(false);
        };

        okButton.onClick = function () {
            var settings;
            try {
                settings = settingsFromControls(controls, unitName);
                if (!settings.showDesired && !settings.showResult &&
                        !settings.showSweep &&
                        !settings.showCentreline &&
                        !settings.showCutterCircles) {
                    throw new Error("Select at least one display option.");
                }
                removePreviewRoot();
                var built = buildPreview(
                    documentRef,
                    sourceData,
                    settings,
                    ensureOutputLayer()
                );
                previewRoot = built.root;
                latestStats = built.stats;
                if (latestStats.renderedPaths === 0) {
                    throw new Error(
                        "No valid router paths could be produced from the selection." +
                        (latestStats.nativeErrorMessage ?
                            " Offset Path error: " +
                                latestStats.nativeErrorMessage : "")
                    );
                }
                acceptedSettings = settings;
                dialog.close(1);
            } catch (error) {
                restoreSelection(documentRef, sourceData.originalSelection);
                alert(error.message, SCRIPT_NAME);
            }
        };

        cancelButton.onClick = function () {
            dialog.close(0);
        };

        var response = dialog.show();
        if (response !== 1) {
            removePreviewRoot();
            if (outputLayer) {
                try {
                    outputLayer.remove();
                } catch (ignoreLayerRemove) {
                }
            }
            restoreSelection(documentRef, sourceData.originalSelection);
            app.redraw();
            return null;
        }

        if (outputLayer && acceptedSettings) {
            var diameterLabel = roundForField(acceptedSettings.diameterValue);
            outputLayer.name = uniqueLayerName(
                documentRef,
                "Router Toolpath - " + diameterLabel + " " +
                    acceptedSettings.unitName + " - " +
                    cutModeShortName(acceptedSettings.cutMode)
            );
            previewRoot.name = "Router toolpath - " + diameterLabel + " " +
                acceptedSettings.unitName + " - " +
                cutModeShortName(acceptedSettings.cutMode);
        }
        restoreSelection(documentRef, sourceData.originalSelection);
        app.redraw();
        return {
            settings: acceptedSettings,
            stats: latestStats,
            layer: outputLayer
        };
    }

    function run() {
        if (typeof app === "undefined" || !app.documents ||
                app.documents.length === 0) {
            alert("Open an Illustrator document and select at least one closed path.",
                SCRIPT_NAME);
            return;
        }
        var documentRef = app.activeDocument;
        if (!documentRef.selection || documentRef.selection.length === 0) {
            alert("Select at least one closed path, compound path, or group.",
                SCRIPT_NAME);
            return;
        }

        var sourceData = collectSelectedPaths(documentRef);
        if (sourceData.paths.length === 0) {
            var noPathsMessage = "No supported closed paths were found.";
            if (sourceData.skippedOpen > 0) {
                noPathsMessage += " Open paths are not supported for profile cuts.";
            }
            alert(noPathsMessage, SCRIPT_NAME);
            return;
        }

        var result = showDialog(documentRef, sourceData);
        if (!result) {
            return;
        }

        var skipped = sourceData.skippedOpen +
            sourceData.skippedUnsupported + sourceData.skippedGuides +
            result.stats.invalidPaths;
        if (skipped > 0 || result.stats.cappedPaths > 0) {
            var warning = "Toolpath created for " +
                result.stats.renderedPaths + " path(s).";
            if (skipped > 0) {
                warning += " " + skipped + " unsupported or invalid item(s) were skipped.";
            }
            if (result.stats.cappedPaths > 0) {
                warning += " " + result.stats.cappedPaths +
                    " very complex path(s) reached the flattening safety limit.";
            }
            alert(warning, SCRIPT_NAME);
        }
    }

    /* Test hook. It is inert in Illustrator and allows geometry regression tests
       to execute the same ES3 core without duplicating the implementation. */
    if (typeof __ROUTER_TOOLPATH_TEST__ !== "undefined" &&
            __ROUTER_TOOLPATH_TEST__) {
        __ROUTER_TOOLPATH_CORE__ = {
            pt: pt,
            signedArea: signedArea,
            pointInPolygon: pointInPolygon,
            offsetClosedPolyline: offsetClosedPolyline,
            machiningOffset: machiningOffset,
            nativeOffsetSequence: nativeOffsetSequence,
            nativeOffsetEffectXML: nativeOffsetEffectXML,
            appendArc: appendArc,
            distance: distance
        };
        return;
    }

    run();
}());
