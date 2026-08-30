#target illustrator

/*
    Roof-Skeleton-Red.jsx

    Fast, geometry-based roof-ridge skeleton generator for Adobe Illustrator.

    Pipeline:
      1. Flatten selected Bezier boundaries adaptively for analysis only.
      2. Sample points along the actual boundary geometry.
      3. Build a Delaunay triangulation of those boundary sites.
      4. Extract interior Voronoi edges supported by separated boundary sites.
      5. Trace and simplify the resulting geometric graph.
      6. Add angle-bisector hips/valleys from sharp corners.

    The optional terminal-fan cleanup removes boundary-facing leaf chains near
    sharp corners, leaving the main centreline ending at the stroke-cap centre.

    No raster grid, bitmap, distance field, or thinning pass is used.
    Original artwork is preserved. Output paths are stroked pure red.

    This is a fast curve-friendly Voronoi/roof-ridge approximation. It is not
    an exact event-driven polygon straight-skeleton implementation.
*/

(function () {
    var SCRIPT_NAME = "Roof Skeleton Red — Geometry v4";
    var OUTPUT_LAYER_NAME = "Roof Skeleton";
    var MAX_BOUNDARY_SITES = 650;
    var MAX_FLATTEN_DEPTH = 12;
    var CORNER_THRESHOLD_DEGREES = 24;
    var BOUNDARY_NEIGHBOUR_RANGE = 3;
    var EPSILON = 0.0000001;

    function main() {
        if (app.documents.length === 0) {
            alert("Open a document and select one or more closed shapes.", SCRIPT_NAME);
            return;
        }

        var doc = app.activeDocument;
        if (!doc.selection || doc.selection.length === 0) {
            alert("Select one or more closed paths, compound paths, groups, or text objects.", SCRIPT_NAME);
            return;
        }

        var options = showOptionsDialog();
        if (!options) {
            return;
        }

        var scaleFactor = getDocumentScaleFactor(doc);
        var requestedSpacing = millimetresToDocumentPoints(options.detailMillimetres, scaleFactor);
        var pruneDistance = millimetresToDocumentPoints(options.pruneMillimetres, scaleFactor);
        var strokeWidth = options.strokePoints / scaleFactor;
        var flattenTolerance = Math.max(requestedSpacing * 0.12, 0.02 / scaleFactor);
        var selectionCopy = copySelection(doc.selection);
        var regions = [];
        var stats = {
            unsupported: 0,
            openPaths: 0,
            emptyRegions: 0,
            adjustedRegions: 0,
            totalSites: 0,
            totalTriangles: 0,
            maxEffectiveSpacing: requestedSpacing
        };

        for (var s = 0; s < selectionCopy.length; s++) {
            collectRegionsFromItem(selectionCopy[s], regions, flattenTolerance, stats);
        }

        if (regions.length === 0) {
            alert("No usable closed outlines were found in the selection.", SCRIPT_NAME);
            return;
        }

        var progress = createProgressPalette(regions.length);
        var allChains = [];
        var processed = 0;

        try {
            for (var r = 0; r < regions.length; r++) {
                updateProgress(progress, r, regions.length,
                    "Solving region " + (r + 1) + " of " + regions.length + " geometrically…");

                var result = buildGeometricSkeleton(
                    regions[r],
                    requestedSpacing,
                    MAX_BOUNDARY_SITES,
                    options.includeCorners,
                    pruneDistance
                );

                if (!result || result.chains.length === 0) {
                    stats.emptyRegions++;
                    continue;
                }

                stats.totalSites += result.siteCount;
                stats.totalTriangles += result.triangleCount;
                if (result.effectiveSpacing > requestedSpacing * 1.0001) {
                    stats.adjustedRegions++;
                    stats.maxEffectiveSpacing = Math.max(stats.maxEffectiveSpacing, result.effectiveSpacing);
                }

                for (var c = 0; c < result.chains.length; c++) {
                    allChains.push(result.chains[c]);
                }
                processed++;
            }

            if (allChains.length === 0) {
                throw new Error("No skeleton network could be generated. Try a smaller geometry detail value.");
            }

            updateProgress(progress, regions.length, regions.length, "Drawing red vector paths…");
            var outputGroup = drawChains(doc, allChains, strokeWidth);
            doc.selection = null;
            outputGroup.selected = true;
            app.redraw();
            closeProgress(progress);

            var message = "Created " + allChains.length + " red skeleton path" +
                (allChains.length === 1 ? "" : "s") + " from " + processed + " region" +
                (processed === 1 ? "" : "s") + ".";

            message += "\n\nGeometry processed: " + stats.totalSites + " boundary sites and " +
                stats.totalTriangles + " Delaunay triangles.";

            if (stats.adjustedRegions > 0) {
                message += "\n\n" + stats.adjustedRegions + " complex region" +
                    (stats.adjustedRegions === 1 ? " was" : "s were") +
                    " automatically capped at " + MAX_BOUNDARY_SITES +
                    " sites. Maximum effective detail: " +
                    formatNumber(documentPointsToMillimetres(stats.maxEffectiveSpacing, scaleFactor), 2) + " mm.";
            }

            if (stats.openPaths > 0 || stats.unsupported > 0 || stats.emptyRegions > 0) {
                message += "\n\nSkipped:";
                if (stats.openPaths > 0) {
                    message += " " + stats.openPaths + " open path" + (stats.openPaths === 1 ? "" : "s") + ";";
                }
                if (stats.unsupported > 0) {
                    message += " " + stats.unsupported + " unsupported item" +
                        (stats.unsupported === 1 ? "" : "s") + ";";
                }
                if (stats.emptyRegions > 0) {
                    message += " " + stats.emptyRegions + " empty/degenerate region" +
                        (stats.emptyRegions === 1 ? "" : "s") + ";";
                }
                message = message.replace(/;$/, ".");
            }

            alert(message, SCRIPT_NAME);
        } catch (error) {
            closeProgress(progress);
            alert("Roof skeleton failed:\n\n" + error.message +
                (error.line ? "\n\nLine: " + error.line : ""), SCRIPT_NAME);
        }
    }

    function showOptionsDialog() {
        var dlg = new Window("dialog", SCRIPT_NAME);
        dlg.orientation = "column";
        dlg.alignChildren = "fill";
        dlg.spacing = 10;
        dlg.margins = 16;

        var intro = dlg.add("statictext", undefined,
            "Create a red roof-ridge skeleton using boundary geometry and Voronoi edges.",
            { multiline: true });
        intro.preferredSize.width = 420;

        var detailGroup = dlg.add("group");
        detailGroup.add("statictext", undefined, "Geometry detail (mm):");
        var detailInput = detailGroup.add("edittext", undefined, "4.0");
        detailInput.characters = 8;
        detailGroup.add("statictext", undefined, "Smaller = finer, but more geometry");

        var pruneGroup = dlg.add("group");
        pruneGroup.add("statictext", undefined, "Remove branches shorter than (mm):");
        var pruneInput = pruneGroup.add("edittext", undefined, "1.0");
        pruneInput.characters = 8;

        var strokeGroup = dlg.add("group");
        strokeGroup.add("statictext", undefined, "Red stroke width (pt):");
        var strokeInput = strokeGroup.add("edittext", undefined, "0.75");
        strokeInput.characters = 8;

        var cornerCheck = dlg.add("checkbox", undefined, "Include roof hips/valleys from sharp corners");
        cornerCheck.value = false;

        var buttonGroup = dlg.add("group");
        buttonGroup.alignment = "right";
        buttonGroup.add("button", undefined, "Cancel", { name: "cancel" });
        buttonGroup.add("button", undefined, "Create", { name: "ok" });

        detailInput.active = true;
        if (dlg.show() !== 1) {
            return null;
        }

        var detail = parseDecimal(detailInput.text);
        var prune = parseDecimal(pruneInput.text);
        var stroke = parseDecimal(strokeInput.text);

        if (!(detail > 0) || detail > 1000) {
            alert("Geometry detail must be greater than 0 mm.", SCRIPT_NAME);
            return showOptionsDialog();
        }
        if (!(prune >= 0)) {
            alert("Branch removal must be 0 mm or greater.", SCRIPT_NAME);
            return showOptionsDialog();
        }
        if (!(stroke > 0)) {
            alert("Stroke width must be greater than 0 pt.", SCRIPT_NAME);
            return showOptionsDialog();
        }

        return {
            detailMillimetres: detail,
            pruneMillimetres: prune,
            strokePoints: stroke,
            includeCorners: cornerCheck.value
        };
    }

    function buildGeometricSkeleton(region, requestedSpacing, maxSites, includeCorners, pruneDistance) {
        var removeTerminalFans = !includeCorners;
        var sampled = buildBoundarySites(region, requestedSpacing, maxSites);
        if (!sampled || sampled.sites.length < 3) {
            return null;
        }

        var triangles = triangulateDelaunay(sampled.sites, sampled.bounds);
        if (triangles.length === 0) {
            return null;
        }

        var graph = buildInteriorVoronoiGraph(
            sampled.sites,
            triangles,
            region,
            sampled.effectiveSpacing,
            sampled.bounds
        );

        collapseCompactCenterArtifacts(graph, sampled.effectiveSpacing);

        if (removeTerminalFans) {
            removeTerminalCornerFans(graph, region.corners, sampled.effectiveSpacing);
            trimCornerSeekingTerminalTails(graph, region.corners, sampled.effectiveSpacing);
        }

        if (pruneDistance > 0) {
            pruneGraphBranches(graph, pruneDistance);
        }

        var chains = traceGraph(graph);
        if (includeCorners && !removeTerminalFans) {
            addGeometricCornerConnections(chains, region.corners, graph, region, sampled.effectiveSpacing);
        }

        if (chains.length === 0 && graph.nodes.length > 0) {
            var isolated = largestRadiusNode(graph.nodes);
            if (isolated) {
                var half = sampled.effectiveSpacing * 0.28;
                chains.push([
                    { x: isolated.x - half, y: isolated.y },
                    { x: isolated.x + half, y: isolated.y }
                ]);
            }
        }

        var simplified = [];
        var simplifyTolerance = sampled.effectiveSpacing * 0.18;
        for (var i = 0; i < chains.length; i++) {
            var chain = simplifyPolyline(chains[i], simplifyTolerance);
            if (chain.length >= 2 && polylineLength(chain) >= sampled.effectiveSpacing * 0.16) {
                simplified.push(chain);
            }
        }

        return {
            chains: simplified,
            siteCount: sampled.sites.length,
            triangleCount: triangles.length,
            effectiveSpacing: sampled.effectiveSpacing
        };
    }

    function removeTerminalCornerFans(graph, corners, spacing) {
        if (!corners || corners.length === 0 || graph.edges.length === 0) {
            return;
        }

        for (var pass = 0; pass < 8; pass++) {
            var edgesToRemove = findTerminalCornerFanEdges(graph, corners, spacing);
            var removed = 0;
            for (var key in edgesToRemove) {
                if (edgesToRemove.hasOwnProperty(key) && graph.edges[Number(key)].active) {
                    graph.edges[Number(key)].active = false;
                    removed++;
                }
            }
            if (removed === 0) {
                break;
            }
        }
    }

    function findTerminalCornerFanEdges(graph, corners, spacing) {
        var adjacency = buildGraphAdjacency(graph);
        var edgesToRemove = {};

        for (var start = 0; start < graph.nodes.length; start++) {
            if (adjacency[start].length !== 1) {
                continue;
            }

            var endpoint = graph.nodes[start];
            var branchEdges = [];
            var currentNode = start;
            var previousEdge = -1;
            var terminalNode = start;
            var terminalDegree = 1;
            var branchLength = 0;
            var guard = 0;

            while (guard < graph.edges.length + 1) {
                var links = adjacency[currentNode];
                var nextLink = null;
                for (var linkIndex = 0; linkIndex < links.length; linkIndex++) {
                    if (links[linkIndex].edge !== previousEdge) {
                        nextLink = links[linkIndex];
                        break;
                    }
                }
                if (!nextLink) {
                    break;
                }

                branchEdges.push(nextLink.edge);
                branchLength += graph.edges[nextLink.edge].length;
                previousEdge = nextLink.edge;
                currentNode = nextLink.other;
                terminalNode = currentNode;
                terminalDegree = adjacency[currentNode].length;

                if (terminalDegree !== 2) {
                    break;
                }
                guard++;
            }

            if (terminalDegree < 3) {
                continue;
            }

            var junction = graph.nodes[terminalNode];
            var nearestCorner = findNearestCorner(endpoint, corners);
            if (!nearestCorner) {
                continue;
            }

            var cornerDistance = Math.sqrt(nearestCorner.distanceSquared);
            var maximumCornerReach = Math.max(spacing * 10, junction.radius * 4.0);
            var maximumBranchLength = Math.max(spacing * 10, junction.radius * 3.1);
            var radiusIsLocal = endpoint.radius <= junction.radius * 1.5 + spacing;
            var directionAlignment = cornerDirectionAlignment(
                junction,
                endpoint,
                nearestCorner.corner
            );

            if (cornerDistance <= maximumCornerReach &&
                branchLength <= maximumBranchLength &&
                radiusIsLocal &&
                directionAlignment >= 0.45) {
                for (var edgeIndex = 0; edgeIndex < branchEdges.length; edgeIndex++) {
                    edgesToRemove[branchEdges[edgeIndex]] = true;
                }
            }
        }
        return edgesToRemove;
    }

    function trimCornerSeekingTerminalTails(graph, corners, spacing) {
        if (!corners || corners.length === 0) {
            return;
        }

        var adjacency = buildGraphAdjacency(graph);
        var edgesToRemove = {};

        for (var start = 0; start < graph.nodes.length; start++) {
            if (adjacency[start].length !== 1) {
                continue;
            }

            var endpoint = graph.nodes[start];
            var nearestCorner = findNearestCorner(endpoint, corners);
            if (!nearestCorner) {
                continue;
            }

            var pathNodes = [start];
            var pathEdges = [];
            var cumulativeLengths = [0];
            var currentNode = start;
            var previousEdge = -1;
            var travelled = 0;
            var scanLimit = spacing * 20;
            var guard = 0;

            while (guard < graph.edges.length + 1 && travelled <= scanLimit) {
                var links = adjacency[currentNode];
                var nextLink = null;
                for (var l = 0; l < links.length; l++) {
                    if (links[l].edge !== previousEdge) {
                        nextLink = links[l];
                        break;
                    }
                }
                if (!nextLink) {
                    break;
                }

                pathEdges.push(nextLink.edge);
                travelled += graph.edges[nextLink.edge].length;
                previousEdge = nextLink.edge;
                currentNode = nextLink.other;
                pathNodes.push(currentNode);
                cumulativeLengths.push(travelled);

                if (adjacency[currentNode].length !== 2) {
                    break;
                }
                guard++;
            }

            if (pathNodes.length < 2) {
                continue;
            }

            var maximumRadius = endpoint.radius;
            for (var n = 1; n < pathNodes.length; n++) {
                maximumRadius = Math.max(maximumRadius, graph.nodes[pathNodes[n]].radius);
            }

            if (!(maximumRadius > endpoint.radius * 1.35 + spacing * 0.15)) {
                continue;
            }

            var cornerDistance = Math.sqrt(nearestCorner.distanceSquared);
            if (cornerDistance > Math.max(spacing * 10, maximumRadius * 4)) {
                continue;
            }

            var targetRadius = maximumRadius * 0.82;
            var maximumTrimLength = Math.max(spacing * 12, maximumRadius * 4);
            var cutNodePosition = -1;
            for (var position = 1; position < pathNodes.length; position++) {
                if (graph.nodes[pathNodes[position]].radius >= targetRadius &&
                    cumulativeLengths[position] <= maximumTrimLength) {
                    cutNodePosition = position;
                    break;
                }
            }

            if (cutNodePosition > 0) {
                for (var edgePosition = 0; edgePosition < cutNodePosition; edgePosition++) {
                    edgesToRemove[pathEdges[edgePosition]] = true;
                }
            }
        }

        for (var key in edgesToRemove) {
            if (edgesToRemove.hasOwnProperty(key)) {
                graph.edges[Number(key)].active = false;
            }
        }
    }

    function findNearestCorner(point, corners) {
        var best = null;
        for (var i = 0; i < corners.length; i++) {
            var d2 = distanceSquared(point, corners[i]);
            if (!best || d2 < best.distanceSquared) {
                best = { corner: corners[i], distanceSquared: d2 };
            }
        }
        return best;
    }

    function cornerDirectionAlignment(junction, endpoint, corner) {
        var branchVector = subtract(endpoint, junction);
        var cornerVector = subtract(corner, junction);
        var branchLength = vectorLength(branchVector);
        var cornerLength = vectorLength(cornerVector);
        if (branchLength <= EPSILON || cornerLength <= EPSILON) {
            return -1;
        }
        return dot(branchVector, cornerVector) / (branchLength * cornerLength);
    }

    function collapseCompactCenterArtifacts(graph, spacing) {
        var adjacency = buildGraphAdjacency(graph);
        var visited = createFilledArray(graph.nodes.length, false);

        for (var start = 0; start < graph.nodes.length; start++) {
            if (visited[start] || adjacency[start].length === 0) {
                continue;
            }

            var queue = [start];
            var componentNodes = [];
            var componentEdges = {};
            visited[start] = true;
            var minX = Number.POSITIVE_INFINITY;
            var minY = Number.POSITIVE_INFINITY;
            var maxX = Number.NEGATIVE_INFINITY;
            var maxY = Number.NEGATIVE_INFINITY;
            var maxRadius = 0;

            while (queue.length > 0) {
                var node = queue.pop();
                componentNodes.push(node);
                minX = Math.min(minX, graph.nodes[node].x);
                minY = Math.min(minY, graph.nodes[node].y);
                maxX = Math.max(maxX, graph.nodes[node].x);
                maxY = Math.max(maxY, graph.nodes[node].y);
                maxRadius = Math.max(maxRadius, graph.nodes[node].radius);

                for (var l = 0; l < adjacency[node].length; l++) {
                    var link = adjacency[node][l];
                    componentEdges[link.edge] = true;
                    if (!visited[link.other]) {
                        visited[link.other] = true;
                        queue.push(link.other);
                    }
                }
            }

            var dx = maxX - minX;
            var dy = maxY - minY;
            var diameter = Math.sqrt(dx * dx + dy * dy);
            if (diameter <= spacing * 6 && maxRadius >= Math.max(diameter * 4, spacing * 8)) {
                for (var edgeKeyValue in componentEdges) {
                    if (componentEdges.hasOwnProperty(edgeKeyValue)) {
                        graph.edges[Number(edgeKeyValue)].active = false;
                    }
                }
            }
        }
    }

    function buildBoundarySites(region, requestedSpacing, maxSites) {
        var bounds = getRegionBounds(region);
        if (!bounds) {
            return null;
        }

        var totalLength = 0;
        var contourLengths = [];
        for (var c = 0; c < region.contours.length; c++) {
            var length = closedPolylineLength(region.contours[c]);
            contourLengths.push(length);
            totalLength += length;
        }

        if (totalLength <= EPSILON) {
            return null;
        }

        var effectiveSpacing = Math.max(requestedSpacing, totalLength / maxSites);
        var sites = [];
        var duplicateTolerance = Math.max(effectiveSpacing * 0.0001, EPSILON);
        var seen = {};

        for (var contourIndex = 0; contourIndex < region.contours.length; contourIndex++) {
            var contour = region.contours[contourIndex];
            var firstSite = sites.length;
            var localOrder = 0;
            var targetCount = Math.max(3, Math.ceil(contourLengths[contourIndex] / effectiveSpacing));
            var contourSamples = sampleClosedContourUniformly(contour, targetCount);

            for (var p = 0; p < contourSamples.length; p++) {
                var x = contourSamples[p].x;
                var y = contourSamples[p].y;
                var key = Math.round(x / duplicateTolerance) + ":" + Math.round(y / duplicateTolerance);
                if (seen[key] === true) {
                    continue;
                }
                seen[key] = true;
                sites.push({
                    x: x,
                    y: y,
                    contourIndex: contourIndex,
                    contourOrder: localOrder,
                    contourCount: 0
                });
                localOrder++;
            }

            var count = sites.length - firstSite;
            for (var fill = firstSite; fill < sites.length; fill++) {
                sites[fill].contourCount = count;
            }
        }

        return {
            sites: sites,
            bounds: bounds,
            effectiveSpacing: effectiveSpacing
        };
    }

    function sampleClosedContourUniformly(contour, sampleCount) {
        var output = [];
        if (contour.length < 2 || sampleCount < 1) {
            return output;
        }

        var edgeLengths = [];
        var perimeter = 0;
        for (var i = 0; i < contour.length; i++) {
            var edgeLength = distance(contour[i], contour[(i + 1) % contour.length]);
            edgeLengths.push(edgeLength);
            perimeter += edgeLength;
        }
        if (perimeter <= EPSILON) {
            return output;
        }

        var edgeIndex = 0;
        var edgeStartDistance = 0;
        for (var sample = 0; sample < sampleCount; sample++) {
            var targetDistance = perimeter * sample / sampleCount;
            while (edgeIndex < edgeLengths.length - 1 &&
                   edgeStartDistance + edgeLengths[edgeIndex] < targetDistance) {
                edgeStartDistance += edgeLengths[edgeIndex];
                edgeIndex++;
            }

            var a = contour[edgeIndex];
            var b = contour[(edgeIndex + 1) % contour.length];
            var length = edgeLengths[edgeIndex];
            var t = length > EPSILON ? (targetDistance - edgeStartDistance) / length : 0;
            output.push({
                x: a.x + (b.x - a.x) * t,
                y: a.y + (b.y - a.y) * t
            });
        }
        return output;
    }

    function triangulateDelaunay(sites, bounds) {
        var count = sites.length;
        var points = [];
        var span = Math.max(bounds.maxX - bounds.minX, bounds.maxY - bounds.minY);
        var jitter = Math.max(span, 1) * 0.0000000001;

        for (var i = 0; i < count; i++) {
            var seed = ((i + 1) * 9301 + 49297) % 233280;
            var seed2 = ((i + 1) * 233 + 12345) % 65521;
            points.push({
                x: sites[i].x + (seed / 233280 - 0.5) * jitter,
                y: sites[i].y + (seed2 / 65521 - 0.5) * jitter
            });
        }

        var centerX = (bounds.minX + bounds.maxX) * 0.5;
        var centerY = (bounds.minY + bounds.maxY) * 0.5;
        var superSize = Math.max(span * 32, 1);
        points.push({ x: centerX - 2 * superSize, y: centerY - superSize });
        points.push({ x: centerX, y: centerY + 2 * superSize });
        points.push({ x: centerX + 2 * superSize, y: centerY - superSize });

        var superA = count;
        var superB = count + 1;
        var superC = count + 2;
        var firstTriangle = makeTriangle(superA, superB, superC, points);
        if (!firstTriangle) {
            return [];
        }
        var triangles = [firstTriangle];

        for (var pointIndex = 0; pointIndex < count; pointIndex++) {
            var point = points[pointIndex];
            var edgeCounts = {};
            var retained = [];

            for (var t = 0; t < triangles.length; t++) {
                var triangle = triangles[t];
                var dx = point.x - triangle.cx;
                var dy = point.y - triangle.cy;
                var insideCircle = dx * dx + dy * dy <= triangle.r2 * (1 + 0.00000001) + EPSILON;
                if (insideCircle) {
                    countEdge(edgeCounts, triangle.a, triangle.b);
                    countEdge(edgeCounts, triangle.b, triangle.c);
                    countEdge(edgeCounts, triangle.c, triangle.a);
                } else {
                    retained.push(triangle);
                }
            }

            for (var key in edgeCounts) {
                if (!edgeCounts.hasOwnProperty(key) || edgeCounts[key].count !== 1) {
                    continue;
                }
                var edge = edgeCounts[key];
                var newTriangle = makeTriangle(edge.a, edge.b, pointIndex, points);
                if (newTriangle) {
                    retained.push(newTriangle);
                }
            }
            triangles = retained;
        }

        var output = [];
        for (var finalIndex = 0; finalIndex < triangles.length; finalIndex++) {
            var finalTriangle = triangles[finalIndex];
            if (finalTriangle.a < count && finalTriangle.b < count && finalTriangle.c < count) {
                output.push(finalTriangle);
            }
        }
        return output;
    }

    function makeTriangle(aIndex, bIndex, cIndex, points) {
        var a = points[aIndex];
        var b = points[bIndex];
        var c = points[cIndex];
        var determinant = 2 * (
            a.x * (b.y - c.y) +
            b.x * (c.y - a.y) +
            c.x * (a.y - b.y)
        );

        var scale = Math.max(
            Math.abs(a.x), Math.abs(a.y), Math.abs(b.x),
            Math.abs(b.y), Math.abs(c.x), Math.abs(c.y), 1
        );
        if (Math.abs(determinant) <= EPSILON * scale) {
            return null;
        }

        var aSquared = a.x * a.x + a.y * a.y;
        var bSquared = b.x * b.x + b.y * b.y;
        var cSquared = c.x * c.x + c.y * c.y;
        var centerX = (
            aSquared * (b.y - c.y) +
            bSquared * (c.y - a.y) +
            cSquared * (a.y - b.y)
        ) / determinant;
        var centerY = (
            aSquared * (c.x - b.x) +
            bSquared * (a.x - c.x) +
            cSquared * (b.x - a.x)
        ) / determinant;
        var dx = centerX - a.x;
        var dy = centerY - a.y;

        return {
            a: aIndex,
            b: bIndex,
            c: cIndex,
            cx: centerX,
            cy: centerY,
            r2: dx * dx + dy * dy
        };
    }

    function buildInteriorVoronoiGraph(sites, triangles, region, spacing, bounds) {
        var graph = { nodes: [], edges: [] };
        var triangleToNode = createFilledArray(triangles.length, -1);
        var mergeTolerance = Math.max(spacing * 0.12, EPSILON * 10);
        var nodeBuckets = {};

        for (var t = 0; t < triangles.length; t++) {
            var triangle = triangles[t];
            var center = { x: triangle.cx, y: triangle.cy };
            if (!pointInRegion(center, region)) {
                continue;
            }

            var radius = distanceToRegionBoundary(center, region);
            if (!(radius >= spacing * 0.28)) {
                continue;
            }

            triangleToNode[t] = findOrCreateGraphNode(
                graph,
                nodeBuckets,
                center,
                radius,
                mergeTolerance
            );
        }

        var sharedEdges = {};
        for (var triangleIndex = 0; triangleIndex < triangles.length; triangleIndex++) {
            if (triangleToNode[triangleIndex] < 0) {
                continue;
            }
            var tri = triangles[triangleIndex];
            addTriangleEdgeReference(sharedEdges, tri.a, tri.b, triangleIndex);
            addTriangleEdgeReference(sharedEdges, tri.b, tri.c, triangleIndex);
            addTriangleEdgeReference(sharedEdges, tri.c, tri.a, triangleIndex);
        }

        var graphEdgeSet = {};
        for (var key in sharedEdges) {
            if (!sharedEdges.hasOwnProperty(key)) {
                continue;
            }
            var shared = sharedEdges[key];
            if (shared.triangles.length !== 2) {
                continue;
            }

            if (boundarySitesAreNeighbours(sites[shared.a], sites[shared.b], BOUNDARY_NEIGHBOUR_RANGE)) {
                continue;
            }

            var nodeA = triangleToNode[shared.triangles[0]];
            var nodeB = triangleToNode[shared.triangles[1]];
            if (nodeA < 0 || nodeB < 0 || nodeA === nodeB) {
                continue;
            }

            var aPoint = graph.nodes[nodeA];
            var bPoint = graph.nodes[nodeB];
            if (!interiorSegmentIsValid(aPoint, bPoint, region)) {
                continue;
            }

            addGraphEdge(graph, nodeA, nodeB, graphEdgeSet);
        }

        return graph;
    }

    function findOrCreateGraphNode(graph, buckets, point, radius, tolerance) {
        var gx = Math.round(point.x / tolerance);
        var gy = Math.round(point.y / tolerance);
        var best = -1;
        var bestDistance = tolerance * tolerance;

        for (var y = gy - 1; y <= gy + 1; y++) {
            for (var x = gx - 1; x <= gx + 1; x++) {
                var key = x + ":" + y;
                var candidates = buckets[key];
                if (!candidates) {
                    continue;
                }
                for (var i = 0; i < candidates.length; i++) {
                    var index = candidates[i];
                    var d2 = distanceSquared(point, graph.nodes[index]);
                    if (d2 <= bestDistance) {
                        bestDistance = d2;
                        best = index;
                    }
                }
            }
        }

        if (best >= 0) {
            graph.nodes[best].radius = Math.max(graph.nodes[best].radius, radius);
            return best;
        }

        var nodeIndex = graph.nodes.length;
        graph.nodes.push({ x: point.x, y: point.y, radius: radius });
        var bucketKey = gx + ":" + gy;
        if (!buckets[bucketKey]) {
            buckets[bucketKey] = [];
        }
        buckets[bucketKey].push(nodeIndex);
        return nodeIndex;
    }

    function addGraphEdge(graph, a, b, edgeSet) {
        var key = edgeKey(a, b);
        if (edgeSet[key] === true) {
            return;
        }
        edgeSet[key] = true;
        graph.edges.push({
            a: a,
            b: b,
            length: distance(graph.nodes[a], graph.nodes[b]),
            active: true
        });
    }

    function pruneGraphBranches(graph, threshold) {
        if (!(threshold > 0)) {
            return;
        }

        for (var pass = 0; pass < 10; pass++) {
            var adjacency = buildGraphAdjacency(graph);
            var removed = false;

            for (var node = 0; node < graph.nodes.length; node++) {
                if (adjacency[node].length !== 1) {
                    continue;
                }

                var branchEdges = [];
                var current = node;
                var previousEdge = -1;
                var length = 0;
                var terminalDegree = 1;
                var guard = 0;

                while (guard < graph.edges.length + 1) {
                    var links = adjacency[current];
                    var nextLink = null;
                    for (var l = 0; l < links.length; l++) {
                        if (links[l].edge !== previousEdge) {
                            nextLink = links[l];
                            break;
                        }
                    }
                    if (!nextLink) {
                        break;
                    }

                    branchEdges.push(nextLink.edge);
                    length += graph.edges[nextLink.edge].length;
                    previousEdge = nextLink.edge;
                    current = nextLink.other;
                    terminalDegree = adjacency[current].length;

                    if (terminalDegree !== 2 || length >= threshold) {
                        break;
                    }
                    guard++;
                }

                if (length < threshold && terminalDegree > 2) {
                    for (var e = 0; e < branchEdges.length; e++) {
                        graph.edges[branchEdges[e]].active = false;
                    }
                    removed = true;
                }
            }

            if (!removed) {
                break;
            }
        }
    }

    function traceGraph(graph) {
        var chains = [];
        var adjacency = buildGraphAdjacency(graph);
        var visited = createFilledArray(graph.edges.length, false);

        for (var node = 0; node < graph.nodes.length; node++) {
            if (adjacency[node].length === 0 || adjacency[node].length === 2) {
                continue;
            }
            for (var l = 0; l < adjacency[node].length; l++) {
                var link = adjacency[node][l];
                if (!visited[link.edge]) {
                    var chain = traceGraphChain(node, link, graph, adjacency, visited);
                    if (chain.length >= 2) {
                        chains.push(chain);
                    }
                }
            }
        }

        for (var edgeIndex = 0; edgeIndex < graph.edges.length; edgeIndex++) {
            if (!graph.edges[edgeIndex].active || visited[edgeIndex]) {
                continue;
            }
            var edge = graph.edges[edgeIndex];
            var cycle = traceGraphChain(
                edge.a,
                { edge: edgeIndex, other: edge.b },
                graph,
                adjacency,
                visited
            );
            if (cycle.length >= 2) {
                chains.push(cycle);
            }
        }

        return chains;
    }

    function traceGraphChain(startNode, firstLink, graph, adjacency, visited) {
        var points = [copyPoint(graph.nodes[startNode])];
        var previousNode = startNode;
        var currentNode = firstLink.other;
        var currentEdge = firstLink.edge;
        visited[currentEdge] = true;
        var guard = 0;

        while (guard < graph.edges.length + 2) {
            points.push(copyPoint(graph.nodes[currentNode]));
            if (currentNode !== startNode && adjacency[currentNode].length !== 2) {
                break;
            }

            var links = adjacency[currentNode];
            var nextLink = null;
            for (var i = 0; i < links.length; i++) {
                if (links[i].edge !== currentEdge && !visited[links[i].edge]) {
                    nextLink = links[i];
                    break;
                }
            }
            if (!nextLink) {
                break;
            }

            previousNode = currentNode;
            currentNode = nextLink.other;
            currentEdge = nextLink.edge;
            visited[currentEdge] = true;
            if (currentNode === startNode) {
                points.push(copyPoint(graph.nodes[currentNode]));
                break;
            }
            guard++;
        }
        return points;
    }

    function buildGraphAdjacency(graph) {
        var adjacency = [];
        for (var n = 0; n < graph.nodes.length; n++) {
            adjacency[n] = [];
        }
        for (var e = 0; e < graph.edges.length; e++) {
            var edge = graph.edges[e];
            if (!edge.active) {
                continue;
            }
            adjacency[edge.a].push({ edge: e, other: edge.b });
            adjacency[edge.b].push({ edge: e, other: edge.a });
        }
        return adjacency;
    }

    function addGeometricCornerConnections(chains, corners, graph, region, spacing) {
        if (!corners || corners.length === 0 || graph.nodes.length === 0) {
            return;
        }

        var adjacency = buildGraphAdjacency(graph);
        var candidates = [];
        for (var n = 0; n < graph.nodes.length; n++) {
            if (adjacency[n].length > 0) {
                candidates.push(graph.nodes[n]);
            }
        }
        if (candidates.length === 0) {
            candidates = graph.nodes;
        }

        for (var c = 0; c < corners.length; c++) {
            var corner = corners[c];
            var direction = chooseInteriorBisector(corner, region, spacing);
            var target = findDirectionalGraphTarget(corner, direction, candidates, region);
            if (!target) {
                target = findNearestVisibleGraphTarget(corner, candidates, region);
            }

            if (target && distance(corner, target) > spacing * 0.18) {
                chains.push([copyPoint(corner), copyPoint(target)]);
            }
        }
    }

    function chooseInteriorBisector(corner, region, spacing) {
        var dx = Number(corner.bisectorX);
        var dy = Number(corner.bisectorY);
        var length = Math.sqrt(dx * dx + dy * dy);
        if (!(length > EPSILON)) {
            return null;
        }
        dx /= length;
        dy /= length;

        var forward = directionInteriorScore(corner, dx, dy, region, spacing);
        var reverse = directionInteriorScore(corner, -dx, -dy, region, spacing);
        if (forward === 0 && reverse === 0) {
            return null;
        }
        return reverse > forward ? { x: -dx, y: -dy } : { x: dx, y: dy };
    }

    function directionInteriorScore(point, dx, dy, region, spacing) {
        var fractions = [0.06, 0.12, 0.22, 0.38];
        var score = 0;
        for (var i = 0; i < fractions.length; i++) {
            var sample = {
                x: point.x + dx * spacing * fractions[i],
                y: point.y + dy * spacing * fractions[i]
            };
            if (pointInRegion(sample, region)) {
                score++;
            }
        }
        return score;
    }

    function findDirectionalGraphTarget(corner, direction, nodes, region) {
        if (!direction) {
            return null;
        }

        var ranked = [];
        for (var i = 0; i < nodes.length; i++) {
            var vx = nodes[i].x - corner.x;
            var vy = nodes[i].y - corner.y;
            var d2 = vx * vx + vy * vy;
            if (d2 <= EPSILON) {
                continue;
            }
            var along = vx * direction.x + vy * direction.y;
            if (along <= 0) {
                continue;
            }
            var cosine = along / Math.sqrt(d2);
            if (cosine < 0.66) {
                continue;
            }
            var perpendicularSquared = Math.max(0, d2 - along * along);
            ranked.push({
                node: nodes[i],
                score: perpendicularSquared * 35 + d2 * 0.015
            });
        }
        ranked.sort(sortByScore);

        for (var r = 0; r < ranked.length && r < 24; r++) {
            if (sampledSegmentInside(corner, ranked[r].node, region)) {
                return ranked[r].node;
            }
        }
        return null;
    }

    function findNearestVisibleGraphTarget(corner, nodes, region) {
        var ranked = [];
        for (var i = 0; i < nodes.length; i++) {
            ranked.push({ node: nodes[i], score: distanceSquared(corner, nodes[i]) });
        }
        ranked.sort(sortByScore);
        for (var r = 0; r < ranked.length && r < 24; r++) {
            if (sampledSegmentInside(corner, ranked[r].node, region)) {
                return ranked[r].node;
            }
        }
        return null;
    }

    function interiorSegmentIsValid(a, b, region) {
        var midpointValue = midpoint(a, b);
        if (!pointInRegion(midpointValue, region)) {
            return false;
        }

        for (var c = 0; c < region.contours.length; c++) {
            var contour = region.contours[c];
            for (var p = 0; p < contour.length; p++) {
                var edgeA = contour[p];
                var edgeB = contour[(p + 1) % contour.length];
                if (segmentsProperlyIntersect(a, b, edgeA, edgeB)) {
                    return false;
                }
            }
        }
        return true;
    }

    function sampledSegmentInside(a, b, region) {
        var sampleCount = 18;
        for (var i = 1; i < sampleCount; i++) {
            var t = i / sampleCount;
            var point = {
                x: a.x + (b.x - a.x) * t,
                y: a.y + (b.y - a.y) * t
            };
            if (!pointInRegion(point, region)) {
                return false;
            }
        }
        return true;
    }

    function pointInRegion(point, region) {
        var inside = false;
        for (var c = 0; c < region.contours.length; c++) {
            var contour = region.contours[c];
            for (var i = 0, j = contour.length - 1; i < contour.length; j = i++) {
                var a = contour[i];
                var b = contour[j];
                var intersects = ((a.y > point.y) !== (b.y > point.y)) &&
                    (point.x < (b.x - a.x) * (point.y - a.y) / (b.y - a.y) + a.x);
                if (intersects) {
                    inside = !inside;
                }
            }
        }
        return inside;
    }

    function distanceToRegionBoundary(point, region) {
        var minimum = Number.POSITIVE_INFINITY;
        for (var c = 0; c < region.contours.length; c++) {
            var contour = region.contours[c];
            for (var i = 0; i < contour.length; i++) {
                var d2 = pointLineDistanceSquared(point, contour[i], contour[(i + 1) % contour.length]);
                if (d2 < minimum) {
                    minimum = d2;
                }
            }
        }
        return Math.sqrt(minimum);
    }

    function segmentsProperlyIntersect(a, b, c, d) {
        if (Math.max(a.x, b.x) < Math.min(c.x, d.x) - EPSILON ||
            Math.max(c.x, d.x) < Math.min(a.x, b.x) - EPSILON ||
            Math.max(a.y, b.y) < Math.min(c.y, d.y) - EPSILON ||
            Math.max(c.y, d.y) < Math.min(a.y, b.y) - EPSILON) {
            return false;
        }

        var o1 = orientation(a, b, c);
        var o2 = orientation(a, b, d);
        var o3 = orientation(c, d, a);
        var o4 = orientation(c, d, b);
        return o1 * o2 < -EPSILON && o3 * o4 < -EPSILON;
    }

    function orientation(a, b, c) {
        return (b.x - a.x) * (c.y - a.y) - (b.y - a.y) * (c.x - a.x);
    }

    function boundarySitesAreNeighbours(a, b, range) {
        if (!a || !b || a.contourIndex !== b.contourIndex || a.contourCount <= 0) {
            return false;
        }
        var difference = Math.abs(a.contourOrder - b.contourOrder);
        difference = Math.min(difference, a.contourCount - difference);
        return difference <= range;
    }

    function countEdge(edgeMap, a, b) {
        var key = edgeKey(a, b);
        if (!edgeMap[key]) {
            edgeMap[key] = { a: Math.min(a, b), b: Math.max(a, b), count: 1 };
        } else {
            edgeMap[key].count++;
        }
    }

    function addTriangleEdgeReference(edgeMap, a, b, triangleIndex) {
        var key = edgeKey(a, b);
        if (!edgeMap[key]) {
            edgeMap[key] = {
                a: Math.min(a, b),
                b: Math.max(a, b),
                triangles: []
            };
        }
        edgeMap[key].triangles.push(triangleIndex);
    }

    function edgeKey(a, b) {
        return a < b ? a + ":" + b : b + ":" + a;
    }

    function largestRadiusNode(nodes) {
        var best = null;
        for (var i = 0; i < nodes.length; i++) {
            if (!best || nodes[i].radius > best.radius) {
                best = nodes[i];
            }
        }
        return best;
    }

    function collectRegionsFromItem(item, regions, flattenTolerance, stats) {
        if (!item) {
            return;
        }

        try {
            if (item.hidden || item.locked) {
                stats.unsupported++;
                return;
            }
        } catch (ignoreVisibility) {}

        var type = item.typename;
        if (type === "PathItem") {
            if (!item.closed) {
                stats.openPaths++;
                return;
            }
            var contour = flattenPathItem(item, flattenTolerance);
            if (contour.length >= 3) {
                regions.push({
                    contours: [contour],
                    corners: extractSharpCorners(item, CORNER_THRESHOLD_DEGREES)
                });
            }
            return;
        }

        if (type === "CompoundPathItem") {
            var contours = [];
            var corners = [];
            for (var p = 0; p < item.pathItems.length; p++) {
                var path = item.pathItems[p];
                if (!path.closed) {
                    stats.openPaths++;
                    continue;
                }
                var compoundContour = flattenPathItem(path, flattenTolerance);
                if (compoundContour.length >= 3) {
                    contours.push(compoundContour);
                    appendArray(corners, extractSharpCorners(path, CORNER_THRESHOLD_DEGREES));
                }
            }
            if (contours.length > 0) {
                regions.push({ contours: contours, corners: corners });
            }
            return;
        }

        if (type === "TextFrame") {
            var duplicate = null;
            var outlines = null;
            try {
                duplicate = item.duplicate();
                outlines = duplicate.createOutline();
                collectRegionsFromItem(outlines, regions, flattenTolerance, stats);
            } catch (textError) {
                stats.unsupported++;
            } finally {
                safeRemove(outlines);
                safeRemove(duplicate);
            }
            return;
        }

        if (type === "GroupItem") {
            if (item.clipped) {
                var clippingItems = [];
                findClippingItems(item, clippingItems);
                if (clippingItems.length > 0) {
                    for (var ci = 0; ci < clippingItems.length; ci++) {
                        collectRegionsFromItem(clippingItems[ci], regions, flattenTolerance, stats);
                    }
                    return;
                }
            }
            for (var g = 0; g < item.pageItems.length; g++) {
                collectRegionsFromItem(item.pageItems[g], regions, flattenTolerance, stats);
            }
            return;
        }

        if (type === "SymbolItem") {
            var symbolDuplicate = null;
            var expanded = null;
            try {
                symbolDuplicate = item.duplicate();
                expanded = symbolDuplicate.breakLink();
                collectRegionsFromItem(expanded || symbolDuplicate, regions, flattenTolerance, stats);
            } catch (symbolError) {
                stats.unsupported++;
            } finally {
                safeRemove(expanded);
                safeRemove(symbolDuplicate);
            }
            return;
        }

        stats.unsupported++;
    }

    function findClippingItems(group, output) {
        for (var i = 0; i < group.pageItems.length; i++) {
            var item = group.pageItems[i];
            if (item.typename === "PathItem" && item.clipping) {
                output.push(item);
            } else if (item.typename === "CompoundPathItem") {
                for (var p = 0; p < item.pathItems.length; p++) {
                    if (item.pathItems[p].clipping) {
                        output.push(item);
                        break;
                    }
                }
            } else if (item.typename === "GroupItem") {
                findClippingItems(item, output);
            }
        }
    }

    function flattenPathItem(pathItem, tolerance) {
        var pathPoints = pathItem.pathPoints;
        var count = pathPoints.length;
        var output = [];
        if (count < 2) {
            return output;
        }

        output.push(pointFromArray(pathPoints[0].anchor));
        var segmentCount = pathItem.closed ? count : count - 1;
        for (var i = 0; i < segmentCount; i++) {
            var next = (i + 1) % count;
            flattenCubic(
                pointFromArray(pathPoints[i].anchor),
                pointFromArray(pathPoints[i].rightDirection),
                pointFromArray(pathPoints[next].leftDirection),
                pointFromArray(pathPoints[next].anchor),
                tolerance * tolerance,
                output,
                0
            );
        }

        if (output.length > 1 && pointsEqual(output[0], output[output.length - 1], tolerance * 0.05)) {
            output.pop();
        }
        return removeConsecutiveDuplicates(output, tolerance * 0.02);
    }

    function flattenCubic(p0, p1, p2, p3, toleranceSquared, output, depth) {
        if (depth >= MAX_FLATTEN_DEPTH ||
            (pointLineDistanceSquared(p1, p0, p3) <= toleranceSquared &&
             pointLineDistanceSquared(p2, p0, p3) <= toleranceSquared)) {
            output.push(copyPoint(p3));
            return;
        }

        var p01 = midpoint(p0, p1);
        var p12 = midpoint(p1, p2);
        var p23 = midpoint(p2, p3);
        var p012 = midpoint(p01, p12);
        var p123 = midpoint(p12, p23);
        var p0123 = midpoint(p012, p123);
        flattenCubic(p0, p01, p012, p0123, toleranceSquared, output, depth + 1);
        flattenCubic(p0123, p123, p23, p3, toleranceSquared, output, depth + 1);
    }

    function extractSharpCorners(pathItem, thresholdDegrees) {
        var result = [];
        var points = pathItem.pathPoints;
        var count = points.length;
        if (!pathItem.closed || count < 3) {
            return result;
        }

        var threshold = thresholdDegrees * Math.PI / 180;
        for (var i = 0; i < count; i++) {
            var previous = (i - 1 + count) % count;
            var next = (i + 1) % count;
            var anchor = pointFromArray(points[i].anchor);
            var left = pointFromArray(points[i].leftDirection);
            var right = pointFromArray(points[i].rightDirection);
            var towardPrevious = distanceSquared(anchor, left) > EPSILON ?
                subtract(left, anchor) : subtract(pointFromArray(points[previous].anchor), anchor);
            var towardNext = distanceSquared(anchor, right) > EPSILON ?
                subtract(right, anchor) : subtract(pointFromArray(points[next].anchor), anchor);
            var lengthA = vectorLength(towardPrevious);
            var lengthB = vectorLength(towardNext);

            if (lengthA <= EPSILON || lengthB <= EPSILON) {
                continue;
            }

            var cosine = dot(towardPrevious, towardNext) / (lengthA * lengthB);
            cosine = Math.max(-1, Math.min(1, cosine));
            var deviation = Math.abs(Math.PI - Math.acos(cosine));
            if (deviation < threshold) {
                continue;
            }

            var bisector = {
                x: towardPrevious.x / lengthA + towardNext.x / lengthB,
                y: towardPrevious.y / lengthA + towardNext.y / lengthB
            };
            var bisectorLength = vectorLength(bisector);
            if (bisectorLength > EPSILON) {
                bisector.x /= bisectorLength;
                bisector.y /= bisectorLength;
            } else {
                bisector.x = 0;
                bisector.y = 0;
            }

            result.push({
                x: anchor.x,
                y: anchor.y,
                bisectorX: bisector.x,
                bisectorY: bisector.y
            });
        }
        return result;
    }

    function simplifyPolyline(points, tolerance) {
        if (points.length <= 2) {
            return points;
        }
        var keep = createFilledArray(points.length, false);
        keep[0] = true;
        keep[points.length - 1] = true;
        var stack = [[0, points.length - 1]];
        var toleranceSquared = tolerance * tolerance;

        while (stack.length > 0) {
            var range = stack.pop();
            var first = range[0];
            var last = range[1];
            var maximum = -1;
            var farthest = -1;

            for (var i = first + 1; i < last; i++) {
                var d2 = pointLineDistanceSquared(points[i], points[first], points[last]);
                if (d2 > maximum) {
                    maximum = d2;
                    farthest = i;
                }
            }

            if (farthest >= 0 && maximum > toleranceSquared) {
                keep[farthest] = true;
                stack.push([first, farthest]);
                stack.push([farthest, last]);
            }
        }

        var output = [];
        for (var p = 0; p < points.length; p++) {
            if (keep[p] && (output.length === 0 ||
                !pointsEqual(output[output.length - 1], points[p], tolerance * 0.02))) {
                output.push(points[p]);
            }
        }
        return output;
    }

    function drawChains(doc, chains, strokeWidth) {
        var layer = getOrCreateOutputLayer(doc, OUTPUT_LAYER_NAME);
        var group = layer.groupItems.add();
        group.name = "Roof Skeleton — Geometry";
        var red = new RGBColor();
        red.red = 255;
        red.green = 0;
        red.blue = 0;

        for (var i = 0; i < chains.length; i++) {
            if (chains[i].length < 2) {
                continue;
            }
            var path = group.pathItems.add();
            var anchors = [];
            for (var p = 0; p < chains[i].length; p++) {
                anchors.push([chains[i][p].x, chains[i][p].y]);
            }
            path.setEntirePath(anchors);
            path.closed = false;
            path.filled = false;
            path.stroked = true;
            path.strokeColor = red;
            path.strokeWidth = strokeWidth;
            path.name = "Roof skeleton";
            try {
                path.strokeCap = StrokeCap.ROUNDENDCAP;
                path.strokeJoin = StrokeJoin.ROUNDENDJOIN;
            } catch (ignoreCaps) {}
        }
        return group;
    }

    function getOrCreateOutputLayer(doc, name) {
        var layer;
        try {
            layer = doc.layers.getByName(name);
        } catch (notFound) {
            layer = doc.layers.add();
            layer.name = name;
        }
        layer.locked = false;
        layer.visible = true;
        return layer;
    }

    function createProgressPalette(total) {
        var win = new Window("palette", SCRIPT_NAME);
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = 14;
        var label = win.add("statictext", undefined, "Preparing geometry…");
        label.preferredSize.width = 360;
        var bar = win.add("progressbar", undefined, 0, Math.max(total, 1));
        bar.preferredSize.width = 360;
        win.show();
        return { window: win, label: label, bar: bar };
    }

    function updateProgress(progress, value, total, message) {
        if (!progress) return;
        progress.bar.maxvalue = Math.max(total, 1);
        progress.bar.value = Math.min(value, total);
        progress.label.text = message;
        progress.window.update();
    }

    function closeProgress(progress) {
        if (!progress || !progress.window) return;
        try { progress.window.close(); } catch (ignore) {}
    }

    function getRegionBounds(region) {
        var minX = Number.POSITIVE_INFINITY;
        var minY = Number.POSITIVE_INFINITY;
        var maxX = Number.NEGATIVE_INFINITY;
        var maxY = Number.NEGATIVE_INFINITY;
        for (var c = 0; c < region.contours.length; c++) {
            for (var p = 0; p < region.contours[c].length; p++) {
                var point = region.contours[c][p];
                minX = Math.min(minX, point.x);
                minY = Math.min(minY, point.y);
                maxX = Math.max(maxX, point.x);
                maxY = Math.max(maxY, point.y);
            }
        }
        if (!isFinite(minX) || !isFinite(minY) || !isFinite(maxX) || !isFinite(maxY)) {
            return null;
        }
        return { minX: minX, minY: minY, maxX: maxX, maxY: maxY };
    }

    function closedPolylineLength(points) {
        var total = 0;
        for (var i = 0; i < points.length; i++) {
            total += distance(points[i], points[(i + 1) % points.length]);
        }
        return total;
    }

    function polylineLength(points) {
        var total = 0;
        for (var i = 1; i < points.length; i++) {
            total += distance(points[i - 1], points[i]);
        }
        return total;
    }

    function sortByScore(a, b) {
        return a.score - b.score;
    }

    function pointLineDistanceSquared(point, lineStart, lineEnd) {
        var dx = lineEnd.x - lineStart.x;
        var dy = lineEnd.y - lineStart.y;
        var lengthSquared = dx * dx + dy * dy;
        if (lengthSquared <= EPSILON) {
            return distanceSquared(point, lineStart);
        }
        var t = ((point.x - lineStart.x) * dx + (point.y - lineStart.y) * dy) / lengthSquared;
        t = Math.max(0, Math.min(1, t));
        return distanceSquared(point, {
            x: lineStart.x + t * dx,
            y: lineStart.y + t * dy
        });
    }

    function distance(a, b) {
        return Math.sqrt(distanceSquared(a, b));
    }

    function distanceSquared(a, b) {
        var dx = a.x - b.x;
        var dy = a.y - b.y;
        return dx * dx + dy * dy;
    }

    function midpoint(a, b) {
        return { x: (a.x + b.x) * 0.5, y: (a.y + b.y) * 0.5 };
    }

    function subtract(a, b) {
        return { x: a.x - b.x, y: a.y - b.y };
    }

    function dot(a, b) {
        return a.x * b.x + a.y * b.y;
    }

    function vectorLength(vector) {
        return Math.sqrt(vector.x * vector.x + vector.y * vector.y);
    }

    function copyPoint(point) {
        return { x: point.x, y: point.y };
    }

    function pointFromArray(array) {
        return { x: Number(array[0]), y: Number(array[1]) };
    }

    function pointsEqual(a, b, tolerance) {
        return distanceSquared(a, b) <= tolerance * tolerance;
    }

    function removeConsecutiveDuplicates(points, tolerance) {
        var output = [];
        for (var i = 0; i < points.length; i++) {
            if (output.length === 0 || !pointsEqual(output[output.length - 1], points[i], tolerance)) {
                output.push(points[i]);
            }
        }
        return output;
    }

    function createFilledArray(length, value) {
        var output = new Array(length);
        for (var i = 0; i < length; i++) {
            output[i] = value;
        }
        return output;
    }

    function copySelection(selection) {
        var output = [];
        for (var i = 0; i < selection.length; i++) {
            output.push(selection[i]);
        }
        return output;
    }

    function appendArray(target, source) {
        for (var i = 0; i < source.length; i++) {
            target.push(source[i]);
        }
    }

    function safeRemove(item) {
        if (!item) return;
        try { item.remove(); } catch (ignore) {}
    }

    function getDocumentScaleFactor(doc) {
        try {
            var factor = Number(doc.scaleFactor);
            if (isFinite(factor) && factor > 0) {
                return factor;
            }
        } catch (ignore) {}
        return 1;
    }

    function millimetresToDocumentPoints(mm, scaleFactor) {
        return mm * 72 / 25.4 / scaleFactor;
    }

    function documentPointsToMillimetres(points, scaleFactor) {
        return points * scaleFactor * 25.4 / 72;
    }

    function parseDecimal(value) {
        return parseFloat(String(value).replace(",", "."));
    }

    function formatNumber(value, decimals) {
        return Number(value).toFixed(decimals).replace(/\.00$/, "");
    }

    if (typeof module !== "undefined" && module.exports) {
        module.exports = {
            buildGeometricSkeleton: buildGeometricSkeleton,
            buildBoundarySites: buildBoundarySites,
            triangulateDelaunay: triangulateDelaunay,
            pointInRegion: pointInRegion,
            simplifyPolyline: simplifyPolyline,
            polylineLength: polylineLength
        };
    }

    if (typeof app !== "undefined") {
        main();
    }
})();
