#target illustrator

/*
Fast Router Radius.jsx

Creates the actual edge a round router cutter can produce from selected closed
paths. The dialog can delete the original selection after successful creation.

Outside / male cut:  +tool radius, -tool radius, expand once
Inside / pocket cut: -tool radius, +tool radius, expand once

Both Offset Path passes use round joins and miter limit 180. They are stacked
as native live effects and expanded together. This eliminates the expensive
intermediate expansion, its temporary object tree, and the second expansion.

Tool diameter is entered at real document size. Illustrator's large-canvas
scaleFactor is applied automatically.

Compatible with Adobe Illustrator ExtendScript (ES3).
*/

(function () {
    var SCRIPT_NAME = "Fast Router Radius";
    var VERSION = "2.2.0";
    var POINTS_PER_MM = 72 / 25.4;

    function trim(value) {
        return String(value).replace(/^\s+|\s+$/g, "");
    }

    function parsePositiveNumber(value) {
        var number = Number(trim(value));
        if (!isFinite(number) || number <= 0) {
            return null;
        }
        return number;
    }

    function documentScale(documentRef) {
        var scale = 1;
        try {
            scale = Number(documentRef.scaleFactor);
        } catch (ignoreScaleFactor) {
            scale = 1;
        }
        if (!isFinite(scale) || scale <= 0) {
            scale = 1;
        }
        return scale;
    }

    function cleanNumber(value) {
        return String(Math.round(value * 100000000) / 100000000);
    }

    function offsetEffectXML(offsetPoints) {
        return '<LiveEffect name="Adobe Offset Path"><Dict data="' +
            'R mlim 180 R ofst ' + cleanNumber(offsetPoints) +
            ' I jntp 0 "/></LiveEffect>';
    }

    function applyOffset(item, offsetPoints) {
        var effectXML = offsetEffectXML(offsetPoints);
        var index;
        try {
            item.applyEffect(effectXML);
            return;
        } catch (primaryError) {
            /* Older Illustrator releases may expose applyEffect only on a
               compound's first path or on children of an expanded group. */
            if (item.typename === "CompoundPathItem" &&
                    item.pathItems.length > 0) {
                item.pathItems[0].applyEffect(effectXML);
                return;
            }
            if (item.typename === "GroupItem") {
                var applied = 0;
                for (index = 0; index < item.pageItems.length; index += 1) {
                    try {
                        if (item.pageItems[index].parent === item) {
                            applyOffset(item.pageItems[index], offsetPoints);
                            applied += 1;
                        }
                    } catch (ignoreChildEffect) {
                    }
                }
                if (applied > 0) {
                    return;
                }
            }
            throw primaryError;
        }
    }

    function uniqueLayerName(documentRef, baseName) {
        var name = baseName;
        var suffix = 2;
        var exists;
        var index;
        do {
            exists = false;
            for (index = 0; index < documentRef.layers.length; index += 1) {
                if (documentRef.layers[index].name === name) {
                    exists = true;
                    name = baseName + " " + suffix;
                    suffix += 1;
                    break;
                }
            }
        } while (exists);
        return name;
    }

    function isClosedPath(pathItem) {
        try {
            return pathItem.closed && !pathItem.guides &&
                !pathItem.clipping && pathItem.pathPoints.length >= 3;
        } catch (ignorePath) {
            return false;
        }
    }

    function isUsableCompound(compoundItem) {
        var index;
        try {
            if (compoundItem.pathItems.length === 0) {
                return false;
            }
            for (index = 0; index < compoundItem.pathItems.length;
                    index += 1) {
                if (!compoundItem.pathItems[index].closed ||
                        compoundItem.pathItems[index].guides) {
                    return false;
                }
            }
            return true;
        } catch (ignoreCompound) {
            return false;
        }
    }

    function isUsableGroup(groupItem) {
        var directChildCount = 0;
        var index;
        try {
            for (index = 0; index < groupItem.pageItems.length; index += 1) {
                var child = groupItem.pageItems[index];
                if (child.parent !== groupItem) {
                    continue;
                }
                directChildCount += 1;
                if (child.typename === "PathItem") {
                    if (!isClosedPath(child)) {
                        return false;
                    }
                } else if (child.typename === "CompoundPathItem") {
                    if (!isUsableCompound(child)) {
                        return false;
                    }
                } else if (child.typename === "GroupItem") {
                    if (!isUsableGroup(child)) {
                        return false;
                    }
                } else {
                    return false;
                }
            }
        } catch (ignoreGroup) {
            return false;
        }
        return directChildCount > 0;
    }

    function collectGeometry(selection) {
        var geometry = [];
        var skipped = 0;

        function visit(item) {
            var index;
            if (!item) {
                return;
            }
            try {
                if (item.hidden || item.locked) {
                    skipped += 1;
                    return;
                }
            } catch (ignoreVisibility) {
            }

            if (item.typename === "PathItem") {
                if (item.parent &&
                        item.parent.typename === "CompoundPathItem") {
                    return;
                }
                if (isClosedPath(item)) {
                    geometry.push(item);
                } else {
                    skipped += 1;
                }
                return;
            }

            if (item.typename === "CompoundPathItem") {
                if (isUsableCompound(item)) {
                    geometry.push(item);
                } else {
                    skipped += 1;
                }
                return;
            }

            if (item.typename === "GroupItem") {
                /* A clean, explicitly grouped selection is one native effect
                   target. This avoids one duplicate and two applyEffect calls
                   for every child while retaining isolation between separate
                   top-level selections. */
                if (isUsableGroup(item)) {
                    geometry.push(item);
                    return;
                }
                for (index = 0; index < item.pageItems.length; index += 1) {
                    try {
                        if (item.pageItems[index].parent === item) {
                            visit(item.pageItems[index]);
                        }
                    } catch (ignoreChild) {
                    }
                }
                return;
            }

            skipped += 1;
        }

        var selectionIndex;
        for (selectionIndex = 0; selectionIndex < selection.length;
                selectionIndex += 1) {
            visit(selection[selectionIndex]);
        }
        return { items: geometry, skipped: skipped };
    }

    function snapshotSelection(selection) {
        var items = [];
        var index;
        for (index = 0; index < selection.length; index += 1) {
            items.push(selection[index]);
        }
        return items;
    }

    function deleteOriginalSelection(items) {
        var removed = 0;
        var failed = 0;
        var index;
        for (index = items.length - 1; index >= 0; index -= 1) {
            try {
                items[index].remove();
                removed += 1;
            } catch (ignoreDeleteFailure) {
                failed += 1;
            }
        }
        return { removed: removed, failed: failed };
    }

    function prepareForOffset(item, sharedBlack) {
        var black = sharedBlack;
        if (!black) {
            black = new RGBColor();
            black.red = 0;
            black.green = 0;
            black.blue = 0;
        }

        function preparePath(path) {
            path.filled = true;
            path.fillColor = black;
            path.stroked = false;
            path.opacity = 100;
        }

        if (item.typename === "PathItem") {
            preparePath(item);
            return;
        }
        if (item.typename === "CompoundPathItem" &&
                item.pathItems.length > 0) {
            /* Compound components share one appearance. Setting its first
               path avoids a slow property-write loop over every subpath. */
            preparePath(item.pathItems[0]);
            return;
        }
        if (item.typename === "GroupItem") {
            var index;
            for (index = 0; index < item.pageItems.length; index += 1) {
                try {
                    if (item.pageItems[index].parent === item) {
                        prepareForOffset(item.pageItems[index], black);
                    }
                } catch (ignoreChildPreparation) {
                }
            }
        }
    }

    function expandTogether(documentRef, items) {
        var index;
        var selected = 0;
        documentRef.selection = null;
        for (index = 0; index < items.length; index += 1) {
            try {
                items[index].selected = true;
                selected += 1;
            } catch (ignoreInvalidItem) {
            }
        }
        if (selected === 0) {
            throw new Error("No offset artwork remained available to expand.");
        }
        app.executeMenuCommand("expandStyle");
        documentRef.selection = null;
    }

    function showSettings(documentRef) {
        var scale = documentScale(documentRef);
        var dialog = new Window("dialog", SCRIPT_NAME + " " + VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.margins = 16;

        var diameterRow = dialog.add("group");
        diameterRow.add("statictext", undefined, "Tool diameter:");
        var diameterInput = diameterRow.add("edittext", undefined, "6");
        diameterInput.characters = 8;
        diameterRow.add("statictext", undefined, "mm");

        var cutRow = dialog.add("group");
        cutRow.add("statictext", undefined, "Cut type:");
        var cutType = cutRow.add("dropdownlist", undefined, [
            "Outside / male",
            "Inside / pocket"
        ]);
        cutType.selection = 0;

        var scaleText = scale === 1 ?
            "Document scale: 1:1" :
            "Document scale: 1:" + cleanNumber(scale) + " (automatic)";
        dialog.add("statictext", undefined, scaleText);

        var deleteOriginal = dialog.add(
            "checkbox", undefined, "Delete original selection"
        );
        deleteOriginal.value = true;

        var buttons = dialog.add("group");
        buttons.alignment = "right";
        buttons.add("button", undefined, "Cancel", { name: "cancel" });
        var createButton = buttons.add(
            "button", undefined, "Create", { name: "ok" }
        );

        createButton.onClick = function () {
            var diameter = parsePositiveNumber(diameterInput.text);
            if (diameter === null) {
                alert("Enter a tool diameter greater than zero.", SCRIPT_NAME);
                return;
            }
            dialog.close(1);
        };

        diameterInput.active = true;
        if (dialog.show() !== 1) {
            return null;
        }

        return {
            diameterMM: parsePositiveNumber(diameterInput.text),
            outside: cutType.selection.index === 0,
            scale: scale,
            deleteOriginal: deleteOriginal.value
        };
    }

    function buildResult(documentRef, geometry, settings) {
        var radiusPoints = settings.diameterMM * POINTS_PER_MM *
            0.5 / settings.scale;
        var firstOffset = settings.outside ? radiusPoints : -radiusPoints;
        var secondOffset = -firstOffset;
        var modeName = settings.outside ? "Outside" : "Inside";
        var baseName = "Router Radius - " + settings.diameterMM +
            " mm - " + modeName;
        var layer = documentRef.layers.add();
        layer.name = uniqueLayerName(documentRef, baseName);
        var root = layer.groupItems.add();
        root.name = baseName;
        var expansionTargets = [];
        var failed = 0;
        var index;
        var sharedBlack = new RGBColor();
        sharedBlack.red = 0;
        sharedBlack.green = 0;
        sharedBlack.blue = 0;

        try {
            for (index = 0; index < geometry.length; index += 1) {
                var seed = null;
                try {
                    seed = geometry[index].duplicate(
                        root,
                        ElementPlacement.PLACEATEND
                    );
                    prepareForOffset(seed, sharedBlack);

                    /* Appearance effects are cumulative. Stacking the reverse
                       offset before expansion produces the same two-pass edge
                       without building an intermediate expanded object tree. */
                    applyOffset(seed, firstOffset);
                    applyOffset(seed, secondOffset);
                    expansionTargets.push(seed);
                } catch (firstError) {
                    failed += 1;
                    if (seed) {
                        try {
                            seed.remove();
                        } catch (ignoreFailedSeedRemove) {
                        }
                    }
                }
            }

            if (expansionTargets.length === 0) {
                throw new Error("Illustrator could not create a result from the selection.");
            }

            /* This is the only Expand Appearance command in the script. */
            expandTogether(documentRef, expansionTargets);

            documentRef.selection = null;
            root.selected = true;
            return {
                created: expansionTargets.length,
                failed: failed,
                layer: layer
            };
        } catch (error) {
            try {
                layer.remove();
            } catch (ignoreLayerRemove) {
            }
            throw error;
        }
    }

    function run() {
        if (typeof app === "undefined" || !app.documents ||
                app.documents.length === 0) {
            alert("Open an Illustrator document first.", SCRIPT_NAME);
            return;
        }

        var documentRef = app.activeDocument;
        if (!documentRef.selection || documentRef.selection.length === 0) {
            alert("Select at least one closed path, compound path, or group.",
                SCRIPT_NAME);
            return;
        }

        var originalSelection = snapshotSelection(documentRef.selection);
        var collected = collectGeometry(originalSelection);
        if (collected.items.length === 0) {
            alert("No supported closed paths were found in the selection.",
                SCRIPT_NAME);
            return;
        }

        var settings = showSettings(documentRef);
        if (!settings) {
            return;
        }

        var previousInteraction = app.userInteractionLevel;
        var result;
        var deletionResult = { removed: 0, failed: 0 };
        var buildError = null;
        try {
            app.userInteractionLevel = UserInteractionLevel.DONTDISPLAYALERTS;
            result = buildResult(documentRef, collected.items, settings);
            if (settings.deleteOriginal) {
                deletionResult = deleteOriginalSelection(originalSelection);
            }
        } catch (error) {
            buildError = error;
        } finally {
            try {
                app.userInteractionLevel = previousInteraction;
            } catch (ignoreInteractionRestore) {
            }
        }

        if (buildError) {
            alert(buildError.message || String(buildError), SCRIPT_NAME);
            return;
        }

        if (collected.skipped > 0 || result.failed > 0 ||
                deletionResult.failed > 0) {
            var warning = "Created " + result.created + " result(s).";
            if (collected.skipped > 0 || result.failed > 0) {
                warning += " " + (collected.skipped + result.failed) +
                    " unsupported or failed item(s) were skipped.";
            }
            if (deletionResult.failed > 0) {
                warning += " " + deletionResult.failed +
                    " original selection item(s) could not be deleted.";
            }
            alert(
                warning,
                SCRIPT_NAME
            );
        }
    }

    run();
}());
