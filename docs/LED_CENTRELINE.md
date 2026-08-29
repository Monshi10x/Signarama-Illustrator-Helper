# LED centreline layout design

## Status and public workflow research

This is an independent implementation. No LED Wizard code, assets, module database, or binary behavior is used. The requested public references are the [LED Wizard product site](https://ledwizard8.com/), [PowerFlow overview](https://docs.ledwizard8.com/doc/lw8/6.2), [module creation/editing](https://docs.ledwizard8.com/doc/lw8/6.3), [PowerFlow property bar](https://docs.ledwizard8.com/doc/lw8/7.1.2), and [official training video](https://www.youtube.com/watch?v=eHTNjMsbKhk). Network access returned HTTP 401 in the implementation environment, so their current contents could not be independently verified here. The only workflow concepts relied on are those explicitly supplied in the feature brief: editable guides precede placement, spacing is centre-to-centre, runs split at a series limit, and excessive wire gaps should be visible.

Bundled profiles are generic placeholders requiring operator verification. Wattage is an estimate from user-entered data, not an electrical, supplier, code-compliance, or photometric recommendation.

## Contracts and coordinates

Geometry uses document-space millimetres in the panel engine. A component has even-odd `contours: [{points:[{x,y}]}]`; counters need no special winding. Engine results use schema version 1 and contain bounded-grid diagnostics, ordered editable guide polylines, placements, series, and warnings. Host payloads convert with the document scale factor and retain source coordinates.

Illustrator ownership metadata is JSON in `note` with `owner: "signarama-led-centreline"`, `schemaVersion: 1`, source/layout/guide/run IDs and settings. Output layers are `LED Guide Paths`, `LED Modules`, `LED Wiring`, and `LED Layout Stats`. Replacement may target only objects with matching ownership metadata.

## Centreline algorithm

`js/led-centreline.js` is UMD/CommonJS pure JavaScript. It:

1. rasterizes all contours with even-odd fill;
2. adapts cell size to a default 250,000-cell cap;
3. computes the exact separable squared Euclidean distance transform;
4. erodes by half the larger module dimension plus edge clearance;
5. applies deterministic Zhang-Suen topology-preserving thinning;
6. traces the 8-neighbour graph between terminals/junctions and prunes only short branches;
7. maps ordered guide pixels back to document coordinates;
8. distributes modules using `ceil(L/S)` intervals so computed spacing is no greater than maximum spacing; and
9. samples every oriented footprint edge plus its centre against the even-odd fill before accepting it.

This bounded raster method is a medial-axis-derived centreline, not repeated contour offsetting. Counters remain outside the mask. Determinism comes from fixed scan and graph traversal order.

## Row planning

The production-enabled implementation is intentionally **single-row**. Local distance values establish whether one row fits. Force 2/3 and multi-row Auto planning remain disabled until gradual topology-aware transitions and full footprint validation have representative Illustrator fixtures. This is preferable to labelling unreliable parallel offsets as complete. Proposed future thresholds are based on local usable diameter: row footprints plus inter-row gaps must all fit, with transition changes limited to one row over a run length of at least two maximum spacings.

## Placement, series, and wiring

For guide usable length `L`, endpoint inset `E`, and maximum spacing `S`, interval count is `max(1, ceil((L-2E)/S))`; actual spacing is `(L-2E)/n`. Stable polyline tangents orient modules. Non-fitting footprints are omitted rather than relaxing clearance. Series split deterministically at the user limit. Gaps over wire reach are warnings and must be styled as unreachable; output does not imply engineering approval.

## Performance and safety

Default budget is 250,000 cells per component and the grid coarsens when necessary. Raster and graph work runs in CEP, not Illustrator. Geometry is batched across the bridge; there are no per-cell/module calls or recurring polls. Invalid, empty and non-finite geometry is rejected. Debug visualization is disabled in shipping builds.

## Operator workflow

1. Select filled live text or outlined/grouped/compound letters.
2. **Create Guide Paths Only** or **Create Layout**.
3. Edit paths on `LED Guide Paths` directly.
4. **Repopulate From Guides** to preserve those edits while replacing owned modules/wiring/stats.
5. Use **Clear Generated Layout** to remove only owned results.

Use generic/local profiles only after verifying physical and electrical values for the actual module.

## Known limitations and manual test

The first production slice enables correct single-row centreline generation only. Automated fixtures cover masks, counters, topology, spacing, containment, series and budgets; multi-row controls must remain disabled. Illustrator was unavailable in the implementation container, so the required Illustrator smoke test (`SIGNARAMA 808`, serif/sans outlines, edited-guide repopulation, forced modes, short wire/series, large canvas, undo/cancel/reopen) remains mandatory before release. Record Illustrator version, OS, elapsed time, module count, and limitations when run.
