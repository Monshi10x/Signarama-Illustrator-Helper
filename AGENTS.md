# LED Centreline and Module Layout for the Signarama Illustrator Helper

## Mission
Implement a production-quality **LEDs** workflow for this Adobe Illustrator CEP extension. Create editable medial-axis-derived LED guide centrelines inside selected filled letter geometry, then populate modules using controlled maximum centre spacing. Public LED Wizard 8 material is workflow reference only; do not copy proprietary code, assets, databases, or supplier data.

## Architecture and engineering constraints
- Preserve `index.html` panel markup, `js/main.js` orchestration, `jsx/hostscript.jsx` ExtendScript bridge, and serialized FIFO `evalScript` calls.
- Host JSX must remain ECMAScript 3 compatible. Heavy geometry belongs in a separate UMD/CommonJS pure JavaScript module shared by CEP and Node tests.
- Preserve `document.scaleFactor`, physical-unit conversion, source coordinates, originals, selection where practical, and unrelated tools.
- Never use repeated inward offsets as a centreline implementation.
- No recurring host polling or per-cell/per-module host calls. Batch extraction and drawing; keep raster work in the panel and yield between components.
- Do not push or merge. Do not publish or create a PR unless explicitly authorized by a higher-priority instruction.

## Research and documentation
Review the public product/workflow references at https://ledwizard8.com/, https://docs.ledwizard8.com/doc/lw8/6.2, https://docs.ledwizard8.com/doc/lw8/6.3, https://docs.ledwizard8.com/doc/lw8/7.1.2, and https://www.youtube.com/watch?v=eHTNjMsbKhk. Record sources, independent design, contracts, even-odd fill, coordinates, algorithms, thresholds, metadata schema, performance limits, tests, and limitations in `docs/LED_CENTRELINE.md`. Do not represent example settings as supplier-approved or photometric guidance.

## Required workflow and UI
Add a top-level **LEDs** tab. Select filled live text/path/compound/group/multiple objects without modifying originals; create first-class editable guide paths; populate/repopulate from current guides; optionally wire and report stats; safely replace only owned generated results.

Provide editable local generic module profiles (brand, series, code/name, body width/height, wattage, voltage, maximum centre spacing, wire reach, series limit), with save/edit/duplicate/delete. Provide can depth, layout mode (Auto, forced 1/2/3, Guides only), max rows, maximum spacing, body-edge clearance, inter-row behavior, transition preference, rotate, draw guides/modules/wires/stats, replace previous, and bounded Advanced geometry controls. Stable unique IDs must persist through existing settings.

Actions: Create Layout, Create Guide Paths Only, Repopulate From Guides, Clear Generated Layout, progress, cancellation between components, and useful warnings. Output layers: `LED Guide Paths`, `LED Modules`, `LED Wiring`, `LED Layout Stats`. Use stable names and versioned `note` metadata. Preserve edited guides during repopulation.

## Geometry and placement
- Duplicate live text to temporary work art and outline only the duplicate; clean up after errors. Support filled PathItem, CompoundPathItem counters, groups, multiple selections, and disconnected components.
- Extract document-coordinate cubic paths adaptively with winding/compound/source data in few calls. Rasterize even-odd masks, compute a bounded Euclidean distance field, erode by module footprint/clearance, skeletonize a genuine medial axis, prune only short spurs, trace graph polylines, smooth/simplify without crossing fill/counters, and map back to document coordinates.
- Target roughly 200k–300k cells per component with adaptive resolution and an accuracy floor. Guard malformed/enormous inputs.
- Auto rows use local distance-field width. If reliable multi-row planning is incomplete, explicitly ship correct single-row only and disable incomplete modes.
- Maximum spacing is centre-to-centre: normally `n=ceil(L/S)` intervals and spacing `L/n<=S`. Apply endpoint inset; validate full oriented rectangles against fill/counters; deduplicate junctions; omit non-fitting modules without implicit clearance relaxation.
- Deterministically split series, draw optional wiring, and flag gaps exceeding wire reach. Do not imply electrical or code approval.

## Output, statistics, and tests
Output exact source coordinates, visible/unlocked/editable subgroups by source/run. Metadata identifies owner, source, settings, guide/run and schema. Report profile, modules/runs, min/max/average spacing, series/splits, estimated user-profile wattage, omissions, unreachable gaps, and schema.

Add deterministic pure tests and fixtures spanning rectangles, rounded/curved/serif/branched/counter/disconnected/variable-width/narrow-neck shapes; spacing, endpoint insets, footprint containment, counter exclusion, pruning/topology, row fit/transitions, series/wire, repeatability, scale and cell budgets. Add static/integration tests for unique IDs, FIFO host calls, entrypoints, source preservation, layer names/metadata, and settings. Run the full test/lint/build/package workflow.

## Acceptance and reporting
Genuine editable centrelines, preserved counters/topology, footprints inside fill, spacing never above max, correct source coordinates/scale, editable-guide repopulation, responsive cancellation, clear errors, unchanged Lightbox behavior, and all tests passing are required. Perform the documented Illustrator smoke test when Illustrator is available; otherwise state that plainly. Final report must lead with outcome and include approach, files, exact tests, manual checks, performance, limitations/next step, branch and commit. Do not claim untested Illustrator behavior as verified.
