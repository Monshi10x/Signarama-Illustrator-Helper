# Dimensions `PARM` investigation

## Executive summary

The Dimensions text path performed an outline conversion for every label solely to obtain bounds. If any operation from text creation through `createOutline()`, bounds reading, translation, or cleanup failed, the text and already-created line artwork remained. Combined operations then ignored the failed side's error and continued.

The fix measures the live text frame first, retains outline conversion only as a guarded fallback, validates every value passed to the affected DOM calls, normalizes `270` degrees to `-90`, creates each label inside its measurement group, and removes that group on failure. Each risky text operation now reports its stage and Illustrator error metadata. Combined operations stop and name the failed side.

## Facts established from repository history

* The risky `duplicate()` / `createOutline()` label measurement path dates to commit `878c6c3` (2026-02-07); it was not introduced during the reported week.
* No Dimensions code changed between 2026-07-31 and the starting commit `d7612f7`.
* Background host interactions did change during the reported window. Commit `94d13ce` (2026-08-06) removed recurring host work, and `ef03bb7` (2026-08-06) routed every remaining `evalScript` invocation, including host-script loading, through one FIFO queue.
* At the starting revision, the queue does not release its in-flight lock until the CEP callback has completed, and all direct `cs.evalScript` callers use that queue.
* The old multi-side implementation ignored error strings returned by `_dim_run()` and counted success using a regular expression.

## Hypotheses and regression window

The repository evidence supports two independent contributors: a long-standing non-atomic text-outline path explains the `PARM` and partial artwork, while increased recurring host activity explains why a latent failure could become more frequent. The likely frequency-regression window is the background-interaction history before the corrective commits `94d13ce..ef03bb7`. A single culprit commit cannot be proven without Illustrator runtime results, and the Dimensions implementation itself is not a recent regression.

No automated bisect was performed because this repository has no Illustrator automation environment or repeatable headless predicate. A manual bisect should compare the parent of the first background polling change with `94d13ce` and `ef03bb7` using the supplied diagnostic script.

## Exact reproduction procedure

1. Install/load the extension and `jsx/hostscript.jsx` in Illustrator.
2. Open `tests/illustrator/dimensions-parm-diagnostic.jsx` and set `repetitions` (default: 50).
3. Run it from Illustrator's script runner. It creates a disposable RGB document and rectangle.
4. The harness runs Top, Bottom, Left, and Right repeatedly, removes each generated Dimensions layer, and records the side, iteration, and exact reported stage on failure.
5. The disposable document is closed without saving and the prior document and surviving selection references are restored.
6. Separately exercise combined buttons and the full document/artwork/settings matrix below in the panel.

## Failure details

Reported Illustrator error: `1346458189 ('PARM')`.

The exact DOM statement for a real-world failure has **not been captured in this non-Illustrator environment**. Before this change, the broad error handling made it impossible to distinguish the statements. The retained diagnostics now identify `createTextFrame`, `setContents`, `resolveFont`, `setTextSize`, `setTextAttributes`, `rotateText`, `readLiveVisibleBounds`, `duplicateTextFallback`, `createOutlineFallback`, `readOutlineBoundsFallback`, or `translateText` with error number, line, file, side, scale, text, anchor, coordinates, and selection count. Runtime Illustrator verification is required before declaring one of these the observed failing operation.

## Why partial output remained

Horizontal and vertical routines created the group and all line/tick/arrow paths before creating the label. A thrown text error bypassed group cleanup. In a combined operation, `_dim_run()` returned an error string, but `atlas_dimensions_runMulti()` did not inspect it and proceeded to the next side. This left line-only groups and allowed an overall-looking result after a partial failure.

## Code changes

* Added finite-number validation and rotation normalization.
* Added stage-specific error details around risky text DOM operations.
* Prefer live `visibleBounds`; use duplicate/outlines only when live bounds are unavailable or malformed.
* Remove temporary outline/copy in `finally`.
* Create text directly inside an atomic measurement group and remove the whole group on failure.
* Restore captured selection in `finally` for selected-object dimension runs.
* Reject invalid host-side sizes and normalize invalid persisted panel payload values.
* Stop multi-side execution at the first error and report the failing side and completed count.
* Added Node regression tests and a non-destructive Illustrator diagnostic harness.

## Test matrix and results

### Automated in this environment

| Test | Result |
| --- | --- |
| Existing and new Node tests | 34/34 passed |
| `js/main.js` syntax | Passed |
| shared Dimensions logic syntax | Passed |
| host JSX syntax after removing the ExtendScript `#target` directive | Passed with Node's parser |
| release validation | Passed |

### Illustrator runtime matrix

Not run: Adobe Illustrator and CEP are unavailable in the execution environment. Therefore the before-fix rate, after-fix rate, affected Illustrator versions, RGB/CMYK and production-document results, large-canvas results, the 50-run combined-side results, and the exact observed failing DOM operation remain unverified. The harness covers 50 runs of each single side; combined buttons, Centre Text, line/replace, angles, area, all selected-artwork types, saved/unsaved files, isolation mode, and settings variants must be recorded by a tester with Illustrator.

## Reproduction rates

* **Before fix:** not measured in this environment; the user-reported failure is recurring and increasingly frequent.
* **After fix:** static/pure-logic tests pass, but no runtime rate is claimed. Illustrator verification is outstanding.

## Known limitations and remaining work

* Live Illustrator results are required to identify the exact failing stage and satisfy the 50-run acceptance matrix.
* `visibleBounds` may itself fail for a host/font combination; the guarded outline fallback remains for that case and reports its exact stage.
* Completed sides of a combined request are intentionally preserved if a later side fails; the returned message explicitly reports partial success. Per-measurement atomicity prevents artifacts for the failed side.
* The Node syntax check is not an ExtendScript semantic test. It removes only the `#target` line and does not reject the ES3-compatible constructs used by the host file.
