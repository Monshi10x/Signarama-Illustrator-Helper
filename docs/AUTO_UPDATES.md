# Automatic updates

## Repository audit and architecture

This project is a **CEP** Illustrator extension, not UXP or hybrid. The evidence is `CSXS/manifest.xml`, the HTML panel, `CSInterface.js`, ExtendScript in `jsx/hostscript.jsx`, and the CSXS 11 runtime declaration. The unchanged plugin ID is `com.signarama.helper`; the initial version is `1.0.0`. No UXP manifest, `.ccx` build, Adobe signing configuration, package manager metadata, release workflow, updater, or packaging script existed before this implementation.

The existing panel settings are written by ExtendScript to an OS user-data location and a development snapshot exists at `.local/panel-settings.json`. Update preferences are deliberately separate from both the extension and the existing panel settings. No Git remote is configured in this checkout, so repository visibility could not be established. The updater uses anonymous public GitHub Release access and contains no credential. If the repository is private, automatic API/download access is intentionally unavailable: publish assets through a protected company endpoint with short-lived URLs, or require the user to download the release in an authenticated browser. Never add a PAT to the extension.

## Design and behavior

The panel checks `https://api.github.com/repos/Monshi10x/Signarama-Illustrator-Helper/releases` asynchronously 1.5 seconds after startup and at most once per 24 hours. **Check for Updates** bypasses that interval. Stable is the default channel and ignores drafts and prereleases; beta allows prereleases. Releases without `update.json`, malformed manifests, disallowed download hosts, wrong plugin IDs/types, or non-newer semantic versions are ignored.

The Updates tab shows checking, current, available, downloading, ready, postponed, skipped, and safe error states. Available releases show installed/available versions, date, size, and the release-notes URL. Major releases are never installed silently. Automatic checking can be disabled. A dismissed version is not shown automatically again unless mandatory or superseded, while a manual check can show it.

The CEP panel downloads a ZIP into the external update data directory and verifies its exact size (when declared) and SHA-256 before enabling installation. It then starts a detached Node updater using an argument array, never a constructed shell command. The updater waits up to 30 minutes for Illustrator to be closed and never force-quits it. It validates every archive entry against absolute paths and traversal, extracts into staging, validates `CSXS/manifest.xml`, plugin ID, version, and `index.html`, and only then renames the current installation to a versioned backup. Directory renames provide atomic replacement where the filesystem permits. Any replacement or post-install validation error restores the backup. The user manually reopens Illustrator after installation.

Update logic is separated into:

- `semver.js`: semantic parsing and comparison.
- `manifest.js`: strict update metadata and release selection.
- `runtime.js`: GitHub checks, external preferences, downloads, hashing, and handoff.
- `archive.js`: archive path safety.
- `updater.js`: wait, stage, validate, back up, replace, log, and roll back.
- `ui.js`: panel state only.

## Paths and retained data

CEP installation paths vary by deployment. Common paths are `%APPDATA%\Adobe\CEP\extensions\Signarama-Illustrator-Helper` on Windows and `~/Library/Application Support/Adobe/CEP/extensions/Signarama-Illustrator-Helper` on macOS. The updater operates on the actual running extension directory rather than assuming either location.

Updater state is outside that replaceable directory:

| Data | Windows | macOS |
| --- | --- | --- |
| Root | `%APPDATA%\Signarama\Illustrator Helper` | `~/Library/Application Support/Signarama/Illustrator Helper` |
| Preferences | `update-preferences.json` under root | same |
| Downloads | `updates\downloads` | `updates/downloads` |
| Staging | `updates\staging` | `updates/staging` |
| Backups | `updates\backups\<version>` | `updates/backups/<version>` |
| Logs | `updates\logs\YYYY-MM-DD.jsonl` | `updates/logs/YYYY-MM-DD.jsonl` |

At most ten update log files are retained. Logs contain non-sensitive status and errors; query strings, auth headers, tokens, licence data, and full private signed URLs are never logged. Existing panel preferences, presets, and local data are not copied into the ZIP or installation swap, so they survive. Update preference schema 1 is merged idempotently with defaults when read.

## Versions, builds, and releases

`package.json` is the version source of truth. Before validation, CI synchronizes both CEP manifest attributes to it so a previously mismatched branch can recover instead of blocking the version job. `npm run lint` then verifies the synchronized values. Every validated push to `main` or `master` automatically increments the patch version, synchronizes both CEP manifest attributes, commits the result with `[skip ci]`, tags that commit, and publishes a stable GitHub Release containing the updater assets. The repository must allow GitHub Actions write access to the protected branch. For a manual version change, keep the files synchronized with:

```bash
npm version patch --no-git-tag-version
npm run version:sync
npm run lint
npm test
npm run build
npm run package
git add package.json package-lock.json CSXS/manifest.xml
git commit -m "Release v$(node -p "require('./package.json').version")"
git tag "v$(node -p "require('./package.json').version")"
git push origin HEAD
git push origin --tags
```

The automatic branch workflow and the manual tag workflow build, package, compute SHA-256, generate `update.json`, and attach the ZIP, manifest, and `SHA256SUMS` to a GitHub Release. The automatic workflow lets the release action create the tag only when it publishes the release, avoiding orphan version tags when packaging fails. It then queries the same unauthenticated GitHub API used by the extension and fails unless the release is publicly discoverable, non-draft, stable, and contains all updater assets. The extension discovers versions from these published releases—not from branch files or version-only commits. The repository itself must be public because the installed extension intentionally contains no GitHub credential. The workflows use only the repository-scoped `GITHUB_TOKEN`; no custom secret is required. Configure Actions with **Read and write permissions** for repository contents and releases. Protect `main`/`master` while allowing this workflow to push its version commit.

For a beta, set a SemVer prerelease in both version locations (for example `1.1.0-beta.1`), commit, and tag it exactly:

```bash
git tag v1.1.0-beta.1
git push origin HEAD v1.1.0-beta.1
```

The workflow marks tags containing `-` as GitHub prereleases and the generator assigns the beta channel. Only users explicitly selecting beta receive them. Return to stable by selecting Stable; no data migration is performed by this updater.

## Rollback and uninstall

Rollback is automatic when extraction, staging validation, replacement, or post-install validation fails. To restore manually, close Illustrator, move the current extension directory aside, and move the desired directory from `updates/backups/<version>` back to the original extension path. Do not merge the directories. Reopen Illustrator and verify the version. Backups are intentionally retained until an administrator removes them after a successful health check.

To uninstall, close Illustrator and remove the extension directory. Optionally remove the external Signarama root above to delete update preferences, packages, backups, and logs; leaving it preserves settings for reinstall.

## Security decisions and limitations

- Only HTTPS and an explicit GitHub host allowlist are accepted; TLS verification remains enabled.
- No GitHub or Adobe credentials are embedded. Public releases are the supported unattended source.
- SHA-256, package size, plugin ID, semantic version, required structure, and ZIP paths are checked before replacement.
- Downloaded content is never executed before verification. `eval` is not used by updater metadata code.
- The updater passes external values as process arguments. ZIP filenames are locally derived from validated SemVer.
- A verified backup must exist before the new directory is installed; failed swaps roll back rather than merge partial files.
- CEP cannot replace loaded files safely, so Illustrator must be closed and manually reopened. The updater will time out rather than force-quit.
- Windows extraction uses built-in PowerShell; macOS uses `/usr/bin/unzip`. Corporate policy or antivirus may block either helper or filesystem replacement, in which case the old installation remains or is rolled back.
- Adobe package signing did not exist in the repository. ZIP hashing provides integrity, but not publisher identity beyond GitHub/TLS. Add ZXP signing in CI if a trusted certificate becomes available.
- Automated tests exercise platform-independent logic and the Unix rollback path. Illustrator integration, Windows, macOS GUI behavior, permissions failures, antivirus behavior, offline mode, and restart must be manually tested on their target systems before claiming production support.

## Manual acceptance checklist

Before production, test offline/current/available checks, skip and remind-later, stable and beta selection, interrupted and bad-checksum downloads, corrupt/traversal archives, Illustrator left open, insufficient permissions, successful replacement, forced rollback, settings retention, and existing panel operations. Repeat on supported Windows and macOS versions with the actual Illustrator build.
