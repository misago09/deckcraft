# SlideWarp navigation work in the DeckCraft fork

This work belongs to `misago09/deckcraft`, forked from `storytold/deckcraft` at
`1dc69720db64323370dda061226becb09ad79fd4`. The existing SlideWarp application,
repository, installation and settings are outside this change.

## Acceptance order

1. Check PPTX display and save/reopen compatibility before navigation changes.
2. Only after that gate passes: middle-button pan, inertia and cursor-centered zoom.
3. Then the entire-workspace minimap, sharing the editor's viewport state.
4. Then Slide Space, built on the existing slide sorter and commands.

An import succeeding, or a self-generated file round-tripping, does not prove
PowerPoint fidelity. Keep failures and untested cases visible. Do not label the
compatibility gate passed while a save changes meaningful content or appearance.

## Isolation and evidence

- Work on `feature/slidewarp-navigation`; preserve the fork's upstream baseline.
- Use generated fixtures and local copies. Do not upload user presentations.
- Put generated PPTX, images and raw logs in ignored `plan/` folders. PowerPoint
  reference images stay in `plan/powerpoint/screenshots/` and are never published.
- Prefer the headless CLI and tests. The PowerPoint reference tool refuses to run
  while an existing PowerPoint process is present, opens its fixture read-only
  without a window, closes only its own presentation, and preserves a session if
  another document appeared during the check.
- The optional local Rust wrapper keeps toolchain/cache changes in a sibling
  directory. It changes environment variables only in its invoking process.
- Keep license and attribution notices. Select our own branding before distributing
  modified application binaries; ArtCraft marks are not part of the code license.

## Background commands

On a configured Rust machine, use the ordinary Cargo commands. On this isolated
Windows checkout, invoke the helper in a separate PowerShell process:

```powershell
./tools/Invoke-LocalRust.ps1 -CargoArguments @('xtask', 'ci')
./tools/Invoke-LocalRust.ps1 -CargoArguments @('test', '--locked', '-p', 'deckcraft-pptx')
./tools/Invoke-LocalRust.ps1 -CargoArguments @('build', '--locked', '-p', 'deckcraft-cli')
python -m unittest discover -s tools -p test_pptx_compatibility.py
python tools/pptx_compatibility.py --cli target/debug/deckcraft-cli.exe --out plan/compatibility/run-01
```

The Python harness requires Python 3.9+ and Pillow. It authors six OOXML slides independently of
DeckCraft's writer: Latin/Hangul text, rotated geometry, a transformed group,
gradient, table, cropped image and an authored comment. It checks text, shape
transforms, page size, media bytes, relationship targets, XML namespaces, a
second save, and before/after render pixels.
It deliberately calls its result **save stability**, not PowerPoint parity.

For the PowerPoint reference, use a new private output directory and no existing
PowerPoint session:

```powershell
./tools/Export-PowerPointReference.ps1 -InputPptx plan/compatibility/run-01/authored-core.pptx -OutputDirectory "$PWD/plan/powerpoint/screenshots/reference-01"
```

## Baseline notes (2026-10-16)

- Rust 1.95.0 x64 MSVC, matching the declared minimum version.
- `cargo xtask ci`: formatting and PDF dependency boundary passed; stopped at the
  pre-existing `manual_range_contains` Clippy warning in
  `crates/model/src/trendline.rs:468`. Subsequent CI phases were not run by that command.
- The shared `storytold/craftrules` repository referenced by upstream instructions
  returned HTTP 404 through both public web access and the authenticated GitHub API.
  The checked-in instructions and feature sources remain available.
- The PPTX suite passed 158 tests; the CLI built successfully. Asset attribution
  (47 files) and dependency layering (28 crates) checks passed independently.
- The new Python harness's three tests and PowerShell syntax checks passed.
- The first full workspace run stopped with seven loopback-socket permission
  errors and two temporary-file access errors in the desktop test binary. A run
  with local socket access and a checkout-local temporary directory, and the WASM
  checks, are still pending. Do not treat the full workspace/CI as passed.
- Compatibility **did not pass**: saving introduces inherited text styles and
  emits malformed modern-comment XML. See the
  [measurements and reproductions](PPTX-COMPATIBILITY-2026-10-16.md).
- A private 166-slide deck rendered identically before/after saving within
  DeckCraft, but also reproduced the comment XML defect. PowerPoint reference
  comparison awaits the user's session closure. Navigation work has not begun.

## Revision history

| Date | Change |
|---|---|
| 2026-10-16 | Isolated fork and compatibility-first workflow established. |
| 2026-10-16 | Authored and private-copy compatibility checks found two save blockers; preserved baseline failures and local evidence. |
