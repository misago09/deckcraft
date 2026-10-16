# PPTX compatibility gate for the SlideWarp fork

**Status: not passed, including PowerPoint reopen and rendering checks.** Tested upstream commit
`1dc69720db64323370dda061226becb09ad79fd4` on Windows x64 on 2026-10-16.
No product code or existing SlideWarp installation was changed. Navigation
features remain conditional on the compatibility gate.

## What was actually tested

- The existing `deckcraft-pptx` suite: **158 passed**.
- An independently authored six-slide OOXML fixture: mixed Latin/Hangul text,
  rotated shapes, a transformed group, gradient, table, cropped diagnostic image,
  and an authored legacy comment. No Office templates or assets were used.
- A private 166-slide presentation, copied locally with the owner's permission.
  Its source SHA-256 was unchanged after testing. No file, image, text, author
  identity, path, or raw report from that document is published here.
- Headless original → PPTX save → second PPTX save, plus original and first-save
  renders through the same renderer used by DeckCraft's editor.
- Actual PowerPoint ordinary read-only open, PNG export, save-copy and reopen,
  with the user's session closed. Reference exports used no visible document
  windows. Installed PowerPoint was 16.0.20430.20146, x64; the COM application
  reported version 16.0. No repair-open option was requested.
- XML namespace well-formedness, relationship targets, source media hashes, page
  dimensions, shape/text/transform signatures and rendered pixels.

The binary was a development build with optimization and debug information,
Rust 1.95.0 x64 MSVC, on Windows Professional 26H2 build 26300.9711 (x64,
as reported locally). Fonts came from `storytold/craft-fonts` commit
`43913056ce82a3fdad0ef69f6b88d1dc922a8104`, explicitly selected for this build;
no host font installation was performed. This is not a release performance
benchmark. The PowerPoint comparisons below demonstrate that visual parity has
not been achieved on these inputs.

## Confirmed blockers

### F-PPTX-01: saving an imported deck inserts inherited text styles

The authored fixture intentionally has empty master text styles and no
presentation default text style. The reader preserves that state, but the
writer substitutes new defaults in `crates/pptx/src/write/mod.rs`, in
`master_xml` and the presentation's `defaultTextStyle` output. Those defaults
include kerning and paragraph settings that the source did not contain.

At 720 × 405 pixels, comparing original and saved renders with a per-channel
difference threshold of 8 produced:

| Slide content | Changed pixels | Fraction |
|---|---:|---:|
| Latin/Hangul text | 2,221 | 0.7617% |
| Rotated geometry | 0 | 0% |
| Transformed group with text | 247 | 0.0847% |
| Gradient | 0 | 0% |
| Table text | 385 | 0.1320% |
| Cropped image | 0 | 0% |

Text content, shape coordinates and original image bytes were retained. The
failure is a change in rendered text placement/spacing, not missing text.
Rendering each *same* original and saved file again gave zero pixel differences,
so this was not renderer nondeterminism. An experiment on a private output copy
removed only the inserted presentation/master text styles: all six images then
matched the original with zero changed pixels. That diagnostic edit is **not a
product fix or a passing compatibility result**.

The later PowerPoint comparison narrows this finding: after only the comment
namespace was corrected in a diagnostic copy, PowerPoint rendered authored
slides 1 and 3 identically to its original input. Only the table slide changed
(107 pixels above the threshold). Thus the three earlier differences are
**DeckCraft editor/save stability failures**, not proof that all three change
when displayed in PowerPoint. A correction must account for PowerPoint's
effective inherited defaults; simply stripping defaults is not established as
the right product fix.

### F-PPTX-02: exported modern comments contain undeclared `a:` prefixes

`modern_comments_xml` declares `xmlns:p188` on `p188:cmLst`, but `tx_body`
emits `a:bodyPr`, `a:p`, `a:r`, `a:rPr` and `a:t` without an `xmlns:a` in scope.
A namespace-aware XML parser rejects
`ppt/comments/modernComment1.xml` with `unbound prefix`.

Both the private deck and the independently authored comment fixture reproduce
this after the first and second saves. Source XML is well formed. Actual
PowerPoint ordinary open rejects both first-save outputs with HRESULT
`0x80070570` (file or directory is corrupted and unreadable).

For causal isolation, separate private copies added only
`xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` to the modern
comment root; all other ZIP part payloads were kept unchanged. Both copies then
opened and exported all slides in PowerPoint. Sham controls rebuilt both ZIPs
with the same tooling and no XML changes; all part payloads matched the saved
inputs, and both still failed with `0x80070570`. This identifies the missing
namespace as the cause of the observed reopen rejection on these inputs. The
product writer remains unchanged; the diagnostic files are not a passing build.

### F-PPTX-03: saved text layout changes in PowerPoint

After the namespace-only correction allowed the private saved deck to open,
**128 of 166** slides differed from the original when both were rendered by
PowerPoint at 780 × 540. The median changed-pixel fraction over all 166 slides
was **0.4941%**, with a maximum of **9.3647%**. The threshold remains a maximum
RGB channel difference greater than 8; files and image sizes all matched.

Private visual inspection confirmed changed word wrapping in table cells and
callout text, words breaking in the middle, and one table extending beyond the
slide's lower boundary. These are not just cross-renderer antialiasing differences:
both sides used the same PowerPoint renderer. PowerPoint's own save/reopen control
had zero changed pixels on every slide.

The exact formatting-loss cause is not isolated. A further private experiment
restored only the source presentation's missing `a:defPPr` to the namespace-fixed
file; all 166 comparison results were unchanged. That single omission alone
does not explain the observed layout failure. No product code was changed for
either experiment.

## PowerPoint comparison matrix

| Input/path | Ordinary open | Render/reopen result |
|---|---|---|
| Authored original, 6 slides | Pass | PowerPoint save-copy → reopen: 6/6 identical |
| Private original copy, 166 slides | Pass | PowerPoint save-copy → reopen: 166/166 identical |
| DeckCraft first save, authored | Fail, `0x80070570` | No reference images; not counted as a display pass |
| DeckCraft first save, private | Fail, `0x80070570` | No reference images; not counted as a display pass |
| Unchanged-part ZIP repack controls, both decks | Fail, `0x80070570` | Repacking alone does not resolve rejection |
| Authored save + namespace-only diagnostic | Pass | 1/6 slides changed; maximum 0.0367% |
| Private save + namespace-only diagnostic | Pass | 128/166 slides changed; maximum 9.3647% |

### F-PPTX-04: rendering of imported originals

Comparing *original* input rendered by DeckCraft against original input rendered
by PowerPoint also found differences on all six authored and all 166 private
slides. For the private deck, median/maximum changed-pixel fractions were
11.7501%/57.2331%. These pixel counts alone are not a fidelity verdict because
different renderers and font substitutions can change antialiasing. Visual
inspection additionally confirmed table-background differences, a Hangul bold
weight difference in the authored fixture, and a different gradient progression
on the authored gradient slide. Those need separate rendering/inheritance work.

## Private-deck observations

All 166 slides rendered before and after saving. Every compared render had zero
changed pixels; text, shape kinds and transforms were unchanged under the
harness's structural signature. All original media hashes survived, and slide
width and height were retained.

The expanded harness passed 174 of 177 checks on this copy. The three non-passing
checks were modern-comment XML well-formedness and the two exact inventories
described next; the failed result is retained, not converted into a passing gate.

Two strict byte/attribute inventories also differed and must be distinguished
from content loss:

- The writer moved an embedded binary into `ppt/media`, increasing that folder's
  item count. No original media hash was lost; OLE activation remains untested.
- The `p:sldSz` `type="A4"` attribute was omitted, while `cx` and `cy` were
  unchanged. The harness records exact metadata separately from dimensions.

The confirmed comment XML failure still blocks the gate, and the later PowerPoint
comparison establishes additional saved-layout failures. Identical DeckCraft
renders had hidden those differences: both before and after can share an import
omission. Animations, notes, comments' semantics and embedded-object activation
have not been fully exercised.

## Reproduction and evidence

Use [the fork workflow](SLIDEWARP-FORK.md) for build commands. With Pillow
installed:

```powershell
python -m unittest discover -s tools -p test_pptx_compatibility.py -v
python tools/pptx_compatibility.py --cli target/debug/deckcraft-cli.exe --out plan/compatibility/authored-check
```

The second command is expected to exit with failure on the baseline: 13 of 17
checks pass; XML well-formedness and three rendered slides fail. It preserves
both input and output plus logs. A local representative file may be supplied
with `--input`; output is restricted to the ignored `plan/` directory. Raw
reports can contain document text and are never suitable for publication.
Previously collected evidence was retained when the XML check was added.

Local evidence folders include `run-01` (initial fixture), `control-original`,
`control-saved`, the style-isolation diagnostic, `run-02` (authored comment
reproducer), and `user-01`/`user-02` (private-copy checks). None is tracked.
PowerPoint reference results, the six-way pixel comparison and the default
paragraph isolation result remain under the private `plan/powerpoint/screenshots/`
tree. Original and copied PPTX SHA-256 values were checked for preservation.

## Remaining validation and implementation order

1. Correct modern-comment namespace output in a separate compatibility change,
   with an independently generated regression fixture. Re-run ordinary PowerPoint
   open on the unmodified product output, not a patched diagnostic copy.
2. Isolate and correct the saved-layout changes and imported/inherited formatting
   differences. Preserve defaults for new documents; do not blindly strip defaults
   based only on DeckCraft self-render comparisons. Repeat the same PowerPoint
   controls and pixel/visual comparisons after each focused change.
3. Check inherited fonts, notes, comment semantics and embedded objects, plus
   animations and other untested behaviors. This harness is not a full OOXML
   schema validator and its shape signature does not cover every semantic property.
4. Only after the compatibility gate passes, implement pan/zoom, minimap and
   Slide Space in that order. No new navigation behavior is claimed here.

## Revision history

| Date | Change |
|---|---|
| 2026-10-16 | Baseline and independent/private-copy checks; two reproduced save blockers, no product changes. |
| 2026-10-16 | Actual PowerPoint comparisons: both saved inputs rejected; namespace-only isolation permits open but exposes saved-layout differences. Self-save controls stable. |
