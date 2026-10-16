# PPTX compatibility gate for the SlideWarp fork

**Status: not passed.** Tested upstream commit
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
- XML namespace well-formedness, relationship targets, source media hashes, page
  dimensions, shape/text/transform signatures and rendered pixels.

The binary was a development build with optimization and debug information,
Rust 1.95.0 x64 MSVC, on Windows Professional 26H2 build 26300.9711 (x64,
as reported locally). Fonts came from `storytold/craft-fonts` commit
`43913056ce82a3fdad0ef69f6b88d1dc922a8104`, explicitly selected for this build;
no host font installation was performed. This is not a release performance
benchmark or a claim of visual parity with PowerPoint.

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

### F-PPTX-02: exported modern comments contain undeclared `a:` prefixes

`modern_comments_xml` declares `xmlns:p188` on `p188:cmLst`, but `tx_body`
emits `a:bodyPr`, `a:p`, `a:r`, `a:rPr` and `a:t` without an `xmlns:a` in scope.
A namespace-aware XML parser rejects
`ppt/comments/modernComment1.xml` with `unbound prefix`.

Both the private deck and the independently authored comment fixture reproduce
this after the first and second saves. Source XML is well formed. DeckCraft
reopening its own output does not establish that other applications can consume
the malformed part. PowerPoint's repair/open behavior is still untested.

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

The confirmed comment XML failure still blocks the gate. Identical DeckCraft
renders can share an import omission, so these results do not prove fidelity to
the original PowerPoint display, animations, notes, comments or embedded objects.

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

## Remaining validation and implementation order

1. Correct inherited-style preservation and modern-comment namespace output in
   a separate compatibility change, with generated regression fixtures. Preserve
   defaults for newly created documents; do not globally remove them.
2. Repeat the same save/render checks, then compare original and saved copies
   against PowerPoint. Its existing user-owned session is preserved; reference
   exports await the user's notice that their work is closed.
3. Check PowerPoint reopen/repair, inherited fonts, notes, comments and embedded
   objects. This harness is not a full OOXML schema validator and its shape
   signature alone does not cover every semantic property.
4. Only after the compatibility gate passes, implement pan/zoom, minimap and
   Slide Space in that order. No new navigation behavior is claimed here.

## Revision history

| Date | Change |
|---|---|
| 2026-10-16 | Baseline and independent/private-copy checks; two reproduced save blockers, no product changes. |
