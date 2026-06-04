# Run / Character Formatting Showcase

Exercises the docx **run** (character-level) property surface. Three files:

- **run-formatting.sh** — builds the document with `officecli`.
- **run-formatting.docx** — generated output.
- **run-formatting.md** — this file.

## Regenerate

```bash
cd examples/word
bash run-formatting.sh
# → run-formatting.docx
```

## Families demonstrated

| Family | Properties |
|---|---|
| Weight & style | `bold`, `italic` |
| Underline | `underline=single\|double\|thick\|dotted\|wave`, `underline.color` |
| Strikethrough | `strike` (single), `dstrike` (double) |
| Case | `caps`, `smallcaps` |
| Vertical align | `superscript`, `subscript` (set on explicit `--type run` children) |
| Color / size / highlight | `color`, `size`, `highlight` |
| Per-script fonts | `font.latin`, `font.eastAsia` (Latin + CJK in one run) |
| Text effects | `emboss`, `imprint`, `outline`, `shadow` |
| Character spacing | `charSpacing`, `position` |
| Language | `lang`, `lang.latin`, `lang.ea`, `lang.cs` (BCP-47 per script) |
| Complex script | `bold.cs`, `italic.cs`, `size.cs`, `rtl`, `direction` |
| Theme fonts | `font.asciiTheme`, `font.hAnsiTheme`, `font.eaTheme`, `font.csTheme` |
| Per-script fonts | `font` (all), `font.latin`, `font.ea`, `font.cs` |
| Run shading / hidden | `shading`, `vanish`, `noproof` |
| w14 text effects | `textFill`, `textOutline`, `w14glow`, `w14reflection`, `w14shadow` |
| Border / kerning / layout | `bdr` (text border), `kern` (kerning), `eastAsianLayout.vert`/`.combine`, `rStyle` (character style) |
| Emphasis & visibility | `em` (着重号 dot/underDot/circle), `effect` (legacy animation), `webHidden`, `fitText`, `snapToGrid`, `specVanish` |

This trio exercises the full settable run property surface — the 43
schema-declared run keys **plus** handler-supported keys that the schema does
not yet enumerate (`kern`, `bdr`, `eastAsianLayout.*`, run-level `rStyle`,
`position`, `underline.color`), **plus** long-tail OOXML run children handled
by the generic typed-attribute fallback (`em`, `effect`, `webHidden`,
`fitText`, `snapToGrid`, `specVanish`). All of them round-trip through
`add` → `get`.

## Mixed runs (super/subscript)

Most lines set run formatting on the paragraph's implicit run. For `E = mc²`
and `H₂O`, separate runs are appended so only part of the line is raised/lowered:

```bash
officecli add file.docx /body --type paragraph --prop text="E = mc"
officecli add file.docx "/body/p[last()]" --type run --prop text=2 --prop superscript=true
```

> The paragraph path `/body/p[last()]` must be quoted in the shell — `[` / `(`
> are shell metacharacters.

> `charSpacing` and `kern` are distinct docx run properties: `charSpacing`
> (w:spacing) adds a fixed gap between every character, while `kern` (w:kern)
> sets the *minimum font size* (in half-points; 28 = 14pt) above which Word
> applies pair kerning. Both round-trip on runs.

> w14 text effects (`textFill`, `textOutline`, `w14*`) now apply via
> `add paragraph --prop ...` too (routed to the implicit run), matching how
> `bold`/`color` already behaved.
