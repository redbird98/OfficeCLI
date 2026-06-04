# Paragraph Formatting Showcase

Exercises the docx **paragraph** property surface. Three files:

- **paragraph-formatting.sh** — builds the document with `officecli`.
- **paragraph-formatting.docx** — generated output.
- **paragraph-formatting.md** — this file.

## Regenerate

```bash
cd examples/word
bash paragraph-formatting.sh
# → paragraph-formatting.docx
```

## Sections

| Section | Properties |
|---|---|
| Alignment | `align=left\|center\|right\|both` |
| Indentation | `indent` (left), `rightIndent`, `firstLineIndent`, `hangingIndent` |
| Spacing | `spaceBefore`, `spaceAfter`, `lineSpacing` |
| Pagination flags | `keepNext`, `keepLines`, `widowControl` |
| Paragraph-level run formatting | `bold`, `italic`, `color`, `size`, `highlight` (applied to every run) |
| Shading | `shading.fill` |
| Paragraph-mark formatting | full `markRPr.*` (bold/italic/strike/underline/size/color/highlight/font.latin/ea/cs) |
| Outline level | `outlineLvl` |
| Run formatting (paragraph-wide) | `strike`, `underline`, `underline.color`, `bold.cs`, `italic.cs`, `size.cs` |
| Spacing/pagination extras | `contextualSpacing`, `lineRule`, `pageBreakBefore`, `wordWrap` |
| Chars-based indent | `firstLineChars`, `hangingChars` |
| Fonts | `font`, `font.latin/ea/cs`, theme fonts (`font.*Theme`), `direction` |
| Styles | `style` (paragraph), `rStyle` (character) |
| Shading variants | `shd`, `shading.val`, `shading.fill`, `shading.color` |
| Tab stops | `tabs` |
| Text frame | `framePr.w/.h/.wrap/.hAnchor/.vAnchor/.hSpace/.vSpace` (floating frame / drop-cap) |
| Paragraph borders | `border=single` / `border=style;size;color` (whole-box `pBdr`) |
| Vertical text alignment | `textAlignment=top\|center\|baseline\|bottom` |
| EastAsian typography | `kinsoku`, `autoSpaceDE`, `autoSpaceDN`, `overflowPunct`, `topLinePunct` |
| Line / indent flags | `suppressLineNumbers`, `suppressAutoHyphens`, `mirrorIndents`, `adjustRightInd`, `snapToGrid` |
| Web / textbox hints | `divId`, `textboxTightWrap` |
| List numbering | `listStyle`, `start`, `numId`, `numLevel` |

This trio exercises the full settable paragraph property surface — the 60
schema-declared paragraph keys **plus** the handler-supported `framePr.*` text
frame, **plus** long-tail OOXML pPr children handled by the generic
typed-attribute fallback (`pBdr` borders, `textAlignment`, the EastAsian
typography toggles, the line/indent flags, and `divId` / `textboxTightWrap`).
All of them round-trip through `add` → `get`.

## Two kinds of "bold" on a paragraph

- **`bold`** applies to every run in the paragraph (and reads back as `bold`).
- **`markRPr.bold`** formats only the paragraph mark (the ¶ pilcrow) — distinct
  from run text, used so appended runs inherit the mark's formatting. They are
  independent: setting one does not surface as the other.

```bash
officecli set file.docx /body/p[1] --prop bold=true          # all runs bold
officecli set file.docx /body/p[1] --prop markRPr.bold=true   # ¶ mark only
```

> `shading.fill` alone now produces schema-valid output — the writer defaults
> the required `w:shd/@val` to `clear`. Pair with `shading.val` for a pattern
> shade.

> Decomposed color keys accept a leading `#` (e.g. `underline.color=#FF0000`,
> `shading.fill=#D9D9D9`) — the dotted-attribute writer strips it, matching the
> curated setters.
