// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class WordBatchEmitter
{

    private static void EmitTable(WordHandler word, string sourcePath, int targetIndex,
                                  List<BatchItem> items, BodyEmitContext? ctx = null,
                                  string? parentTablePath = null,
                                  string containerPath = "/body",
                                  int depth = 0)
    {
        // CONSISTENCY(dos-hardening): nested-table emission recurses with no
        // structural bound; a crafted deeply-nested table would otherwise hang
        // (or overflow the stack) during `dump`. See DocumentLimits.
        OfficeCli.Core.DocumentLimits.EnsureDepth(depth);

        // BUG-R11A(BUG1): bump the document-order table ordinal BEFORE the
        // empty-table early-return so the count never desyncs from the
        // `(//w:tbl)[N]` selectors used by cell-SDT raw-sets. EmitTable recurses
        // in DFS document order, so this matches the target's //w:tbl indexing.
        int tableOrdinal = 0;
        if (ctx != null) tableOrdinal = ++ctx.TableOrdinalBox[0];

        var tableNode = word.Get(sourcePath);
        var rows = (tableNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "row")
            .ToList();
        if (rows.Count == 0) return;

        // Column count must cover the widest row including colspan effects.
        // Format["cols"] reflects gridCol; per-row effective width is
        // sum(colspan or 1) over each cell. Take the max so a first row
        // with merged cells (visible cell count < grid width) doesn't
        // truncate the table shape and break later `set tc[N]` rows.
        var rowEffectiveWidths = new List<int>(rows.Count);
        var rowCellNodes = new List<List<DocumentNode>>(rows.Count);
        var rowNodes = new List<DocumentNode>(rows.Count);
        foreach (var rowChild in rows)
        {
            var rowNode = word.Get(rowChild.Path);
            rowNodes.Add(rowNode);
            var cells = (rowNode.Children ?? new List<DocumentNode>())
                .Where(c => c.Type == "cell")
                .ToList();
            rowCellNodes.Add(cells);
            int width = 0;
            foreach (var cell in cells)
            {
                int span = 1;
                if (cell.Format.TryGetValue("colspan", out var sp) &&
                    int.TryParse(sp?.ToString(), out var n) && n > 0)
                {
                    span = n;
                }
                width += span;
            }
            rowEffectiveWidths.Add(width);
        }
        int colsFromRows = rowEffectiveWidths.Count > 0 ? rowEffectiveWidths.Max() : 0;
        int colsFromGrid = 0;
        if (tableNode.Format.TryGetValue("cols", out var gridColObj) &&
            int.TryParse(gridColObj?.ToString(), out var gridCols))
        {
            colsFromGrid = gridCols;
        }
        // Format["cols"] back-fills from first-row cell count when source has
        // no <w:tblGrid> at all, so it can't tell us "source had zero gridCol".
        // _gridCols is the unbiased count (Navigation emits 0 when TableGrid
        // is missing or empty). EmitTable uses this to drive the gridCols=0
        // opt-out on the dumped `add table`.
        int actualGridCols = colsFromGrid;
        if (tableNode.Format.TryGetValue("_gridCols", out var actualGridObj) &&
            int.TryParse(actualGridObj?.ToString(), out var ag))
        {
            actualGridCols = ag;
        }
        int cols = Math.Max(colsFromGrid, colsFromRows);
        if (cols == 0) return;

        var tableProps = FilterEmittableProps(tableNode.Format);
        // Strip the revision-marker surface keys so they don't ride on
        // `add table` — they're consumed by a follow-up EmitTrackChangeMarker
        // call below. Without this, AddTable's schema fallback would create
        // a phantom <w:tblPrChange> on the new table, and the follow-up
        // `set trackChange.author=...` would then trip the
        // "element already has a pending tblPrChange" guard.
        tableProps.Remove("tblPrChange.author");
        tableProps.Remove("tblPrChange.date");
        tableProps["rows"] = rows.Count.ToString();
        tableProps["cols"] = cols.ToString();
        // Source had no <w:tblGrid> or an empty one — cells (if any) carry
        // their own tcW, or the table is auto-fit. Without an explicit
        // `gridCols=0`, AddTable would seed `cols` default GridColumn entries
        // which ReadCellProps then back-fills as per-cell widths on the next
        // dump, producing N×M extra `set tc width=…` rows the source never
        // had (test.docx tbl[1]). Signal AddTable to leave tblGrid empty.
        if (actualGridCols == 0)
            tableProps["gridCols"] = "0";
        // Source had no <w:tblW> — surface a `skipTblW=true` user-facing
        // flag (mirrors `gridCols=0`). AddTable's default-tblW stamp
        // path defers to this when set, so replay won't grow a phantom
        // <w:tblW>. Skip when source had any explicit width (auto / dxa /
        // pct) — those round-trip through the existing `width=` key.
        bool sourceHadNoTblW = tableNode.Format.TryGetValue("_noTblW", out var noTblW)
            && noTblW is bool b && b;
        // BUG-DUMP-AUTOFITW: a table is autofit unless layout is explicitly
        // "fixed" (OOXML default, and what ReadTableProps emits when the
        // <w:tblLayout> element is absent). EmitTable passes this to
        // ExtractCellOnlyProps so tcW-less cells in an autofit table keep
        // their tcW-less state (no fabricated width) — only fixed-layout
        // tables and cells with real source tcW emit a width.
        bool tableIsAutofit = !(tableNode.Format.TryGetValue("layout", out var layoutObj)
            && layoutObj is string layoutStr
            && layoutStr.Equals("fixed", StringComparison.OrdinalIgnoreCase));
        if (sourceHadNoTblW && !tableProps.ContainsKey("width"))
            tableProps["skipTblW"] = "true";
        // Drop the internal-only markers from emitted props (BatchItem.Props
        // never carries them; only Navigation→EmitTable consumes them).
        tableProps.Remove("_gridCols");
        tableProps.Remove("_noTblW");
        // BUG-BORDER-PARTIAL: AddTable seeds all 6 default borders and overlays user
        // props on top, so a partial border spec (e.g. only border.top +
        // border.bottom for a banner-line table) replays as 6 single-borders.
        // If the source table emits only a subset of the 6 sides, prepend an
        // explicit `border=none` wipe so the visible result round-trips.
        // CONSISTENCY(border-default-overlay).
        //
        // The same fix applies to the zero-sides case: source tables with no
        // <w:tblBorders> at all (Word treats as no rules) used to replay as
        // 6 single-borders because EmitTable emitted no border prop and
        // AddTable's default-overlay won. The second dump then saw the
        // stamped borders and emitted six border.* props that the first
        // dump didn't — a 6× length asymmetry per affected table. Extend
        // the wipe to fire whenever no per-side / no-border-all key is
        // present in source's emit.
        {
            var sideKeys = new[] { "border.top", "border.bottom", "border.left",
                "border.right", "border.insideH", "border.insideV" };
            int presentSides = sideKeys.Count(s => tableProps.ContainsKey(s));
            bool hasBorderAll = tableProps.ContainsKey("border") || tableProps.ContainsKey("border.all");
            // BUG-STYLE-BORDER: the none-wipe (and AddTable's default-border
            // seed it counteracts) exist for tables that are GENUINELY
            // borderless (no inline <w:tblBorders>, no style). A table whose
            // borders come from a TABLE STYLE (a <w:tblStyle w:val="…"/> ref,
            // no inline borders) ALSO emits <6 per-side keys here — but it must
            // NOT get any inline override: writing all-none borders renders it
            // borderless, and AddTable's default single-border seed paints a
            // generic thin grid that masks the style's GridTable5Dark design.
            // For style-driven tables, instead tell AddTable to skip the
            // default-border seed entirely (skipDefaultBorders) so the style's
            // borders apply unmodified; any explicit inline border.* keys the
            // source DID carry still overlay on top. Genuinely borderless,
            // style-less tables keep the original none-wipe idempotency fix.
            // CONSISTENCY(border-default-overlay).
            bool hasTableStyleRef = tableProps.TryGetValue("style", out var styRef)
                && !string.IsNullOrEmpty(styRef);
            if (hasTableStyleRef)
            {
                // Suppress AddTable's generic 6-side single-border seed. The
                // table style supplies borders; explicit inline border.* keys
                // (if any) still apply via ApplyTableBorders after the seed is
                // skipped. No all-none wipe — that would defeat the style.
                tableProps["skipDefaultBorders"] = "true";
                // Still collapse 6 identical explicit per-side keys to the
                // compact form for idempotency (fall through to the else-if
                // below would be skipped otherwise).
                if (presentSides == 6 && !hasBorderAll)
                {
                    var firstS = tableProps[sideKeys[0]];
                    if (sideKeys.All(s => tableProps[s] == firstS))
                    {
                        foreach (var s in sideKeys) tableProps.Remove(s);
                        tableProps["border"] = firstS;
                    }
                }
            }
            else if (presentSides < 6 && !hasBorderAll)
            {
                // Use the canonical "style;size" form: ApplyTableBorders'
                // ParseBorderValue defaults size=4 for `none`, so writing
                // `none;4` matches what the round-trip produces (six explicit
                // <w:none w:sz="4"> elements collapsed by the all-same fold
                // below). Without the `;4`, the FIRST dump emits `border=none`
                // and the SECOND dump emits `border=none;4` — non-idempotent
                // value shape.
                tableProps["border"] = "none;4";
            }
            // Symmetric collapse: when all 6 sides carry the IDENTICAL folded
            // value (same style + sz + color + space), prefer the compact
            // `border=<v>` form so dump round-trips that started from
            // "no <w:tblBorders>" (whose first emit becomes `border=none`)
            // re-emit the same single key after replay rather than fanning
            // out to six explicit per-side rows. ApplyTableBorders interprets
            // `border=<v>` as "set all 6 sides to <v>", so the visible result
            // is identical either way.
            else if (presentSides == 6 && !hasBorderAll)
            {
                var first = tableProps[sideKeys[0]];
                if (sideKeys.All(s => tableProps[s] == first))
                {
                    foreach (var s in sideKeys) tableProps.Remove(s);
                    tableProps["border"] = first;
                }
            }
        }
        // Nested tables sit inside a parent table cell; AddTable accepts
        // /body/tbl[N]/tr[M]/tc[K] as a parent. Outer-level tables target
        // /body. parentTablePath, when set, is a cell target path
        // (/body/tbl[X]/tr[Y]/tc[Z]) that we emit nested tables under.
        var tableParentPath = parentTablePath ?? containerPath;

        // tblPrChange round-trip — D1 emit shape.
        //
        // OLD path was: (1) `add table` with all props (rows/cols/width/
        // borders/...) AND (2) a follow-up no-op `set` carrying only
        // `revision.author=…` to re-stamp the marker. The second step's
        // BeginTrackChangeIfRequested snapshotted the just-finalized tblPr
        // as the "before" state — but that's also the "after" state since
        // nothing changed between (1) and (2). Word's reviewing pane then
        // silently dropped the revision because before==after.
        //
        // NEW path: when the source carried a tblPrChange, split into
        //   (1) `add table` with structural-only keys (rows / cols /
        //       gridCols / skipTblW) so the seeded tblPr is bare,
        //   (2) `set table` with EVERY other prop the source had + the
        //       attribution. BeginTrackChangeIfRequested now snapshots the
        //       bare-tblPr from (1) as "before" and captures the props
        //       applied in (2) as the diff — Word's reviewing pane sees a
        //       real change and surfaces it as Alice's tracked edit.
        //
        // Snapshot is still over-attributed (the source might only have
        // changed `width`, but all 10 current props get marked as
        // "changed" because we can't recover which subset was the real
        // edit). Over-attribute is the lesser evil; current behavior
        // under-attributes to silent loss.
        var tblPrAuthor = TryStringFormat(tableNode.Format, "tblPrChange.author");
        var tblPrDate = TryStringFormat(tableNode.Format, "tblPrChange.date");
        bool hasTblPrChange = !string.IsNullOrEmpty(tblPrAuthor);

        Dictionary<string, string> tableAddProps;
        Dictionary<string, string>? tableSetProps = null;
        if (hasTblPrChange)
        {
            tableAddProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            tableSetProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var (k, v) in tableProps)
            {
                if (k is "rows" or "cols" or "gridCols")
                    tableAddProps[k] = v;
                else
                    tableSetProps[k] = v;
            }
            // Force AddTable to seed an empty tblPr (no default tblW, no
            // default 6-border block). Without these, AddTable's defaults
            // tend to coincide with the source's set values (auto-fit
            // tblW = grid sum; single-4 borders are the most common case),
            // making the follow-up set a no-op in OOXML terms — the
            // snapshot then equals current state and Word's reviewing
            // pane silently hides the tblPrChange. With the seed
            // suppressed, the snapshot captures the bare tblPr and the
            // set's props produce a real diff Word will surface.
            tableAddProps["skipTblW"] = "true";
            tableAddProps["skipDefaultBorders"] = "true";
            tableSetProps["revision.type"] = "format";
            tableSetProps["revision.author"] = tblPrAuthor!;
            if (!string.IsNullOrEmpty(tblPrDate))
                tableSetProps["revision.date"] = tblPrDate!;
        }
        else
        {
            tableAddProps = tableProps;
        }

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = tableParentPath,
            Type = "table",
            Props = tableAddProps
        });

        // BUG-R3 (nested-table multi-instance): a single cell may hold MORE
        // THAN ONE nested table stacked back-to-back (LibreOffice export
        // splits a logical table across several <w:tbl> siblings). The old
        // `tbl[1]` target hardcoded the FIRST nested table for every nested
        // emit, so the 2nd..Nth tables' per-cell `set tc[K]` ops resolved
        // against table #1 (wrong rows/cols) and produced thousands of
        // "Path not found …/tbl[1]/tr[N]/tc[K]". `add table` on a cell parent
        // appends (tbl[1], tbl[2], …, verified), so the just-added nested
        // table is always the cell's LAST table — mirror the outer-table
        // `tbl[last()]` convention so each nested table addresses itself.
        var tablePath = parentTablePath != null
            ? $"{parentTablePath}/tbl[last()]"
            : $"{containerPath}/tbl[last()]";

        if (tableSetProps != null && tableSetProps.Count > 0)
        {
            items.Add(new BatchItem
            {
                Command = "set",
                Path = tablePath,
                Props = tableSetProps
            });
        }
        for (int r = 0; r < rows.Count; r++)
        {
            // Emit row-level properties (header / height / height.rule) as a
            // `set` on the row path — `add table` only seeds rows, it doesn't
            // surface per-row props (BUG-ROWPROPS). Without this, `dump→batch`
            // silently strips repeating-header rows and explicit row heights.
            var rowNode = rowNodes[r];
            // trPrChange D1 round-trip: fold the attribution into the
            // row's prop-set step so BeginTrackChangeIfRequested snapshots
            // the bare-trPr "before" vs the props-applied "after". If the
            // source had a trPrChange but no row-only props, still emit
            // the bare-attribution set so the marker exists (over-attributed
            // tail case, same lesser-evil as the tblPrChange edge case).
            var rowProps = ExtractRowOnlyProps(rowNode.Format);
            bool rowHadRevision = FoldRevisionIntoProps(rowNode.Format, "trPrChange", rowProps);
            if (rowProps.Count > 0 || rowHadRevision)
            {
                items.Add(new BatchItem
                {
                    Command = "set",
                    Path = $"{tablePath}/tr[{r + 1}]",
                    Props = rowProps
                });
            }
            var cells = rowCellNodes[r];
            for (int c = 0; c < cells.Count; c++)
            {
                var cellNode = word.Get(cells[c].Path);
                var cellTargetPath = $"{tablePath}/tr[{r + 1}]/tc[{c + 1}]";

                // BUG-DUMP-R26-7: the global-ordinal XPath of THIS cell, used by
                // EmitParagraph's inline raw-set fallbacks (rich field result,
                // nested SDT, VML textbox) to append verbatim content into the
                // correct cell instead of falling back to the lossy typed emit.
                // Only carried for /document-hosted (body) tables — the same
                // restriction the cell-SDT raw-set's rawPart uses; header/footer
                // -hosted cell raw-set targeting is out of scope this round.
                string? cellRawXPath = containerPath == "/body"
                    ? $"(//w:tbl)[{tableOrdinal}]/w:tr[{r + 1}]/w:tc[{c + 1}]"
                    : null;

                // Cell-level tcPr properties (fill, valign, width, borders,
                // padding, colspan, …) are surfaced on cellNode.Format but
                // were previously dropped — only the inner paragraph was
                // emitted. Push them via a `set` on the cell path before
                // the paragraph emits so cell shading / merges / widths
                // round-trip. Skip keys that EmitParagraph will re-apply
                // to the first paragraph (align/direction/run leak-throughs)
                // to avoid double-application.
                // tcPrChange D1 round-trip: fold attribution into the
                // cell's prop-set step (mirrors the row branch above).
                var cellProps = ExtractCellOnlyProps(cellNode.Format, tableIsAutofit);
                bool cellHadRevision = FoldRevisionIntoProps(cellNode.Format, "tcPrChange", cellProps);
                if (cellProps.Count > 0 || cellHadRevision)
                {
                    // CONSISTENCY(tblgrid-preserve): tcW values in the source
                    // are allowed to disagree with the gridCol widths (Word
                    // renders by tcW; tblGrid is a layout hint). Suppress
                    // Set.tc's tblGrid-sync side effect so AddTable's
                    // authoritative colWidths survives subsequent per-cell
                    // width sets.
                    if (cellProps.ContainsKey("width"))
                        cellProps["skipGridSync"] = "true";
                    items.Add(new BatchItem
                    {
                        Command = "set",
                        Path = cellTargetPath,
                        Props = cellProps
                    });
                }

                // Each cell carries auto-generated paragraphs (Add table seeds
                // one empty paragraph per cell). Update the first one in place
                // and append further paragraphs as fresh adds. Nested tables
                // and paragraphs are emitted in document order so footnote/
                // chart cursors (carried in ctx) advance correctly through
                // the table cell content. Without ctx threading, body-level
                // footnote/chart references after a table would resolve
                // against the wrong note text.
                var cellChildren = cellNode.Children ?? new List<DocumentNode>();
                int cellParaIdx = 0;
                int nestedTblIdx = 0;
                bool firstParaSeen = false;

                // BUG-DUMP-NESTED-TBL-TRAILING: OOXML requires every cell to
                // end with a paragraph (not a table). When a cell would
                // otherwise end with a table, the SDK auto-inserts a trailing
                // paragraph on save — so the cell's LAST paragraph following
                // a nested table is structurally auto-present on the target
                // side too, regardless of whether source's iteration already
                // used its autoPresent slot on a leading paragraph. Without
                // this, source [table, p] dumps `set p[last()]`
                // (autoPresent=true) but target [auto-p, table, p] re-dumps
                // `set p[1]` + `add p` and diverges by one row.
                int trailingAutoP = -1;
                for (int k = cellChildren.Count - 1; k >= 0; k--)
                {
                    var ct = cellChildren[k].Type;
                    if (ct != "paragraph" && ct != "p") continue;
                    if (k > 0 && cellChildren[k - 1].Type == "table")
                        trailingAutoP = k;
                    break;
                }

                for (int k = 0; k < cellChildren.Count; k++)
                {
                    var cc = cellChildren[k];
                    if (cc.Type == "paragraph" || cc.Type == "p")
                    {
                        cellParaIdx++;
                        // BUG-R4 (DBF-R4-02): a display equation (<m:oMathPara>)
                        // inside a cell surfaces here as a plain paragraph child
                        // whose Get returns an empty paragraph — EmitParagraph
                        // would emit `set p[N]` with no content and the formula
                        // would be lost. Mirror the body walker's typed routing:
                        // detect the oMathPara-wrapper and emit `add equation`
                        // targeting the cell paragraph instead. `add equation`
                        // (display) on an existing cell paragraph appends the
                        // m:oMathPara into it, reproducing the wrapper shape.
                        var cellEq = word.TryGetDisplayEquationAtParagraph(cc.Path);
                        if (cellEq != null)
                        {
                            bool eqIsTrailingAutoP = k == trailingAutoP;
                            // First cell paragraph (or the SDK auto-trailing one)
                            // reuses an auto-present seeded paragraph; otherwise
                            // create a fresh host paragraph for the equation.
                            if (firstParaSeen && !eqIsTrailingAutoP)
                                items.Add(new BatchItem
                                {
                                    Command = "add",
                                    Parent = cellTargetPath,
                                    Type = "paragraph",
                                });
                            EmitCellDisplayEquation(cellEq,
                                $"{cellTargetPath}/p[{cellParaIdx}]", items);
                            firstParaSeen = true;
                            continue;
                        }
                        bool isTrailingAutoP = k == trailingAutoP;
                        // BUG-DUMP-R26-7: publish THIS cell's raw-set XPath so the
                        // paragraph's inline raw-set fallbacks target the right
                        // cell. Re-set per paragraph because a preceding nested
                        // table recursion overwrote the box with its own cell.
                        if (ctx != null) ctx.CurrentCellXPathBox[0] = cellRawXPath;
                        EmitParagraph(word, cc.Path, cellTargetPath, cellParaIdx, items,
                                      autoPresent: !firstParaSeen || isTrailingAutoP, ctx);
                        firstParaSeen = true;
                    }
                    else if (cc.Type == "table")
                    {
                        nestedTblIdx++;
                        EmitTable(word, cc.Path, nestedTblIdx, items, ctx,
                                  parentTablePath: cellTargetPath, depth: depth + 1);
                    }
                    else if (cc.Type == "sdt" && ctx != null)
                    {
                        // BUG-R11A(BUG1): a block-level <w:sdt> that is a direct
                        // child of this cell. Previously the cell walk recognised
                        // only paragraphs and nested tables, so the SDT (and its
                        // inner content) was dropped on dump. Emit it via the
                        // shared cell-SDT helper (typed `add sdt` for text-shaped
                        // controls, raw-set verbatim for rich block content).
                        // The raw-set xpath resolves to THIS cell by the table's
                        // document-order ordinal plus the current row/cell index.
                        var rawPart = containerPath == "/body" ? "/document" : containerPath;
                        var cellXPath = $"(//w:tbl)[{tableOrdinal}]/w:tr[{r + 1}]/w:tc[{c + 1}]";
                        EmitCellSdt(word, cc.Path, cellTargetPath, cellXPath, rawPart,
                                    cellHasContent: firstParaSeen, items, ctx);
                    }
                }

                // BUG-DUMP-R2-NESTED-LEAD: a cell whose FIRST source child is a
                // table has no leading source paragraph to reuse the empty
                // paragraph `add table` auto-seeds, so that seed survives as a
                // phantom blank line above the nested table (source [table, p] →
                // replay [p, table, p]). The trailing paragraph already landed
                // on the SDK's auto-trailing paragraph via trailingAutoP, so the
                // leading seed (cell's p[1]) is unconsumed — remove it. Validates
                // clean either way; this restores source structure.
                if (cellChildren.Count > 0 && cellChildren[0].Type == "table")
                {
                    items.Add(new BatchItem
                    {
                        Command = "remove",
                        Path = $"{cellTargetPath}/p[1]",
                    });
                }
            }
            // Trim trailing cells when source row is underfilled (sum of
            // source spans < gridCols). AddTable seeds `cols` cells per row;
            // `set tc[i] colspan=N` removes excess cells DOWN TO gridCols but
            // also PADS UP TO gridCols when the post-set total is short — so
            // a source row like [colspan=3] in a 4-col grid lands at 2 cells
            // post-replay (1 spanning + 1 pad). Source-shape preservation
            // demands removing (gridCols - sum_of_source_spans) trailing
            // cells AFTER all per-cell sets. The remove path is non-padding,
            // so the final cell count matches source. CONSISTENCY(table-row-
            // cell-count).
            int excessTrail = cols - rowEffectiveWidths[r];
            for (int e = 0; e < excessTrail; e++)
            {
                items.Add(new BatchItem
                {
                    Command = "remove",
                    Path = $"{tablePath}/tr[{r + 1}]/tc[last()]",
                });
            }
        }
        // BUG-DUMP-R26-7: clear the cell-XPath context once this table is fully
        // emitted so body/header/footer content AFTER the table (or a parent
        // cell's content after a nested table) doesn't inherit a stale cell
        // address. A parent cell re-publishes its own XPath before its next
        // paragraph (see the per-paragraph set above), so null here is safe.
        if (ctx != null) ctx.CurrentCellXPathBox[0] = null;
    }

    // BUG-R4 (DBF-R4-02): emit a typed `add equation` (display) targeting a cell
    // paragraph path. Mirrors TryEmitDisplayEquation (WordBatchEmitter.Paragraph.cs)
    // but for an arbitrary cell-paragraph parent (TryEmitDisplayEquation is hard-
    // coded to parent "/body"). `add equation` on an existing cell paragraph
    // appends the m:oMathPara into it, reproducing the source wrapper shape.
    private static void EmitCellDisplayEquation(DocumentNode eqNode, string parentPath, List<BatchItem> items)
    {
        var mode = eqNode.Format.TryGetValue("mode", out var m) ? m?.ToString() : "display";
        var eqProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["mode"] = string.IsNullOrEmpty(mode) ? "display" : mode!
        };
        if (!string.IsNullOrEmpty(eqNode.Text))
            eqProps["formula"] = eqNode.Text!;
        if (eqNode.Format.TryGetValue("align", out var eqAlign)
            && eqAlign != null && !string.IsNullOrEmpty(eqAlign.ToString()))
            eqProps["align"] = eqAlign.ToString()!;
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = parentPath,
            Type = "equation",
            Props = eqProps
        });
    }

    // Cell Format includes both true tcPr keys and "leaked" keys read from
    // the first inner paragraph/run (align, direction, font, size, bold, …).
    // EmitParagraph re-emits those for the first paragraph, so emitting them
    // here too would double-apply. Whitelist genuine cell-level keys only.
    private static readonly HashSet<string> CellOnlyKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "fill", "width", "valign", "vmerge", "hmerge", "colspan", "nowrap", "textDirection",
        "cnfStyle",
        // BUG-DUMP-CELLTAIL: forward the long-tail tcPr toggles so dump→batch
        // round-trips them (Add.Table.cs / Set.cs apply both).
        "hideMark", "tcFitText",
    };

    private static Dictionary<string, string> ExtractCellOnlyProps(
        Dictionary<string, object?> raw, bool tableIsAutofit)
    {
        // BUG-DUMP-AUTOFITW: drop the fabricated width when the table is
        // autofit AND this cell's width was derived from tblGrid (no source
        // <w:tcW>). Re-emitting a synthetic tcW over-constrains Word's column
        // solver and shifts boundaries. Fixed-layout tables and cells with a
        // real source tcW (no _widthDerived marker) are unchanged — they keep
        // their width + the existing skipGridSync handling.
        bool widthDerived = raw.TryGetValue("_widthDerived", out var wd)
            && wd is bool wdb && wdb;
        bool suppressWidth = tableIsAutofit && widthDerived;
        var filtered = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var (key, val) in raw)
        {
            if (suppressWidth && key.Equals("width", StringComparison.OrdinalIgnoreCase))
                continue;
            if (CellOnlyKeys.Contains(key) ||
                key.StartsWith("border.", StringComparison.OrdinalIgnoreCase) ||
                key.StartsWith("padding.", StringComparison.OrdinalIgnoreCase) ||
                key.StartsWith("shading.", StringComparison.OrdinalIgnoreCase))
            {
                filtered[key] = val;
            }
        }
        // BUG-DUMP21-02: when shading.* sub-keys are present, the
        // FilterEmittableProps shading-fold will emit a folded `shading`
        // key carrying val+fill+color. The legacy `fill` alias surfaced by
        // ReadCellProps duplicates the same color and would cause Set tc
        // to apply the bare-color form on top of the folded shading,
        // overwriting val/color. Drop it here so only the canonical folded
        // form replays.
        if (filtered.Keys.Any(k => k.StartsWith("shading.", StringComparison.OrdinalIgnoreCase)))
        {
            filtered.Remove("fill");
        }
        return FilterEmittableProps(filtered);
    }

    // Row-level keys emitted by Navigation.ReadRowProps. Used by EmitTable
    // so dump→batch round-trips header rows / heights / cantSplit. Cell
    // children are emitted separately via ExtractCellOnlyProps.
    private static readonly HashSet<string> RowOnlyKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "header", "height", "cantSplit", "cnfStyle",
        // BUG-DUMP-R24-1: row-level <w:jc> (whole-row alignment).
        "rowAlign",
        // BUG-DUMP-R24-4: per-row <w:tblPrEx> overrides (verbatim element).
        "tblPrEx",
    };

    /// <summary>Read a string-valued key from a DocumentNode.Format dict
    /// (Format values are typed as <c>object?</c>). Returns null when
    /// the key is missing, the value is null, or the string is empty.
    /// Used by the D1 round-trip path to detect whether a host element
    /// carried a `*PrChange` marker in source.</summary>
    private static string? TryStringFormat(Dictionary<string, object?> format, string key)
    {
        if (!format.TryGetValue(key, out var obj) || obj == null) return null;
        var s = obj.ToString();
        return string.IsNullOrEmpty(s) ? null : s;
    }

    /// <summary>Fold the source's `<paramref name="prefix"/>.author` /
    /// `.date` keys into an existing prop bag as a `revision.type=format`
    /// + `revision.author` + `revision.date` triplet. Called by the
    /// row / cell emit paths so the structural-prop `set` and the
    /// revision attribution travel in one batch step — the
    /// BeginTrackChangeIfRequested snapshot then captures the bare-pr
    /// "before" state vs the just-applied "after" state, producing a
    /// real *PrChange diff Word's reviewing pane will surface (instead
    /// of the legacy two-step emit's `before==after` lie).
    ///
    /// Returns true when a revision was folded in. Mutates
    /// <paramref name="props"/> in place.</summary>
    private static bool FoldRevisionIntoProps(
        Dictionary<string, object?> format,
        string prefix,
        Dictionary<string, string> props)
    {
        var author = TryStringFormat(format, $"{prefix}.author");
        if (author == null) return false;
        props["revision.type"] = "format";
        props["revision.author"] = author;
        var date = TryStringFormat(format, $"{prefix}.date");
        if (date != null) props["revision.date"] = date;
        return true;
    }

    private static Dictionary<string, string> ExtractRowOnlyProps(Dictionary<string, object?> raw)
    {
        var filtered = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        // BUG-DUMP-R25-1: translate the readback's height + height.rule into the
        // rule-specific apply key. Absent height.rule = AUTO row-sizing → bare
        // `height` (no @w:hRule injected). Exact/atLeast map to the explicit
        // keys so the source rule round-trips faithfully.
        string? heightRule = null;
        if (raw.TryGetValue("height.rule", out var ruleObj))
            heightRule = ruleObj?.ToString();
        foreach (var (key, val) in raw)
        {
            if (!RowOnlyKeys.Contains(key)) continue;
            if (string.Equals(key, "height", StringComparison.OrdinalIgnoreCase))
            {
                if (string.Equals(heightRule, "exact", StringComparison.OrdinalIgnoreCase))
                    filtered["height.exact"] = val;
                else if (string.Equals(heightRule, "atLeast", StringComparison.OrdinalIgnoreCase))
                    filtered["height.atleast"] = val;
                else
                    filtered["height"] = val;
            }
            else
            {
                filtered[key] = val;
            }
        }
        return FilterEmittableProps(filtered);
    }
}
