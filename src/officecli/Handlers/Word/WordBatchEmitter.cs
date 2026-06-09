// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

/// <summary>
/// Walks an opened handler's document tree and emits a sequence of BatchItem
/// rows that, when replayed against a blank document of the same format,
/// reconstruct the original document.
///
/// <para>
/// This is the core of the `officecli dump --format batch` pipeline. The
/// emit relies on the OOXML schema reflection fallback in
/// <see cref="TypedAttributeFallback"/> + <see cref="GenericXmlQuery"/>:
/// any leaf property that Get reads can be re-applied via Add/Set, so
/// emit just transcribes Format keys directly without per-property
/// allowlisting.
/// </para>
///
/// <para>
/// Scope (v0.5): docx body paragraphs (with run formatting) + tables (single
/// paragraph + single run per cell, common case). Resources (styles,
/// numbering, theme, headers, footers, sections, comments, footnotes,
/// endnotes) and richer cell contents are NOT yet emitted — follow-up
/// passes will add them.
/// </para>
/// </summary>
public static partial class WordBatchEmitter
{
    /// <summary>
    /// Captured at emit time when a paragraph carries content we cannot round-trip
    /// through the existing handler vocabulary (currently: OLE/embedded-object
    /// runs whose recreate path needs an external src file the emitted batch
    /// has no carrier for). The host paragraph itself is emitted; the
    /// unsupported run is dropped silently from `items` but recorded here so
    /// the CLI can surface a warning bundle to the caller — mirroring pptx's
    /// <see cref="PptxBatchEmitter.UnsupportedWarning"/> contract.
    /// </summary>
    public sealed record DocxUnsupportedWarning(string Element, string Path, string Reason);

    /// <summary>
    /// Emit a batch sequence for a subtree of a Word document.
    /// <para>
    /// Path semantics: dump scopes purely to "what's under this path".
    /// `/` = whole document including all parts (styles, numbering, theme,
    /// settings, body, headers/footers, comments). A subtree path like
    /// `/body/p[5]` emits only that paragraph — styles/numbering/theme are
    /// NOT included because they live at sibling paths (`/styles`,
    /// `/numbering`, etc.), not under the requested subtree. References
    /// such as `style=Heading1` or `numId=3` are emitted as-is; replay
    /// onto a target document that already defines them works, otherwise
    /// the reference falls back to the target's defaults.
    /// </para>
    /// <para>
    /// Known limitations of subtree (non-`/`) dumps:
    /// — Footnote/endnote/chart references inside the emitted paragraph
    ///   resolve to the first N items in the source document's notes/charts,
    ///   not the original positions (cursors start at 0). Use `/` if the
    ///   subtree contains such references.
    /// — Image rels (rIds) reference the source package; the resource itself
    ///   is not bundled.
    /// </para>
    /// </summary>
    /// <summary>
    /// Back-compat overload — returns only the items, discarding the warning
    /// list. Used by the rich test corpus that pre-dates the warning channel.
    /// New callers (CommandBuilder.Dump, ResidentServer) should use
    /// <see cref="EmitWordWithWarnings(WordHandler, string)"/> so the
    /// envelope can surface unsupported-element warnings.
    /// </summary>
    public static List<BatchItem> EmitWord(WordHandler word, string path)
        => EmitWordWithWarnings(word, path).Items;

    public static (List<BatchItem> Items, List<DocxUnsupportedWarning> Warnings) EmitWordWithWarnings(WordHandler word, string path)
    {
        if (string.IsNullOrEmpty(path))
            throw new CliException("dump path cannot be empty. Use '/' for the full document or a subtree path like /body/p[1].")
                { Code = "invalid_path" };
        if (path == "/") return EmitWordWithWarnings(word);

        var items = new List<BatchItem>();
        var warnings = new List<DocxUnsupportedWarning>();
        switch (path.ToLowerInvariant())
        {
            case "/theme": EmitThemeRaw(word, items, warnings); return (items, warnings);
            case "/settings": EmitSettingsRaw(word, items); return (items, warnings);
            case "/numbering": EmitNumberingRaw(word, items); return (items, warnings);
            case "/styles": EmitStyles(word, items); return (items, warnings);
            case "/body":
                EmitBody(word, items, warnings, new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase));
                return (items, warnings);
        }

        // Reject bare /body/p and /body/tbl (no [N]). WordHandler.Get resolves
        // bare name segments to FirstOrDefault, which would silently dump the
        // first paragraph/table — almost never what the caller meant.
        var lastSeg = path.Substring(path.LastIndexOf('/') + 1);
        if (string.Equals(lastSeg, "p", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(lastSeg, "tbl", StringComparison.OrdinalIgnoreCase))
        {
            throw new CliException(
                $"dump path not supported: {path} (missing index predicate). " +
                "Supported: /, /body, /body/p[N], /body/tbl[N], /theme, /settings, /numbering, /styles")
            { Code = "unsupported_path" };
        }

        // Reject deep paths (e.g. /body/tbl[1]/tr[1]/tc[1]/p[1]). The dispatch
        // below assumes parent="/body" and would silently emit a wrongly
        // re-parented node. Supported subtree paths at this point are
        // /body/p[N] or /body/tbl[N] — exactly 2 segments below root.
        var segments = path.Split('/', StringSplitOptions.RemoveEmptyEntries);
        if (segments.Length > 2)
        {
            throw new CliException(
                $"dump path not supported: {path} (nested below /body). " +
                "Supported: /, /body, /body/p[N], /body/tbl[N], /theme, /settings, /numbering, /styles")
            { Code = "unsupported_path" };
        }

        DocumentNode node;
        try { node = word.Get(path); }
        catch (Exception ex)
        {
            throw new CliException($"dump path not found: {path} ({ex.Message})") { Code = "path_not_found" };
        }

        if (node.Type != "paragraph" && node.Type != "p" && node.Type != "table")
        {
            throw new CliException(
                $"dump path not supported: {path} (type={node.Type}). " +
                "Supported: /, /body, /body/p[N], /body/tbl[N], /theme, /settings, /numbering, /styles")
            { Code = "unsupported_path" };
        }

        var ctx = new BodyEmitContext(
            FootnoteTexts: word.Query("footnote").Select(n => n.Text ?? "").ToList(),
            EndnoteTexts: word.Query("endnote").Select(n => n.Text ?? "").ToList(),
            FootnoteCursor: new NoteCursor(),
            EndnoteCursor: new NoteCursor(),
            ChartSpecs: word.Query("chart").Select(c =>
            {
                var full = word.Get(c.Path);
                return new ChartSpec(full.Format, full.Children ?? new List<DocumentNode>());
            }).ToList(),
            ChartCursor: new NoteCursor(),
            ParaIdToTargetIdx: new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase),
            DeferredBookmarks: new List<BatchItem>(),
            TextboxCounters: new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase),
            TableOrdinalBox: new int[1],
            MovePairIds: word.BuildMovePairIdMap(),
            Warnings: warnings);

        if (node.Type == "table")
            EmitTable(word, path, 1, items, ctx);
        else
            EmitParagraph(word, path, "/body", 1, items, autoPresent: false, ctx);

        items.AddRange(ctx.DeferredBookmarks);
        return (items, warnings);
    }

    /// <summary>
    /// Back-compat overload — returns only the items. See
    /// <see cref="EmitWord(WordHandler, string)"/> for the rationale.
    /// </summary>
    public static List<BatchItem> EmitWord(WordHandler word)
        => EmitWordWithWarnings(word).Items;

    /// <summary>Emit a batch sequence for a Word document (full document, equivalent to path "/").</summary>
    public static (List<BatchItem> Items, List<DocxUnsupportedWarning> Warnings) EmitWordWithWarnings(WordHandler word)
    {
        var items = new List<BatchItem>();
        var warnings = new List<DocxUnsupportedWarning>();

        // Phase order matters: resources first so body refs (style=Heading1,
        // numId=3, etc.) resolve when the paragraph adds reach them on replay.
        // Numbering must come BEFORE styles — list-style definitions
        // (Heading paragraphs with numPr) reference numId values, so style
        // adds that carry `numId=N` need /numbering to already hold N.
        EmitNumberingRaw(word, items);
        EmitStyles(word, items);
        // docDefaults (inside styles.xml) round-trips verbatim via raw-set —
        // must follow EmitStyles so it overwrites the blank's stamped block
        // rather than being clobbered by it. See EmitDocDefaultsRaw.
        EmitDocDefaultsRaw(word, items);
        EmitThemeRaw(word, items, warnings);
        EmitSettingsRaw(word, items);
        EmitSection(word, items);
        // Headers/footers run AFTER body: multi-section docs now emit
        // `add header parent="/section[N]"` (see EmitHeaderFooterPart), and
        // the /section[N] resolver only finds the carrier paragraph after
        // EmitBody has added it. Without body in place, every /section[N]
        // resolved to the body-level sectPr (the last section's), so
        // adding header type=default to two different sections collided
        // ("already exists in this section"). Body→header direction has
        // no replay-time dependency: header parts (PAGE/PAGEREF fields,
        // etc.) resolve their cross-refs at render time, not at batch-
        // apply time.
        var paraIdToTargetIdx = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        EmitBody(word, items, warnings, paraIdToTargetIdx);
        EmitHeadersFooters(word, items, warnings);
        EmitComments(word, items, paraIdToTargetIdx);
        // CONSISTENCY(markRPr-inherit-opt-out): dump emits each run's props
        // verbatim from the source; we never want AddRun's UX-convenience
        // markRPr→rPr type-fill to add a w:rFonts (or any other) child the
        // source never had. Stamp the opt-out on every emitted `add r` once
        // here, instead of threading it through five EmitParagraph branches.
        foreach (var it in items)
        {
            if (it.Command == "add" && (it.Type == "r" || it.Type == "run"))
            {
                it.Props ??= new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (!it.Props.ContainsKey("noMarkRPrInherit") && !it.Props.ContainsKey("nomarkrprinherit"))
                    it.Props["noMarkRPrInherit"] = "true";
            }
        }
        // R11 aux-parts: surface a warning per package part the dump surface
        // does not round-trip (customXml, glossary, webSettings, fontTable,
        // embedded fonts, modern-comment metadata, user docProps). Silent
        // data loss is worse than a noisy warning — the warning channel
        // lets agents/users see exactly which content vanished on dump.
        // Scoped to full-document dumps only; subtree paths (`/body/p[5]`,
        // `/styles`, etc.) intentionally do NOT include sibling parts, so
        // warning about them every time would be noise.
        EmitAuxiliaryPartsScan(word, warnings);
        // BUG-R7B(BUG1): footnotes/endnotes round-trip through the reference
        // run that anchors them (emitted as `add footnote`/`add endnote` while
        // walking the body). A note body present in footnotes.xml/endnotes.xml
        // with NO anchoring reference in the document is an orphan — Word never
        // displays it and `add footnote` (which always inserts a reference)
        // cannot recreate it without fabricating an anchor the source never
        // had. Rather than silently drop such note bodies, surface a warning
        // so the loss is visible (mirrors EmitAuxiliaryPartsScan's philosophy:
        // a noisy warning beats silent data loss).
        WarnOrphanNotes(word, items, warnings);
        return (items, warnings);
    }

    // BUG-R7B(BUG1): warn on note bodies that have no anchoring reference run.
    // `add footnote`/`add endnote` items are emitted one-per-reference during
    // the body walk; comparing that count against the number of real note
    // bodies (excluding the reserved separator / continuationSeparator notes,
    // ids -1/0, which Query already filters out) exposes orphans.
    private static void WarnOrphanNotes(WordHandler word, List<BatchItem> items,
                                        List<DocxUnsupportedWarning> warnings)
    {
        WarnOrphanNotesOfKind(word, items, warnings, "footnote");
        WarnOrphanNotesOfKind(word, items, warnings, "endnote");
    }

    private static void WarnOrphanNotesOfKind(WordHandler word, List<BatchItem> items,
                                              List<DocxUnsupportedWarning> warnings, string kind)
    {
        List<DocumentNode> notes;
        try { notes = word.Query(kind); }
        catch { return; }
        if (notes.Count == 0) return;

        var emitted = items.Count(it =>
            it.Command == "add"
            && string.Equals(it.Type, kind, StringComparison.OrdinalIgnoreCase));
        if (emitted >= notes.Count) return;

        // The first `emitted` notes are the ones a reference recovered (Query
        // and the body walk both run in document order); the remainder are
        // orphans. Report each so its lost text is visible.
        foreach (var orphan in notes.Skip(emitted))
        {
            warnings.Add(new DocxUnsupportedWarning(
                Element: kind,
                Path: orphan.Path,
                Reason: $"{kind} body has no anchoring reference in the document; orphan note dropped on dump (text: \"{Truncate(orphan.Text)}\")"));
        }
    }

    private static string Truncate(string? s, int max = 60)
    {
        s ??= "";
        return s.Length <= max ? s : s.Substring(0, max) + "…";
    }

    private static string? ExtractParaId(string anchorPath)
    {
        var m = System.Text.RegularExpressions.Regex.Match(anchorPath, @"@paraId=([0-9A-Fa-f]+)");
        return m.Success ? m.Groups[1].Value : null;
    }

    // Root-level keys that round-trip via `set /`. Includes section page
    // layout, document protection, doc-level grid + defaults. Excludes
    // metadata that auto-updates on save (created/modified timestamps,
    // lastModifiedBy, package author/title — those re-stamp anyway).
    private static readonly HashSet<string> RootScalarKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        // Section page layout (mirrors body's trailing sectPr)
        "pageWidth", "pageHeight", "orientation",
        "marginTop", "marginBottom", "marginLeft", "marginRight",
        // pgMar header/footer-from-edge distances and binding gutter. Without
        // these the round-tripped sectPr fell back to the blank's defaults,
        // silently dropping the source's header/footer spacing.
        "marginHeader", "marginFooter", "marginGutter",
        "pageStart", "pageNumFmt",
        // BUG-DUMP11-01: chapter-numbering attributes on w:pgNumType.
        "chapStyle", "chapSep",
        "titlePage", "direction", "rtlGutter",
        // BUG-DUMP-SECT-VALIGN: vertical text alignment on page (w:vAlign).
        // Set / accepts it; without inclusion here the body section reverted
        // to top on round-trip.
        "vAlign",
        // pgBorders shorthand ('box' / 'none') — Set materialises four
        // matching sides; Get/Navigation surfaces the presence. Without
        // this key the round-trip silently dropped page borders.
        "pgBorders",
        // BUG-DUMP11-03: <w:noEndnote/> section flag.
        "noEndnote",
        "lineNumbers", "lineNumberCountBy",
        // BUG-DUMP11-02: lnNumType/@w:start (first line number when counting).
        "lineNumberStart",
        // Multi-column section layout. Get exposes these as canonical keys
        // (columns, columnSpace, columns.equalWidth) and Set's case table
        // accepts all three (WordHandler.Set.SectionLayout.cs). Without them
        // here, multi-column documents silently revert to single column on
        // round-trip.
        "columns", "columnSpace",
        // BUG-R3: explicit per-column widths/spaces for an unequal-width
        // (equalWidth="false") layout. Get now surfaces these on the body-level
        // section (mirroring per-section readback) and Set / accepts them; without
        // them here the emitted <w:cols equalWidth="false"> had no <w:col>
        // children on replay, silently collapsing uneven columns to equal width.
        // They don't match the "columns." prefix group so must be listed.
        "colWidths", "colSpaces",
        // Document-level final-section break type (oddPage / evenPage /
        // continuous). Set / accepts section.type but the canonical Get
        // surfaces it bare; emit so the trailing sectPr's type survives.
        "section.type",
        // Document protection
        "protection", "protectionEnforced",
        // BUG-DUMP10-03: document-level page background color
        // (<w:document><w:background w:color="…"/>). Set already accepts
        // this canonical key (WordHandler.Add.cs:565); without inclusion
        // here, dump silently dropped the page background on round-trip.
        "background",
        // Document grid (CJK-aware line layout)
        "charSpacingControl",
        // pPrDefault CJK toggles — without these, Word inserts an automatic
        // space between Latin runs and adjacent CJK glyphs ("2025年" →
        // "2025 年"). Templates that explicitly disable autoSpaceDE/DN
        // depend on these surviving the round-trip.
        "kinsoku", "overflowPunct", "autoSpaceDE", "autoSpaceDN",
    };

    // Dotted-prefix groups that round-trip wholesale via `set /`. Each
    // sub-key is forwarded as-is; the schema-reflection layer routes the
    // dotted path into the right OOXML target.
    private static readonly string[] RootPrefixGroups = new[]
    {
        "docDefaults.",
        "docGrid.",
        // columns.equalWidth / columns.separator etc. roundtrip via the
        // canonical dotted form Get already emits.
        "columns.",
        // Page-border per-side detail + offsetFrom. Get/Navigation surfaces
        // pgBorders.<side> + pgBorders.<side>.sz/.color/.space + pgBorders.offsetFrom;
        // EmitSection folds the per-side sub-keys into the colon... err,
        // semicolon-encoded STYLE;SIZE;COLOR;SPACE form Set's pgborders.<side>
        // case expects. Without this prefix the detailed keys were dropped and
        // page borders collapsed to the box default on round-trip.
        "pgBorders.",
        // BUG-DUMP-SECT-FOOTNOTE: section-level footnote/endnote numbering
        // (footnotePr.numFmt / .numRestart / .numStart / .pos, same for
        // endnotePr). Get/Navigation surfaces these dotted keys; Set / routes
        // them through TrySetFootnoteEndnoteNumProps. Without this prefix the
        // trailing section's footnote numbering reverted to decimal/continuous.
        "footnotePr.", "endnotePr.",
    };

    // Captured once per process: blank doc's `Get("/")` root Format, normalized
    // to string values. Used by EmitSection to skip keys whose source value
    // matches what BlankDocCreator stamps — those keys would otherwise leak
    // from blank into the replay target and re-appear on the next dump,
    // breaking dump-then-replay-then-dump idempotency.
    private static readonly Lazy<IReadOnlyDictionary<string, string>> _blankRootBaseline =
        new(ComputeBlankRootBaseline);

    private static IReadOnlyDictionary<string, string> ComputeBlankRootBaseline()
    {
        var tempPath = Path.Combine(Path.GetTempPath(),
            $"officecli_blank_baseline_{Guid.NewGuid():N}.docx");
        try
        {
            OfficeCli.BlankDocCreator.Create(tempPath);
            using var handler = new OfficeCli.Handlers.WordHandler(tempPath, editable: false);
            var root = handler.Get("/");
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var (k, v) in root.Format)
            {
                if (v == null) continue;
                var s = v switch { bool b => b ? "true" : "false", _ => v.ToString() ?? "" };
                if (s.Length > 0) result[k] = s;
            }
            return result;
        }
        catch
        {
            // If baseline computation fails (test harness with no temp path
            // access, etc.), fall back to an empty baseline. EmitSection then
            // behaves as it did before this change.
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }
        finally
        {
            try { File.Delete(tempPath); } catch { }
        }
    }

    private sealed record ChartSpec(Dictionary<string, object?> Format, IReadOnlyList<DocumentNode> Series);

    private sealed record BodyEmitContext(
        List<string> FootnoteTexts,
        List<string> EndnoteTexts,
        NoteCursor FootnoteCursor,
        NoteCursor EndnoteCursor,
        List<ChartSpec> ChartSpecs,
        NoteCursor ChartCursor,
        Dictionary<string, int>? ParaIdToTargetIdx,
        // BUG-DUMP10-04: cross-paragraph bookmarks (endPara > 0) need to be
        // emitted *after* every host paragraph already exists on replay,
        // because AddBookmark relocates the BookmarkEnd to siblings[N+endPara]
        // and that sibling does not exist yet during the in-order walk.
        // EmitParagraph stashes the deferred `add bookmark` rows here;
        // EmitBody appends them once all paragraphs are emitted.
        List<BatchItem> DeferredBookmarks,
        // BUG-DUMP-TXBX: per-host textbox counter. Keyed by the host path
        // ("/body", "/body/tbl[1]/tr[1]/tc[1]", "/header[N]", "/footer[N]").
        // TryEmitPictureRun bumps this when it identifies a textbox-bearing
        // Drawing and uses the post-bump value as N for /<host>/textbox[N].
        // Matches the CountTextboxesInHost selector on the Add side so dump
        // and Add-side indexing stay in lockstep.
        Dictionary<string, int> TextboxCounters,
        // BUG-R11A(BUG1): document-order ordinal of the table currently being
        // emitted, used to build a `(//w:tbl)[N]` raw-set xpath when injecting a
        // block <w:sdt> that is a direct child of a table cell. Single-element
        // mutable box (records are immutable) bumped at the top of every
        // EmitTable call; because EmitTable recurses in DFS document order and
        // every `add table` row is appended to `items` ahead of that table's
        // cell content, N is the stable 1-based `//w:tbl` document-order index
        // the cell-SDT raw-set resolves against at replay time.
        int[] TableOrdinalBox,
        // CONSISTENCY(move-range-markers): map from each moveFrom/moveTo run's
        // own w:id to the SHARED pairing id its bracketing range-marker w:name
        // implies (see WordHandler.BuildMovePairIdMap). EmitPlainOrHyperlinkRun
        // rewrites a moveFrom/moveTo run's revision.id through this map so both
        // halves emit one shared id — AddRun then re-brackets each half with
        // Move_{id} range markers and the moveFrom pairs with its moveTo.
        Dictionary<string, string> MovePairIds,
        // R10-bug1: collected during the body walk whenever an emit helper
        // identifies content it cannot round-trip through the existing
        // handler vocabulary (OLE runs without a carrier for the embedded
        // payload, etc). Mirrors pptx's <see cref="PptxBatchEmitter.SlideEmitContext.Unsupported"/>.
        List<DocxUnsupportedWarning> Warnings);

    private static void EmitBody(WordHandler word, List<BatchItem> items,
                                 List<DocxUnsupportedWarning> warnings,
                                 Dictionary<string, int>? paraIdToTargetIdx = null)
    {
        // BUG-DUMP-X6-02: word.Get("/body") raises "Path not found: /body" on
        // a zip lacking word/document.xml. Surface a CliException pointing at
        // the file rather than leaking an internal path the user never asked
        // for (common when dumping "/" on a corrupt or non-Word zip).
        //
        // BUG-R4B(BUG1): the original catch swallowed EVERY non-CliException as
        // "document.xml is missing", which is misleading when the part IS
        // present but a value inside it failed to parse (e.g. a decimal-valued
        // integer attribute like w:tblInd w:w="0.0"). Distinguish the two: the
        // genuine missing-part path bottoms out in "Path not found: /body";
        // anything else is a real read/parse failure and must surface its own
        // message so the user can see the actual cause.
        DocumentNode bodyNode;
        try
        {
            bodyNode = word.Get("/body");
        }
        catch (Exception ex) when (ex is not CliException)
        {
            var isMissingPart = ex is ArgumentException
                && ex.Message.Contains("Path not found", StringComparison.OrdinalIgnoreCase);
            if (isMissingPart)
                throw new CliException(
                    "dump failed: word/document.xml is missing — the file may not be a valid Word document")
                    { Code = "invalid_document" };
            throw new CliException(
                $"dump failed: could not read word/document.xml — {ex.Message}")
                { Code = "invalid_document" };
        }
        if (bodyNode.Children == null) return;

        // Footnotes/endnotes are referenced by runs (rStyle=FootnoteReference)
        // inside body paragraphs but the run carries no id back to the
        // notes part. We assume notes are listed in document order matching
        // reference order — the typical case since AddFootnote/AddEndnote
        // allocate ids sequentially.
        // Charts: query("chart") returns /chart[N] in document order, which
        // matches the order chart-bearing runs appear in body. Pre-resolve
        // each chart's properties + series children so EmitParagraph can
        // emit a typed `add chart` row when it walks across each ref.
        var charts = word.Query("chart");
        var chartSpecs = charts.Select(c =>
        {
            var full = word.Get(c.Path);
            return new ChartSpec(full.Format, full.Children ?? new List<DocumentNode>());
        }).ToList();

        var ctx = new BodyEmitContext(
            FootnoteTexts: word.Query("footnote").Select(n => n.Text ?? "").ToList(),
            EndnoteTexts: word.Query("endnote").Select(n => n.Text ?? "").ToList(),
            FootnoteCursor: new NoteCursor(),
            EndnoteCursor: new NoteCursor(),
            ChartSpecs: chartSpecs,
            ChartCursor: new NoteCursor(),
            ParaIdToTargetIdx: paraIdToTargetIdx,
            DeferredBookmarks: new List<BatchItem>(),
            TextboxCounters: new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase),
            TableOrdinalBox: new int[1],
            MovePairIds: word.BuildMovePairIdMap(),
            Warnings: warnings);

        // Cross-paragraph fields (a real cached TOC, an IF/REF whose result
        // spans paragraphs, …) can't be modelled by the per-paragraph typed
        // emit — the opener's fldChar(begin) has no matching end in its own
        // paragraph, so TryEmitTocParagraph/AddField would collapse it to a
        // self-contained placeholder and drop the first entry's cached content
        // while orphaning the rest of the result. Raw-pass every paragraph of
        // such a span verbatim instead. Ranges are inclusive 1-based /body/p[N]
        // positions matching pIndex below.
        var spanStartToEnd = new Dictionary<int, int>();
        foreach (var (s, e) in word.GetCrossParagraphFieldSpanRanges())
            spanStartToEnd[s] = e;
        int? activeSpanEnd = null;

        int pIndex = 0, tblIndex = 0;
        foreach (var child in bodyNode.Children)
        {
            switch (child.Type)
            {
                case "paragraph":
                case "p":
                    // BUG-X4-FUZZ-1: display-mode equations surface in
                    // bodyNode.Children as type="paragraph" but the path
                    // resolver addresses them as /body/oMathPara[N], NOT as
                    // /body/p[N]. Incrementing pIndex for them would offset
                    // every subsequent inline-child path (hyperlink/footnote/
                    // run) by +1 per preceding equation, breaking round-trip.
                    // Detect the wrapper via path and route to EmitParagraph
                    // without bumping pIndex — EmitParagraph's equation branch
                    // re-emits the equation as `add /body --type equation`.
                    if (child.Path.Contains("/oMathPara[", StringComparison.OrdinalIgnoreCase))
                    {
                        EmitParagraph(word, child.Path, "/body", pIndex + 1, items, autoPresent: false, ctx);
                    }
                    else
                    {
                        pIndex++;
                        if (activeSpanEnd == null && spanStartToEnd.TryGetValue(pIndex, out var spEnd))
                            activeSpanEnd = spEnd;
                        if (activeSpanEnd != null)
                        {
                            EmitCrossParagraphFieldMember(word, child, pIndex, items, ctx);
                            if (pIndex >= activeSpanEnd.Value) activeSpanEnd = null;
                        }
                        else
                        {
                            EmitParagraph(word, child.Path, "/body", pIndex, items, autoPresent: false, ctx);
                        }
                    }
                    break;
                case "table":
                    tblIndex++;
                    EmitTable(word, child.Path, tblIndex, items, ctx);
                    break;
                case "section":
                case "sectPr":
                    // The body always carries one trailing sectPr that the
                    // blank document already provides; for v0.5 we rely on
                    // that default and skip emitting section properties.
                    // Section emit is a follow-up.
                    break;
                case "sdt":
                    EmitSdt(word, child.Path, items, ctx);
                    break;
                case "bookmark":
                    // Standalone body-level <w:bookmarkStart> (e.g. an anchor
                    // added with `add /body --type bookmark`). Inline bookmarks
                    // inside paragraphs are handled by EmitParagraph; without
                    // this case, body-level bookmark anchors were silently
                    // dropped on dump.
                    {
                        // Use the child node's own Format, NOT word.Get(child.Path):
                        // Navigation assigns every body-direct-child bookmarkStart
                        // the same path (/body/bookmarkStart[1]), so re-resolving by
                        // path returns the FIRST bookmark for all of them — a doc
                        // with body-level TableOfContents + TableOfFigures +
                        // TableOfTables anchors emitted TableOfContents three times
                        // (the dup adds then failed "already exists") and dropped
                        // the other two. child.Format already carries the correct
                        // per-bookmark name/id/endPara.
                        var bmProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        if (child.Format.TryGetValue("name", out var nm)
                            && nm != null && !string.IsNullOrEmpty(nm.ToString()))
                            bmProps["name"] = nm.ToString()!;
                        else
                            break; // BookmarkStart with no name is unusable
                        // BUG-DUMP-BMSPAN: a content-wrapping bookmark splits
                        // into a positioned `open=true` start here and a
                        // separate `end=true` op at the matching bookmarkEnd's
                        // own DOM position (handled by the `bookmarkEnd` case
                        // below). This supersedes the old `endPara` offset for
                        // body-direct bookmarks: the End is replayed at its real
                        // position rather than relocated by paragraph count.
                        if (child.Format.TryGetValue("_spanOpen", out var so)
                            && so is bool bso && bso)
                            bmProps["open"] = "true";
                        else if (child.Format.TryGetValue("endPara", out var ep)
                            && ep != null && ep.ToString() is { Length: > 0 } eps && eps != "0")
                            bmProps["endPara"] = eps;
                        items.Add(new BatchItem
                        {
                            Command = "add",
                            Parent = "/body",
                            Type = "bookmark",
                            Props = bmProps
                        });
                    }
                    break;
                case "bookmarkEnd":
                    // BUG-DUMP-BMSPAN: a NAMED body-direct bookmarkEnd closes a
                    // content-wrapping bookmark opened with `open=true`. Replay a
                    // positioned `end=true` op here so the End lands after the
                    // wrapped paragraph in document order, keeping the range
                    // non-empty (REF/PAGEREF/TOC anchors survive). An unnamed end
                    // node belongs to an empty bookmark whose combined start op
                    // already recreated the pair — emit nothing.
                    if (child.Format.TryGetValue("name", out var beNm)
                        && beNm != null && beNm.ToString() is { Length: > 0 } beNs)
                    {
                        items.Add(new BatchItem
                        {
                            Command = "add",
                            Parent = "/body",
                            Type = "bookmark",
                            Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                            {
                                ["name"] = beNs,
                                ["end"] = "true",
                            }
                        });
                    }
                    break;
                case "equation":
                    // BUG-DUMP13-03: a bare <m:oMathPara> direct child of
                    // <w:body> (not wrapped in a w:p) surfaces in
                    // bodyNode.Children as type="equation". Without this case
                    // it fell to `default: break` and was silently dropped.
                    // Mirror the EmitParagraph equation branch shape.
                    {
                        var eqFull = word.Get(child.Path);
                        var mode = eqFull.Format.TryGetValue("mode", out var m) ? m?.ToString() : "display";
                        var eqProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        {
                            ["mode"] = string.IsNullOrEmpty(mode) ? "display" : mode
                        };
                        if (!string.IsNullOrEmpty(eqFull.Text))
                            eqProps["formula"] = eqFull.Text!;
                        // BUG-DUMP19-02: forward block-equation alignment.
                        if (eqFull.Format.TryGetValue("align", out var eqAlign)
                            && eqAlign != null && !string.IsNullOrEmpty(eqAlign.ToString()))
                            eqProps["align"] = eqAlign.ToString()!;
                        items.Add(new BatchItem
                        {
                            Command = "add",
                            Parent = "/body",
                            Type = "equation",
                            Props = eqProps
                        });
                    }
                    break;
                case "altChunk":
                    // <w:altChunk r:id="…"/> embeds an alternate-format payload
                    // (HTML, RTF, plain text, …) by relationship into the body.
                    // The payload part itself surfaces in EmitAuxiliaryPartsScan
                    // as an `auxiliaryPart` warning, but the body element that
                    // references it would otherwise drop silently — emit a
                    // dedicated warning so the loss is visible in the
                    // dump-warning bundle without the user having to correlate
                    // the aux-part path back to a missing body reference.
                    ctx.Warnings.Add(new DocxUnsupportedWarning(
                        Element: "altChunk",
                        Path: child.Path,
                        Reason: "alternate-format chunk reference dropped on dump (no curated emit path; the referenced payload part is reported separately)"));
                    break;
                default:
                    // Unknown body-level child types — skip for v0.5.
                    break;
            }
        }

        // BUG-DUMP10-04: flush deferred cross-paragraph bookmark rows. They
        // are emitted last so AddBookmark sees the full sibling list when
        // walking forward to the BookmarkEnd's target paragraph.
        items.AddRange(ctx.DeferredBookmarks);
    }

    // Raw-pass one paragraph belonging to a cross-paragraph field span (see
    // WordHandler.GetCrossParagraphFieldSpanRanges). The verbatim <w:p> is
    // inserted immediately before the body's trailing sectPr — the same slot
    // `add p` lands in (AppendBodyParaFast appends before sectPr) — so a span
    // emitted in document order interleaves correctly with the surrounding
    // typed paragraph adds. pIndex still advances per member so subsequent
    // paragraphs' /body/p[N] targets stay aligned with their real position.
    private static void EmitCrossParagraphFieldMember(
        WordHandler word, DocumentNode child, int pIndex,
        List<BatchItem> items, BodyEmitContext ctx)
    {
        // Comments / cross-paragraph anchors keyed by paraId still need this
        // paragraph's position to resolve on replay.
        if (ctx.ParaIdToTargetIdx != null
            && child.Format.TryGetValue("paraId", out var pid) && pid != null)
            ctx.ParaIdToTargetIdx[pid.ToString()!] = pIndex;

        var rawP = word.GetElementXml(child.Path);
        if (string.IsNullOrEmpty(rawP))
        {
            // Couldn't read the element XML — fall back to the typed emit so the
            // paragraph at least keeps its visible text (degraded field).
            EmitParagraph(word, child.Path, "/body", pIndex, items, autoPresent: false, ctx);
            return;
        }
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/document",
            Xpath = "/w:document/w:body/w:sectPr",
            Action = "before",
            Xml = rawP
        });
    }
}
