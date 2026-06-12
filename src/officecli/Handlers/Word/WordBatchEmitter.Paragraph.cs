// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class WordBatchEmitter
{

    /// <summary>
    /// Emit a paragraph at the target index under <paramref name="parentPath"/>.
    /// When <paramref name="autoPresent"/> is true, the parent already has a
    /// pre-existing paragraph at that index (e.g. an auto-created table cell
    /// paragraph); we issue a `set` instead of a fresh `add` so the existing
    /// paragraph gets reused rather than duplicated.
    /// </summary>
    private static void EmitParagraph(WordHandler word, string sourcePath, string parentPath,
                                      int targetIndex, List<BatchItem> items, bool autoPresent,
                                      BodyEmitContext? ctx = null)
    {
        var pNode = word.Get(sourcePath);

        if (TryEmitDisplayEquation(pNode, parentPath, autoPresent, items)) return;

        // Track source paraId -> target index BEFORE any early-return path
        // (section break, TOC, …). Comments anchored on a section-break or
        // TOC paragraph would otherwise miss the mapping and fall back to
        // /body/p[1], silently retargeting the comment.
        if (ctx?.ParaIdToTargetIdx != null && parentPath == "/body" &&
            pNode.Format.TryGetValue("paraId", out var earlyParaId) && earlyParaId != null)
        {
            ctx.ParaIdToTargetIdx[earlyParaId.ToString()!] = targetIndex;
        }

        if (TryEmitInlineSectionBreak(word, pNode, parentPath, items, ctx)) return;
        if (TryEmitTocParagraph(pNode, parentPath, items)) return;
        if (TryEmitTextboxOnlyParagraph(word, pNode, parentPath, autoPresent, items, ctx)) return;

        var props = FilterEmittableProps(pNode.Format);
        // BUG-DUMP-R44-6: paraMarkIns.* must round-trip as a MARK-ONLY tracked
        // insertion — <w:pPr><w:rPr><w:ins/></w:rPr></w:pPr> on the pilcrow
        // ALONE — never as a content insertion that wraps the (plain) run text.
        // The former path here rewrote paraMarkIns.* into a bare revision.author
        // (no revision.type); AddParagraph's bare-attribution branch treats that
        // as "this whole paragraph was inserted" and wraps every auto-created
        // <w:r> in <w:ins>, promoting plain run text to a tracked insertion it
        // never was (Reject Changes would then delete text the source keeps).
        // Pass the paraMarkIns.* keys through verbatim instead — AddParagraph's
        // dedicated paraMarkIns block stamps the mark rPr only, exactly mirroring
        // the paraMarkDel.* handling just below. A genuine run/paragraph content
        // insertion still arrives as revision.type=ins on the run/paragraph and
        // wraps correctly (unaffected by this branch).
        // (No remove here — let paraMarkIns.* pass through to AddParagraph.)
        // BUG-DUMP-R43-8: pPrChange's PreviousParagraphProperties snapshot now
        // round-trips verbatim via revision.beforeXml (set by the pPrChange
        // readback in Navigation.cs and consumed by AddParagraph). The former
        // revision.beforeLost warn-and-drop path is retired — the prior-pPr
        // payload is preserved, not lost. (A defensive strip remains in case a
        // stale dump still carries the legacy key.)
        props.Remove("revision.beforeLost", out var _);
        // paraMarkDel.* — surfaces the dump path through AddParagraph's
        // paraMarkDel block (added alongside this readback). Pass the keys
        // through unchanged; AddParagraph allowlists the prefix and consumes
        // them at end-of-function. Don't fold into revision.* — that
        // namespace already routes to ins/format paths and a paragraph can
        // legitimately carry BOTH paraMarkDel (¶ join) and revision.type=
        // format (pPrChange).
        // (No remove here — let the props pass through to AddParagraph.)
        // BUG-DUMP26-01: numId/numLevel that came from style inheritance
        // (ResolveNumPrFromStyle, no direct w:numPr on the paragraph) must
        // not ride on `add p` — the style already supplies them, and emitting
        // them would semantically promote inherited→explicit on replay.
        // Mirrors the first-run hoist precedent for run-character props
        // inherited from styles.
        bool numInherited = pNode.Format.TryGetValue("numInherited", out var niVal)
            && string.Equals(niVal?.ToString(), "true", StringComparison.OrdinalIgnoreCase);
        if (numInherited)
        {
            props.Remove("numId");
            props.Remove("numLevel");
            props.Remove("numFmt");
            props.Remove("listStyle");
            props.Remove("start");
        }
        // When a paragraph carries numId, the abstractNum/num pair is already
        // in /numbering (raw-set wholesale by EmitNumberingRaw). Forwarding
        // numFmt/listStyle/start to AddParagraph triggers ad-hoc
        // numbering-definition creation in WordHandler.Add — Word allocates
        // a fresh numId (1→9, 2→16, …) and the paragraph references the
        // new one, orphaning the original abstract numbering's level rPr
        // (color, bold, custom marker text). Drop those keys so the
        // paragraph just attaches by numId+numLevel to the existing def.
        if (props.ContainsKey("numId"))
        {
            props.Remove("numFmt");
            props.Remove("listStyle");
            props.Remove("start");
        }
        // BUG-R4F-02: a paragraph may carry a numId that does not resolve to any
        // <w:num> in /numbering (dangling reference). This is valid OOXML — Word
        // renders the paragraph, just without a list marker — but the Add-side
        // dangling-numId guard (WordHandler.Add.Text.cs) rejects it, and because
        // `add p` is atomic the whole paragraph (TEXT included) is lost on replay.
        // Drop the numbering props so the `add p` succeeds with its text, and
        // surface a warning so the dropped numbering is visible. The numbering is
        // kept whenever the numId IS defined (the common case). Mirrors the
        // out-of-window font-size / zero-width-column clamps in Filters.cs.
        if (props.TryGetValue("numId", out var numIdStr)
            && int.TryParse(numIdStr, out var numIdInt)
            && numIdInt > 0
            && !word.IsNumIdDefined(numIdInt))
        {
            props.Remove("numId");
            props.Remove("numLevel");
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "numId",
                Path: pNode.Path,
                Reason: $"paragraph references numId={numIdInt} which is not defined in /numbering (dangling reference); the numbering was dropped so the paragraph text survives dump→batch round-trip"));
        }
        // Collapse non-TOC field chains (fldChar(begin) + instrText(" PAGE ")
        // + fldChar(separate) + display run(s) + fldChar(end)) into a single
        // synthetic "field" entry. Without this collapse, the subsequent
        // `runs` filter sees only the cached display run and emits the field
        // value as static text — PAGE/REF/SEQ/HYPERLINK/NUMPAGES degrade to
        // their evaluated string and stop auto-updating (BUG-X2-05 / X2-1).
        var fieldEntries = CollapseFieldChains(pNode.Children ?? new List<DocumentNode>());
        // R14-bug1+2: each legacy form field embeds a BookmarkStart/End of its
        // own name. AddFormField recreates that bookmark internally — if the
        // emit pipeline also drops an `add bookmark name=X` row before the
        // `add formfield name=X`, AddFormField throws on the duplicate. Filter
        // bookmarks whose name matches a sibling formfield synth's ffName.
        var formFieldNames = fieldEntries
            .Where(e => e.Type == "formfield" && e.Format.TryGetValue("ffName", out _))
            .Select(e => e.Format["ffName"]?.ToString() ?? "")
            .Where(n => !string.IsNullOrEmpty(n))
            .ToHashSet(StringComparer.Ordinal);
        if (formFieldNames.Count > 0)
        {
            fieldEntries = fieldEntries
                .Where(e => !(e.Type == "bookmark"
                    && e.Format.TryGetValue("name", out var bn)
                    && bn != null
                    && formFieldNames.Contains(bn.ToString() ?? "")))
                .ToList();
        }
        // BUG-DUMP5-01/02: include break-typed children in the same ordered
        // list as runs so document-order is preserved on emit.
        var runs = fieldEntries
            .Where(c => c.Type == "run" || c.Type == "r" || c.Type == "picture" || c.Type == "field" || c.Type == "formfield" || c.Type == "ptab" || c.Type == "break"
                || c.Type == "equation"
                || c.Type == "tab"
                || c.Type == "bookmark"
                || c.Type == "bookmarkEnd"
                // BUG-DUMP-PERM: ranged editing-permission markers are
                // positioned paragraph children — keep them in the ordered run
                // list so TryEmitPermRun replays them at their source offset.
                || c.Type == "permStart"
                || c.Type == "permEnd"
                // BUG-DUMP-RUBY: ruby (phonetic guide) child surfaces the
                // verbatim <w:r><w:ruby> XML for a raw-set append.
                || c.Type == "ruby"
                // BUG-DUMP-R42-9: bdo (bidirectional override) child surfaces the
                // verbatim <w:bdo> wrapper XML for a raw-set append.
                || c.Type == "bdo"
                // BUG-DUMP-R43-7: dir (bidirectional embedding) child surfaces the
                // verbatim <w:dir> wrapper XML for a raw-set append.
                || c.Type == "dir"
                // R10-bug1: include ole children so TryEmitOleRun can fire
                // a warning instead of letting them be silently filtered
                // out of the run list (full round-trip is a backlog item).
                || c.Type == "ole")
            .ToList();
        var breaks = runs.Where(c => c.Type == "break").ToList();
        var bookmarks = (pNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "bookmark"
                && !(c.Format.TryGetValue("name", out var bn) && bn != null
                    && formFieldNames.Contains(bn.ToString() ?? "")))
            .ToList();
        var inlineSdts = (pNode.Children ?? new List<DocumentNode>())
            .Where(c => c.Type == "sdt")
            .ToList();

        bool collapseSingleRun = ShouldCollapseSingleRun(word, runs, breaks.Count, bookmarks.Count, inlineSdts.Count);
        pNode.Format.TryGetValue("tabs", out var pTabs);

        if (collapseSingleRun)
        {
            if (runs.Count == 1)
            {
                // BUG-DUMP-R35-2: a wrapper-flattened run that collapses into the
                // paragraph's own `text` prop bypasses EmitPlainOrHyperlinkRun, so
                // the deterministic smartTag/customXml flatten warning never fired
                // for a single-run wrapped paragraph (text survived, loss silent).
                // CONSISTENCY(wrapper-flatten-warning): same emit as
                // EmitPlainOrHyperlinkRun's _wrapperFlattened branch.
                WarnWrapperFlattened(runs[0], ctx);
                var runProps = FilterEmittableProps(runs[0].Format);
                foreach (var (k, v) in runProps)
                {
                    if (!props.ContainsKey(k)) props[k] = v;
                }
                if (!string.IsNullOrEmpty(runs[0].Text))
                    props["text"] = runs[0].Text!;
            }

            if (autoPresent)
            {
                if (props.Count > 0)
                {
                    items.Add(new BatchItem
                    {
                        Command = "set",
                        Path = $"{parentPath}/p[last()]",
                        Props = props
                    });
                }
            }
            else
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = parentPath,
                    Type = "p",
                    Props = props.Count > 0 ? props : null
                });
            }
            EmitTabStops($"{parentPath}/p[last()]", pTabs, items);
            return;
        }

        // Multi-run paragraph: emit the paragraph empty first, then add each
        // run as an explicit child. See BUG-DUMP-HOIST in
        // StripRunCharacterPropsFromParagraph — for multi-run paragraphs the
        // firstRun hoist would re-apply formatting to every sibling on
        // replay, so strip run-level keys before emit.
        //
        // BUG-DUMP-R42-4: but a RUN-LESS paragraph that nonetheless reaches this
        // path (e.g. it carries a bookmark marker, which ShouldCollapseSingleRun
        // keeps off the collapse path so TryEmitBookmarkRun replays the marker)
        // has NO source run to hoist from — Navigation's firstRun-fallback read
        // those bare size/size.cs/bold.cs/font.* keys straight off the paragraph
        // MARK rPr (the ¶ glyph's formatting), not off a run. Stripping them here
        // drops the markRPr entirely, leaving `add p {}` and collapsing a
        // cover-page paragraph's large-size ¶ on rebuild. Only strip when a real
        // text/format-bearing run exists (the genuine hoist source); a run-less
        // paragraph keeps its bare markRPr keys, same as the non-bookmark empty
        // paragraph that rides the collapse path with full markRPr.
        // BUG-DUMP-R26: a field chain swallows the paragraph's text runs in
        // CollapseFieldChains, so a field-result paragraph has NO run-typed
        // children left — yet the paragraph node's bare character keys were
        // harvested (firstRun-fallback) from the field's RESULT runs, and the
        // field emit (raw-set verbatim / add field) replays that formatting
        // itself. Leaving the harvested keys on `add p` duplicates them onto
        // the ¶ mark on rebuild (<w:b/> count 2). Field entries are therefore
        // format-bearing hoist sources too.
        bool hasFormatBearingRun = runs.Any(c => c.Type == "run" || c.Type == "r" || c.Type == "field");
        if (hasFormatBearingRun)
            StripRunCharacterPropsFromParagraph(props);
        if (autoPresent)
        {
            if (props.Count > 0)
            {
                items.Add(new BatchItem
                {
                    Command = "set",
                    Path = $"{parentPath}/p[last()]",
                    Props = props
                });
            }
        }
        else
        {
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = parentPath,
                Type = "p",
                Props = props.Count > 0 ? props : null
            });
        }

        var paraTargetPath = $"{parentPath}/p[last()]";
        EmitTabStops(paraTargetPath, pTabs, items);

        // BUG-DUMP4-06 / BUG-R12B(BUG1): emit each inline SdtRun child AT its
        // real intra-paragraph position interleaved with the runs, not hoisted
        // ahead of them. Navigation appends every SdtRun at the tail of
        // pNode.Children (it sits outside the DOM-ordered run/bookmark/eq merge),
        // so emitting all SDTs before the runs loop scrambled document order: a
        // content control sitting between two runs ("Video [sdt] a powerful…")
        // came back as "[sdt] Video a powerful…". Recover each child's true
        // document rank from the source paragraph XML (top-level child order),
        // then flush each SDT just before the first run whose rank exceeds it.
        var childDocOrder = ComputeParagraphChildDocOrder(word, pNode.Path);
        // Local SDT emit (formerly the standalone foreach above). Returns after
        // appending the appropriate `add sdt` / rich raw-set op to `items`.
        void EmitInlineSdt(DocumentNode sdt)
        {
            // BUG-R12A(BUG1): an inline/run-level <w:sdt> whose content carries
            // more than one run OR any run-level rPr (bold/color/size/…) cannot
            // round-trip through the flat `add sdt text=` path — AddSdt seeds a
            // single unformatted run from `text`, so "FIRSTSECOND" comes back as
            // one plain run and the bold+red is lost. Mirror the R11 block-SDT
            // fix: raw-set the <w:sdt> verbatim into the just-emitted host
            // paragraph (a run-level SDT has no inner <w:p>, so the rich-BLOCK
            // detector doesn't apply; use a run-level richness check). Restricted
            // to /body hosts + no external rels (same constraints as the inline
            // textbox raw-set) so dangling r:id/r:embed can't be produced; other
            // hosts fall back to the flat text emit.
            // BUG-DUMP-R26-7: rich/nested inline SDTs now round-trip verbatim in
            // header/footer/cell hosts too (ResolveRawSetHost), not only /body.
            if (TryEmitRichInlineSdt(word, sdt, parentPath, items, ctx))
                return;
            var sdtProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            // BUG-DUMP-SDTPROPS: forward the form-control sdtPr children the typed
            // emit previously dropped (lock / placeholder / date-picker value /
            // combo+dropdown selection). Same whitelist as the block-SDT path
            // (EmitSdtTyped) so inline and block controls round-trip identically.
            foreach (var key in SdtTypedEmitKeys)
            {
                if (sdt.Format.TryGetValue(key, out var v) && v != null)
                {
                    var s = v.ToString() ?? "";
                    if (s.Length > 0) sdtProps[key] = s;
                }
            }
            if (!string.IsNullOrEmpty(sdt.Text))
                sdtProps["text"] = sdt.Text!;
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "sdt",
                Props = sdtProps
            });
        }

        // BUG-DUMP6-05: collapse N runs that share a hyperlink wrapper into
        // one synthetic hyperlink-typed entry — see CoalesceHyperlinkRuns.
        runs = CoalesceHyperlinkRuns(runs);
        // BUG-D1-MULTIDRAWING-HOST: when this paragraph hosts ≥2 drawing-
        // bearing runs (side-by-side card layout), every textbox must attach
        // to the SAME host paragraph just emitted by the `add p` above so the
        // side-by-side relationship survives round-trip. A single drawing
        // either reached the wrapper-coalesce shortcut (children.Count == 1)
        // OR shares its source paragraph with sibling text/runs — in the
        // latter case attaching to the host paragraph is still correct
        // (preserves the inline relationship that the source had).
        int drawingBearingCount = runs.Count(r =>
        {
            if (r.Type == "picture") return true;
            if (r.Type != "run" && r.Type != "r") return false;
            var probe = word.GetElementXml(r.Path);
            return !string.IsNullOrEmpty(probe) && IsTextboxDrawing(probe);
        });
        // BUG-DUMP-R25-7: a header/footer paragraph that hosts a textbox drawing
        // (e.g. the centred page-number box in a footer) must keep the drawing
        // run INSIDE its styled host paragraph. The body-only
        // TryEmitTextboxOnlyParagraph shortcut does not fire for header/footer
        // hosts, so the textbox otherwise fell through to AddTextbox creating a
        // NEW unstyled paragraph — splitting the footer into (a) an empty
        // pStyle-carrying paragraph and (b) an unstyled drawing paragraph that
        // reverted to Normal's taller exact line height, growing the reserved
        // footer area and reflowing the body (+1 rendered page). Attaching to
        // the just-emitted paraTargetPath (which carries the pPr/pStyle/ind/jc)
        // keeps the drawing in its styled host. Restrict to non-body hosts —
        // /body single-drawing paragraphs already round-trip via the
        // wrapper-coalesce shortcut, and the ≥2 side-by-side case is unchanged.
        string? sharedAttachPara =
            (drawingBearingCount >= 2
             || (drawingBearingCount == 1 && parentPath != "/body"
                 && IsHeaderFooterHost(parentPath)))
            ? paraTargetPath : null;
        // BUG-R14B: hyperlink rows already emitted at this parent belong to
        // earlier paragraphs (paraTargetPath is the same "/body/p[last()]"
        // literal for every <w:p>). Capture that count so multi-run hyperlinks
        // in THIS paragraph re-index from 1 — see EmitStructuredHyperlink.
        int hlBaseline = items.Count(it => it.Type == "hyperlink"
            && string.Equals(it.Parent, paraTargetPath, StringComparison.Ordinal));
        // BUG-R12B(BUG1): inline SDTs sorted by their source document rank,
        // each flushed just before the first run that sits after it. A rank of
        // int.MaxValue (no XML position recovered) falls through to the post-loop
        // tail flush — same behavior as the old hoist-to-end ordering, so a
        // paragraph whose XML can't be probed degrades to the prior shape rather
        // than dropping the SDT.
        var pendingSdts = inlineSdts
            .Select(s => (sdt: s, rank: ChildDocRank(childDocOrder, s.Path)))
            .OrderBy(t => t.rank)
            .ToList();
        int sdtCursor = 0;
        foreach (var run in runs)
        {
            int runRank = ChildDocRank(childDocOrder, run.Path);
            while (sdtCursor < pendingSdts.Count && pendingSdts[sdtCursor].rank < runRank)
            {
                EmitInlineSdt(pendingSdts[sdtCursor].sdt);
                sdtCursor++;
            }
            if (TryEmitBookmarkRun(run, paraTargetPath, items, ctx)) continue;
            if (TryEmitPermRun(run, paraTargetPath, items)) continue;
            if (TryEmitPgNumRun(word, run, parentPath, items, ctx)) continue;
            if (TryEmitDateFieldRun(word, run, parentPath, items, ctx)) continue;
            if (TryEmitHyphenRun(word, run, parentPath, items, ctx)) continue;
            if (TryEmitRubyRun(run, parentPath, paraTargetPath, items, ctx)) continue;
            if (TryEmitBdoRun(run, parentPath, items, ctx)) continue;
            if (TryEmitDirRun(run, parentPath, items, ctx)) continue;
            if (TryEmitBreakRun(word, run, parentPath, paraTargetPath, items, ctx)) continue;
            if (TryEmitTabRun(run, paraTargetPath, items)) continue;
            if (TryEmitPtabRun(run, paraTargetPath, items)) continue;
            if (TryEmitEquationRun(word, run, paraTargetPath, items)) continue;
            if (TryEmitFormFieldRun(run, paraTargetPath, items)) continue;
            if (TryEmitFieldRun(word, run, paraTargetPath, parentPath, items, ctx)) continue;
            // OLE/embedded-object runs surface as type="ole" (see CreateOleNode
            // in WordHandler.ImageHelpers.cs). TryEmitOleRun base64-inlines the
            // embedded payload + icon and the VML frame metadata into a
            // self-contained `add ole` (picture-run style), so the object
            // round-trips with no external file; it warns only when the payload
            // can't be resolved.
            if (TryEmitOleRun(run, paraTargetPath, items, ctx, word)) continue;
            if (TryEmitPictureRun(word, run, paraTargetPath, parentPath, targetIndex, items, ctx, sharedAttachPara)) continue;
            if (TryEmitNoteRefRun(word, run, paraTargetPath, items, ctx)) continue;
            if (TryEmitMixedBreakRun(word, run, parentPath, items, ctx)) continue;
            EmitPlainOrHyperlinkRun(run, paraTargetPath, items, ctx, hlBaseline);
        }
        // Flush any SDTs that sit after the last run (or whose rank could not be
        // recovered from the XML — int.MaxValue lands here).
        for (; sdtCursor < pendingSdts.Count; sdtCursor++)
            EmitInlineSdt(pendingSdts[sdtCursor].sdt);
    }

    // BUG-R12B(BUG1): recover the document-order rank of each top-level
    // paragraph child (run / sdt / bookmark / …) from the source paragraph XML.
    // Navigation surfaces inline SdtRun children at the tail of pNode.Children
    // (outside the DOM-ordered run merge), so the only place the true
    // interleaving survives is the OOXML element order itself. Returns a map
    // from positional child segment ("r[2]", "sdt[1]", "bookmark[1]") to a
    // 0-based document rank. Depth-tracked so a <w:r> nested INSIDE a <w:sdt>
    // (the SDT's own content run) is not counted as a sibling.
    private static Dictionary<string, int> ComputeParagraphChildDocOrder(WordHandler word, string? paragraphPath)
    {
        var map = new Dictionary<string, int>(StringComparer.Ordinal);
        if (string.IsNullOrEmpty(paragraphPath)) return map;
        string xml;
        try { xml = word.GetElementXml(paragraphPath) ?? ""; }
        catch { return map; }
        if (string.IsNullOrEmpty(xml)) return map;
        // Strip the leading <w:p …> open and trailing </w:p> so we walk only the
        // paragraph's content; track nesting depth to keep to top-level children.
        var perNameIdx = new Dictionary<string, int>(StringComparer.Ordinal);
        int rank = 0;
        int depth = 0; // depth relative to the paragraph's content (0 = direct child)
        bool seenParaOpen = false;
        // Match element opens/closes/self-closes for the w: and m: namespaces
        // (m:oMathPara / m:oMath surface as paragraph children too).
        foreach (System.Text.RegularExpressions.Match m in
                 System.Text.RegularExpressions.Regex.Matches(xml, @"<(/?)(?:w|m):([A-Za-z]+)\b[^>]*?(/?)>"))
        {
            var closing = m.Groups[1].Value == "/";
            var name = m.Groups[2].Value;
            var selfClose = m.Groups[3].Value == "/";
            if (!seenParaOpen)
            {
                // The first open we encounter is the paragraph element itself.
                if (!closing) { seenParaOpen = true; }
                continue;
            }
            if (closing) { depth--; continue; }
            if (depth == 0)
            {
                // Direct child of the paragraph — assign the next document rank
                // under its OOXML local name (matches the /r[N], /sdt[N], …
                // path segments Navigation builds).
                var seg = name switch
                {
                    "r" => "r",
                    "sdt" => "sdt",
                    "bookmarkStart" => "bookmark",
                    "bookmarkEnd" => "bookmarkEnd",
                    "hyperlink" => "hyperlink",
                    "fldSimple" => "field",
                    _ => name,
                };
                int idx = perNameIdx.TryGetValue(seg, out var c) ? c : 0;
                perNameIdx[seg] = idx + 1;
                map[$"{seg}[{idx + 1}]"] = rank++;
            }
            if (!selfClose) depth++;
        }
        return map;
    }

    // Look up a child node's document rank from the trailing positional path
    // segment (e.g. "/body/p[…]/r[2]" → "r[2]"). Falls back to int.MaxValue when
    // the path uses a non-positional segment (e.g. r[@…]) or isn't in the map,
    // so unrecoverable children sort to the tail rather than to the front.
    private static int ChildDocRank(Dictionary<string, int> docOrder, string? path)
    {
        if (string.IsNullOrEmpty(path) || docOrder.Count == 0) return int.MaxValue;
        int slash = path.LastIndexOf('/');
        var seg = slash >= 0 ? path[(slash + 1)..] : path;
        return docOrder.TryGetValue(seg, out var r) ? r : int.MaxValue;
    }

    // ── Extracted helpers (behavior unchanged from inline original) ──

    private static bool TryEmitDisplayEquation(DocumentNode pNode, string parentPath, bool autoPresent, List<BatchItem> items)
    {
        // Display-mode equations (<m:oMathPara>) surface in EmitBody's
        // bodyNode.Children as type=paragraph, but a direct Get on the
        // path returns type=equation with the LaTeX-ish formula in
        // DocumentNode.Text. EmitParagraph would otherwise emit an empty
        // `add p` and lose the entire formula. Route to typed
        // `add /body --type equation` instead.
        if (pNode.Type != "equation" || parentPath != "/body" || autoPresent) return false;
        var mode = pNode.Format.TryGetValue("mode", out var m) ? m?.ToString() : "display";
        var eqProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["mode"] = string.IsNullOrEmpty(mode) ? "display" : mode
        };
        if (!string.IsNullOrEmpty(pNode.Text))
            eqProps["formula"] = pNode.Text!;
        // BUG-DUMP19-02: forward block-equation alignment.
        if (pNode.Format.TryGetValue("align", out var eqAlign)
            && eqAlign != null && !string.IsNullOrEmpty(eqAlign.ToString()))
            eqProps["align"] = eqAlign.ToString()!;
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = "/body",
            Type = "equation",
            Props = eqProps
        });
        return true;
    }

    private static bool TryEmitInlineSectionBreak(WordHandler word, DocumentNode pNode, string parentPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // Inline section break: a paragraph carrying <w:sectPr> is the
        // OOXML representation of a mid-document section boundary.
        // AddSection on /body produces this same shape, so we emit
        // `add /body --type section` (which creates a fresh break paragraph)
        // rather than emitting a regular `add p`. The companion
        // sectionBreak.* keys map back to AddSection's prop vocabulary.
        if (parentPath != "/body" ||
            !pNode.Format.TryGetValue("sectionBreak", out var breakKind) || breakKind == null)
            return false;
        {
            // BUG-DUMP-R31-1: a childless mid-document <w:sectPr/> must round-trip
            // bare — no fabricated <w:type>, no default pgSz/pgMar. Navigation
            // emits sectionBreak.type ONLY when the source carried a real
            // <w:type>, and sectionBreak.empty=true when the sectPr was truly
            // childless. Emit `type` only when the source had one; forward the
            // `empty` flag so AddSection produces a bare sectPr instead of
            // stamping the body section's geometry.
            bool sourceHadType = pNode.Format.ContainsKey("sectionBreak.type");
            var sectProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (sourceHadType)
                sectProps["type"] = breakKind.ToString() ?? "nextPage";
            foreach (var (k, v) in pNode.Format)
            {
                if (!k.StartsWith("sectionBreak.", StringComparison.OrdinalIgnoreCase)) continue;
                if (v == null) continue;
                var keyTail = k["sectionBreak.".Length..];
                // `sectionBreak.type` already drove the `type` key above; don't
                // also emit a stray `type` (it's the same value) — skip the alias.
                if (string.Equals(keyTail, "type", StringComparison.OrdinalIgnoreCase)) continue;
                var s = v switch { bool b => b ? "true" : "false", _ => v.ToString() ?? "" };
                if (s.Length > 0) sectProps[keyTail] = s;
            }
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = "/body",
                Type = "section",
                Props = sectProps
            });
            // BUG-DUMP-R37-1: AddSection only captures the sectPr geometry
            // (page/margin/grid/column via sectionBreak.* keys). The HOST
            // paragraph that carries the <w:sectPr> in its pPr also holds its
            // own paragraph-level properties (spacing/ind/jc/widowControl/…) and
            // a paragraph-mark rPr (bold/size/color/font.ea) — the empty
            // section-terminating paragraph's mark-rPr size/spacing set its
            // rendered line height, so dropping them shifts pagination near
            // section boundaries. Re-apply them with a `set` on the rebuilt
            // section paragraph (now at /body/p[last()]), mirroring how a normal
            // paragraph round-trips its pPr: reuse FilterEmittableProps (same
            // pPr/mark-rPr vocabulary AddParagraph/Set understand) and strip the
            // sectionBreak.* keys already consumed by `add section`. On a
            // run-less section paragraph the bare bold/size/color/font.* keys
            // land on the ParagraphMarkRunProperties (SetElement Paragraph branch).
            var sectPProps = FilterEmittableProps(pNode.Format);
            sectPProps.Remove("sectionBreak");
            foreach (var k in sectPProps.Keys
                         .Where(k => k.StartsWith("sectionBreak.", StringComparison.OrdinalIgnoreCase))
                         .ToList())
                sectPProps.Remove(k);
            if (sectPProps.Count > 0)
            {
                items.Add(new BatchItem
                {
                    Command = "set",
                    Path = "/body/p[last()]",
                    Props = sectPProps
                });
            }
            // BUG-DUMP4-04: a section-break paragraph can also carry visible
            // text runs (the carrier paragraph is just a regular paragraph
            // with sectPr in its pPr). AddSection appends a fresh paragraph
            // at /body/p[targetIndex]; emit each text-bearing run as
            // `add r` against that paragraph.
            var carrierRuns = (pNode.Children ?? new List<DocumentNode>())
                .Where(c =>
                {
                    // BUG-DUMP7-11: include inline w:sdt carrier children.
                    if (c.Type == "sdt") return true;
                    // Anchored cover art surfaces as type="picture" when the
                    // drawing carries an image blip — include it for the
                    // drawing-carrier branch below.
                    if (c.Type != "run" && c.Type != "r" && c.Type != "picture") return false;
                    if (!string.IsNullOrEmpty(c.Text)) return true;
                    // BUG-DUMP5-08 / BUG-R7B(BUG1): include empty footnote /
                    // endnote reference runs (their visible text comes via the
                    // typed emit branch below, not from c.Text). Detect via the
                    // actual reference child, not the arbitrary rStyle name.
                    var rsv = c.Format.TryGetValue("rStyle", out var rsraw) ? rsraw?.ToString() : null;
                    if (ClassifyNoteRefRun(word, c, rsv) != NoteRefKind.None)
                        return true;
                    // A cover-page section paragraph can host anchored
                    // drawings (textboxes, picture-filled shapes) in
                    // otherwise text-less runs; dropping them deleted the
                    // whole cover art. Include drawing/pict-bearing runs.
                    var crx = word.GetElementXml(c.Path);
                    if (!string.IsNullOrEmpty(crx)
                        && (crx.Contains("<w:drawing", StringComparison.Ordinal)
                            || crx.Contains("<w:pict", StringComparison.Ordinal)))
                        return true;
                    return false;
                })
                .ToList();
            if (carrierRuns.Count > 0)
            {
                var carrierPath = $"/body/p[last()]";
                foreach (var run in carrierRuns)
                {
                    // BUG-DUMP7-11: inline SDT carrier — same prop whitelist
                    // as the body-paragraph inline-SDT branch.
                    if (run.Type == "sdt")
                    {
                        var sdtCarrierProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        foreach (var key in new[] { "type", "alias", "tag", "items", "format" })
                        {
                            if (run.Format.TryGetValue(key, out var v) && v != null)
                            {
                                var s = v.ToString() ?? "";
                                if (s.Length > 0) sdtCarrierProps[key] = s;
                            }
                        }
                        if (!string.IsNullOrEmpty(run.Text))
                            sdtCarrierProps["text"] = run.Text!;
                        items.Add(new BatchItem
                        {
                            Command = "add",
                            Parent = carrierPath,
                            Type = "sdt",
                            Props = sdtCarrierProps
                        });
                        continue;
                    }
                    var rStyle = run.Format.TryGetValue("rStyle", out var rs) ? rs?.ToString() : null;
                    // BUG-R7B(BUG1): detect via the actual reference child, not
                    // the (arbitrary) rStyle name.
                    var carrierNoteKind = ctx != null
                        ? ClassifyNoteRefRun(word, run, rStyle) : NoteRefKind.None;
                    // BUG-R12A(BUG3): structural note-body emit (per-run rPr +
                    // multi-paragraph), same as TryEmitNoteRefRun. The cursor is
                    // pre-incremented to the 1-based source/target note index.
                    if (carrierNoteKind == NoteRefKind.Footnote)
                    {
                        int idx = ++ctx!.FootnoteCursor.Index;
                        EmitNoteReference(word, "footnote", idx, idx, carrierPath, items, run);
                        continue;
                    }
                    if (carrierNoteKind == NoteRefKind.Endnote)
                    {
                        int idx = ++ctx!.EndnoteCursor.Index;
                        EmitNoteReference(word, "endnote", idx, idx, carrierPath, items, run);
                        continue;
                    }
                    // Drawing/pict-bearing carrier run: ship via the
                    // inlined-parts carrier when it references parts
                    // (rel ids rewritten on replay); a rel-less drawing
                    // raw-sets verbatim into the section paragraph.
                    var carrierRunXml = word.GetElementXml(run.Path);
                    if (!string.IsNullOrEmpty(carrierRunXml)
                        && (carrierRunXml.Contains("<w:drawing", StringComparison.Ordinal)
                            || carrierRunXml.Contains("<w:pict", StringComparison.Ordinal)))
                    {
                        if (word.GetDrawingShapeEmitData(run.Path) is { } csData)
                        {
                            items.Add(new BatchItem
                            {
                                Command = "add",
                                Parent = carrierPath,
                                Type = "drawingshape",
                                Props = PackInlinedPartsProps(csData),
                            });
                            continue;
                        }
                        if (word.GetVmlShapeEmitData(run.Path) is { } cvData)
                        {
                            items.Add(new BatchItem
                            {
                                Command = "add",
                                Parent = carrierPath,
                                Type = "vmlshape",
                                Props = PackInlinedPartsProps(cvData),
                            });
                            continue;
                        }
                        items.Add(new BatchItem
                        {
                            Command = "raw-set",
                            Part = "/document",
                            Xpath = "/w:document/w:body/w:p[last()]",
                            Action = "append",
                            Xml = carrierRunXml
                        });
                        continue;
                    }
                    var rProps = FilterEmittableProps(run.Format);
                    if (!string.IsNullOrEmpty(run.Text))
                        rProps["text"] = run.Text!;
                    items.Add(new BatchItem
                    {
                        Command = "add",
                        Parent = carrierPath,
                        Type = "r",
                        Props = rProps
                    });
                }
            }
            return true;
        }
    }

    /// <summary>
    /// BUG-DUMP-TXBX-WRAPPER: a body paragraph whose only meaningful child is
    /// a textbox-bearing Drawing run (textboxes ship inside
    /// <c>&lt;mc:AlternateContent&gt;</c>, so Get reports the run as
    /// type=&quot;run&quot; with no Format hints) used to emit BOTH an empty
    /// <c>add p</c> wrapper AND a typed <c>add textbox</c> row. On replay
    /// AddTextbox creates its own host paragraph, leaving the target with
    /// one extra empty paragraph per textbox. Detect the
    /// textbox-only-paragraph shape here and emit only the textbox row.
    /// </summary>
    private static bool TryEmitTextboxOnlyParagraph(
        WordHandler word, DocumentNode pNode, string parentPath, bool autoPresent,
        List<BatchItem> items, BodyEmitContext? ctx)
    {
        // Wrapper coalescing only makes sense at /body — header/footer/cell
        // hosts of a textbox have their own pattern and we don't want to
        // skip wrapping paragraphs that carry visible run formatting.
        if (parentPath != "/body" || autoPresent) return false;
        if (!string.IsNullOrEmpty(pNode.Text)) return false;
        var children = pNode.Children ?? new List<DocumentNode>();
        // Need exactly one drawing-bearing child (run / picture) and nothing
        // else. Bookmarks / sdts / breaks need the paragraph wrapper to
        // anchor against and must not coalesce.
        if (children.Count != 1) return false;
        var run = children[0];
        // Source-side: AlternateContent wraps the drawing so Get reports the
        // run as plain "run"/"r" with no Format hints.
        // Target-side (after AddTextbox replay): Drawing sits directly under
        // Run with no AlternateContent, so Get reports it as "picture".
        // Both shapes must collapse here — otherwise source and target dumps
        // disagree on whether to emit the `add p` wrapper and the round-trip
        // drift grows on every textbox.
        if (run.Type != "run" && run.Type != "r" && run.Type != "picture") return false;
        // Don't gate on run.Text here: picture/textbox runs surface their
        // docPr name in DocumentNode.Text (e.g. "文本框 1") which is not
        // visible body text — it doesn't disqualify the wrapper-coalesce.
        var rawXml = word.GetElementXml(run.Path);
        if (string.IsNullOrEmpty(rawXml) || !IsTextboxDrawing(rawXml)) return false;
        // BUG-DUMP-R26-6: a legacy VML textbox round-trips via a verbatim
        // raw-set APPEND into the host paragraph (TryEmitTextbox), which needs
        // an `add p` to exist first. The host-less shortcut here (which relies
        // on AddTextbox creating its own modern host) would leave the raw-set
        // with no /body/p[last()] target. Bail so the normal EmitParagraph flow
        // emits `add p`, then routes the run through TryEmitPictureRun →
        // TryEmitTextbox, which raw-sets the VML into the just-added paragraph.
        if (IsVmlTextbox(rawXml)) return false;
        // Delegate to the same emit path TryEmitPictureRun uses so geometry
        // props + inner-paragraph recursion stay identical.
        return TryEmitTextbox(word, run, rawXml, parentPath, items, ctx);
    }

    private static bool TryEmitTocParagraph(DocumentNode pNode, string parentPath, List<BatchItem> items)
    {
        // TOC field-bearing paragraph: a fldChar(begin) + instrText("TOC ...")
        // + fldChar(separate) + placeholder run + fldChar(end) chain. Get
        // exposes only the placeholder text on the parent paragraph, so
        // emitting a regular `add p text=...` would drop the field structure
        // entirely and Word would no longer auto-update the TOC on open.
        if (parentPath != "/body" || pNode.Children == null) return false;
        var instrChild = pNode.Children
            .FirstOrDefault(c => c.Type == "instrText"
                && (c.Format.TryGetValue("instruction", out var iv)
                    && iv?.ToString()?.TrimStart().StartsWith("TOC", StringComparison.OrdinalIgnoreCase) == true));
        if (instrChild == null) return false;
        var instr = instrChild.Format["instruction"]!.ToString()!;
        var tocProps = ParseTocInstruction(instr);
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = "/body",
            Type = "toc",
            Props = tocProps
        });
        return true;
    }

    private static bool ShouldCollapseSingleRun(WordHandler word, List<DocumentNode> runs, int breaksCount, int bookmarksCount, int inlineSdtsCount)
    {
        // Single-run / no-run paragraph: collapse run formatting into the
        // paragraph's prop bag (the schema-reflection layer accepts run-level
        // keys on a paragraph and routes them through ApplyRunFormatting).
        if (runs.Count > 1) return false;
        if (breaksCount > 0 || bookmarksCount > 0 || inlineSdtsCount > 0) return false;
        if (runs.Count == 0) return true;
        // BUG-DUMP-PERM: a paragraph holding a ranged editing-permission marker
        // must stay on the explicit-run path so TryEmitPermRun replays the
        // marker at its offset; collapsing into `add p` would drop it.
        if (runs.Any(rr => rr.Type == "permStart" || rr.Type == "permEnd")) return false;
        // BUG-DUMP-R40-7: a run-LESS paragraph whose only content is a
        // bookmark/bookmarkEnd marker (the close end of a bookmark that spans
        // into this paragraph) must stay on the explicit-run path so
        // TryEmitBookmarkRun replays the `add bookmark end=true` op. The
        // `bookmarks` collector (pNode.Children where Type=="bookmark") only
        // catches the START side, so a lone bookmarkEnd left bookmarksCount==0
        // and the single-run collapse folded the marker's name= into `add p`
        // props (leaking a phantom paragraph name and dropping the bookmarkEnd
        // entirely — the bookmark was left unterminated). Bookmark markers carry
        // no paragraph-level text/format, so there is nothing to collapse anyway.
        if (runs.Any(rr => rr.Type == "bookmark" || rr.Type == "bookmarkEnd")) return false;
        var r = runs[0];
        // Picture / ptab runs need their own typed `add` rows.
        if (r.Type == "picture" || r.Type == "ptab") return false;
        // BUG-DUMP-R25-2: a sole run whose only content is a tab CHARACTER
        // (<w:r><w:tab/></w:r>, surfaced as Type=="tab" with empty Text) must
        // stay on the explicit-run path so TryEmitTabRun replays `add r
        // text="\t"`. Collapsing into `add p` flattens it via GetRunText and
        // — with no <w:t> text to carry — the tab character vanishes. In
        // multi-run paragraphs the tab survives because sibling run ops keep
        // the structure; only this sole-run case was lost.
        if (r.Type == "tab") return false;
        // OLE / embedded-object runs must stay on the explicit-run path so
        // TryEmitOleRun fires. Collapsing folds the ole metadata (progId,
        // fileSize, drawAspect, …) into `add p` props that AddParagraph
        // silently ignores — the embedded object vanishes with no warning.
        // The explicit path surfaces the documented "ole run dropped" warning
        // (full binary round-trip is a separate, deferred feature) and keeps
        // the host paragraph intact.
        if (r.Type == "ole") return false;
        // A single run carrying a drawing/shape payload (textbox, connector
        // line, autoshape — surfaced as a plain "run" when AlternateContent-
        // wrapped) must NOT collapse into `add p`: collapsing flattens the
        // paragraph and silently drops the shape. Keep it on the explicit-run
        // path so TryEmitPictureRun preserves it (typed textbox or raw-set).
        if ((r.Type == "run" || r.Type == "r"))
        {
            var probe = word.GetElementXml(r.Path);
            if (!string.IsNullOrEmpty(probe) && IsDrawingBearingRun(probe)) return false;
            // BUG-DUMP-R24-3: a sole run carrying a page/column <w:br> MIXED with
            // <w:t> text must stay on the explicit-run path so TryEmitMixedBreakRun
            // raw-passes the verbatim <w:r>. Collapsing to `add p text="…"`
            // flattens the run via GetRunText (which drops page/column breaks —
            // they have no \n representation), silently losing the break.
            if (!string.IsNullOrEmpty(probe)
                && System.Text.RegularExpressions.Regex.IsMatch(probe, @"<w:br\b[^>]*\bw:type=""(?:page|column)""")
                && System.Text.RegularExpressions.Regex.IsMatch(probe, @"<w:t[\s>]"))
                return false;
        }
        // R14-bug1+2: legacy form field synth — needs its own typed
        // `add formfield` row; collapsing into `add p` flattens the
        // ffData wrapper into paragraph props that AddParagraph ignores.
        if (r.Type == "formfield") return false;
        // BUG-DUMP6-05 / BUG-DUMP10-05: hyperlink-wrapped run (url, anchor,
        // or tooltip-only via isHyperlink sentinel) must re-emit as
        // `add hyperlink` — `add p` does not consume url/anchor.
        if (r.Format.ContainsKey("url") || r.Format.ContainsKey("anchor")
            || r.Format.ContainsKey("isHyperlink")) return false;
        // BUG-FUZZ-2 / BUG-R7B(BUG1): footnote/endnote reference runs need the
        // typed `add footnote/endnote` branch; AddParagraph doesn't consume the
        // reference. Detect via the actual reference child, not the arbitrary
        // rStyle name (real docs use ids like "a5", not "FootnoteReference").
        {
            var srStyle = r.Format.TryGetValue("rStyle", out var srraw) ? srraw?.ToString() : null;
            if (ClassifyNoteRefRun(word, r, srStyle) != NoteRefKind.None)
                return false;
        }
        // BUG-W14-EFFECTS / BUG-DUMP5-09 / 7-01 / 5-10: run-level w14 effects /
        // OpenType properties / sym / trackChange — AddParagraph's
        // ApplyRunFormatting fallback has no cases for these; they'd
        // surface as UNSUPPORTED on replay.
        if (r.Format.ContainsKey("w14shadow")
            || r.Format.ContainsKey("textOutline")
            || r.Format.ContainsKey("textFill")
            || r.Format.ContainsKey("w14glow")
            || r.Format.ContainsKey("w14reflection")
            || r.Format.ContainsKey("ligatures")
            || r.Format.ContainsKey("numForm")
            || r.Format.ContainsKey("numSpacing")
            || r.Format.ContainsKey("revision.type")
            // BUG-DUMP-PGNUM: a sole run containing <w:pgNum/> (even one that
            // ALSO carries <w:t> text and/or a <w:cr/>) must stay on the
            // explicit-run path so TryEmitPgNumRun raw-passes the verbatim
            // <w:r>. Collapsing to `add p text="…"` flattens the run into
            // paragraph props and silently drops the pgNum placeholder.
            || r.Format.ContainsKey("_hasPgNum")
            // BUG-DUMP-DATEFIELD: a run containing a date-component placeholder
            // (<w:dayLong/> etc.) must stay on the explicit-run path so
            // TryEmitDateFieldRun raw-passes the verbatim <w:r>. Collapsing to
            // `add p text="…"` flattens the run into paragraph props and persists
            // GetRunText's "[dayLong]" sentinel as literal text, losing the
            // element Word substitutes the date against.
            || r.Format.ContainsKey("_hasDateField")
            // BUG-DUMP-R40-3: a run containing <w:noBreakHyphen/>/<w:softHyphen/>
            // must stay on the explicit-run path so TryEmitHyphenRun raw-passes
            // the verbatim <w:r>. Collapsing to `add p text="…"` persists
            // GetRunText's glyph as literal text and drops the structural hyphen.
            || r.Format.ContainsKey("_hasHyphen")
            || r.Format.ContainsKey("sym")) return false;
        // BUG-RSHD-PROMOTE: a sole run carrying run-level character shading
        // (<w:rPr><w:shd>) must NOT collapse into `add p`. AddParagraph routes
        // `shading`/`shd` to PARAGRAPH properties (<w:pPr><w:shd>) only — there
        // is no run-level shading routing on `add p` (BUG-DUMP22-03 deliberately
        // suppresses stamping pPr shading onto the inline run). Collapsing the
        // run's shading.* keys would therefore hoist a tight character highlight
        // into a full page-width paragraph band. Unlike `bdr`/`highlight` —
        // which `add p` forwards to ApplyRunFormatting (rPr) so they keep their
        // run-level identity in the collapse — `shading` semantics diverge
        // between pPr and rPr, so the only correct round-trip is the explicit
        // `add r shading=…` path (WordHandler.Add.Text.cs routes a run's
        // `shading`/`shd` to rPr). Keep this run on the explicit-run path.
        // Genuine paragraph-level shading (read from pPr onto the paragraph
        // node, not the run) still rides on `add p` unaffected.
        if (r.Format.Keys.Any(k => k.StartsWith("shading.", StringComparison.OrdinalIgnoreCase)))
            return false;
        // BUG-FIELD-COLLAPSE: a synthetic field run carries `instruction=…` —
        // collapse would lose the field chain on replay.
        if (r.Type == "field") return false;
        // BUG-DUMP-RUBY: a sole ruby (phonetic-guide) run must stay on the
        // explicit-run path so TryEmitRubyRun raw-sets the <w:ruby> verbatim.
        // Collapsing into `add p` would fold the base text into a plain run
        // and drop the <w:ruby>/<w:rt>/<w:rubyBase> wrapper.
        if (r.Type == "ruby") return false;
        // BUG-DUMP-R42-9: a sole <w:bdo> (bidirectional override) child must stay
        // on the explicit-run path so TryEmitBdoRun raw-sets the wrapper verbatim.
        // Collapsing into `add p` would fold the inner text into a plain run and
        // drop the <w:bdo> wrapper (and its load-bearing w:val direction).
        if (r.Type == "bdo") return false;
        // BUG-DUMP-R43-7: a sole <w:dir> (bidirectional embedding) child must stay
        // on the explicit-run path so TryEmitDirRun raw-sets the wrapper verbatim.
        if (r.Type == "dir") return false;
        // BUG-DUMP7-03: inline equation must emit `add equation` explicitly.
        if (r.Type == "equation") return false;
        // BUG-DUMP-R42-5: a sole run carrying its own reading direction
        // (<w:rPr><w:rtl/></w:rtl>, surfaced as Format["direction"]) must NOT
        // collapse into `add p`. The run-prop hoist would copy `direction` into
        // the paragraph prop bag, and AddParagraph routes `direction` through
        // ApplyDirectionCascade — stamping a paragraph-level <w:bidi/> AND a ¶
        // mark <w:rtl/>, flipping the whole paragraph's base direction. rtl is a
        // RUN-level property; keep it on the explicit-run path so it re-emits as
        // `add r --prop direction=…` (run rPr only). A genuine paragraph-level
        // <w:bidi/> rides on the paragraph node's own `direction` key (read from
        // pPr.BiDi) and is unaffected.
        if (r.Format.ContainsKey("direction")) return false;
        return true;
    }

    private static bool TryEmitBookmarkRun(DocumentNode run, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // BUG-DUMP25-01: bookmark child emitted in DOM order so a
        // BookmarkStart between runs survives round-trip at its original
        // intra-paragraph offset.
        //
        // BUG-DUMP-BMSPAN: a content-WRAPPING bookmark (`_spanOpen=true`, set
        // by Navigation when the range holds runs/equations/fields) is split
        // into TWO positioned ops: an `open=true` start here (places only
        // <w:bookmarkStart>) and a matching `end=true` op emitted at the
        // BookmarkEnd's own DOM position (a separate `bookmarkEnd` child, even
        // a downstream paragraph for cross-paragraph spans). Document-order
        // replay then keeps the wrapped content INSIDE the range so
        // REF/PAGEREF/TOC anchors survive. A zero-length bookmark (no
        // `_spanOpen`) keeps the single combined op so it stays empty.
        if (run.Type == "bookmarkEnd")
        {
            var endProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (run.Format.TryGetValue("name", out var enm) && enm != null
                && enm.ToString() is { Length: > 0 } ens)
                endProps["name"] = ens;
            else
                return true; // unnamed end marker — start emit recreates pair
            endProps["end"] = "true";
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "bookmark",
                Props = endProps
            });
            return true;
        }
        if (run.Type != "bookmark") return false;
        var bmProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (run.Format.TryGetValue("name", out var bmName) && bmName != null)
        {
            var s = bmName.ToString();
            if (!string.IsNullOrEmpty(s)) bmProps["name"] = s;
        }
        if (bmProps.Count == 0) return true; // skip unnamed/anonymous bookmarks
        if (run.Format.TryGetValue("_spanOpen", out var sp) && sp is bool bsp && bsp)
            bmProps["open"] = "true";
        // BUG-DUMP-R32-4: forward colFirst/colLast for an inline
        // table-column-range bookmark (mirrors the body-direct case).
        ForwardBookmarkColRange(run.Format, bmProps);
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = paraTargetPath,
            Type = "bookmark",
            Props = bmProps
        });
        return true;
    }

    // BUG-DUMP-R32-4: copy a bookmark node's colFirst/colLast (a rectangular
    // table-column-range bookmark) into the `add bookmark` props so AddBookmark
    // re-stamps w:colFirst/w:colLast. Shared by the body-direct and inline
    // bookmark emit paths.
    private static void ForwardBookmarkColRange(
        Dictionary<string, object?> format, Dictionary<string, string> bmProps)
    {
        if (format.TryGetValue("colFirst", out var cf)
            && cf?.ToString() is { Length: > 0 } cfs)
            bmProps["colFirst"] = cfs;
        if (format.TryGetValue("colLast", out var cl)
            && cl?.ToString() is { Length: > 0 } cls)
            bmProps["colLast"] = cls;
    }

    // BUG-DUMP-PERM: emit a ranged editing-permission marker (<w:permStart>/
    // <w:permEnd>) at its DOM position so an editable-region delimiter survives
    // round-trip. Mirrors TryEmitBookmarkRun: each marker is a positioned
    // `add permStart`/`add permEnd` op carrying its source attributes.
    private static bool TryEmitPermRun(DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        if (run.Type != "permStart" && run.Type != "permEnd") return false;
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var key in run.Type == "permStart"
            ? new[] { "id", "edGrp", "ed", "colFirst", "colLast" }
            : new[] { "id" })
        {
            if (run.Format.TryGetValue(key, out var v) && v != null)
            {
                var s = v.ToString() ?? "";
                if (s.Length > 0) props[key] = s;
            }
        }
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = paraTargetPath,
            Type = run.Type,
            Props = props
        });
        return true;
    }

    private static bool TryEmitBreakRun(WordHandler word, DocumentNode run, string parentPath, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // BUG-DUMP5-01/02: a soft <w:br/> with NO type attribute is a line
        // break, not a page break — fall back to type=line. Emitted inline
        // from the unified runs loop so each break stays at its source
        // position instead of being hoisted to the front of the paragraph.
        if (run.Type != "break") return false;
        var breakType = run.Format.TryGetValue("breakType", out var bt) ? bt?.ToString() : null;
        // BUG-R12B(BUG3): a break run can carry its own <w:rPr> (<w:noProof/>,
        // font, color, …). Navigation strips every typography/proofing key from
        // a `break` node (TypographyOnlyKeys), and `add pagebreak` builds a bare
        // <w:r><w:br/></w:r> with no rPr, so the run-properties were dropped on
        // round-trip. The break has no scalar add-API for arbitrary rPr, so —
        // mirroring the ruby / rich-inline-SDT raw-set fallback — re-insert the
        // verbatim <w:r><w:rPr>…</w:rPr><w:br/></w:r> when the source run carries
        // a non-empty rPr. Restricted to /body hosts with no external rels (same
        // constraints as the other raw-set fallbacks) so no r:id/r:embed can
        // dangle; everything else stays on the lossless typed `add pagebreak`.
        if (parentPath == "/body")
        {
            var rawXml = word.RawElementXml(run.Path);
            if (!string.IsNullOrEmpty(rawXml)
                && BreakRunHasMeaningfulRunProps(rawXml!)
                && !HasExternalRelRef(rawXml!))
            {
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = "/document",
                    Xpath = "/w:document/w:body/w:p[last()]",
                    Action = "append",
                    Xml = rawXml!
                });
                return true;
            }
        }
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = paraTargetPath,
            Type = "pagebreak",
            Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["type"] = string.IsNullOrEmpty(breakType) ? "line" : breakType!
            }
        });
        return true;
    }

    // BUG-DUMP-R24-3: a page/column break that lives MIXED inside a
    // text-bearing run (<w:r><w:t>Before</w:t><w:br w:type="page"/><w:t>After
    // </w:t></w:r>) is dropped on dump→batch. GetRunText flattens the run text
    // to "BeforeAfter" (page/column <w:br> has no \n representation, unlike a
    // textWrapping break) and the `add r` replay loses the break entirely.
    // A break that is the SOLE content of its run already round-trips via the
    // Navigation `break`-node upgrade + TryEmitBreakRun. Here we cover only the
    // mixed case: when the run is a plain text run whose verbatim XML carries a
    // page/column <w:br>, re-insert the whole <w:r> verbatim via a raw-set
    // append (mirrors the rich-break / ruby / pgNum raw-set fallback), so text
    // AND the break — with full run formatting — survive intact.
    private static bool TryEmitMixedBreakRun(WordHandler word, DocumentNode run, string parentPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // Only plain text runs reach here (break-only runs are Type=="break"
        // and consumed by TryEmitBreakRun upstream). Picture/field/etc. runs
        // were already handled by their own TryEmit* probes before this point.
        if (run.Type is not ("run" or "r")) return false;
        var rawXml = word.RawElementXml(run.Path);
        if (string.IsNullOrEmpty(rawXml)) return false;
        // Mixed only: a page/column <w:br> alongside a <w:t> (text) child. A
        // break-only run never has a <w:t>, so it won't match and stays on the
        // typed path. textWrapping/line breaks (no w:type, or w:type="textWrapping")
        // already round-trip as \n and must NOT be raw-set here.
        bool hasPageOrColumnBreak = System.Text.RegularExpressions.Regex.IsMatch(
            rawXml, @"<w:br\b[^>]*\bw:type=""(?:page|column)""");
        if (!hasPageOrColumnBreak) return false;
        bool hasText = System.Text.RegularExpressions.Regex.IsMatch(rawXml, @"<w:t[\s>]");
        if (!hasText) return false;
        // BUG-DUMP-R24-3 (FIX-3 follow-up): the verbatim raw-set targets
        // /w:document/w:body/w:p[last()]; only a body host has that addressable
        // last() paragraph. A mixed-break run inside a header/footer/cell would
        // mis-anchor. Return FALSE (not true) so the run falls through to the
        // normal text path (EmitPlainOrHyperlinkRun), which preserves the run's
        // <w:t> content via GetRunText — only the page/column <w:br> is dropped.
        // Returning true here would consume the run and DROP ITS TEXT TOO,
        // which is worse than the pre-fix flatten-but-keep-text behaviour.
        // Surface the deterministic "break lost" warning, then defer.
        if (parentPath != "/body")
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "break",
                Path: run.Path,
                Reason: "page/column break mixed inside a text run within a header/footer/table cell could not be serialized for round-trip; the break is dropped from the replayed document (the run's text is preserved)"));
            return false;
        }
        if (HasExternalRelRef(rawXml!)) return false;
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/document",
            Xpath = "/w:document/w:body/w:p[last()]",
            Action = "append",
            Xml = rawXml!
        });
        return true;
    }

    // A break run is "rich" when it carries a non-empty <w:rPr> (noProof, font,
    // color, …) that the typed `add pagebreak` path cannot reproduce. An empty
    // <w:rPr/> (or none at all) is lossless on the typed path, so only a
    // populated rPr forces the verbatim raw-set fallback.
    private static bool BreakRunHasMeaningfulRunProps(string runXml)
    {
        var m = System.Text.RegularExpressions.Regex.Match(
            runXml, @"<w:rPr\b[^>]*?(?:/>|>(.*?)</w:rPr>)", System.Text.RegularExpressions.RegexOptions.Singleline);
        if (!m.Success) return false;
        // Self-closing <w:rPr/> has no inner content → nothing to preserve.
        var inner = m.Groups[1].Value;
        return !string.IsNullOrWhiteSpace(inner);
    }

    private static bool TryEmitRubyRun(DocumentNode run, string parentPath, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // BUG-DUMP-RUBY: a <w:ruby> (CJK phonetic guide) has no scalar
        // representation on add/set — its <w:rt> (furigana) and <w:rubyBase>
        // (base text) carry independent runs with their own rPr, and the
        // <w:rubyPr> alignment/hps/hpsRaise/hpsBaseText/lid attributes have no
        // add-API. Re-insert the captured <w:r><w:ruby>…</w:r> verbatim via a
        // raw-set append, mirroring the textbox/connector raw-set fallback in
        // TryEmitPictureRun (same shape: append the run's outer XML to the
        // just-emitted host paragraph at /w:document/w:body/w:p[last()]).
        if (run.Type != "ruby") return false;
        var rawXml = run.Format.TryGetValue("_rawRubyXml", out var rx) ? rx?.ToString() : null;
        if (string.IsNullOrEmpty(rawXml))
            return true; // nothing to emit (shouldn't happen) — consumed anyway
        // The verbatim raw-set targets /w:document/w:body/w:p[last()]; only a
        // body host has that addressable last() paragraph. A ruby inside a
        // header/footer/cell would need a different anchor — flag the loss
        // rather than mis-anchoring (full non-body round-trip is a backlog
        // item; same conservatism as the non-body textbox raw-set).
        if (parentPath != "/body")
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "ruby",
                Path: run.Path,
                Reason: "ruby (phonetic guide) inside a header/footer/table cell could not be serialized for round-trip; the base text is lost from the replayed document"));
            return true;
        }
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/document",
            Xpath = "/w:document/w:body/w:p[last()]",
            Action = "append",
            Xml = rawXml!
        });
        return true;
    }

    private static bool TryEmitBdoRun(DocumentNode run, string parentPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // BUG-DUMP-R42-9: a <w:bdo> (bidirectional override — forces the visual
        // RTL/LTR character ordering of its wrapped runs) has no scalar add/set
        // representation; the typed `add r` path drops the wrapper, losing the
        // load-bearing w:val direction. Re-insert the captured <w:bdo>…</w:bdo>
        // verbatim via a raw-set append, mirroring the ruby/pgNum raw-set
        // fallback (same shape: append the wrapper's outer XML to the just-
        // emitted host paragraph at /w:document/w:body/w:p[last()]).
        if (run.Type != "bdo") return false;
        var rawXml = run.Format.TryGetValue("_rawBdoXml", out var rx) ? rx?.ToString() : null;
        if (string.IsNullOrEmpty(rawXml))
            return true; // nothing to emit (shouldn't happen) — consumed anyway
        // The verbatim raw-set targets /w:document/w:body/w:p[last()]; only a
        // body host has that addressable last() paragraph. A bdo inside a
        // header/footer/cell would need a different anchor — flag the loss
        // rather than mis-anchoring (same conservatism as the ruby fallback).
        if (parentPath != "/body")
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "bdo",
                Path: run.Path,
                Reason: "bidirectional override (w:bdo) inside a header/footer/table cell could not be serialized for round-trip; the wrapped runs survive flattened but the forced character ordering is lost from the replayed document"));
            return true;
        }
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/document",
            Xpath = "/w:document/w:body/w:p[last()]",
            Action = "append",
            Xml = rawXml!
        });
        return true;
    }

    private static bool TryEmitDirRun(DocumentNode run, string parentPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // BUG-DUMP-R43-7: a <w:dir> (bidirectional embedding — sets the bidi
        // embedding direction of its wrapped runs, distinct from <w:bdo> override
        // and the run-level <w:rtl> toggle) has no scalar add/set representation;
        // the typed `add r` path drops the wrapper, losing the load-bearing w:val
        // direction. Re-insert the captured <w:dir>…</w:dir> verbatim via a
        // raw-set append, mirroring the bdo fallback exactly.
        if (run.Type != "dir") return false;
        var rawXml = run.Format.TryGetValue("_rawDirXml", out var rx) ? rx?.ToString() : null;
        if (string.IsNullOrEmpty(rawXml))
            return true; // nothing to emit (shouldn't happen) — consumed anyway
        // Only a body host has an addressable last() paragraph; a dir inside a
        // header/footer/cell would mis-anchor — flag the loss instead (same
        // conservatism as the bdo/ruby fallbacks).
        if (parentPath != "/body")
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "dir",
                Path: run.Path,
                Reason: "bidirectional embedding (w:dir) inside a header/footer/table cell could not be serialized for round-trip; the wrapped runs survive flattened but the embedding direction is lost from the replayed document"));
            return true;
        }
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/document",
            Xpath = "/w:document/w:body/w:p[last()]",
            Action = "append",
            Xml = rawXml!
        });
        return true;
    }

    private static bool TryEmitPgNumRun(WordHandler word, DocumentNode run, string parentPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // BUG-DUMP-PGNUM: a run containing <w:pgNum/> (page-number placeholder)
        // has no scalar add/set representation — the typed `add r` path drops the
        // <w:pgNum/> entirely (text/other content survives but the placeholder
        // vanishes with no warning). Mirroring the rich-break / ruby raw-set
        // fallback, re-insert the verbatim <w:r> via a raw-set append so the
        // <w:pgNum/> — and any co-located <w:cr/> the typed path would demote to
        // <w:br/> — survive the round-trip. RunToNode stamps Format["_hasPgNum"].
        if (!run.Format.ContainsKey("_hasPgNum")) return false;
        // Only the /body host has the addressable last() paragraph anchor the
        // raw-set targets; a header/footer-hosted pgNum needs a different anchor
        // (backlog, same conservatism as the non-body ruby/textbox fallback).
        if (parentPath != "/body")
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "pgNum",
                Path: run.Path,
                Reason: "page-number placeholder (w:pgNum) inside a header/footer/table cell could not be serialized for round-trip; the placeholder is lost from the replayed document"));
            return true;
        }
        var rawXml = word.RawElementXml(run.Path);
        if (string.IsNullOrEmpty(rawXml)) return true; // nothing to emit
        // Same external-rel guard as the other raw-set fallbacks — a dangling
        // r:id/r:embed would not resolve in the rebuilt document.
        if (HasExternalRelRef(rawXml!))
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "pgNum",
                Path: run.Path,
                Reason: "page-number placeholder run carries an external relationship reference and could not be serialized verbatim for round-trip"));
            return true;
        }
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/document",
            Xpath = "/w:document/w:body/w:p[last()]",
            Action = "append",
            Xml = rawXml!
        });
        return true;
    }

    private static bool TryEmitDateFieldRun(WordHandler word, DocumentNode run, string parentPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // BUG-DUMP-DATEFIELD: a run containing a Word date-component placeholder
        // (<w:dayLong/> / <w:dayShort/> / <w:monthLong/> / <w:monthShort/> /
        // <w:yearLong/> / <w:yearShort/>) has no scalar add/set representation —
        // the typed `add r` path persists GetRunText's "[dayLong]" human sentinel
        // as literal <w:t> text and drops the element entirely. Mirroring the
        // pgNum raw-set fallback, re-insert the verbatim <w:r> via a raw-set
        // append so the date element — and any co-located <w:t> text in the same
        // run — survive the round-trip. RunToNode stamps Format["_hasDateField"].
        if (!run.Format.ContainsKey("_hasDateField")) return false;
        // Only the /body host has the addressable last() paragraph anchor the
        // raw-set targets; a header/footer-hosted date field needs a different
        // anchor (backlog, same conservatism as the pgNum/ruby/textbox fallback).
        if (parentPath != "/body")
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "dateField",
                Path: run.Path,
                Reason: "date-component placeholder (w:dayLong/w:monthLong/…) inside a header/footer/table cell could not be serialized for round-trip; the placeholder is lost from the replayed document"));
            return true;
        }
        var rawXml = word.RawElementXml(run.Path);
        if (string.IsNullOrEmpty(rawXml)) return true; // nothing to emit
        // Same external-rel guard as the other raw-set fallbacks — a dangling
        // r:id/r:embed would not resolve in the rebuilt document.
        if (HasExternalRelRef(rawXml!))
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "dateField",
                Path: run.Path,
                Reason: "date-component placeholder run carries an external relationship reference and could not be serialized verbatim for round-trip"));
            return true;
        }
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/document",
            Xpath = "/w:document/w:body/w:p[last()]",
            Action = "append",
            Xml = rawXml!
        });
        return true;
    }

    private static bool TryEmitHyphenRun(WordHandler word, DocumentNode run, string parentPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // BUG-DUMP-R40-3: a run containing <w:noBreakHyphen/> (non-breaking
        // hyphen) or <w:softHyphen/> (discretionary hyphen) — siblings of <w:t>
        // inside the run — has no scalar add/set representation. The typed
        // `add r`/`add p text=` path persists GetRunText's Unicode glyph
        // (U+2011 / U+00AD) as literal <w:t> text and drops the structural hyphen
        // element. Mirroring the pgNum/dateField raw-set fallback, re-insert the
        // verbatim <w:r> via a raw-set append so the hyphen element — and any
        // co-located <w:t> text in the same run — survive the round-trip.
        // RunToNode stamps Format["_hasHyphen"].
        if (!run.Format.ContainsKey("_hasHyphen")) return false;
        // Only the /body host has the addressable last() paragraph anchor the
        // raw-set targets; a header/footer/cell-hosted hyphen needs a different
        // anchor (backlog, same conservatism as the pgNum/dateField fallback).
        if (parentPath != "/body")
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "hyphen",
                Path: run.Path,
                Reason: "non-breaking/soft hyphen (w:noBreakHyphen/w:softHyphen) inside a header/footer/table cell could not be serialized for round-trip; the structural hyphen is lost from the replayed document (the run's text is preserved)"));
            return true;
        }
        var rawXml = word.RawElementXml(run.Path);
        if (string.IsNullOrEmpty(rawXml)) return true; // nothing to emit
        // Same external-rel guard as the other raw-set fallbacks — a dangling
        // r:id/r:embed would not resolve in the rebuilt document.
        if (HasExternalRelRef(rawXml!))
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "hyphen",
                Path: run.Path,
                Reason: "hyphen run carries an external relationship reference and could not be serialized verbatim for round-trip"));
            return true;
        }
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = "/document",
            Xpath = "/w:document/w:body/w:p[last()]",
            Action = "append",
            Xml = rawXml!
        });
        return true;
    }

    private static bool TryEmitTabRun(DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        // BUG-DUMP14-02: tab-only run (<w:r><w:tab/></w:r>) surfaces as
        // type="tab" with empty Text. AddText splits "\t" into TabChar, so
        // emit `add r text="\t"` to round-trip the tab character.
        if (run.Type != "tab") return false;
        var tabParent = ResolveHyperlinkParent(run, paraTargetPath, items);
        var tabProps = FilterEmittableProps(run.Format);
        tabProps["text"] = "\t";
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = tabParent,
            Type = "r",
            Props = tabProps
        });
        return true;
    }

    private static bool TryEmitPtabRun(DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        // BUG-PTAB: ptab (positional tab) — Navigation surfaces its own run
        // type with align/relativeTo/leader on Format. Without an explicit
        // emit branch the runs filter would drop it and round-trip would
        // silently lose right-align/header-style tabs.
        if (run.Type != "ptab") return false;
        // BUG-DUMP-TABRPR: carry the ptab run's own rPr (font / size / szCs /
        // …) alongside its align/relativeTo/leader. Like a tab, a positional
        // tab paints a leader in the run's font and contributes to line
        // height, so its typography is meaningful — RunToNode keeps it on
        // run.Format and AddPtab applies it on replay.
        var ptabProps = FilterEmittableProps(run.Format);
        if (run.Format.TryGetValue("align", out var pAlign) && pAlign != null)
            ptabProps["alignment"] = pAlign.ToString() ?? "";
        if (run.Format.TryGetValue("relativeTo", out var pRel) && pRel != null)
            ptabProps["relativeTo"] = pRel.ToString() ?? "";
        if (run.Format.TryGetValue("leader", out var pLead) && pLead != null)
            ptabProps["leader"] = pLead.ToString() ?? "";
        var ptabParent = ResolveHyperlinkParent(run, paraTargetPath, items);
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = ptabParent,
            Type = "ptab",
            Props = ptabProps.Count > 0 ? ptabProps : null
        });
        return true;
    }

    // Build `add hyperlink` props from a source hyperlink node so a hyperlink
    // that carries no emittable runs (e.g. one wrapping only an <m:oMath>) can
    // still be materialized. Returns null when the hyperlink has no addressable
    // destination (neither url nor anchor) — the caller then routes the child
    // under the paragraph so its content still survives.
    private static Dictionary<string, string>? TryBuildHyperlinkAddProps(WordHandler word, string srcHlPath)
    {
        DocumentNode hl;
        try { hl = word.Get(srcHlPath); }
        catch { return null; }
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var key in new[] { "url", "anchor", "tooltip", "tgtFrame", "history" })
        {
            if (hl.Format.TryGetValue(key, out var v) && v != null)
            {
                var s = v.ToString();
                if (!string.IsNullOrEmpty(s)) props[key] = s!;
            }
        }
        return (props.ContainsKey("url") || props.ContainsKey("anchor")) ? props : null;
    }

    private static bool TryEmitEquationRun(WordHandler word, DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        // BUG-DUMP7-03: inline <m:oMath> as paragraph child. Get surfaces it
        // as type="equation" with mode=inline and the LaTeX-ish formula in
        // Text. AddEquation accepts a paragraph parent for inline mode.
        // BUG-DUMP15-04: m:oMath inside w:hyperlink surfaces with a
        // hyperlink-scoped path (.../p[N]/hyperlink[K]/equation[M]). Strip
        // the trailing /equation[M] segment so the emitted Parent places the
        // equation INSIDE the hyperlink on replay.
        if (run.Type != "equation") return false;
        var eqMode = run.Format.TryGetValue("mode", out var emv) ? emv?.ToString() : "inline";
        var eqProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["mode"] = string.IsNullOrEmpty(eqMode) ? "inline" : eqMode!
        };
        // Always emit `formula` (even when empty); ToLatex may legitimately
        // return "" for minimal m:oMath.
        eqProps["formula"] = run.Text ?? "";
        var eqParent = paraTargetPath;
        if (!string.IsNullOrEmpty(run.Path))
        {
            var idxEq = run.Path.LastIndexOf("/equation[", StringComparison.Ordinal);
            if (idxEq > 0)
            {
                var derived = run.Path.Substring(0, idxEq);
                var hlIdx = derived.LastIndexOf("/hyperlink[", StringComparison.Ordinal);
                if (hlIdx > 0)
                {
                    // BUG-DBF-R3-02: the equation lives inside a <w:hyperlink>.
                    // A hyperlink is normally materialized via its child runs, but
                    // a hyperlink whose ONLY content is an <m:oMath> has no runs —
                    // so no `add hyperlink` row is emitted and replaying
                    // `add equation parent=…/hyperlink[K]` fails ("path not
                    // found"), dropping the math. Emit the missing `add hyperlink`
                    // (fetched from the source) before the equation so its parent
                    // exists; if the hyperlink can't be resolved, fall back to the
                    // paragraph so the math survives (the rare link wrapper is lost
                    // rather than the content).
                    var srcHlPath = run.Path.Substring(0, idxEq); // …/hyperlink[K]
                    var hlSeg = derived.Substring(hlIdx); // /hyperlink[K]
                    int alreadyEmitted = items.Count(it => it.Type == "hyperlink"
                        && string.Equals(it.Parent, paraTargetPath, StringComparison.Ordinal));
                    var rebasedHl = paraTargetPath + hlSeg;
                    int wantK = 0;
                    var kStr = hlSeg.Length > 11 ? hlSeg[11..^1] : "";
                    int.TryParse(kStr, out wantK);
                    if (alreadyEmitted < wantK)
                    {
                        var hlProps = TryBuildHyperlinkAddProps(word, srcHlPath);
                        if (hlProps != null)
                        {
                            items.Add(new BatchItem
                            {
                                Command = "add",
                                Parent = paraTargetPath,
                                Type = "hyperlink",
                                Props = hlProps
                            });
                            eqParent = rebasedHl;
                        }
                        // else: leave eqParent = paraTargetPath (math survives).
                    }
                    else
                    {
                        eqParent = rebasedHl;
                    }
                }
            }
        }
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = eqParent,
            Type = "equation",
            Props = eqProps
        });
        return true;
    }

    // R14-bug1+2: legacy form field synth from CollapseFieldChains. The
    // begin run's <w:ffData> was unpacked onto ff*-prefixed Format keys
    // (ffName, ffType, ffDefault, ffMaxLength, ffChecked, ffItems,
    // ffHelpText, ffStatusText, ffEntryMacro, ffExitMacro, ffCalcOnExit,
    // ffTextType, ffTextFormat, ffCheckBoxSize, ffEnabled). Map back to
    // AddFormField's accepted keys (drop the `ff` prefix, lowercase the
    // first letter) so replay rebuilds the FieldChar + FormFieldData +
    // Bookmark chain with every original wrapper intact.
    private static bool TryEmitFormFieldRun(DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        if (run.Type != "formfield") return false;
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        // CONSISTENCY(formfield-keys): AddFormField accepts type/formfieldtype,
        // name, default, maxLength, checked, items, helpText, statusText,
        // entryMacro, exitMacro, calcOnExit, textType, textFormat,
        // checkBoxSize, enabled. Mirror the AddFormField accepted vocabulary.
        foreach (var (k, v) in run.Format)
        {
            if (v == null) continue;
            if (!k.StartsWith("ff", StringComparison.OrdinalIgnoreCase)) continue;
            // Strip the "ff" prefix and lowercase the first remaining char so
            // ffName→name, ffDefault→default, etc.
            if (k.Length <= 2) continue;
            var bare = char.ToLowerInvariant(k[2]) + k.Substring(3);
            var sv = v.ToString();
            if (string.IsNullOrEmpty(sv)) continue;
            props[bare] = sv!;
        }
        // Preserve cached display text where AddFormField would otherwise
        // emit the placeholder symbol (text input) or "false" (checkbox).
        if (!string.IsNullOrEmpty(run.Text) && !props.ContainsKey("text"))
            props["text"] = run.Text!;
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = paraTargetPath,
            Type = "formfield",
            Props = props,
        });
        return true;
    }

    private static bool TryEmitFieldRun(WordHandler word, DocumentNode run, string paraTargetPath, string parentPath, List<BatchItem> items, BodyEmitContext? ctx = null)
    {
        if (run.Type != "field") return false;
        // BUG-DUMP-R26-2: a field whose cached result has multiple distinctly-
        // formatted runs can't round-trip through `add field` (single-rPr model
        // collapses the runs and leaks the first run's bold onto the fldChar
        // markers). CollapseFieldChains flagged it and stashed the field-slice
        // run paths (begin..end). Raw-set the whole chain verbatim so per-run
        // formatting survives.
        // BUG-DUMP-R26-7: fire for /body, header/footer AND table-cell hosts
        // (ResolveRawSetHost), not only /body — previously a rich field result
        // in a header/footer/cell silently fell to the lossy typed emit.
        if (run.Format.TryGetValue("_richFieldResult", out var rfr) && rfr is bool rfrB && rfrB
            && run.Format.TryGetValue("_fieldSlicePaths", out var spObj) && spObj is string spStr
            && !string.IsNullOrEmpty(spStr)
            && ResolveRawSetHost(parentPath, ctx) is { } fieldHost)
        {
            var sb = new System.Text.StringBuilder();
            bool ok = true;
            foreach (var p in spStr.Split('\n'))
            {
                if (string.IsNullOrEmpty(p)) continue;
                var xml = word.GetElementXml(p);
                if (string.IsNullOrEmpty(xml)) { ok = false; break; }
                sb.Append(xml);
            }
            if (ok && sb.Length > 0)
            {
                var chainXml = sb.ToString();
                // BUG-DUMP-R26-7 (PART B): an external relationship inside the
                // cached result (e.g. a hyperlink r:id) would DANGLE in the
                // rebuilt part — the raw-set can't recreate the rel. Do NOT fall
                // silently to the lossy extractor; emit a deterministic warning
                // naming the loss. Full rel preservation is a separate effort.
                if (HasExternalRelRef(chainXml))
                {
                    ctx?.Warnings.Add(new DocxUnsupportedWarning(
                        Element: "field.richResult",
                        Path: run.Path,
                        Reason: "field cached result with per-run formatting AND an external relationship (hyperlink/image) cannot round-trip verbatim; the relationship target is not carried through dump→batch, so the formatted result is flattened to uniform text on replay"));
                    // Fall through to the typed path (preserves instruction +
                    // cached value with uniform formatting; loss is now visible).
                }
                else
                {
                    items.Add(new BatchItem
                    {
                        Command = "raw-set",
                        Part = fieldHost.Part,
                        Xpath = fieldHost.XPath,
                        Action = "append",
                        Xml = chainXml
                    });
                    return true;
                }
            }
            // Reconstruction failed — fall through to the typed path so the
            // field still round-trips (cached value + uniform formatting).
        }
        // BUG-DUMP-R26-7 (PART B): a single-run field result that wraps a
        // hyperlink (external rel) isn't "rich" (so the verbatim raw-set above
        // doesn't fire) but the typed path below still drops the hyperlink
        // wrapper + rel silently (text + bold survive). Emit a deterministic
        // warning so the loss is visible. (The multi-run rich+rel case is
        // already warned on the raw-set branch above, so gate on NOT-rich to
        // avoid a double warning.)
        if (run.Format.TryGetValue("_fieldResultHasExternalRel", out var frer)
            && frer is bool frerB && frerB
            && !(run.Format.TryGetValue("_richFieldResult", out var rfr2) && rfr2 is bool rfr2B && rfr2B))
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "field.richResult",
                Path: run.Path,
                Reason: "field cached result wraps a hyperlink whose external relationship target is not carried through dump→batch; the hyperlink is flattened (its link is dropped) on replay"));
        }
        // Synthetic field entry from CollapseFieldChains. Format carries
        // `instruction` (raw fldSimple/instrText) and Text holds the cached
        // display value. AddField parses the instruction and rebuilds the
        // fldChar chain on replay.
        // BUG-DUMP18-02: w:fldSimple / fldChar-chain field inside w:hyperlink
        // should replay INSIDE the hyperlink — but only when a prior
        // `add hyperlink` row actually landed at the target paragraph
        // (BUG-DUMP9-03 fldSimple-only hyperlinks never surface a hyperlink
        // row, and routing the field there would fail path lookup on replay).
        if (run.Type != "field") return false;
        // R10-bug7: CollapseFieldChains flagged a nested field (IF/REF with
        // an inner DATE/PAGE/MERGEFIELD branch). AddField rebuilds a flat
        // begin/instr/separate/display/end chain and cannot model the
        // nested branches — emitting an `add field` row here would either
        // throw (parser sees garbage), drop the inner branches, OR merge
        // the inner instruction into the outer expression. Backlog item:
        // teach AddField to accept a tree representation. Until then, the
        // cheapest correct behavior is to flag the loss in envelope
        // warnings — same model as the OLE warning above — so callers
        // don't ship a doc with the IF false-branch silently stripped.
        // R10-bug8: malformed field (begin without matching end) surfaces as
        // a synth from CollapseFieldChains with _unmatchedFieldBegin=true.
        // Same warning model as _nestedField — preserve cached display,
        // flag the partial instruction in envelope.warnings.
        if (run.Format.TryGetValue("_unmatchedFieldBegin", out var ufbObj) && ufbObj is bool ufbB && ufbB)
        {
            if (ctx != null)
            {
                var partialInstr = run.Format.TryGetValue("instruction", out var pIv)
                    ? pIv?.ToString() ?? "" : "";
                ctx.Warnings.Add(new DocxUnsupportedWarning(
                    Element: "field.unmatched_begin",
                    Path: run.Path,
                    Reason: $"fldChar(begin) without matching end; partial instruction='{partialInstr}' dropped"));
            }
            if (!string.IsNullOrEmpty(run.Text))
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = paraTargetPath,
                    Type = "r",
                    Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["text"] = run.Text!
                    }
                });
            }
            return true;
        }
        if (run.Format.TryGetValue("_nestedField", out var nfObj) && nfObj is bool nfB && nfB)
        {
            // BUG-DUMP-R43-5: round-trip the nested field's begin..end run slice
            // verbatim via raw-set so the full fldChar/instrText sequence (and the
            // inner field) survives, instead of collapsing to cached display text.
            // Each slice run carries a resolvable source Path; concatenate their
            // OuterXml and append to the just-emitted host paragraph at
            // /w:document/w:body/w:p[last()] (the same last()-relative attach the
            // ruby/bdo raw-set fallbacks use). Only a /body host has an
            // addressable last() paragraph — a nested field inside a header/footer/
            // cell falls back to the legacy warning+cached-text path.
            var slicePaths = run.Format.TryGetValue("_nestedFieldSlicePaths", out var nfSpObj)
                ? nfSpObj as List<string> : null;
            if (parentPath == "/body" && slicePaths is { Count: > 0 })
            {
                var sb = new System.Text.StringBuilder();
                bool allResolved = true;
                foreach (var sp in slicePaths)
                {
                    var rx = word.GetElementXml(sp);
                    if (string.IsNullOrEmpty(rx)) { allResolved = false; break; }
                    sb.Append(rx);
                }
                if (allResolved && sb.Length > 0)
                {
                    items.Add(new BatchItem
                    {
                        Command = "raw-set",
                        Part = "/document",
                        Xpath = "/w:document/w:body/w:p[last()]",
                        Action = "append",
                        Xml = sb.ToString()
                    });
                    return true;
                }
            }
            // Fallback (non-body host or unresolvable slice): preserve the cached
            // display and flag the loss, as before.
            if (ctx != null)
            {
                ctx.Warnings.Add(new DocxUnsupportedWarning(
                    Element: "field.nested",
                    Path: run.Path,
                    Reason: "nested field (begin inside a field's branch) inside a header/footer/table cell could not be serialized for round-trip; cached display preserved but inner field codes dropped"));
            }
            // Still emit the cached display so the paragraph isn't empty.
            if (!string.IsNullOrEmpty(run.Text))
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = paraTargetPath,
                    Type = "r",
                    Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["text"] = run.Text!
                    }
                });
            }
            return true;
        }
        var instr = run.Format.TryGetValue("instruction", out var iv)
            ? iv?.ToString() ?? "" : "";
        var fieldProps = BuildFieldAddProps(instr, run.Text ?? "");
        if (fieldProps != null
            && run.Format.TryGetValue("_noFieldSeparator", out var nfs)
            && nfs is bool nfsB && nfsB)
        {
            fieldProps["noSeparator"] = "true";
        }
        // BUG-DUMP-R37-4: forward the source field's lock bit so AddField
        // recreates it locked (begin fldChar @w:fldLock). Carried on the synth
        // Format by CollapseFieldChains (complex) / Navigation (fldSimple).
        if (fieldProps != null
            && run.Format.TryGetValue("fldLock", out var flkv) && flkv != null
            && string.Equals(flkv.ToString(), "true", StringComparison.OrdinalIgnoreCase))
        {
            fieldProps["fldLock"] = "true";
        }
        // BUG-DUMP-R24-2: source field had a separator but an empty cached
        // result. Pass an explicit empty `text` so AddField emits an empty
        // result run instead of fabricating a «name» placeholder.
        if (fieldProps != null
            && !fieldProps.ContainsKey("text")
            && run.Format.TryGetValue("_emptyFieldResult", out var efr)
            && efr is bool efrB && efrB)
        {
            fieldProps["text"] = "";
        }
        // BUG-R12A(BUG1): forward the captured cached-result-run formatting
        // (stashed under `_resultFmt.` by CollapseFieldChains) onto the `add
        // field` prop bag. AddField applies font/size/bold/color uniformly to
        // every rebuilt field run, so a bold/red/20pt PAGE field round-trips
        // styled instead of as a plain "1". Keys AddField can't express
        // (italic/underline/strike) are surfaced as a warning — the field still
        // round-trips (instruction + cached value preserved), only the extra
        // emphasis is lost; raw-set field passthrough is a backlog item.
        if (fieldProps != null)
        {
            var unsupportedFmt = new List<string>();
            foreach (var (k, v) in run.Format)
            {
                if (v == null) continue;
                if (!k.StartsWith("_resultFmt.", StringComparison.OrdinalIgnoreCase)) continue;
                var bare = k["_resultFmt.".Length..];
                if (!FieldAddSupportedFormatKeys.Contains(bare))
                {
                    unsupportedFmt.Add(bare);
                    continue;
                }
                // Map font.latin/font.ascii/font.hAnsi onto AddField's `font`.
                var target = bare.StartsWith("font.", StringComparison.OrdinalIgnoreCase) ? "font" : bare;
                if (!fieldProps.ContainsKey(target))
                {
                    var s = v switch { bool b => b ? "true" : "false", _ => v.ToString() ?? "" };
                    if (s.Length > 0) fieldProps[target] = s;
                }
            }
            if (unsupportedFmt.Count > 0 && ctx != null)
            {
                ctx.Warnings.Add(new DocxUnsupportedWarning(
                    Element: "field.resultFormat",
                    Path: run.Path,
                    Reason: $"cached field-result run formatting ({string.Join(", ", unsupportedFmt)}) cannot be expressed via add field; field instruction + bold/color/size/font preserved, extra emphasis dropped"));
            }
        }
        // BUG-DUMP-DELFIELD: forward the revision attribution that
        // CollapseFieldChains propagated from the <w:del>/<w:ins> wrapper onto
        // the synth. AddField wraps the rebuilt field chain in the matching
        // <w:del>/<w:ins> when it sees revision.type — so a deleted HYPERLINK
        // round-trips as a deletion (delInstrText + delText inside <w:del>)
        // rather than collapsing to live text. Forward to both the `add field`
        // and the plain-text fallback below.
        if (fieldProps != null)
        {
            foreach (var rk in new[] { "revision.type", "revision.author", "revision.date", "revision.id" })
            {
                if (fieldProps.ContainsKey(rk)) continue;
                if (run.Format.TryGetValue(rk, out var rv) && rv != null)
                {
                    var s = rv.ToString();
                    if (!string.IsNullOrEmpty(s)) fieldProps[rk] = s;
                }
            }
        }
        var fldParent = paraTargetPath;
        string? candidateHlParent = null;
        if (!string.IsNullOrEmpty(run.Path))
        {
            var idxFld = run.Path.LastIndexOf("/field[", StringComparison.Ordinal);
            if (idxFld > 0)
            {
                var derived = run.Path.Substring(0, idxFld);
                if (derived.Contains("/hyperlink["))
                    candidateHlParent = derived;
            }
        }
        // fldChar-chain fields surface with a flat /…/r[N] path; the
        // hyperlink hint is in Format._hyperlinkParent.
        if (candidateHlParent == null
            && run.Format.TryGetValue("_hyperlinkParent", out var fhlpObj)
            && fhlpObj != null)
        {
            var hint = fhlpObj.ToString();
            if (!string.IsNullOrEmpty(hint)) candidateHlParent = hint;
        }
        if (candidateHlParent != null)
        {
            // Re-base the candidate path onto paraTargetPath and verify a
            // prior `add hyperlink` row landed under that same paragraph.
            const string hlMarker = "/hyperlink[";
            var hlIdxStart = candidateHlParent.LastIndexOf(hlMarker, StringComparison.Ordinal);
            if (hlIdxStart > 0)
            {
                var hlEnd = candidateHlParent.IndexOf(']', hlIdxStart);
                if (hlEnd > hlIdxStart)
                {
                    var kStr = candidateHlParent.Substring(hlIdxStart + hlMarker.Length,
                        hlEnd - hlIdxStart - hlMarker.Length);
                    if (int.TryParse(kStr, out var kIdx))
                    {
                        var rebased = paraTargetPath
                            + candidateHlParent.Substring(hlIdxStart);
                        int emittedHls = items.Count(it => it.Type == "hyperlink"
                            && string.Equals(it.Parent, paraTargetPath, StringComparison.Ordinal));
                        if (emittedHls >= kIdx)
                            fldParent = rebased;
                    }
                }
            }
        }
        if (fieldProps != null)
        {
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = fldParent,
                Type = "field",
                Props = fieldProps
            });
        }
        else if (!string.IsNullOrEmpty(run.Text))
        {
            // Unparseable instruction — fall back to plain text so the
            // paragraph still renders the cached value rather than going empty.
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = fldParent,
                Type = "r",
                Props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { ["text"] = run.Text! }
            });
        }
        return true;
    }

    private static bool TryEmitPictureRun(WordHandler word, DocumentNode run, string paraTargetPath, string parentPath, int targetIndex, List<BatchItem> items, BodyEmitContext? ctx, string? sharedAttachPara = null)
    {
        // Drawing-bearing runs surface as type="picture" regardless of
        // whether the Drawing wraps an image (Blip) or a chart (c:chart).
        // Try the image path first; if no embedded image part the run is a
        // chart anchor — pull the next pre-resolved ChartSpec and emit a
        // typed `add chart` row.
        // Drawings wrapped in <mc:AlternateContent>/<mc:Choice> surface as a
        // plain "run" node (Run.GetFirstChild<Drawing>() returns null because
        // the Drawing lives inside the AlternateContent wrapper), so we also
        // accept "run" / "r" when the raw XML carries an obvious textbox
        // marker. Non-drawing runs without those markers short-circuit out
        // of the textbox/picture path immediately.
        if (run.Type != "picture")
        {
            if (run.Type != "run" && run.Type != "r") return false;
            var probeXml = word.GetElementXml(run.Path);
            if (string.IsNullOrEmpty(probeXml)) return false;
            bool isTextbox = IsTextboxDrawing(probeXml);
            // A genuine text/hyperlink run (no drawing payload at all) belongs
            // to the plain-run path — short-circuit out so EmitPlainOrHyperlink
            // run handles it. Only drawing-bearing runs continue here.
            if (!isTextbox && !IsDrawingBearingRun(probeXml)) return false;
            // BUG-R3 (dump emits `add textbox` into a table cell): see the
            // sibling site below — a cell-hosted textbox must attach to its
            // containing paragraph (paraTargetPath inside the cell), never as a
            // direct <w:tc> child. LibreOffice exports textboxes wrapped in
            // <mc:AlternateContent>, so they reach this "run"-typed branch.
            string? attachParaAc = sharedAttachPara
                ?? (parentPath.Contains("/tc[", StringComparison.Ordinal) ? paraTargetPath : null);
            if (isTextbox && TryEmitTextbox(word, run, probeXml, parentPath, items, ctx, attachParaAc, paraTargetPath))
                return true;
            // AlternateContent-wrapped non-textbox shapes — connector lines
            // (<wps:cNvCnPr>, e.g. a letterhead separator), autoshapes, groups
            // — and textboxes whose typed emit failed fall back to a raw-set
            // append so the shape survives round-trip instead of vanishing.
            // BUG-DUMP-R47-1: the host now resolves via ResolveRawSetHost
            // (body / header / footer / table cell), mirroring the wps:wsp
            // DrawingML-shape path below — a cell/header/footer-hosted legacy
            // VML/non-textbox shape was previously dropped because this gate
            // hardcoded parentPath == "/body" and the body anchor. The external-rel
            // guards are unchanged: drawings carrying an r:embed/r:id we can't
            // re-anchor still fall through to the warn+drop below.
            // BUG-R1-01 (preserved): ResolveRawSetHost("/body") returns exactly
            // ("/document", "/w:document/w:body/w:p[last()]") — attach to the
            // most-recently-emitted paragraph via last(), NOT the navigation
            // pIndex (targetIndex). pIndex deliberately does not advance for an
            // oMathPara-bearing paragraph (EmitBody: display equations resolve as
            // /body/oMathPara[N], not /body/p[N]), but in the LITERAL XML the
            // equation is a real <w:p> wrapping <m:oMathPara>. So a numeric
            // w:p[{pIndex}] under-counts by one per preceding equation and the
            // shape lands on the wrong paragraph. The host paragraph was just
            // added by this same EmitParagraph call, so it is always the last
            // w:p at replay time — the same last()-relative attach the typed
            // picture/textbox path uses. So body-hosted shapes round-trip
            // byte-identically to the previous hardcoded anchor.
            if (!probeXml.Contains("r:embed") && !probeXml.Contains("r:id")
                && !DrawingHasUnreconstructableRel(probeXml)
                && ctx != null
                && ResolveRawSetHost(parentPath, ctx) is { } drawHost)
            {
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = drawHost.Part,
                    Xpath = drawHost.XPath,
                    Action = "append",
                    Xml = probeXml
                });
                return true;
            }
            // Drawing-bearing but not safely raw-set-able inline (lives in a
            // header/footer/cell, or carries an external relationship we can't
            // re-anchor). Flag the loss rather than silently dropping it.
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                "drawing", run.Path,
                "non-textbox drawing/shape could not be serialized for round-trip; it will be missing from the replayed document"));
            return true;
        }
        // BUG-DUMP-R31-4: a run whose drawing is a <wps:wsp> DrawingML SHAPE
        // (preset geometry + outline + spPr fill — e.g. an ellipse with a
        // blipFill image and a red a:ln) surfaces as type="picture" because
        // GetImageBinary finds the blipFill's embedded image. Routed through the
        // `add picture` path below it is FLATTENED to a plain <pic:pic> rect: the
        // prstGeom, the a:ln outline and the shape semantics are all lost. A
        // wps:wsp shape must round-trip AS A SHAPE — raw-set its drawing verbatim
        // so geometry + outline + fill structure survive — not be downgraded to a
        // picture. Distinguish a genuine inline <pic:pic> (keep the picture path)
        // from a wps:wsp shape (raw-set passthrough). Mirrors the VML-textbox /
        // non-textbox-shape raw-set convention.
        {
            var shapeXml = word.GetElementXml(run.Path);
            if (!string.IsNullOrEmpty(shapeXml)
                && IsWpsShapeDrawing(shapeXml)
                && ctx != null
                && ResolveRawSetHost(parentPath, ctx) is { } shapeHost)
            {
                // The shape's blipFill references an embedded image part via
                // r:embed; a verbatim raw-set would dangle it. Ship the run
                // through the inlined-parts carrier (verbatim runXml +
                // part{N} image bytes, rel ids rewritten on replay) so the
                // bitmap fill survives — same shape as the vmlshape carrier.
                // A wps:wsp with NO external rel (plain fill) raw-sets verbatim.
                if (HasExternalRelRef(shapeXml)
                    && word.GetDrawingShapeEmitData(run.Path) is { } shpData)
                {
                    items.Add(new BatchItem
                    {
                        Command = "add",
                        Parent = paraTargetPath,
                        Type = "drawingshape",
                        Props = PackInlinedPartsProps(shpData),
                    });
                    return true;
                }
                // Fallback (unresolvable reference): scrub the blipFill rel
                // (neutral solidFill placeholder keeps the shape valid) and
                // warn that the image bitmap is dropped — geometry and outline
                // still round-trip.
                var scrubbed = ScrubDrawingBlipFillRels(shapeXml, out var blipDropped);
                if (blipDropped)
                    ctx.Warnings.Add(new DocxUnsupportedWarning(
                        Element: "shape.blipFill",
                        Path: run.Path,
                        Reason: "DrawingML shape (wps:wsp) image fill (a:blipFill r:embed) dropped — shape-image round-trip is not supported; the shape's geometry and outline are preserved and the fill was replaced with a neutral placeholder to keep the rebuilt drawing valid"));
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = shapeHost.Part,
                    Xpath = shapeHost.XPath,
                    Action = "append",
                    Xml = scrubbed
                });
                return true;
            }
        }
        var binary = word.GetImageBinary(run.Path);
        if (binary.HasValue)
        {
            var (bytes, contentType) = binary.Value;
            var dataUri = $"data:{contentType};base64,{Convert.ToBase64String(bytes)}";
            var picProps = FilterEmittableProps(run.Format);
            picProps.Remove("id");
            picProps.Remove("contentType");
            picProps.Remove("fileSize");
            picProps["src"] = dataUri;
            // BUG-DUMP-R28-1: the picture node's width/height come from
            // CreateImageNode formatted as 1-decimal CENTIMETRES (EMU /
            // EmuPerCmF, "F1"). Replaying that cm string through ParseEmu snaps
            // cx/cy back to a 360000-EMU (0.1cm) grid, shifting every inline
            // drawing by up to ~17800 EMU. The cumulative drift over many
            // drawings moves where page breaks land — the whole document's
            // text reflows vertically. Emit the EXACT cx/cy straight from
            // <wp:extent> as "<emu>emu" so ParseEmu reconstructs the original
            // value byte-for-byte (mirrors the textbox path, which already
            // emits raw EMU). AddPicture applies the same parsed EMU to both
            // <wp:extent> and the matching <a:ext> in <a:xfrm>. Covers inline
            // drawings in /body AND inside table cells (same emit path).
            var picXml = word.GetElementXml(run.Path);
            if (!string.IsNullOrEmpty(picXml))
            {
                var extMatch = System.Text.RegularExpressions.Regex.Match(
                    picXml,
                    @"wp:extent\b[^>]*\bcx=""(\d+)""[^>]*\bcy=""(\d+)""");
                if (extMatch.Success)
                {
                    picProps["width"] = extMatch.Groups[1].Value + "emu";
                    picProps["height"] = extMatch.Groups[2].Value + "emu";
                }
                // BUG-DUMP-R29-1: Word's inline-picture layout HEIGHT depends on
                // <wp:effectExtent l/t/r/b> (the drawing's visual overflow/effect
                // margin) even when <wp:extent> is identical. The rebuild
                // hardcoded effectExtent to 0/0/0/0 (CreateImageRun /
                // CreateAnchorImageRun), so each affected drawing rendered
                // ~35px shorter, pulling downstream content up across page
                // boundaries — a visible document-wide vertical drift. Capture
                // the source l/t/r/b as a single "l,t,r,b" EMU prop so AddPicture
                // can restore it byte-for-byte. Absent (or all-zero) effectExtent
                // is the interactive default and needs no prop.
                var eeMatch = System.Text.RegularExpressions.Regex.Match(
                    picXml,
                    @"wp:effectExtent\b[^>]*\bl=""(-?\d+)""[^>]*\bt=""(-?\d+)""[^>]*\br=""(-?\d+)""[^>]*\bb=""(-?\d+)""");
                if (eeMatch.Success
                    && !(eeMatch.Groups[1].Value == "0" && eeMatch.Groups[2].Value == "0"
                         && eeMatch.Groups[3].Value == "0" && eeMatch.Groups[4].Value == "0"))
                {
                    picProps["effectExtent"] =
                        $"{eeMatch.Groups[1].Value},{eeMatch.Groups[2].Value}," +
                        $"{eeMatch.Groups[3].Value},{eeMatch.Groups[4].Value}";
                }
                // BUG-DUMP-R39-1: an anchored (floating) picture positioned with
                // an absolute offset stores it as <wp:positionH><wp:posOffset>EMU.
                // CreateImageNode emits Format["hPosition"]/["vPosition"] as
                // 1-decimal CENTIMETRES (EmuPerCmF, "F1"). Replaying that cm string
                // through AddPicture's ParseEmu snaps the offset back to a
                // 360000-EMU (0.1cm) grid — e.g. posOffset 1234567 -> 1224000,
                // visibly shifting the floating image. Mirror the R28/R38 raw-EMU
                // pattern: pull the EXACT posOffset straight from the source XML,
                // scoped to its own positionH/positionV block so H->hPosition and
                // V->vPosition map correctly, and emit "<emu>emu" so ParseEmu
                // reconstructs the original offset byte-for-byte. Only override
                // when a posOffset is present — anchors using <wp:align>
                // (left/center/right) carry no posOffset, and inline pictures have
                // no positionH/V at all, so the regex simply won't match (safe).
                var hPosMatch = System.Text.RegularExpressions.Regex.Match(
                    picXml,
                    @"<wp:positionH\b[^>]*>.*?<wp:posOffset>(-?\d+)</wp:posOffset>.*?</wp:positionH>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                if (hPosMatch.Success)
                    picProps["hPosition"] = hPosMatch.Groups[1].Value + "emu";
                var vPosMatch = System.Text.RegularExpressions.Regex.Match(
                    picXml,
                    @"<wp:positionV\b[^>]*>.*?<wp:posOffset>(-?\d+)</wp:posOffset>.*?</wp:positionV>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                if (vPosMatch.Success)
                    picProps["vPosition"] = vPosMatch.Groups[1].Value + "emu";
                // BUG-DUMP-R45-4: the reconstructor rebuilds <a:blip>/<pic:spPr>
                // from a FIXED subset, silently dropping image-level visual
                // content the source carried — recolor/alpha inside <a:blip>
                // (duotone/biLevel/alphaModFix) and the spPr drop-shadow
                // (<a:effectLst><a:outerShdw>). Mirror the chart verbatim-capture
                // pattern: grab them as verbatim XML props so AddPicture can
                // re-inject at schema-correct positions. Only set each prop when
                // the source actually has that content (a plain picture stays
                // plain — no spurious empty effectLst / blip children).
                var blipInner = CapturePicBlipInnerXml(picXml);
                if (!string.IsNullOrEmpty(blipInner))
                    picProps["blipEffects"] = blipInner!;
                var spEffectLst = CapturePicSpPrEffectLst(picXml);
                if (!string.IsNullOrEmpty(spEffectLst))
                    picProps["spEffects"] = spEffectLst!;
                // The fixed spPr rebuild also drops xfrm flip flags
                // (<a:xfrm flipH="1"> — mirrored logos), a content extent
                // (<a:ext>) that legitimately differs from the frame's
                // wp:extent, bwMode, and explicit <a:noFill/>/<a:ln> blocks.
                // Capture the whole <pic:spPr> verbatim; AddPicture swaps its
                // rebuilt spPr for this block (and then skips the narrower
                // spEffects injection — the effectLst already rides inside).
                var spPrWhole = System.Text.RegularExpressions.Regex.Match(
                    picXml,
                    @"<pic:spPr[^>]*?>.*?</pic:spPr>|<pic:spPr[^>]*?/>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                if (spPrWhole.Success)
                    picProps["spPrXml"] = spPrWhole.Value;
                // Anchor wrap distances (distT/distB/distL/distR — the gap
                // between a floating image and the text wrapping around it).
                // CreateAnchorImageRun hardcodes T/B=0, L/R=114300; a figure
                // with asymmetric distances shifted every adjacent line.
                // Capture as "T,B,L,R" so AddPicture restores them. Inline
                // pictures have no <wp:anchor>, so the match simply won't fire.
                var anchorMatch = System.Text.RegularExpressions.Regex.Match(
                    picXml, @"<wp:anchor\b([^>]*)>");
                if (anchorMatch.Success)
                {
                    string DistAttr(string n) =>
                        System.Text.RegularExpressions.Regex.Match(
                            anchorMatch.Groups[1].Value, n + "=\"(\\d+)\"") is { Success: true } mm
                            ? mm.Groups[1].Value : "0";
                    var wt = DistAttr("distT"); var wb = DistAttr("distB");
                    var wl = DistAttr("distL"); var wr = DistAttr("distR");
                    if (!(wt == "0" && wb == "0" && wl == "114300" && wr == "114300"))
                        picProps["wrapDist"] = $"{wt},{wb},{wl},{wr}";
                }
            }
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "picture",
                Props = picProps
            });
            return true;
        }

        // Only consume a ChartSpec if the run is genuinely a chart. Picture-
        // typed runs that aren't images can also be background images, OLE
        // objects, SmartArt, watermark anchors etc — falling through
        // unconditionally would misalign chart positions.
        if (ctx != null && word.IsChartRun(run.Path)
            && ctx.ChartCursor.Index < ctx.ChartSpecs.Count)
        {
            var spec = ctx.ChartSpecs[ctx.ChartCursor.Index];
            ctx.ChartCursor.Index++;
            var chartProps = BuildChartProps(spec);
            // BUG-DUMP-R38-1: the chart node's width/height come from
            // WordHandler.Query formatted as 1-decimal CENTIMETRES (cx/cy /
            // EmuPerCmF, "F1"). Replaying that cm string through ParseEmu in
            // AddChart snaps the <wp:extent cx cy> back to a 360000-EMU (0.1cm)
            // grid — e.g. cx=5486400 (15.24cm) -> 5472000 (15.2cm), a 14400-EMU
            // width loss that visibly reflows the chart's title/bars/labels.
            // Emit the EXACT cx/cy straight from <wp:extent> as "<emu>emu" so
            // ParseEmu reconstructs the original value byte-for-byte. Mirrors
            // the inline-picture path (R28) above. effectExtent is out of scope:
            // AddChart accepts no effectExtent prop and these chart anchors
            // carry all-zero effectExtent.
            var chartXml = word.GetElementXml(run.Path);
            if (!string.IsNullOrEmpty(chartXml))
            {
                var chartExtMatch = System.Text.RegularExpressions.Regex.Match(
                    chartXml,
                    @"wp:extent\b[^>]*\bcx=""(\d+)""[^>]*\bcy=""(\d+)""");
                if (chartExtMatch.Success)
                {
                    chartProps["width"] = chartExtMatch.Groups[1].Value + "emu";
                    chartProps["height"] = chartExtMatch.Groups[2].Value + "emu";
                }
            }
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "chart",
                Props = chartProps
            });
            return true;
        }
        // Drawing without image part and not a chart — most likely a wps
        // shape. BUG-DUMP-TXBX: textbox-bearing drawings get a typed
        // `add textbox` row plus recursive inner-paragraph/run emits so
        // round-trip preserves structure (raw-set fallback was emitting
        // BOTH the full <w:drawing> XML AND flattening the textbox's
        // inner runs back onto the host paragraph). Non-textbox shapes
        // still fall through to the raw-set append.
        var rawXml = word.GetElementXml(run.Path);
        if (!string.IsNullOrEmpty(rawXml) && IsTextboxDrawing(rawXml))
        {
            // BUG-R3 (dump emits `add textbox` into a table cell): a textbox is
            // a drawing carried by a run inside a paragraph, never a direct
            // <w:tc> child — `add textbox` with a bare cell parent is rejected
            // ("table cells only accept paragraphs, tables, or SDTs"). When the
            // host is a cell, attach the textbox to the run's containing
            // paragraph (paraTargetPath, which lives inside the cell) so
            // AddTextbox's paragraph-attach path (ResolveDrawingHostFromParagraph)
            // anchors it correctly — mirrors the body single-drawing path. Use
            // any already-computed sharedAttachPara (side-by-side multi-drawing
            // host) first; otherwise fall back to paraTargetPath for cells.
            string? attachPara = sharedAttachPara
                ?? (parentPath.Contains("/tc[", StringComparison.Ordinal) ? paraTargetPath : null);
            if (TryEmitTextbox(word, run, rawXml, parentPath, items, ctx, attachPara, paraTargetPath))
                return true;
        }
        // BUG-R3 (linked external image): a <w:drawing> carrying an
        // unreconstructable relationship (linked-image r:link, SmartArt
        // r:dm/r:lo/r:qs/r:cs) must NOT be raw-set verbatim — its relationship
        // target isn't recreated, so the replayed drawing would dangle and fail
        // [Semantic] validation. Drop it cleanly and surface the loss (mirrors
        // the SmartArt/non-textbox warn-and-skip on the AlternateContent path
        // above) instead of falling through to a silent raw-set or a silent
        // return.
        // SmartArt: the diagram's data/layout/quickStyle/colors parts (plus the
        // data part's rendered-drawing child) ship base64-inlined in a
        // self-contained `add diagram`, mirroring the `add activex` carrier —
        // previously the whole diagram was warn-dropped as unreconstructable.
        if (!string.IsNullOrEmpty(rawXml) && rawXml.Contains("relIds")
            && word.GetDiagramEmitData(run.Path) is { } dgmData)
        {
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "diagram",
                Props = PackInlinedPartsProps(dgmData),
            });
            return true;
        }
        if (!string.IsNullOrEmpty(rawXml) && DrawingHasUnreconstructableRel(rawXml))
        {
            string detail = rawXml.Contains("r:link")
                ? "linked external image (<a:blip r:link>) references an external relationship that is not carried through dump→batch; the image will be missing from the replayed document"
                : "drawing references a relationship (SmartArt/diagram) that cannot be reconstructed; it will be missing from the replayed document";
            ctx?.Warnings.Add(new DocxUnsupportedWarning("drawing", run.Path, detail));
            return true;
        }
        if (!string.IsNullOrEmpty(rawXml) &&
            parentPath == "/body" &&
            !rawXml.Contains("r:embed") && !rawXml.Contains("r:id"))
        {
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = "/document",
                // BUG-R1-01 (second site, plain <w:drawing> wps shape): same
                // last()-relative attach as above — pIndex under-counts literal
                // w:p when a display equation precedes this shape's paragraph.
                Xpath = "/w:document/w:body/w:p[last()]",
                Action = "append",
                Xml = rawXml
            });
        }
        return true;
    }

    // SmartArt (diagram) drawings reference their data / layout / colors /
    // quickStyle parts via r:dm / r:lo / r:qs / r:cs — relationships, but NOT
    // r:embed/r:id. The dump never reconstructs those parts, so raw-setting the
    // drawing verbatim leaves dangling relationships: the SDK validator NREs in
    // RelationshipTypeConstraint and real Word refuses to open the file ("may be
    // corrupt"). Treat such drawings as unreconstructable so the caller flags a
    // loss warning instead of emitting a corrupt file. (Plain shapes/connectors
    // carry no relationships and still round-trip via raw-set.)
    private static bool DrawingHasUnreconstructableRel(string xml) =>
        xml.Contains("r:dm") || xml.Contains("r:lo") ||
        xml.Contains("r:qs") || xml.Contains("r:cs") ||
        // BUG-R3 (linked external image): <a:blip r:link="rIdN"> references an
        // EXTERNAL image relationship (TargetMode="External"), not an embedded
        // image part. The dump never recreates that relationship, so raw-setting
        // the drawing verbatim leaves a dangling r:link → [Semantic] "the
        // relationship referenced by r:link does not exist". Full linked-image
        // round-trip (carrying the external rel) is a separate feature; treat
        // the drawing as unreconstructable so the raw-set sites skip it cleanly
        // and the caller surfaces a loss warning (mirrors the SmartArt path).
        xml.Contains("r:link");

    // BUG-DUMP-R45-4: capture the inner XML of the FIRST <a:blip> inside a
    // picture's drawing (the recolor/alpha children — duotone / biLevel /
    // alphaModFix / lum* / clrChange — that AddPicture's fixed blip rebuild
    // drops). The r:embed is an ATTRIBUTE of <a:blip>, not a child, so it is
    // preserved automatically by AddPicture and excluded here. Returns null
    // when the blip is self-closing or empty (a plain picture stays plain).
    private static string? CapturePicBlipInnerXml(string picXml)
    {
        // Match the first non-self-closing <a:blip …>…</a:blip> and grab the
        // inner content. A self-closing <a:blip … /> has no children → skip.
        var m = System.Text.RegularExpressions.Regex.Match(
            picXml,
            @"<a:blip\b[^>]*?>(.*?)</a:blip>",
            System.Text.RegularExpressions.RegexOptions.Singleline);
        if (!m.Success) return null;
        var inner = m.Groups[1].Value.Trim();
        return inner.Length > 0 ? inner : null;
    }

    // BUG-DUMP-R45-4: capture the verbatim <a:effectLst>…</a:effectLst> sitting
    // inside the picture's <pic:spPr> (the drop-shadow / glow / reflection that
    // AddPicture's fixed spPr rebuild drops). Returns null when no effectLst is
    // present so a plain picture stays plain.
    private static string? CapturePicSpPrEffectLst(string picXml)
    {
        // Scope to the <pic:spPr> block first so we don't accidentally grab an
        // effectLst from an unrelated sibling (none today, but keeps the regex
        // honest if spPr structure grows).
        var spMatch = System.Text.RegularExpressions.Regex.Match(
            picXml,
            @"<pic:spPr\b[^>]*?>(.*?)</pic:spPr>",
            System.Text.RegularExpressions.RegexOptions.Singleline);
        var scope = spMatch.Success ? spMatch.Groups[1].Value : picXml;
        var m = System.Text.RegularExpressions.Regex.Match(
            scope,
            @"<a:effectLst\b[^>]*?>.*?</a:effectLst>|<a:effectLst\b[^>]*?/>",
            System.Text.RegularExpressions.RegexOptions.Singleline);
        return m.Success ? m.Value : null;
    }

    // BUG-DUMP-R26-6: a legacy VML textbox is a <w:pict> carrying a
    // <v:textbox> (with <w:txbxContent>) or a <v:shape type="#_x0000_t202">
    // (the VML textbox preset). Distinct from the modern DrawingML textbox
    // (wps:wsp/wps:txbx), which the typed `add textbox` path handles. We
    // detect VML so it can round-trip verbatim via raw-set instead of being
    // force-converted (and emptied) through the DrawingML emit.
    private static bool IsVmlTextbox(string rawXml)
    {
        if (!rawXml.Contains("<w:pict", StringComparison.Ordinal)
            && !rawXml.Contains("<v:", StringComparison.Ordinal))
            return false;
        return rawXml.Contains("<v:textbox", StringComparison.Ordinal)
            || rawXml.Contains("_x0000_t202", StringComparison.Ordinal);
    }

    private static bool IsTextboxDrawing(string rawXml)
    {
        // Mirrors WordHandler.CountTextboxesInHost / Navigation's textbox
        // selector — a textbox is a wps:wsp with txBox=1 cNvSpPr or a
        // wps:txbx child carrying w:txbxContent.
        return rawXml.Contains("txBox=\"1\"")
            || rawXml.Contains("<wps:txbx")
            || rawXml.Contains("txbxContent");
    }

    /// <summary>
    /// True when a run carries any drawing/shape payload (image, chart,
    /// textbox, connector line, autoshape, group) — i.e. a
    /// <c>&lt;w:drawing&gt;</c>, legacy VML <c>&lt;w:pict&gt;</c>, or an
    /// <c>&lt;mc:AlternateContent&gt;</c>/<c>&lt;wps:&gt;</c> wrapper around one.
    /// Used to decide whether a non-textbox run still needs preserving via a
    /// raw-set append rather than being treated as a plain text run.
    /// </summary>
    private static bool IsDrawingBearingRun(string rawXml)
    {
        return rawXml.Contains("<w:drawing")
            || rawXml.Contains("<w:pict")
            || rawXml.Contains("<mc:AlternateContent")
            || rawXml.Contains("<wps:");
    }

    /// <summary>
    /// BUG-DUMP-R31-4: True when a run's drawing is a modern DrawingML SHAPE
    /// (<c>&lt;wps:wsp&gt;</c> with preset geometry / spPr) rather than a plain
    /// inline <c>&lt;pic:pic&gt;</c> picture. Such a shape may carry a blipFill
    /// image (so it surfaces as type="picture"), but its geometry + outline are
    /// shape semantics that the picture-emit path would flatten away. Textboxes
    /// (which carry their own typed/raw-set path) are excluded — they are handled
    /// upstream by IsTextboxDrawing.
    /// </summary>
    private static bool IsWpsShapeDrawing(string rawXml)
    {
        if (string.IsNullOrEmpty(rawXml)) return false;
        if (IsTextboxDrawing(rawXml)) return false;          // textbox has its own path
        if (!rawXml.Contains("<wps:wsp", StringComparison.Ordinal)) return false;
        // A genuine shape carries a preset/custom geometry or its own shape
        // properties — that's what the picture path cannot represent.
        return rawXml.Contains("prstGeom", StringComparison.Ordinal)
            || rawXml.Contains("custGeom", StringComparison.Ordinal)
            || rawXml.Contains("<wps:spPr", StringComparison.Ordinal);
    }

    /// <summary>
    /// BUG-DUMP-R31-4: replace every <c>&lt;a:blipFill&gt;</c> whose blip
    /// carries a relationship reference (r:embed / r:link) with a neutral
    /// <c>&lt;a:solidFill&gt;&lt;a:srgbClr val="FFFFFF"/&gt;&lt;/a:solidFill&gt;</c>
    /// placeholder. The dump cannot carry the referenced image binary + part
    /// rels through a raw-set, so a verbatim passthrough would dangle the rel and
    /// corrupt the rebuilt drawing. Scrubbing keeps the shape a schema-valid
    /// filled shape (geometry + outline intact) at the cost of the image bitmap.
    /// Mirrors StripDanglingThemeBlipRefs in WordBatchEmitter.Resources.cs.
    /// </summary>
    private static string ScrubDrawingBlipFillRels(string xml, out bool dropped)
    {
        dropped = false;
        if (string.IsNullOrEmpty(xml) || !xml.Contains("blipFill", StringComparison.Ordinal))
            return xml;
        if (!xml.Contains(":embed=", StringComparison.Ordinal)
            && !xml.Contains(":link=", StringComparison.Ordinal))
            return xml;
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(xml);
            if (doc.Root == null) return xml;
            var aNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/drawingml/2006/main";
            var rNs = (System.Xml.Linq.XNamespace)"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            foreach (var blipFill in doc.Descendants(aNs + "blipFill").ToList())
            {
                var blip = blipFill.Element(aNs + "blip");
                var hasRel = blip != null && blip.Attributes().Any(a => a.Name.Namespace == rNs);
                if (!hasRel) continue;
                var placeholder = new System.Xml.Linq.XElement(aNs + "solidFill",
                    new System.Xml.Linq.XElement(aNs + "srgbClr",
                        new System.Xml.Linq.XAttribute("val", "FFFFFF")));
                blipFill.ReplaceWith(placeholder);
                dropped = true;
            }
            if (!dropped) return xml;
            return doc.Root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        catch
        {
            return xml;
        }
    }

    /// <summary>
    /// BUG-DUMP-TXBX: emit a typed <c>add textbox</c> row for the host of
    /// the current drawing run, followed by recursive inner-paragraph/run
    /// emits under <c>/&lt;host&gt;/textbox[N]</c>. Geometry props
    /// (width/height/wrap/anchor.x/anchor.y/fill) are extracted from the
    /// raw drawing XML so the rebuilt textbox keeps its layout.
    /// </summary>
    private static bool TryEmitTextbox(WordHandler word, DocumentNode run, string rawXml,
                                       string parentPath, List<BatchItem> items, BodyEmitContext? ctx,
                                       string? attachParaPath = null, string? paraTargetPath = null)
    {
        if (ctx == null) return false;

        // BUG-DUMP-R26-6: a LEGACY VML textbox (<w:pict> with <v:shape
        // type="#_x0000_t202"> / <v:textbox><w:txbxContent>) is a different
        // shape family than the modern DrawingML box `add textbox` produces.
        // The typed emit below parses DrawingML namespaces (wp:/wps:/a:) that
        // don't exist in VML, so it extracted ZERO props (no text, fill, or
        // stroke) and Navigation can't surface the VML txbxContent under
        // /<host>/textbox[N] for the recursive inner-content emit — the box came
        // back as an empty modern textbox with its content + fillcolor/
        // strokecolor gone. The faithful (and lossless) round-trip is a verbatim
        // raw-set of the whole <w:pict> run into the just-emitted host paragraph
        // — mirrors the non-textbox-shape / rich-inline-SDT raw-set append.
        // BUG-DUMP-R26-7: fire for /body, header/footer AND table-cell hosts
        // (ResolveRawSetHost), not only /body — VML page-number / watermark
        // boxes in headers are the most common real case.
        if (IsVmlTextbox(rawXml) && ResolveRawSetHost(parentPath, ctx) is { } vmlHost)
        {
            // BUG-DUMP-R26-7 (PART B): a VML shape with an external relationship
            // (e.g. <v:imagedata r:id> referencing an image part) would dangle in
            // the rebuilt part. Don't silently flatten — emit a deterministic
            // warning naming the loss, then fall through. Plain VML textboxes
            // (fillcolor/strokecolor, no r:id) still round-trip verbatim.
            if (HasExternalRelRef(rawXml))
            {
                // The rel refs are resolvable (hyperlinks inside the textbox
                // content, v:imagedata image parts): ship them through the
                // inlined-parts carrier so the shape round-trips with its
                // relationships recreated — previously warn-dropped.
                var vmlParent = attachParaPath ?? paraTargetPath;
                var vmlData = vmlParent != null ? word.GetVmlShapeEmitData(run.Path) : null;
                if (vmlData != null)
                {
                    items.Add(new BatchItem
                    {
                        Command = "add",
                        Parent = vmlParent,
                        Type = "vmlshape",
                        Props = PackInlinedPartsProps(vmlData),
                    });
                    return true;
                }
                ctx.Warnings.Add(new DocxUnsupportedWarning(
                    Element: "textbox.vmlContent",
                    Path: run.Path,
                    Reason: "legacy VML shape with an external relationship (image r:id / linked content) cannot round-trip verbatim; the relationship target is not carried through dump→batch, so the shape content is dropped on replay"));
            }
            else
            {
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = vmlHost.Part,
                    Xpath = vmlHost.XPath,
                    Action = "append",
                    Xml = rawXml
                });
                return true;
            }
        }

        // Only emit a typed `add textbox` for hosts AddTextbox itself
        // supports: /body, /body/tbl[..]/tc[N], /header[N], /footer[N].
        // Other parents fall through to the raw-set append.
        string hostPath = parentPath;
        if (!IsTextboxHostPath(hostPath)) return false;

        // BUG-D1-MULTIDRAWING-HOST: when N textboxes share a source
        // paragraph (side-by-side card layout), attach each to the same
        // already-emitted host paragraph (attachParaPath = /body/p[last()])
        // instead of /body — otherwise AddTextbox creates a fresh host per
        // textbox and the side-by-side layout fans out into N stacked
        // paragraphs. The textbox INDEX still scopes to hostPath so
        // /body/textbox[K] addressing remains continuous across the doc.
        string emitParent = attachParaPath ?? hostPath;

        // Allocate next 1-based textbox index for this host.
        int n = ctx.TextboxCounters.TryGetValue(hostPath, out var prev) ? prev + 1 : 1;
        ctx.TextboxCounters[hostPath] = n;
        string textboxPath = hostPath == "/" ? "/textbox[" + n + "]" : $"{hostPath}/textbox[{n}]";

        // Extract geometry / wrap / fill / anchor from the drawing XML so the
        // rebuilt textbox keeps its layout. Conservative best-effort — any
        // attribute we can't parse falls back to AddTextbox's defaults.
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(rawXml);
            System.Xml.Linq.XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
            System.Xml.Linq.XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

            var anchor = doc.Descendants(wp + "anchor").FirstOrDefault()
                      ?? (System.Xml.Linq.XElement?)doc.Descendants(wp + "inline").FirstOrDefault();
            var extent = doc.Descendants(wp + "extent").FirstOrDefault();
            if (extent != null)
            {
                var cx = extent.Attribute("cx")?.Value;
                var cy = extent.Attribute("cy")?.Value;
                if (!string.IsNullOrEmpty(cx)) props["width"] = cx + "emu";
                if (!string.IsNullOrEmpty(cy)) props["height"] = cy + "emu";
            }
            if (anchor != null)
            {
                var posHEl = anchor.Element(wp + "positionH");
                var posVEl = anchor.Element(wp + "positionV");
                var posH = posHEl?.Element(wp + "posOffset")?.Value;
                var posV = posVEl?.Element(wp + "posOffset")?.Value;
                // A floating drawing positions each axis EITHER by an absolute
                // posOffset OR by a relative <wp:align> (left/center/right /
                // top/bottom). Capture the align form too — without it a
                // center-aligned textbox round-trips as posOffset=0 (i.e. jumps
                // hard left). posOffset wins when both somehow appear.
                var alignH = posHEl?.Element(wp + "align")?.Value;
                var alignV = posVEl?.Element(wp + "align")?.Value;
                if (!string.IsNullOrEmpty(posH)) props["anchor.x"] = posH + "emu";
                else if (!string.IsNullOrEmpty(alignH)) props["hAlign"] = alignH;
                if (!string.IsNullOrEmpty(posV)) props["anchor.y"] = posV + "emu";
                else if (!string.IsNullOrEmpty(alignV)) props["vAlign"] = alignV;
                // Anchor reference frames. AddTextbox hardcodes column/paragraph;
                // without forwarding these a textbox anchored relativeFrom="page"
                // round-trips as relativeFrom="paragraph" and floats off-position
                // (the source posOffset is measured from a different origin).
                var hRel = posHEl?.Attribute("relativeFrom")?.Value;
                var vRel = posVEl?.Attribute("relativeFrom")?.Value;
                if (!string.IsNullOrEmpty(hRel)) props["hRelative"] = hRel;
                if (!string.IsNullOrEmpty(vRel)) props["vRelative"] = vRel;
                // wrap token
                if (anchor.Element(wp + "wrapSquare") != null) props["wrap"] = "square";
                else if (anchor.Element(wp + "wrapTight") != null) props["wrap"] = "tight";
                else if (anchor.Element(wp + "wrapTopAndBottom") != null) props["wrap"] = "topAndBottom";
                else if (anchor.Element(wp + "wrapNone") != null) props["wrap"] = "none";
            }
            var spPr = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "spPr");
            // Geometry preset (rect default). roundRect etc. otherwise reverted
            // to a sharp rectangle on rebuild.
            var prst = spPr?.Element(a + "prstGeom")?.Attribute("prst")?.Value;
            if (!string.IsNullOrEmpty(prst) && prst != "rect") props["geometry"] = prst;
            // Rotation (<a:xfrm rot>, 60000ths of a degree). Round-trips raw.
            var rot = spPr?.Element(a + "xfrm")?.Attribute("rot")?.Value;
            if (!string.IsNullOrEmpty(rot) && rot != "0") props["rotation"] = rot;
            // Fill: solidFill (with optional alpha) or gradFill inside wps:spPr.
            var solidFill = spPr?.Element(a + "solidFill");
            var srgbClr = solidFill?.Element(a + "srgbClr");
            if (srgbClr?.Attribute("val")?.Value is string fillHex && !string.IsNullOrEmpty(fillHex))
            {
                props["fill"] = fillHex;
                var alpha = srgbClr.Element(a + "alpha")?.Attribute("val")?.Value;
                if (!string.IsNullOrEmpty(alpha)) props["fill.opacity"] = alpha;
            }
            else
            {
                var grad = spPr?.Element(a + "gradFill");
                if (grad != null)
                {
                    var stops = grad.Element(a + "gsLst")?.Elements(a + "gs")
                        .Select(gs => (gs.Element(a + "srgbClr")?.Attribute("val")?.Value, gs.Attribute("pos")?.Value))
                        .Where(t => !string.IsNullOrEmpty(t.Item1))
                        .Select(t => $"{t.Item1}@{t.Item2 ?? "0"}");
                    if (stops != null)
                    {
                        var joined = string.Join(";", stops);
                        if (!string.IsNullOrEmpty(joined)) props["fill.gradient"] = joined;
                    }
                }
            }
            // Shadow (<a:effectLst><a:outerShdw>): emit the compact tuple the
            // add path understands so the drop shadow round-trips faithfully.
            var shdw = spPr?.Element(a + "effectLst")?.Element(a + "outerShdw");
            if (shdw != null)
            {
                var sClr = shdw.Element(a + "srgbClr");
                props["shadow"] = string.Join(";",
                    shdw.Attribute("blurRad")?.Value ?? "50800",
                    shdw.Attribute("dist")?.Value ?? "38100",
                    shdw.Attribute("dir")?.Value ?? "5400000",
                    sClr?.Attribute("val")?.Value ?? "000000",
                    sClr?.Element(a + "alpha")?.Attribute("val")?.Value ?? "40000");
            }
            // Line / border (<a:ln>): width (EMU, round-trips through ParseEmu),
            // solidFill color, and dash style. Without this the textbox outline
            // was dropped entirely on dump→batch — borders vanished and content
            // reflowed. <a:noFill/> means the box explicitly has no border.
            var ln = spPr?.Element(a + "ln");
            if (ln != null)
            {
                if (ln.Element(a + "noFill") != null)
                {
                    props["line.style"] = "none";
                }
                else
                {
                    var lnW = ln.Attribute("w")?.Value;
                    if (!string.IsNullOrEmpty(lnW)) props["line.width"] = lnW;
                    var lnClr = ln.Element(a + "solidFill")?.Element(a + "srgbClr")?.Attribute("val")?.Value;
                    if (!string.IsNullOrEmpty(lnClr)) props["line.color"] = lnClr;
                    // a:prstDash@val is already an OOXML dash name; AddTextbox's
                    // MapDashStyle accepts it (lower-cased) and re-emits it.
                    var dash = ln.Element(a + "prstDash")?.Attribute("val")?.Value;
                    if (!string.IsNullOrEmpty(dash)) props["line.style"] = dash;
                }
            }
            // bodyPr text insets. AddTextbox hardcodes Word defaults
            // (91440/45720); a source with zero insets (common on tight
            // letterhead title boxes) otherwise loses ~0.2in of usable width
            // and its text rewraps. Forward each present inset verbatim (EMU).
            var bodyPr = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "bodyPr");
            if (bodyPr != null)
            {
                foreach (var (attr, key) in new[] { ("lIns", "inset.left"), ("tIns", "inset.top"), ("rIns", "inset.right"), ("bIns", "inset.bottom") })
                {
                    var v = bodyPr.Attribute(attr)?.Value;
                    if (!string.IsNullOrEmpty(v)) props[key] = v;
                }
                // Vertical text flow + vertical anchor. Without these a vertical
                // (eaVert) box renders as char-wrapped horizontal text and a
                // centered box anchors to the top.
                var vert = bodyPr.Attribute("vert")?.Value;
                if (!string.IsNullOrEmpty(vert) && vert != "horz") props["textDirection"] = vert;
                var bAnchor = bodyPr.Attribute("anchor")?.Value;
                if (!string.IsNullOrEmpty(bAnchor) && bAnchor != "t") props["textAnchor"] = bAnchor;
                // BUG-DUMP-R25-6: wps:bodyPr/@wrap is the IN-shape text-wrap
                // mode (distinct from wp:wrapNone/wp:wrapSquare, the around-shape
                // wrap forwarded via `wrap`). AddTextbox hardcoded wrap="square";
                // a source wrap="none"/"tight"/"through" was clobbered to square.
                // Forward the value verbatim. The source bodyPr ALWAYS exists
                // here (we're inside `if (bodyPr != null)`), so emit an explicit
                // empty-string sentinel when @wrap is ABSENT — AddTextbox then
                // omits the attribute entirely, preserving absence instead of
                // re-injecting the schema default. (A `null` key, by contrast,
                // means "interactive caller never touched wrap" → legacy square.)
                var bodyWrap = bodyPr.Attribute("wrap")?.Value;
                props["bodyWrap"] = bodyWrap ?? "";
                // BUG-DUMP-R25-6: <a:spAutoFit/> (resize shape to fit text) is a
                // child of bodyPr. AddTextbox never emitted it, so the box kept
                // its fixed cy instead of shrinking to the content — same
                // vertical-extent / reflow defect. Forward as an autoFit flag.
                if (bodyPr.Elements().Any(e => e.Name.LocalName == "spAutoFit"))
                    props["autoFit"] = "true";
            }
            // docPr name → alt
            var docPr = doc.Descendants(wp + "docPr").FirstOrDefault();
            var altName = docPr?.Attribute("name")?.Value;
            if (!string.IsNullOrEmpty(altName) && altName != "Text Box") props["alt"] = altName;
        }
        catch
        {
            // Parsing failures: still emit the `add textbox` row with whatever
            // we managed to extract; defaults cover the rest.
        }

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = emitParent,
            Type = "textbox",
            Props = props.Count > 0 ? props : null
        });

        // Recurse over inner content. Get on /<host>/textbox[N] returns the
        // <w:txbxContent>; its children are the inner <w:p>. AddTextbox auto-
        // seeds one empty <w:p>, so the first source paragraph uses set-on-
        // existing (autoPresent: true) and the rest emit as fresh adds.
        try
        {
            var txbxNode = word.Get(textboxPath);
            var children = txbxNode.Children ?? new List<DocumentNode>();
            int innerPIdx = 0;
            int innerTblIdx = 0;
            bool firstParaSeen = false;
            foreach (var child in children)
            {
                if (child.Type == "paragraph" || child.Type == "p")
                {
                    innerPIdx++;
                    // The generic fallback fabricates child paths from the
                    // OOXML LocalName ("/body/txbxContent[N]/p[M]") which the
                    // Navigation layer can't re-resolve — the user-facing
                    // path segment is "textbox", not "txbxContent". Use the
                    // canonical /body/textbox[N]/p[M] form instead.
                    var sourceParaPath = $"{textboxPath}/p[{innerPIdx}]";
                    EmitParagraph(word, sourceParaPath, textboxPath, innerPIdx, items,
                                  autoPresent: !firstParaSeen, ctx);
                    firstParaSeen = true;
                }
                else if (child.Type == "table" || child.Type == "tbl")
                {
                    // BUG-D1-TXBX-TABLE: tables nested INSIDE a textbox were
                    // silently dropped on dump — the children loop only
                    // recognised paragraph types. Reuse EmitTable with the
                    // textbox path as containerPath so the resulting
                    // `add table` rows target /body/textbox[N]/tbl[K]
                    // (AddTable already accepts a TextBoxContent parent).
                    innerTblIdx++;
                    var sourceTblPath = $"{textboxPath}/tbl[{innerTblIdx}]";
                    EmitTable(word, sourceTblPath, innerTblIdx, items, ctx,
                              parentTablePath: null, containerPath: textboxPath);
                }
                else if (child.Type == "sdt")
                {
                    // BUG-R5B(BUG1): a block-level <w:sdt> nested inside a
                    // textbox (the canonical centered page-number footer:
                    // textbox → sdt → <w:p> with a PAGE field) was silently
                    // dropped — the inner walk only recognised paragraph and
                    // table types, so the SDT (and the field inside it) vanished
                    // on dump→batch and the footer lost its page number. There
                    // is no typed `add sdt` parent for a txbxContent host, and a
                    // PAGE-field SDT is "rich" content the typed text emit cannot
                    // reproduce anyway. Inject the SDT verbatim via raw-set into
                    // the just-created textbox's <w:txbxContent> (mirrors the
                    // rich-block-SDT raw-set in EmitSdt). The part is the host's
                    // own part (/document for body, /header[N] or /footer[N]
                    // otherwise); the (//w:txbxContent)[n] index selects the
                    // textbox we just emitted.
                    var sdtRaw = word.RawElementXml(child.Path);
                    if (!string.IsNullOrEmpty(sdtRaw))
                    {
                        var rawPart = hostPath == "/body" ? "/document" : hostPath;
                        // `add textbox` auto-seeds one empty <w:p> in the
                        // txbxContent. To keep document order, an SDT that
                        // precedes every source paragraph (the canonical
                        // page-number footer: sdt-then-empty-p) is prepended so
                        // it lands ahead of the seed paragraph; an SDT that
                        // follows already-emitted paragraphs is appended.
                        items.Add(new BatchItem
                        {
                            Command = "raw-set",
                            Part = rawPart,
                            Xpath = $"(//w:txbxContent)[{n}]",
                            Action = firstParaSeen ? "append" : "prepend",
                            Xml = sdtRaw
                        });
                    }
                    else
                    {
                        ctx.Warnings.Add(new DocxUnsupportedWarning(
                            Element: "textbox.sdt",
                            Path: child.Path,
                            Reason: "content control nested inside a textbox could not be serialized for round-trip; it will be missing from the replayed document"));
                    }
                }
            }
        }
        catch
        {
            // If the inner walk fails for any reason, the typed `add textbox`
            // still landed — round-trip recreates an empty textbox with the
            // right geometry, which beats the previous double-emit.
        }
        return true;
    }

    private static bool IsTextboxHostPath(string parentPath)
    {
        // Matches ResolveDrawingHost: /body, /body/tbl[..]/tr[..]/tc[N],
        // /header[N], /footer[N]. Reject anything else so non-supported
        // hosts fall through to the raw-set append.
        if (string.Equals(parentPath, "/body", StringComparison.Ordinal)) return true;
        if (parentPath.StartsWith("/header[", StringComparison.Ordinal)
            && parentPath.EndsWith("]", StringComparison.Ordinal)
            && !parentPath.Substring(8).Contains('/')) return true;
        if (parentPath.StartsWith("/footer[", StringComparison.Ordinal)
            && parentPath.EndsWith("]", StringComparison.Ordinal)
            && !parentPath.Substring(8).Contains('/')) return true;
        if (parentPath.Contains("/tc[", StringComparison.Ordinal)
            && parentPath.EndsWith("]", StringComparison.Ordinal)) return true;
        return false;
    }

    // BUG-DUMP-R25-7: a bare /header[N] or /footer[N] host (NOT a cell, which
    // already attaches its textbox to the containing paragraph via the
    // "/tc[" branch in TryEmitPictureRun). Used to keep a single footer/header
    // textbox inside its styled host paragraph instead of letting AddTextbox
    // fork a new unstyled paragraph.
    private static bool IsHeaderFooterHost(string parentPath)
    {
        // Reject nested (cell) paths: a bare /footer[N] has its only '/' at 0.
        if (parentPath.IndexOf('/', 1) >= 0) return false;
        return (parentPath.StartsWith("/header[", StringComparison.Ordinal)
                || parentPath.StartsWith("/footer[", StringComparison.Ordinal))
            && parentPath.EndsWith("]", StringComparison.Ordinal);
    }

    // BUG-DUMP-R26-7: resolve the (Part, XPath) for a raw-set APPEND into the
    // last paragraph of the host identified by <paramref name="parentPath"/>.
    // The verbatim-content raw-set fallbacks (rich field result FIX-2, nested
    // SDT FIX-5, VML textbox FIX-6) previously fired only for /body; this
    // extends them to header / footer / table-cell hosts so the content
    // survives there too instead of falling silently to the lossy extractor.
    //   /body                         -> ("/document",  "/w:document/w:body/w:p[last()]")
    //   /header[N]                    -> ("/header[N]", "/w:hdr/w:p[last()]")
    //   /footer[N]                    -> ("/footer[N]", "/w:ftr/w:p[last()]")
    //   table cell (ctx box set)      -> ("/document",  "(//w:tbl)[N]/w:tr[M]/w:tc[K]/w:p[last()]")
    // Returns null when the host isn't one we can address (the caller then
    // keeps its existing behaviour — typically the lossy typed emit).
    private static (string Part, string XPath)? ResolveRawSetHost(string parentPath, BodyEmitContext? ctx)
    {
        if (parentPath == "/body")
            return ("/document", "/w:document/w:body/w:p[last()]");
        if (IsHeaderFooterHost(parentPath))
        {
            var root = parentPath.StartsWith("/header[", StringComparison.Ordinal) ? "/w:hdr" : "/w:ftr";
            return (parentPath, $"{root}/w:p[last()]");
        }
        // Table cell: EmitTable stashes the current cell's global-ordinal XPath
        // in the context box while walking the cell's paragraphs. Append into
        // that cell's last paragraph. The container part is the document for a
        // body table; header/footer-hosted tables aren't carried here (the box
        // is only set for /document-hosted tables) so we conservatively skip.
        if (parentPath.Contains("/tc[", StringComparison.Ordinal)
            && ctx?.CurrentCellXPathBox is { } box && box[0] is { } cellXPath)
        {
            return ("/document", $"{cellXPath}/w:p[last()]");
        }
        return null;
    }

    /// <summary>
    /// R10-bug1: OLE / embedded-object runs (Type=="ole"). Full round-trip:
    /// pull the embedded payload + VML icon binaries and the frame metadata
    /// (progId / DrawAspect / dimensions / alt-name) and emit a self-contained
    /// `add ole` carrying the bytes as data: URIs (mirrors picture-run base64
    /// inlining). AddOle rebuilds the embedded part, icon part and &lt;w:object&gt;
    /// wrapper from these — no external src file needed.
    ///
    /// If the payload can't be resolved (orphaned relationship / unreadable
    /// part), fall back to keeping the host paragraph and emitting a warning
    /// naming the path, rather than silently dropping the object.
    /// </summary>
    private static bool TryEmitOleRun(DocumentNode run, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx, WordHandler word)
    {
        if (run.Type != "ole") return false;

        var data = word.GetOleEmitData(run.Path);
        if (data != null)
        {
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["src"] = $"data:{data.EmbeddedContentType};base64,{Convert.ToBase64String(data.EmbeddedBytes)}",
                ["oleKind"] = data.OleKind,
                ["contentType"] = data.EmbeddedContentType,
            };
            if (!string.IsNullOrEmpty(data.EmbeddedExt)) props["embedExt"] = data.EmbeddedExt;
            if (!string.IsNullOrEmpty(data.ProgId)) props["progId"] = data.ProgId!;
            if (!string.IsNullOrEmpty(data.Display)) props["display"] = data.Display!;
            if (!string.IsNullOrEmpty(data.Width)) props["width"] = data.Width!;
            if (!string.IsNullOrEmpty(data.Height)) props["height"] = data.Height!;
            if (!string.IsNullOrEmpty(data.Name)) props["name"] = data.Name!;
            if (data.IconBytes is { Length: > 0 })
                props["icon"] = $"data:{data.IconContentType ?? "image/png"};base64,{Convert.ToBase64String(data.IconBytes)}";
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "ole",
                Props = props,
            });
            return true;
        }

        // An ActiveX form control (<w:object> hosting <w:control r:id> with a
        // VML preview image, no o:OLEObject) has no embedded OLE payload, so
        // GetOleEmitData returns null — previously every such control (radio
        // buttons, checkboxes in form templates) was warn-dropped and its table
        // cell rebuilt empty. Emit a self-contained `add activex` instead:
        // verbatim run XML plus every referenced part's bytes base64-inlined,
        // mirroring the `add ole` data: URI design.
        var axData = word.GetActiveXEmitData(run.Path);
        if (axData != null)
        {
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "activex",
                Props = PackInlinedPartsProps(axData),
            });
            return true;
        }

        if (ctx != null)
        {
            var progId = run.Format.TryGetValue("progId", out var pid) ? pid?.ToString() : null;
            var reason = progId != null
                ? $"ole run dropped (progId={progId}); embedded payload could not be resolved for round-trip"
                : "ole run dropped; embedded payload could not be resolved for round-trip";
            ctx.Warnings.Add(new DocxUnsupportedWarning(
                Element: "ole",
                Path: run.Path,
                Reason: reason));
        }
        return true;
    }

    // Shared prop packing for the inlined-parts carriers (`add activex`,
    // `add diagram`): verbatim run XML + one part{N}.relId/part{N}.data pair
    // per referenced package part, with part{N}.child{M}.* for nested parts.
    private static Dictionary<string, string> PackInlinedPartsProps(WordHandler.ActiveXEmitData data)
    {
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["runXml"] = data.RunXml,
        };
        int pi = 0;
        foreach (var part in data.Parts)
        {
            pi++;
            props[$"part{pi}.relId"] = part.RelId;
            props[$"part{pi}.data"] =
                $"data:{part.ContentType};base64,{Convert.ToBase64String(part.Bytes)}";
            int ci = 0;
            foreach (var child in part.Children)
            {
                ci++;
                props[$"part{pi}.child{ci}.relId"] = child.RelId;
                props[$"part{pi}.child{ci}.data"] =
                    $"data:{child.ContentType};base64,{Convert.ToBase64String(child.Bytes)}";
            }
        }
        int ei = 0;
        foreach (var ext in data.Externals)
        {
            ei++;
            props[$"ext{ei}.relId"] = ext.RelId;
            props[$"ext{ei}.type"] = ext.Type;
            props[$"ext{ei}.target"] = ext.Target;
        }
        return props;
    }

    private static bool TryEmitNoteRefRun(WordHandler word, DocumentNode run, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // Footnote/endnote reference runs carry a <w:footnoteReference> /
        // <w:endnoteReference> child. Emit them as a typed footnote/endnote
        // add anchored on the host paragraph and pull the body text from the
        // pre-resolved ordered list — see BodyEmitContext for the
        // document-order assumption.
        //
        // BUG-R7B(BUG1): detection cannot rely on rStyle == "FootnoteReference":
        // that is only the default style name. Real documents reference the
        // note with an arbitrary style id (e.g. rStyle="a5"), so the run's
        // rStyle never matched and the reference (plus its note body) was
        // silently dropped — Get surfaces no Format key for the reference
        // element, so probe the raw run XML for the actual reference child.
        if (ctx == null) return false;
        var rStyle = run.Format.TryGetValue("rStyle", out var rs) ? rs?.ToString() : null;
        var noteKind = ClassifyNoteRefRun(word, run, rStyle);
        // BUG-R12A(BUG3): emit the note body STRUCTURALLY (per-run rPr +
        // multi-paragraph) instead of flattening it to one `text` prop. The
        // cursor is the 0-based document-order reference index; source AND target
        // note are the (cursor+1)-th note (Query and the body walk both run in
        // source order, one `add <kind>` per reference).
        if (noteKind == NoteRefKind.Footnote)
        {
            int idx = ++ctx.FootnoteCursor.Index; // 1-based source/target index
            EmitNoteReference(word, "footnote", idx, idx, paraTargetPath, items, run);
            return true;
        }
        if (noteKind == NoteRefKind.Endnote)
        {
            int idx = ++ctx.EndnoteCursor.Index; // 1-based source/target index
            EmitNoteReference(word, "endnote", idx, idx, paraTargetPath, items, run);
            return true;
        }
        return false;
    }

    private enum NoteRefKind { None, Footnote, Endnote }

    // BUG-R7B(BUG1): a run is a footnote/endnote reference when it contains a
    // <w:footnoteReference>/<w:endnoteReference> child — the rStyle is only a
    // weak hint (defaults to FootnoteReference/EndnoteReference but is an
    // arbitrary style id in real documents). Probe the raw XML; fall back to
    // the rStyle name when raw XML is unavailable.
    private static NoteRefKind ClassifyNoteRefRun(WordHandler word, DocumentNode run, string? rStyle)
    {
        if (run.Type != "run" && run.Type != "r") return NoteRefKind.None;
        var raw = word.GetElementXml(run.Path);
        if (!string.IsNullOrEmpty(raw))
        {
            if (raw.Contains("footnoteReference", StringComparison.Ordinal)) return NoteRefKind.Footnote;
            if (raw.Contains("endnoteReference", StringComparison.Ordinal)) return NoteRefKind.Endnote;
            return NoteRefKind.None;
        }
        if (string.Equals(rStyle, "FootnoteReference", StringComparison.OrdinalIgnoreCase)) return NoteRefKind.Footnote;
        if (string.Equals(rStyle, "EndnoteReference", StringComparison.OrdinalIgnoreCase)) return NoteRefKind.Endnote;
        return NoteRefKind.None;
    }

    // BUG-DUMP-R35-2: deterministic flatten warning for a run synthesized from
    // inside a <w:smartTag>/<w:customXml> wrapper (Navigation marks it with
    // Format["_wrapperFlattened"]). Shared by the per-run emit path
    // (EmitPlainOrHyperlinkRun) and the single-run paragraph collapse path so
    // the wrapper loss is never silent regardless of which shape the
    // paragraph takes. CONSISTENCY(wrapper-flatten-warning).
    private static void WarnWrapperFlattened(DocumentNode run, BodyEmitContext? ctx)
    {
        if (run.Format.TryGetValue("_wrapperFlattened", out var wfObj)
            && wfObj is bool wfB && wfB && ctx != null)
        {
            ctx.Warnings.Add(new DocxUnsupportedWarning(
                Element: "smartTag/customXml",
                Path: run.Path,
                Reason: "inline smartTag/customXml wrapper flattened on dump→batch round-trip; the wrapped run text and formatting are preserved, only the wrapper element is dropped"));
        }
    }

    private static void EmitPlainOrHyperlinkRun(DocumentNode run, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx = null, int hlBaseline = 0)
    {
        // BUG-R12A(BUG1): a hyperlink wrapper with >1 run or any per-run rPr was
        // stashed by CoalesceHyperlinkRuns with its original runs in Children.
        // Emit the wrapper carrying the FIRST run's text + rPr, then one
        // structured `add run` per remaining run targeting the hyperlink path so
        // each run's bold/color/size/font survives round-trip (the flat
        // `add hyperlink text=` path would flatten them into one unformatted run).
        // Mirrors the R9 comment-body structured-add-run fix; raw-set is not an
        // option here because the <w:hyperlink> r:id relationship can't be
        // recreated by verbatim XML injection.
        if (run.Format.TryGetValue("_hlStructured", out var hlsObj) && hlsObj is bool hlsB && hlsB
            && run.Children is { Count: > 0 } hlRuns)
        {
            EmitStructuredHyperlink(hlRuns, paraTargetPath, items, ctx, hlBaseline);
            return;
        }
        var rProps = FilterEmittableProps(run.Format);
        if (!string.IsNullOrEmpty(run.Text))
            rProps["text"] = run.Text!;
        // BUG-DUMP-R35-2: a run synthesized from inside a <w:smartTag>/
        // <w:customXml> wrapper. We FLATTEN the wrapper (drop the smartTag/
        // customXml element) but PRESERVE the inner run's text + formatting —
        // consistent with how Word often strips these and with the project's
        // flatten precedents. Surface a deterministic warning so the wrapper
        // loss isn't silent (matches the external-rel / picBullet convention).
        WarnWrapperFlattened(run, ctx);
        // CONSISTENCY(move-range-markers): a moveFrom/moveTo run's own w:id in
        // the source usually differs from its paired half (the pairing lives on
        // the bracketing range markers' shared w:name, not on the run id). Rewrite
        // both halves to one SHARED revision.id so AddRun's WrapRunAsMove* helpers
        // synthesize Move_{id} range markers that pair the moveFrom to its moveTo.
        if (ctx is { MovePairIds.Count: > 0 }
            && rProps.TryGetValue("revision.type", out var mvType)
            && (mvType.Equals("moveFrom", StringComparison.OrdinalIgnoreCase)
                || mvType.Equals("moveTo", StringComparison.OrdinalIgnoreCase))
            && rProps.TryGetValue("revision.id", out var mvId)
            && ctx.MovePairIds.TryGetValue(mvId, out var sharedId))
        {
            rProps["revision.id"] = sharedId;
        }
        // BUG-DUMP-R43-8: rPrChange's PreviousRunProperties snapshot now
        // round-trips verbatim via revision.beforeXml (see EmitParagraph
        // counterpart). The former warn-and-drop path is retired — the
        // prior-rPr payload is preserved. Defensive strip only, for stale dumps.
        rProps.Remove("revision.beforeLost");

        // Hyperlink-wrapped run: Get flattens a <w:hyperlink>'s child run
        // into a regular run-typed node but copies the resolved URL onto
        // Format["url"]. AddRun does not consume `url` — emitting type="r"
        // would silently drop the hyperlink wrapper. Re-emit as a typed
        // `add hyperlink` so the <w:hyperlink>+rel-relationship round-trip
        // rebuilds correctly.
        // CONSISTENCY(docx-hyperlink-canonical-url): canonical key is `url`
        // on both Get readback and Add input.
        if (rProps.ContainsKey("url") || rProps.ContainsKey("anchor")
            || rProps.ContainsKey("isHyperlink"))
        {
            // AddHyperlink writes its own color/underline defaults from theme;
            // drop the inferred `color: hyperlink` / `underline: single` Get
            // echoes back so we don't override those defaults. Track the drops:
            // a dropped key means the SOURCE had an explicit element that the
            // defaults reproduce — the "inherit" sentinel below must NOT fire
            // for it, or the explicit single underline / theme color vanishes.
            bool hlColorDropped = false, hlUnderlineDropped = false;
            if (rProps.TryGetValue("color", out var hlColor)
                && string.Equals(hlColor, "hyperlink", StringComparison.OrdinalIgnoreCase))
            {
                rProps.Remove("color");
                hlColorDropped = true;
            }
            if (rProps.TryGetValue("underline", out var hlUl)
                && string.Equals(hlUl, "single", StringComparison.OrdinalIgnoreCase))
            {
                rProps.Remove("underline");
                hlUnderlineDropped = true;
            }
            rProps.Remove("isHyperlink");
            // Bare <w:hyperlink> wrapper with no url/anchor/tooltip/tgtFrame
            // /history carries no round-trippable property — AddHyperlink
            // would reject it. Fall through and emit as a plain run.
            if (!rProps.ContainsKey("url") && !rProps.ContainsKey("anchor")
                && !rProps.ContainsKey("tooltip") && !rProps.ContainsKey("tgtFrame")
                && !rProps.ContainsKey("tgtframe") && !rProps.ContainsKey("history"))
            {
                items.Add(new BatchItem
                {
                    Command = "add",
                    Parent = paraTargetPath,
                    Type = "r",
                    Props = rProps.Count > 0 ? rProps : null
                });
                return;
            }
            // The SOURCE run has no <w:color>/<w:u> at all (a TOC leader row
            // that deliberately looks like plain text). Without the sentinel,
            // AddHyperlink stamps its theme-blue + single-underline defaults
            // and the dotted leader comes back blue. Injected only on the real
            // `add hyperlink` path — the bare-wrapper fallback above emits a
            // plain `add r`, whose color parser must not see "inherit".
            if (!hlColorDropped && !rProps.ContainsKey("color")) rProps["color"] = "inherit";
            if (!hlUnderlineDropped && !rProps.ContainsKey("underline")) rProps["underline"] = "inherit";
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "hyperlink",
                Props = rProps,
            });
            return;
        }
        items.Add(new BatchItem
        {
            Command = "add",
            Parent = paraTargetPath,
            Type = "r",
            Props = rProps.Count > 0 ? rProps : null
        });
    }

    // BUG-R12A(BUG1): emit a hyperlink wrapper + one structured `add run` per
    // wrapped run so per-run rPr survives. The first run materializes the
    // wrapper (AddHyperlink seeds the wrapper's first run from `text` + rPr
    // props); subsequent runs are appended via `add r` targeting the
    // hyperlink[K] path (verified working — AddRun accepts a hyperlink parent
    // and preserves bold/color/size/font). hlIndex is the 1-based ordinal of
    // this hyperlink among the rows already emitted under the host paragraph.
    // BUG-R14B: hlBaseline is the number of hyperlink rows already present
    // at paraTargetPath *before this paragraph's own run-emit pass began*.
    // paraTargetPath is the literal "/body/p[last()]" for every paragraph, so
    // a raw items.Count() of hyperlink rows at that parent also tallies the
    // hyperlinks from previously-emitted paragraphs — at replay time those
    // live under earlier <w:p> elements, not the current last() paragraph, so
    // the current paragraph's hyperlinks re-index from 1. Subtracting the
    // baseline yields the wrapper's LIVE 1-based index inside this paragraph,
    // which is what the trailing `add r` rows must target.
    private static void EmitStructuredHyperlink(List<DocumentNode> hlRuns, string paraTargetPath,
                                                List<BatchItem> items, BodyEmitContext? ctx, int hlBaseline = 0)
    {
        // Build the wrapper add from the first run's props (url/anchor/tooltip/…
        // + the run's own rPr). Reuse the existing flat-emit logic by routing
        // the first run through EmitPlainOrHyperlinkRun with the structured flag
        // cleared, so the theme-default scrubbing + bare-wrapper fallback stay
        // identical.
        var first = hlRuns[0];
        var firstClone = new DocumentNode
        {
            Path = first.Path,
            Type = "run",
            // A tab-first hyperlink (TOC leader): the wrapper's text "\t" is
            // split into a real <w:tab/> by AddText, so the leader stays the
            // hyperlink's first child in source order.
            Text = first.Type == "tab" ? "\t" : first.Text,
            Format = new Dictionary<string, object?>(first.Format, StringComparer.OrdinalIgnoreCase),
        };
        firstClone.Format.Remove("_hlStructured");
        int hlBefore = items.Count(it => it.Type == "hyperlink"
            && string.Equals(it.Parent, paraTargetPath, StringComparison.Ordinal));
        EmitPlainOrHyperlinkRun(firstClone, paraTargetPath, items, ctx);
        int hlAfter = items.Count(it => it.Type == "hyperlink"
            && string.Equals(it.Parent, paraTargetPath, StringComparison.Ordinal));
        // If the first run did not materialize a hyperlink row (bare wrapper with
        // no url/anchor/tooltip → AddHyperlink can't represent it, so it fell back
        // to a plain run), the wrapper is gone — emit the remaining runs as plain
        // runs too so their text/rPr still survive.
        if (hlAfter == hlBefore)
        {
            for (int k = 1; k < hlRuns.Count; k++)
            {
                var clone = new DocumentNode
                {
                    Path = hlRuns[k].Path,
                    Type = hlRuns[k].Type,
                    Text = hlRuns[k].Text,
                    Format = new Dictionary<string, object?>(hlRuns[k].Format, StringComparer.OrdinalIgnoreCase),
                };
                clone.Format.Remove("_hlStructured");
                clone.Format.Remove("url");
                clone.Format.Remove("anchor");
                EmitPlainOrHyperlinkRun(clone, paraTargetPath, items, ctx);
            }
            return;
        }
        // BUG-R14B: live in-paragraph index = total hyperlink rows at this
        // parent minus the count that belonged to earlier paragraphs.
        var hlPath = $"{paraTargetPath}/hyperlink[{hlAfter - hlBaseline}]";
        for (int k = 1; k < hlRuns.Count; k++)
        {
            var rProps = FilterEmittableProps(hlRuns[k].Format);
            // The hyperlink-wrapper keys belong to the <w:hyperlink>, not its
            // child runs — strip them so `add r` doesn't choke / re-wrap.
            // BUG-DUMP-R43-4: rStyle/rstyle is a PER-RUN character style (e.g.
            // Internetlink), NOT a hyperlink-wrapper attribute. It lives in
            // HyperlinkWrapperOnlyKeys only so a sole rStyle-bearing run stays on
            // the lossless fast path; here, on the structured (multi-run) path,
            // each trailing run must re-emit its OWN rStyle (the first run already
            // got it via the wrapper add). Keep it — strip only the genuine
            // wrapper attrs.
            foreach (var wk in HyperlinkWrapperOnlyKeys)
            {
                if (string.Equals(wk, "rStyle", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(wk, "rstyle", StringComparison.OrdinalIgnoreCase))
                    continue;
                rProps.Remove(wk);
            }
            // Drop the theme-default echoes Get stamps on every hyperlink run;
            // the wrapper's own run already carries the real Hyperlink style.
            if (rProps.TryGetValue("color", out var c)
                && string.Equals(c, "hyperlink", StringComparison.OrdinalIgnoreCase))
                rProps.Remove("color");
            if (rProps.TryGetValue("underline", out var u)
                && string.Equals(u, "single", StringComparison.OrdinalIgnoreCase))
                rProps.Remove("underline");
            if (hlRuns[k].Type == "tab")
                rProps["text"] = "\t";
            else if (!string.IsNullOrEmpty(hlRuns[k].Text))
                rProps["text"] = hlRuns[k].Text!;
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = hlPath,
                Type = "r",
                Props = rProps.Count > 0 ? rProps : null
            });
        }
    }

    // BUG-R12A(BUG1): raw-set a rich inline (run-level) <w:sdt> into the host
    // paragraph so its per-run formatting survives. Returns true when it emitted
    // a raw-set (or a warning + fall-through is desired); false when the SDT is
    // simple enough for the flat `add sdt text=` fast path. The host paragraph
    // was just added by EmitParagraph, so it is the last <w:p> in the body at
    // replay time — the same last()-relative attach the inline-textbox raw-set
    // uses.
    private static bool TryEmitRichInlineSdt(WordHandler word, DocumentNode sdt,
                                             string parentPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        var rawXml = word.RawElementXml(sdt.Path);
        if (string.IsNullOrEmpty(rawXml) || !IsRichInlineSdt(rawXml!)) return false;
        // BUG-DUMP-R26-7: target /body, header/footer OR table-cell hosts, not
        // only /body. When the host isn't raw-set-addressable, fall back to the
        // typed emit (no regression vs the old /body-only guard).
        if (ResolveRawSetHost(parentPath, ctx) is not { } sdtHost) return false;
        // External relationship references (hyperlink r:id, image r:embed/r:link)
        // would dangle in the rebuilt part — raw injection does not recreate the
        // matching rels. Emit a deterministic warning naming the loss instead of
        // silently flattening (BUG-DUMP-R26-7 PART B), then fall back to the flat
        // text emit. Full rel preservation is a separate, larger effort.
        if (HasExternalRelRef(rawXml!))
        {
            ctx?.Warnings.Add(new DocxUnsupportedWarning(
                Element: "sdt.richContent",
                Path: sdt.Path,
                Reason: "inline content control with formatted/nested content AND an external relationship (hyperlink/image) cannot round-trip verbatim; the relationship target is not carried through dump→batch, so the control is flattened to text on replay"));
            return false;
        }
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = sdtHost.Part,
            Xpath = sdtHost.XPath,
            Action = "append",
            Xml = rawXml
        });
        return true;
    }

    // A run-level <w:sdt> is "rich" when its sdtContent carries formatting the
    // text-only typed emit can't reproduce: more than one run, any run-level
    // <w:rPr>, a hyperlink/field, or an intra-run break/tab. Single plain run →
    // flat `add sdt text=` stays lossless. (Distinct from IsRichBlockSdt, which
    // keys on inner <w:p> count — a run-level SDT has no inner paragraph.)
    private static bool IsRichInlineSdt(string sdtXml)
    {
        // BUG-DUMP-R27-4: a run-level SDT carrying a repeatingSection /
        // repeatingSectionItem / docPartObj sdtPr marker must raw-set verbatim
        // (the typed path can't express the special type). See
        // HasSpecialSdtTypeMarker in WordBatchEmitter.Resources.cs.
        if (HasSpecialSdtTypeMarker(sdtXml))
            return true;
        // BUG-DUMP-R26-5: nested inline SDT. The outer <w:sdt> wraps one or more
        // child <w:sdt> in its sdtContent (L1>L2>L3 tag/id/alias nesting). The
        // flat `add sdt text=` path seeds a single run from the innermost text
        // and drops every nesting level's tag/id/alias plus the structure. A
        // second <w:sdt> anywhere in the XML means at least one nested control,
        // so raw-set the whole tree verbatim to preserve depth + per-level props.
        if (System.Text.RegularExpressions.Regex.Matches(sdtXml, "<w:sdt[ >]").Count > 1)
            return true;
        // More than one content run.
        if (System.Text.RegularExpressions.Regex.Matches(sdtXml, "<w:r[ >]").Count > 1)
            return true;
        // Any run-level run properties (bold/color/size/font/…) inside content.
        if (sdtXml.Contains("<w:rPr", StringComparison.Ordinal)
            // exclude sdtPr/endPr-only rPr false positives: those live in
            // <w:sdtPr>/<w:sdtEndPr><w:rPr>; a content rPr sits under <w:r>.
            && System.Text.RegularExpressions.Regex.IsMatch(sdtXml, "<w:r[ >].*?<w:rPr"))
            return true;
        return sdtXml.Contains("<w:hyperlink", StringComparison.Ordinal)
            || sdtXml.Contains("<w:fldChar", StringComparison.Ordinal)
            || sdtXml.Contains("w:instrText", StringComparison.Ordinal)
            || sdtXml.Contains("<w:fldSimple", StringComparison.Ordinal)
            || sdtXml.Contains("<w:drawing", StringComparison.Ordinal)
            || sdtXml.Contains("<w:br", StringComparison.Ordinal)
            || sdtXml.Contains("<w:tab", StringComparison.Ordinal)
            || sdtXml.Contains("<w:cr", StringComparison.Ordinal);
    }

    // Collapse OOXML complex field chains (fldChar(begin) + instrText + …
    // + fldChar(end)) into a single synthetic "field" DocumentNode with
    // Format["instruction"] (raw code) and Text (cached display value).
    // Non-field children pass through untouched in original order. The TOC
    // chain is handled by the dedicated EmitParagraph branch above and never
    // reaches this collapsing step (early-return in that branch).
    // BUG-DUMP6-05: collapse consecutive runs sharing the same url/anchor
    // into a single synthetic node so dump emits ONE `add hyperlink` per
    // <w:hyperlink>, regardless of how many runs the source wrapped. The
    // synthesized node carries the merged Text (for AddHyperlink's `text`
    // prop) and the shared url/anchor/Hyperlink-style format keys.
    // Mirrors the field-emit hyperlink-parent rebase logic for tab/ptab runs.
    // Navigation marks tab-only runs that live inside w:hyperlink with a
    // Format["_hyperlinkParent"] hint (e.g. /body/p[1]/hyperlink[2]); without
    // re-routing on emit they would replay under the bare paragraph and lose
    // the hyperlink wrapper. The candidate-verify step (a prior `add hyperlink`
    // row must have landed under paraTargetPath) avoids dangling paths when
    // the hyperlink has no emittable runs and so was never added.
    private static string ResolveHyperlinkParent(DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        string? candidateHlParent = null;
        if (run.Format.TryGetValue("_hyperlinkParent", out var hlpObj) && hlpObj != null)
        {
            var hint = hlpObj.ToString();
            if (!string.IsNullOrEmpty(hint)) candidateHlParent = hint;
        }
        if (candidateHlParent == null) return paraTargetPath;

        const string hlMarker = "/hyperlink[";
        var hlIdxStart = candidateHlParent.LastIndexOf(hlMarker, StringComparison.Ordinal);
        if (hlIdxStart <= 0) return paraTargetPath;
        var hlEnd = candidateHlParent.IndexOf(']', hlIdxStart);
        if (hlEnd <= hlIdxStart) return paraTargetPath;
        var kStr = candidateHlParent.Substring(hlIdxStart + hlMarker.Length,
            hlEnd - hlIdxStart - hlMarker.Length);
        if (!int.TryParse(kStr, out var kIdx)) return paraTargetPath;
        var rebased = paraTargetPath + candidateHlParent.Substring(hlIdxStart);
        int emittedHls = items.Count(it => it.Type == "hyperlink"
            && string.Equals(it.Parent, paraTargetPath, StringComparison.Ordinal));
        return emittedHls >= kIdx ? rebased : paraTargetPath;
    }

    // BUG-R12A(BUG1): hyperlink-wrapper keys + the theme defaults AddHyperlink
    // re-applies on its own. A run carrying ONLY these keys has no per-run
    // formatting worth preserving structurally, so the flat `add hyperlink
    // text=` fast path stays lossless. Any OTHER key (bold, color≠hyperlink,
    // size, font, …) means the run's rPr must survive — route to structured emit.
    private static readonly HashSet<string> HyperlinkWrapperOnlyKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "url", "anchor", "tooltip", "tgtFrame", "tgtframe", "history",
        "docLocation", "doclocation", "isHyperlink", "_hyperlinkParent",
        "rStyle", "rstyle",
    };

    // True when a hyperlink-wrapped run carries run-level formatting (rPr) that
    // AddHyperlink's flat `text=` path cannot reproduce. The theme-default
    // color=hyperlink / underline=single echoes that Get stamps back are NOT
    // real formatting (AddHyperlink re-applies them), so they don't count.
    private static bool HyperlinkRunHasRunFormatting(DocumentNode run)
    {
        var props = FilterEmittableProps(run.Format);
        foreach (var (k, v) in props)
        {
            if (HyperlinkWrapperOnlyKeys.Contains(k)) continue;
            if (string.Equals(k, "text", StringComparison.OrdinalIgnoreCase)) continue;
            if (string.Equals(k, "color", StringComparison.OrdinalIgnoreCase)
                && string.Equals(v, "hyperlink", StringComparison.OrdinalIgnoreCase)) continue;
            if (string.Equals(k, "underline", StringComparison.OrdinalIgnoreCase)
                && string.Equals(v, "single", StringComparison.OrdinalIgnoreCase)) continue;
            return true;
        }
        return false;
    }

    private static List<DocumentNode> CoalesceHyperlinkRuns(List<DocumentNode> runs)
    {
        var result = new List<DocumentNode>(runs.Count);
        int i = 0;
        while (i < runs.Count)
        {
            var run = runs[i];
            string? url = null, anchor = null, hlParent = null;
            // Tab runs INSIDE a hyperlink (a TOC row whose link wraps the
            // leader tab + page number) carry the same anchor/_hyperlinkParent
            // markers — include them so the wrapper add is emitted before (and
            // contains) the tab. Excluding them split the link: the tab run
            // raced ahead of the wrapper (its parent guard miscounted across
            // paragraphs sharing the literal /body/p[last()] parent) and the
            // replay dropped it.
            if (run.Type == "run" || run.Type == "r" || run.Type == "tab")
            {
                if (run.Format.TryGetValue("url", out var u))
                    url = u?.ToString();
                if (run.Format.TryGetValue("anchor", out var a))
                    anchor = a?.ToString();
                // `_hyperlinkParent` (e.g. `/body/p[1]/hyperlink[2]`) is the
                // unique-per-wrapper marker NodeBuilder stamps on every run
                // hosted inside a `<w:hyperlink>`. Two sibling hyperlinks
                // sharing the same target URL (e.g. two "Read more →" links
                // pointing at /pricing) carry identical `url`/`anchor` but
                // distinct `_hyperlinkParent` paths, so the URL-only equality
                // pre-R12 collapsed them into one item whose text was
                // "Read more →Read more →" — silent loss of the second link.
                if (run.Format.TryGetValue("_hyperlinkParent", out var hp))
                    hlParent = hp?.ToString();
            }
            if (string.IsNullOrEmpty(url) && string.IsNullOrEmpty(anchor))
            {
                result.Add(run);
                i++;
                continue;
            }
            // Walk forward over consecutive runs with the same url/anchor AND
            // the same hyperlink-wrapper path. Differing `_hyperlinkParent`
            // values mark a hyperlink boundary even when URL/anchor match.
            int j = i + 1;
            var group = new List<DocumentNode> { run };
            var sb = new System.Text.StringBuilder(run.Text ?? "");
            while (j < runs.Count)
            {
                var next = runs[j];
                if (next.Type != "run" && next.Type != "r" && next.Type != "tab") break;
                next.Format.TryGetValue("url", out var nUrlObj);
                next.Format.TryGetValue("anchor", out var nAncObj);
                next.Format.TryGetValue("_hyperlinkParent", out var nHlpObj);
                var nUrl = nUrlObj?.ToString();
                var nAnchor = nAncObj?.ToString();
                var nHlParent = nHlpObj?.ToString();
                if (!string.Equals(nUrl, url, StringComparison.Ordinal)) break;
                if (!string.Equals(nAnchor, anchor, StringComparison.Ordinal)) break;
                if (!string.Equals(nHlParent, hlParent, StringComparison.Ordinal)) break;
                group.Add(next);
                sb.Append(next.Type == "tab" ? "\t" : (next.Text ?? ""));
                j++;
            }
            // BUG-R12A(BUG1): the flat `add hyperlink text=Part1Part2` fast path
            // drops every per-run rPr on the wrapped runs (a 2nd bold+red run
            // round-trips as plain text). Keep the fast path ONLY for the lossless
            // case: a single run with no run-level formatting. Otherwise stash the
            // original group runs in Children and flag the node so
            // EmitPlainOrHyperlinkRun emits the wrapper + structured `add run`
            // rows (mirrors the R9 comment-body fix), preserving each run's rPr.
            bool needsStructured = group.Count > 1 || group.Any(HyperlinkRunHasRunFormatting);
            var merged = new DocumentNode
            {
                Path = run.Path,
                // Normalize to "run": a group whose FIRST member is a tab must
                // not surface as a tab-typed node, or TryEmitTabRun would
                // intercept it ahead of the hyperlink emit.
                Type = "run",
                Text = sb.ToString(),
                Format = new Dictionary<string, object?>(run.Format, StringComparer.OrdinalIgnoreCase),
            };
            if (needsStructured)
            {
                merged.Format["_hlStructured"] = true;
                merged.Children = group;
            }
            result.Add(merged);
            i = j;
        }
        return result;
    }

    // BUG-DUMP-HOIST: run-level character properties that WordHandler.Navigation
    // surfaces on the paragraph node (via the firstRun fallback) but which must
    // NOT ride on `add p` for multi-run paragraphs — every individual run gets
    // its own `add r` carrying its real props.
    private static readonly HashSet<string> RunCharacterPropsHoistedFromFirstRun = new(StringComparer.OrdinalIgnoreCase)
    {
        "bold", "italic", "size", "color", "underline", "underline.color",
        "strike", "highlight",
        "font.latin", "font.ea", "font.ascii", "font.hAnsi",
        // complex-script siblings populated by ReadComplexScriptRunFormatting
        "bold.cs", "italic.cs", "size.cs", "font.cs",
    };

    private static void StripRunCharacterPropsFromParagraph(Dictionary<string, string> props)
    {
        foreach (var k in RunCharacterPropsHoistedFromFirstRun)
            props.Remove(k);
    }

    // Layer per-stop `add tab` rows under a parent path that already has the
    // host paragraph/style created. tabs is the flat List<Dict> Get exposes.
    private static void EmitTabStops(string parentPath, object? tabsVal, List<BatchItem> items)
    {
        if (tabsVal is not IEnumerable<Dictionary<string, object?>> list) return;
        foreach (var t in list)
        {
            var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (t.TryGetValue("pos", out var p) && p != null) props["pos"] = p.ToString() ?? "";
            if (t.TryGetValue("val", out var v) && v != null) props["val"] = v.ToString() ?? "";
            if (t.TryGetValue("leader", out var l) && l != null) props["leader"] = l.ToString() ?? "";
            if (props.Count == 0 || !props.ContainsKey("pos")) continue;
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = parentPath,
                Type = "tab",
                Props = props
            });
        }
    }
}
