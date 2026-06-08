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

        if (TryEmitInlineSectionBreak(pNode, parentPath, items, ctx)) return;
        if (TryEmitTocParagraph(pNode, parentPath, items)) return;
        if (TryEmitTextboxOnlyParagraph(word, pNode, parentPath, autoPresent, items, ctx)) return;

        var props = FilterEmittableProps(pNode.Format);
        // paraMarkIns.* → AddParagraph's bare trackChange.author/date form.
        // A bare trackChange.author (no `trackChange=<kind>` literal) on
        // `add p` produces both <w:pPr><w:rPr><w:ins/></w:rPr></w:pPr> and
        // wraps each auto-created content run in <w:ins>. We pass NO `text=`
        // here so step 1 only stamps the paragraph mark; subsequent `add r`
        // steps (which already carry their own trackChange=ins) wrap the
        // content. Guarded against clashing with a sibling pPrChange — if
        // both surface, paraMarkIns wins (its absence on round-trip is more
        // visually obvious in Word's revision UI).
        if (props.Remove("paraMarkIns.author", out var pmiAuthor))
        {
            props["revision.author"] = pmiAuthor;
            // Strip revision.type=format / revision.date from a sibling pPrChange
            // so the bare-attribution path on AddParagraph fires instead of the
            // pPrChange path. pPrChange round-trip on paragraphs that are
            // ALSO newly inserted is a corner case we accept losing for now
            // (rare in practice; pPrChange semantics overlap with paraMarkIns
            // on a fresh paragraph anyway).
            props.Remove("revision.type");
        }
        if (props.Remove("paraMarkIns.date", out var pmiDate))
        {
            if (!props.ContainsKey("revision.date"))
                props["revision.date"] = pmiDate;
        }
        props.Remove("paraMarkIns.id");
        // revision.beforeLost — Get-side flag set when pPrChange's
        // PreviousParagraphProperties carries a real pre-change pPr snapshot
        // that Add v1 can't reconstruct. Strip the key (Add wouldn't accept
        // it) and emit a dump warning so callers know the Reject-Change
        // restoration state will not survive replay.
        if (props.Remove("revision.beforeLost", out var _) && ctx != null)
        {
            ctx.Warnings.Add(new DocxUnsupportedWarning(
                Element: "revision.before",
                Path: pNode.Path,
                Reason: "pPrChange previous-properties snapshot dropped on dump→batch round-trip; Word's Reject-Change restoration state will be lost on replay"));
        }
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

        // BUG-DUMP4-06: emit inline SdtRun children before the runs loop.
        foreach (var sdt in inlineSdts)
        {
            var sdtProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var key in new[] { "type", "alias", "tag", "items", "format" })
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
        string? sharedAttachPara = drawingBearingCount >= 2 ? paraTargetPath : null;
        foreach (var run in runs)
        {
            if (TryEmitBookmarkRun(run, paraTargetPath, items, ctx)) continue;
            if (TryEmitBreakRun(run, paraTargetPath, items)) continue;
            if (TryEmitTabRun(run, paraTargetPath, items)) continue;
            if (TryEmitPtabRun(run, paraTargetPath, items)) continue;
            if (TryEmitEquationRun(run, paraTargetPath, items)) continue;
            if (TryEmitFormFieldRun(run, paraTargetPath, items)) continue;
            if (TryEmitFieldRun(run, paraTargetPath, items, ctx)) continue;
            // OLE/embedded-object runs surface as type="ole" (see CreateOleNode
            // in WordHandler.ImageHelpers.cs). TryEmitOleRun base64-inlines the
            // embedded payload + icon and the VML frame metadata into a
            // self-contained `add ole` (picture-run style), so the object
            // round-trips with no external file; it warns only when the payload
            // can't be resolved.
            if (TryEmitOleRun(run, paraTargetPath, items, ctx, word)) continue;
            if (TryEmitPictureRun(word, run, paraTargetPath, parentPath, targetIndex, items, ctx, sharedAttachPara)) continue;
            if (TryEmitNoteRefRun(run, paraTargetPath, items, ctx)) continue;
            EmitPlainOrHyperlinkRun(run, paraTargetPath, items, ctx);
        }
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

    private static bool TryEmitInlineSectionBreak(DocumentNode pNode, string parentPath, List<BatchItem> items, BodyEmitContext? ctx)
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
            var sectProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["type"] = breakKind.ToString() ?? "nextPage"
            };
            foreach (var (k, v) in pNode.Format)
            {
                if (!k.StartsWith("sectionBreak.", StringComparison.OrdinalIgnoreCase)) continue;
                if (v == null) continue;
                var keyTail = k["sectionBreak.".Length..];
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
                    if (c.Type != "run" && c.Type != "r") return false;
                    if (!string.IsNullOrEmpty(c.Text)) return true;
                    // BUG-DUMP5-08: include empty rStyle-bearing footnote /
                    // endnote refs (their visible text comes via the typed
                    // emit branch below, not from c.Text).
                    if (c.Format.TryGetValue("rStyle", out var rsv)
                        && rsv != null
                        && (string.Equals(rsv.ToString(), "FootnoteReference", StringComparison.OrdinalIgnoreCase)
                            || string.Equals(rsv.ToString(), "EndnoteReference", StringComparison.OrdinalIgnoreCase)))
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
                    if (ctx != null && rStyle == "FootnoteReference")
                    {
                        var noteText = ctx.FootnoteCursor.Index < ctx.FootnoteTexts.Count
                            ? ctx.FootnoteTexts[ctx.FootnoteCursor.Index] : "";
                        ctx.FootnoteCursor.Index++;
                        items.Add(new BatchItem
                        {
                            Command = "add",
                            Parent = carrierPath,
                            Type = "footnote",
                            Props = new() { ["text"] = noteText }
                        });
                        continue;
                    }
                    if (ctx != null && rStyle == "EndnoteReference")
                    {
                        var noteText = ctx.EndnoteCursor.Index < ctx.EndnoteTexts.Count
                            ? ctx.EndnoteTexts[ctx.EndnoteCursor.Index] : "";
                        ctx.EndnoteCursor.Index++;
                        items.Add(new BatchItem
                        {
                            Command = "add",
                            Parent = carrierPath,
                            Type = "endnote",
                            Props = new() { ["text"] = noteText }
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
        var r = runs[0];
        // Picture / ptab runs need their own typed `add` rows.
        if (r.Type == "picture" || r.Type == "ptab") return false;
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
        // BUG-FUZZ-2: footnote/endnote reference runs need the typed
        // `add footnote/endnote` branch; AddParagraph doesn't consume rStyle.
        if (r.Format.TryGetValue("rStyle", out var srStyle)
            && (string.Equals(srStyle?.ToString(), "FootnoteReference", StringComparison.OrdinalIgnoreCase)
                || string.Equals(srStyle?.ToString(), "EndnoteReference", StringComparison.OrdinalIgnoreCase)))
            return false;
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
            || r.Format.ContainsKey("sym")) return false;
        // BUG-FIELD-COLLAPSE: a synthetic field run carries `instruction=…` —
        // collapse would lose the field chain on replay.
        if (r.Type == "field") return false;
        // BUG-DUMP7-03: inline equation must emit `add equation` explicitly.
        if (r.Type == "equation") return false;
        return true;
    }

    private static bool TryEmitBookmarkRun(DocumentNode run, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // BUG-DUMP25-01: bookmark child emitted in DOM order so a
        // BookmarkStart between runs survives round-trip at its original
        // intra-paragraph offset. Deferred bookmarks (endPara=true) are
        // pushed onto ctx.DeferredBookmarks so the End sibling can land in
        // a downstream paragraph.
        if (run.Type != "bookmark") return false;
        var bmProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (run.Format.TryGetValue("name", out var bmName) && bmName != null)
        {
            var s = bmName.ToString();
            if (!string.IsNullOrEmpty(s)) bmProps["name"] = s;
        }
        if (bmProps.Count == 0) return true; // skip unnamed/anonymous bookmarks
        bool deferred = false;
        if (run.Format.TryGetValue("endPara", out var bmEnd) && bmEnd != null)
        {
            var s = bmEnd.ToString();
            if (!string.IsNullOrEmpty(s) && s != "0")
            {
                bmProps["endPara"] = s;
                deferred = true;
            }
        }
        var bmItem = new BatchItem
        {
            Command = "add",
            Parent = paraTargetPath,
            Type = "bookmark",
            Props = bmProps
        };
        if (deferred && ctx != null)
            ctx.DeferredBookmarks.Add(bmItem);
        else
            items.Add(bmItem);
        return true;
    }

    private static bool TryEmitBreakRun(DocumentNode run, string paraTargetPath, List<BatchItem> items)
    {
        // BUG-DUMP5-01/02: a soft <w:br/> with NO type attribute is a line
        // break, not a page break — fall back to type=line. Emitted inline
        // from the unified runs loop so each break stays at its source
        // position instead of being hoisted to the front of the paragraph.
        if (run.Type != "break") return false;
        var breakType = run.Format.TryGetValue("breakType", out var bt) ? bt?.ToString() : null;
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

    private static bool TryEmitEquationRun(DocumentNode run, string paraTargetPath, List<BatchItem> items)
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
                if (derived.Contains("/hyperlink["))
                    eqParent = derived;
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

    private static bool TryEmitFieldRun(DocumentNode run, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx = null)
    {
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
            if (ctx != null)
            {
                ctx.Warnings.Add(new DocxUnsupportedWarning(
                    Element: "field.nested",
                    Path: run.Path,
                    Reason: "nested field (begin inside a field's branch) cannot round-trip through add field; cached display preserved but inner field codes dropped"));
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
            if (isTextbox && TryEmitTextbox(word, run, probeXml, parentPath, items, ctx, sharedAttachPara))
                return true;
            // AlternateContent-wrapped non-textbox shapes — connector lines
            // (<wps:cNvCnPr>, e.g. a letterhead separator), autoshapes, groups
            // — and textboxes whose typed emit failed fall back to a raw-set
            // append so the shape survives round-trip instead of vanishing.
            if (parentPath == "/body" && !probeXml.Contains("r:embed") && !probeXml.Contains("r:id")
                && !DrawingHasUnreconstructableRel(probeXml))
            {
                items.Add(new BatchItem
                {
                    Command = "raw-set",
                    Part = "/document",
                    // BUG-R1-01: attach to the most-recently-emitted paragraph
                    // via last(), NOT the navigation pIndex (targetIndex). pIndex
                    // deliberately does not advance for an oMathPara-bearing
                    // paragraph (EmitBody: display equations resolve as
                    // /body/oMathPara[N], not /body/p[N]), but in the LITERAL XML
                    // the equation is a real <w:p> wrapping <m:oMathPara>. So a
                    // numeric w:p[{pIndex}] under-counts by one per preceding
                    // equation and the shape lands on the wrong paragraph. The
                    // host paragraph was just added by this same EmitParagraph
                    // call, so it is always the last w:p at replay time — the same
                    // last()-relative attach the typed picture/textbox path uses.
                    Xpath = "/w:document/w:body/w:p[last()]",
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
            if (TryEmitTextbox(word, run, rawXml, parentPath, items, ctx, sharedAttachPara))
                return true;
        }
        if (!string.IsNullOrEmpty(rawXml) &&
            parentPath == "/body" &&
            !rawXml.Contains("r:embed") && !rawXml.Contains("r:id") &&
            !DrawingHasUnreconstructableRel(rawXml))
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
        xml.Contains("r:qs") || xml.Contains("r:cs");

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
    /// BUG-DUMP-TXBX: emit a typed <c>add textbox</c> row for the host of
    /// the current drawing run, followed by recursive inner-paragraph/run
    /// emits under <c>/&lt;host&gt;/textbox[N]</c>. Geometry props
    /// (width/height/wrap/anchor.x/anchor.y/fill) are extracted from the
    /// raw drawing XML so the rebuilt textbox keeps its layout.
    /// </summary>
    private static bool TryEmitTextbox(WordHandler word, DocumentNode run, string rawXml,
                                       string parentPath, List<BatchItem> items, BodyEmitContext? ctx,
                                       string? attachParaPath = null)
    {
        if (ctx == null) return false;

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
                if (!string.IsNullOrEmpty(posH)) props["anchor.x"] = posH + "emu";
                if (!string.IsNullOrEmpty(posV)) props["anchor.y"] = posV + "emu";
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

    private static bool TryEmitNoteRefRun(DocumentNode run, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx)
    {
        // Footnote/endnote reference runs are empty <w:r> elements with
        // rStyle = FootnoteReference / EndnoteReference. Emit them as a
        // typed footnote/endnote add anchored on the host paragraph and
        // pull the body text from the pre-resolved ordered list — see
        // BodyEmitContext for the document-order assumption.
        if (ctx == null) return false;
        var rStyle = run.Format.TryGetValue("rStyle", out var rs) ? rs?.ToString() : null;
        if (rStyle == "FootnoteReference")
        {
            var noteText = ctx.FootnoteCursor.Index < ctx.FootnoteTexts.Count
                ? ctx.FootnoteTexts[ctx.FootnoteCursor.Index]
                : "";
            ctx.FootnoteCursor.Index++;
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "footnote",
                Props = new() { ["text"] = noteText }
            });
            return true;
        }
        if (rStyle == "EndnoteReference")
        {
            var noteText = ctx.EndnoteCursor.Index < ctx.EndnoteTexts.Count
                ? ctx.EndnoteTexts[ctx.EndnoteCursor.Index]
                : "";
            ctx.EndnoteCursor.Index++;
            items.Add(new BatchItem
            {
                Command = "add",
                Parent = paraTargetPath,
                Type = "endnote",
                Props = new() { ["text"] = noteText }
            });
            return true;
        }
        return false;
    }

    private static void EmitPlainOrHyperlinkRun(DocumentNode run, string paraTargetPath, List<BatchItem> items, BodyEmitContext? ctx = null)
    {
        var rProps = FilterEmittableProps(run.Format);
        if (!string.IsNullOrEmpty(run.Text))
            rProps["text"] = run.Text!;
        // revision.beforeLost — see EmitParagraph counterpart. Strip from
        // the emitted props (AddRun would flag it UNSUPPORTED) and surface
        // a dump warning so the lost rPrChange snapshot is visible to the
        // caller.
        if (rProps.Remove("revision.beforeLost") && ctx != null)
        {
            ctx.Warnings.Add(new DocxUnsupportedWarning(
                Element: "revision.before",
                Path: run.Path,
                Reason: "rPrChange previous-properties snapshot dropped on dump→batch round-trip; Word's Reject-Change restoration state will be lost on replay"));
        }

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
            // echoes back so we don't override those defaults.
            if (rProps.TryGetValue("color", out var hlColor)
                && string.Equals(hlColor, "hyperlink", StringComparison.OrdinalIgnoreCase))
                rProps.Remove("color");
            if (rProps.TryGetValue("underline", out var hlUl)
                && string.Equals(hlUl, "single", StringComparison.OrdinalIgnoreCase))
                rProps.Remove("underline");
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

    private static List<DocumentNode> CoalesceHyperlinkRuns(List<DocumentNode> runs)
    {
        var result = new List<DocumentNode>(runs.Count);
        int i = 0;
        while (i < runs.Count)
        {
            var run = runs[i];
            string? url = null, anchor = null, hlParent = null;
            if (run.Type == "run" || run.Type == "r")
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
            var sb = new System.Text.StringBuilder(run.Text ?? "");
            while (j < runs.Count)
            {
                var next = runs[j];
                if (next.Type != "run" && next.Type != "r") break;
                next.Format.TryGetValue("url", out var nUrlObj);
                next.Format.TryGetValue("anchor", out var nAncObj);
                next.Format.TryGetValue("_hyperlinkParent", out var nHlpObj);
                var nUrl = nUrlObj?.ToString();
                var nAnchor = nAncObj?.ToString();
                var nHlParent = nHlpObj?.ToString();
                if (!string.Equals(nUrl, url, StringComparison.Ordinal)) break;
                if (!string.Equals(nAnchor, anchor, StringComparison.Ordinal)) break;
                if (!string.Equals(nHlParent, hlParent, StringComparison.Ordinal)) break;
                sb.Append(next.Text ?? "");
                j++;
            }
            if (j == i + 1)
            {
                // No coalescing — emit the single run as-is.
                result.Add(run);
            }
            else
            {
                var merged = new DocumentNode
                {
                    Path = run.Path,
                    Type = run.Type,
                    Text = sb.ToString(),
                    Format = new Dictionary<string, object?>(run.Format, StringComparer.OrdinalIgnoreCase),
                };
                result.Add(merged);
            }
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
