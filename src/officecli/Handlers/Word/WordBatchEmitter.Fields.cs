// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class WordBatchEmitter
{

    // BUG-R12A(BUG1): run-level formatting keys a cached field-result run may
    // carry that are worth preserving on round-trip. AddField can apply
    // font/size/bold/color uniformly to the rebuilt field runs; italic /
    // underline are captured too so the raw-set fallback (when AddField can't
    // express them) can be chosen. Keep narrow — paragraph-level/derived keys
    // (effective.*, alignment, …) are NOT result-run formatting.
    private static readonly HashSet<string> FieldResultFormatKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "bold", "italic", "color", "size", "font", "font.latin", "font.ascii", "font.hAnsi",
        "underline", "strike",
        // BUG-DUMP-FIELDVALIGN: field-wide vertical alignment (superscript /
        // subscript) — the common case is a cross-reference citation mark
        // ([1],[2]…) whose every run (begin/instr/sep/result/end) shares the
        // same <w:vertAlign w:val="superscript"/>. The result run reflects it,
        // so capturing it here (and applying it uniformly via AddField below)
        // restores the superscript on round-trip.
        "superscript", "subscript",
    };

    // AddField's --prop vocabulary for field-run formatting. A captured result
    // rPr made up exclusively of these keys round-trips losslessly through the
    // typed `add field` path; anything else (italic/underline/strike/…) means
    // the field needs a raw-set passthrough to keep full fidelity.
    private static readonly HashSet<string> FieldAddSupportedFormatKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        "bold", "color", "size", "font", "font.latin", "font.ascii", "font.hAnsi",
        // BUG-DUMP-FIELDVALIGN: AddField applies vertAlign (superscript /
        // subscript) uniformly to every rebuilt field run, so these are
        // losslessly expressible through the typed `add field` path.
        "superscript", "subscript",
    };

    private static bool FieldRunHasFormatting(DocumentNode run)
    {
        foreach (var (k, v) in run.Format)
        {
            if (v == null) continue;
            if (FieldResultFormatKeys.Contains(k)) return true;
        }
        return false;
    }

    // BUG-DUMP-R26-2: the cached field result is "rich" when its post-separate
    // text runs don't all share one formatting signature. AddField's single-rPr
    // model can only apply ONE rPr to all rebuilt result runs, so a result like
    // "Bold "(b) + "Red "(color) + "Italic"(i) collapses to one run with the
    // first run's bold leaked onto every run (and onto the fldChar markers).
    // When this returns true the emitter routes the whole field chain through a
    // verbatim raw-set instead. Single-run results (the common case) and an
    // empty result are NOT rich — they round-trip fine through `add field`.
    private static bool ResultRunsAreRich(List<DocumentNode> resultRuns)
    {
        // Only text-bearing runs count; markers / empty rPr-only runs don't
        // contribute a distinct visible segment.
        var textRuns = resultRuns.Where(r => !string.IsNullOrEmpty(r.Text)).ToList();
        if (textRuns.Count <= 1) return false;
        string Sig(DocumentNode r)
        {
            var parts = new List<string>();
            foreach (var k in FieldResultFormatKeys)
            {
                if (r.Format.TryGetValue(k, out var v) && v != null)
                {
                    var s = v switch { bool b => b ? "1" : "0", _ => v.ToString() ?? "" };
                    if (s.Length > 0 && s != "0" && s != "false") parts.Add(k + "=" + s);
                }
            }
            parts.Sort(StringComparer.Ordinal);
            return string.Join(";", parts);
        }
        var first = Sig(textRuns[0]);
        return textRuns.Any(r => Sig(r) != first);
    }

    // Track-change attribution keys a revision wrapper (<w:del>/<w:ins>/
    // <w:moveFrom>/<w:moveTo>) stamps on each run it wraps. A field collapsed
    // out of such runs must carry them onto its synth so the emitter re-wraps
    // the rebuilt field in the same revision marker.
    private static readonly string[] RevisionAttributionKeys =
    {
        "revision.type", "revision.author", "revision.date", "revision.id",
    };

    private static void PropagateRevisionKeys(DocumentNode from, DocumentNode to)
    {
        foreach (var rk in RevisionAttributionKeys)
        {
            if (to.Format.ContainsKey(rk)) continue;
            if (from.Format.TryGetValue(rk, out var rv) && rv != null)
                to.Format[rk] = rv;
        }
    }

    private static List<DocumentNode> CollapseFieldChains(List<DocumentNode> children)
    {
        var result = new List<DocumentNode>();
        for (int i = 0; i < children.Count; i++)
        {
            var c = children[i];
            bool isBegin = c.Type == "fieldChar"
                && c.Format.TryGetValue("fieldCharType", out var fct)
                && string.Equals(fct?.ToString(), "begin", StringComparison.OrdinalIgnoreCase);
            if (!isBegin)
            {
                result.Add(c);
                continue;
            }

            // Walk forward to find instruction text and end marker.
            // R10-bug7: track nesting depth so an inner field (e.g. DATE
            // wrapped inside an outer IF's true/false branch) does NOT have
            // its instrText flattened into the outer instruction string —
            // that flattening silently merged the inner field's code into
            // the outer IF's expression, destroyed the false-branch
            // boundary, and produced an instruction the IF parser could
            // not round-trip.
            string instruction = "";
            string display = "";
            bool sawSeparate = false;
            bool sawNestedField = false;
            int end = -1;
            int depth = 1;
            // BUG-R12A(BUG1): capture run-level formatting on the cached result
            // run(s). The flat `add field` path dropped it, so a bold/red/20pt
            // PAGE field round-tripped as plain "1". AddField applies uniform
            // font/size/bold/color to all field runs, so carrying the result
            // run's rPr forward restores the styled field on replay. (Distinct
            // result-run formatting per segment is rare; we capture the first
            // formatted display run, matching AddField's single-rPr model.)
            DocumentNode? firstFormattedResult = null;
            // BUG-DUMP-R26-2: collect EVERY post-separate result run so we can
            // detect a rich (multi-run, heterogeneously-formatted) cached result.
            // AddField's single-rPr model collapses such a result to one run and
            // applies the FIRST run's bold to all of them (and leaks it onto the
            // begin/instr/separate/end fldChar runs). When the result carries >1
            // distinctly-formatted run, we round-trip the whole field chain
            // verbatim via raw-set instead (see TryEmitFieldRun).
            var resultRuns = new List<DocumentNode>();
            // BUG-DUMP-FIELDVALIGN: field-wide vertical alignment (superscript /
            // subscript) is uniform across EVERY run of the field — a citation
            // mark whose begin/instr/separate/result/end runs all carry the same
            // <w:vertAlign>. The post-separate `firstFormattedResult` capture
            // misses two real shapes: (1) an empty result run that sits BEFORE
            // the separator (Word emits a stray rPr-only run between instr and
            // separate — its rPr still defines the field's vertAlign), and (2) a
            // field whose post-separate result is empty (no text run to read).
            // The begin/instr/sep/end marker NODES have their vertAlign stripped
            // (TypographyOnlyKeys noise-suppression in RunToNode), so the only
            // non-stripped carrier is a `run`/`r`-typed node anywhere in the
            // chain. Scan ALL depth-1 result runs (regardless of sawSeparate or
            // text content) for a vertAlign and stash it so it rides on the field
            // op even when no formatted post-separate text run exists.
            string? fieldVertAlign = null;
            for (int j = i + 1; j < children.Count; j++)
            {
                var k = children[j];
                if (k.Type == "instrText")
                {
                    // Only the OUTERMOST instrText belongs in this field's
                    // instruction. Inner instrText (depth > 1) is part of a
                    // nested field whose collapse target is the outer
                    // begin/end pair we're walking inside.
                    if (depth == 1)
                    {
                        if (k.Format.TryGetValue("instruction", out var iv) && iv != null)
                            instruction += iv.ToString();
                        else if (!string.IsNullOrEmpty(k.Text))
                            instruction += k.Text;
                    }
                }
                else if (k.Type == "fieldChar"
                    && k.Format.TryGetValue("fieldCharType", out var ft))
                {
                    var ftStr = ft?.ToString();
                    if (string.Equals(ftStr, "begin", StringComparison.OrdinalIgnoreCase))
                    {
                        // Nested field opens. The outer field can no longer
                        // round-trip through AddField (AddField rebuilds a
                        // flat begin/instr/sep/display/end chain and has no
                        // model for nested branches). Mark and keep
                        // counting until the matching outer end.
                        sawNestedField = true;
                        depth++;
                    }
                    else if (string.Equals(ftStr, "separate", StringComparison.OrdinalIgnoreCase))
                    {
                        if (depth == 1) sawSeparate = true;
                    }
                    else if (string.Equals(ftStr, "end", StringComparison.OrdinalIgnoreCase))
                    {
                        depth--;
                        if (depth == 0)
                        {
                            end = j;
                            break;
                        }
                    }
                }
                else if ((k.Type == "run" || k.Type == "r") && depth == 1)
                {
                    // Cached display segments after fldChar(separate). Concatenate
                    // their text. At depth>1 the run belongs to the nested
                    // field's cached display and is consumed by its own collapse
                    // pass after the outer field is rolled back.
                    if (!string.IsNullOrEmpty(k.Text)) display += k.Text;
                    // BUG-R12A(BUG1): remember the first display run that carries
                    // real run-level formatting so TryEmitFieldRun can forward it
                    // to AddField (which applies it to the rebuilt field runs).
                    if (sawSeparate && firstFormattedResult == null && FieldRunHasFormatting(k))
                        firstFormattedResult = k;
                    // BUG-DUMP-R26-2: remember all post-separate result runs (for
                    // the rich-result heterogeneity check below).
                    if (sawSeparate) resultRuns.Add(k);
                    // BUG-DUMP-FIELDVALIGN: capture vertAlign from ANY result run
                    // in the chain (independent of sawSeparate / text), so an
                    // empty pre-separate rPr-only run or a field with an empty
                    // post-separate result still carries the field-wide
                    // superscript/subscript onto the op.
                    if (fieldVertAlign == null)
                    {
                        if (k.Format.TryGetValue("superscript", out var supv) && supv is bool sb && sb)
                            fieldVertAlign = "superscript";
                        else if (k.Format.TryGetValue("subscript", out var subv) && subv is bool sbb && sbb)
                            fieldVertAlign = "subscript";
                    }
                }
            }
            if (end < 0)
            {
                // R10-bug8: malformed field — fldChar(begin) with no matching
                // end. The previous "fall back to passing through" path
                // returned the bare fldChar(begin) node, which the run-list
                // filter in EmitParagraph then silently dropped (fieldChar
                // is not in the allowlist). Surface a synthetic field
                // entry carrying the partial instruction so TryEmitFieldRun
                // can attach an envelope warning instead. The cached
                // display (any runs accumulated before we ran out of input)
                // is preserved so the paragraph keeps its visible text.
                var malformedSynth = new DocumentNode
                {
                    Path = c.Path,
                    Type = "field",
                    Text = display,
                    Format = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["instruction"] = instruction.Trim(),
                        ["_unmatchedFieldBegin"] = true,
                    }
                };
                result.Add(malformedSynth);
                continue;
            }
            if (sawNestedField)
            {
                // Nested-field branch: the AddField rebuild path cannot
                // represent IF/REF/MERGEFIELD with embedded child fields.
                // Round-trip through a raw-set passthrough so the nested
                // structure survives byte-for-byte. The host paragraph's
                // emit already creates the paragraph; the raw-set append
                // is wired below in TryEmitFieldRun via the
                // `_rawFieldSlice` Format hint. Synthesize a sentinel
                // entry that the field-emit branch routes to raw-set
                // instead of AddField.
                var rawSynth = new DocumentNode
                {
                    Path = c.Path,
                    Type = "field",
                    Text = display,
                    Format = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["instruction"] = instruction.Trim(),
                        ["_nestedField"] = true,
                        ["_fieldChildStart"] = i,
                        ["_fieldChildEnd"] = end,
                    }
                };
                result.Add(rawSynth);
                i = end;
                continue;
            }
            // R14-bug1+2: legacy form field — the begin run carries
            // <w:ffData>, surfaced by Navigation as ffName/ffType/ffDefault/
            // ffMaxLength/ffChecked/… on the fieldChar node. The plain `field`
            // synth would drive BuildFieldAddProps through its default arm,
            // emit `instr=FORMTEXT`, and AddField (via the formtext delegate
            // arm) would rebuild a /formfield with NONE of the original
            // ffData props (name, default, maxLength, items, helpText, …).
            // Route to a `formfield` synth instead so TryEmitFieldRun emits
            // `add formfield` carrying the full payload.
            if (c.Format.TryGetValue("hasFormFieldData", out var hffd)
                && hffd is bool hffdB && hffdB)
            {
                var ffSynth = new DocumentNode
                {
                    Path = c.Path,
                    Type = "formfield",
                    Text = display,
                    Format = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["instruction"] = instruction.Trim()
                    }
                };
                // Carry every ff* prop forward so TryEmitFormFieldRun can
                // map them onto AddFormField's prop bag.
                foreach (var (fk, fv) in c.Format)
                {
                    if (fv == null) continue;
                    if (fk.StartsWith("ff", StringComparison.OrdinalIgnoreCase))
                        ffSynth.Format[fk] = fv;
                }
                result.Add(ffSynth);
                i = end;
                continue;
            }
            var synth = new DocumentNode
            {
                Path = c.Path,
                Type = "field",
                Text = display,
                Format = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                {
                    ["instruction"] = instruction.Trim()
                }
            };
            // BUG-DUMP-R26-2: a rich (multi-run, heterogeneously-formatted)
            // cached result cannot round-trip through `add field` — its single
            // rPr model collapses the runs and leaks the first run's bold onto
            // every result run AND the begin/instr/separate/end fldChar markers.
            // Flag it and stash the field-slice run paths so TryEmitFieldRun
            // raw-sets the whole begin..end chain verbatim, preserving per-run
            // formatting. Empty / single-run results stay on the typed path.
            if (sawSeparate && ResultRunsAreRich(resultRuns))
            {
                var slicePaths = new List<string>();
                for (int s = i; s <= end; s++)
                {
                    var sp = children[s].Path;
                    if (!string.IsNullOrEmpty(sp)) slicePaths.Add(sp);
                }
                if (slicePaths.Count > 0)
                {
                    synth.Format["_richFieldResult"] = true;
                    synth.Format["_fieldSlicePaths"] = string.Join("\n", slicePaths);
                }
            }
            // BUG-DUMP-R26-7 (PART B): a field cached result that wraps a
            // HYPERLINK (result run carries a `url`, i.e. an external r:id rel)
            // does NOT round-trip the hyperlink through the typed `add field`
            // path — the link wrapper is dropped (text + bold survive, the rel
            // is lost) and bold leaks onto the fldChar markers. This is a silent
            // loss even when the result is a single run (so ResultRunsAreRich is
            // false). Flag it so TryEmitFieldRun emits a deterministic warning.
            // Full hyperlink-in-field-result preservation is a separate effort.
            if (sawSeparate && resultRuns.Any(r =>
                    r.Format.TryGetValue("url", out var u) && u != null
                    && !string.IsNullOrEmpty(u.ToString())))
            {
                synth.Format["_fieldResultHasExternalRel"] = true;
            }
            // Source field has no <w:fldChar w:fldCharType="separate"/> — it's
            // the begin+instr+end shape (Word recomputes the result on open).
            // Flag this so EmitField on the field branch can pass `text=""`
            // explicitly to AddField, which short-circuits AddField's default
            // placeholder ("1" for PAGE etc.) and emits the same separator-
            // less shape. Without this flag, the second dump surfaces a
            // phantom `text="1"` key that the source never had.
            if (!sawSeparate)
                synth.Format["_noFieldSeparator"] = true;
            // BUG-DUMP-R24-2: source field HAS a separator (sawSeparate) but the
            // cached result is empty (no result run between separate and end).
            // Without an explicit signal, AddField fabricates a «name»
            // placeholder for REF/MERGEFIELD/STYLEREF/DOCPROPERTY because
            // `text` is absent from the prop bag. Flag the empty-but-present
            // result so TryEmitFieldRun passes `text=""` — AddField then emits
            // an empty result run, faithfully preserving "no cached result".
            else if (string.IsNullOrEmpty(display))
                synth.Format["_emptyFieldResult"] = true;
            // BUG-R12A(BUG1): carry the cached result run's run-level formatting
            // (bold/italic/color/size/font/font.latin) under a `_resultFmt.`
            // prefix so TryEmitFieldRun can map AddField-supported keys onto the
            // `add field` prop bag. Forwarded verbatim — TryEmitFieldRun decides
            // which keys AddField can honour and which need a raw-set fallback.
            if (firstFormattedResult != null)
            {
                foreach (var (fk, fv) in firstFormattedResult.Format)
                {
                    if (fv == null) continue;
                    if (FieldResultFormatKeys.Contains(fk))
                        synth.Format["_resultFmt." + fk] = fv;
                }
            }
            // BUG-DUMP-FIELDVALIGN: forward the field-wide vertAlign captured
            // from any result run in the chain. Stash it under the same
            // `_resultFmt.` channel TryEmitFieldRun already drains so it maps
            // onto the `add field` superscript/subscript prop — covering the
            // empty-result / pre-separate-rPr-run shapes the post-separate
            // firstFormattedResult scan misses. Only set when the
            // firstFormattedResult path didn't already carry it (no override).
            if (fieldVertAlign != null
                && !synth.Format.ContainsKey("_resultFmt.superscript")
                && !synth.Format.ContainsKey("_resultFmt.subscript"))
            {
                synth.Format["_resultFmt." + fieldVertAlign] = true;
            }
            // BUG-DUMP18-02: propagate hyperlink-scope hint from the begin
            // run so the field-emit branch can target the hyperlink parent
            // on replay.
            if (c.Format.TryGetValue("_hyperlinkParent", out var hlp) && hlp != null)
                synth.Format["_hyperlinkParent"] = hlp;
            // BUG-DUMP-DELFIELD: a field wrapped in a <w:del>/<w:ins> revision
            // collapses N runs (begin/instr/separate/result/end) into one synth.
            // Every constituent run carries the same revision.* attribution from
            // the shared wrapper; the synth must inherit it so TryEmitFieldRun
            // re-emits `add field` with revision.type=del/ins (rebuilding the
            // <w:del>/<w:ins> + <w:delInstrText>/<w:delText> on replay). Without
            // this the deletion is dropped and tracked-deleted field text is
            // resurrected as live document text. Read from the begin run (c) —
            // all runs in the chain share the one wrapper.
            PropagateRevisionKeys(c, synth);
            result.Add(synth);
            i = end;
        }
        return result;
    }

    // Build the prop bag AddField consumes from a parsed field instruction.
    // Returns null when the instruction is empty or its first token is not a
    // known field code; the caller falls back to a plain-text run for the
    // cached display value so the paragraph still renders.
    private static Dictionary<string, string>? BuildFieldAddProps(string instruction, string display)
    {
        if (string.IsNullOrWhiteSpace(instruction)) return null;
        var trimmed = instruction.Trim();
        // First whitespace-separated token is the field code.
        var firstSpace = trimmed.IndexOfAny(new[] { ' ', '\t' });
        var code = (firstSpace < 0 ? trimmed : trimmed[..firstSpace]).ToUpperInvariant();
        var rest = firstSpace < 0 ? "" : trimmed[(firstSpace + 1)..].Trim();

        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["fieldType"] = code
        };
        switch (code)
        {
            case "PAGE":
            case "NUMPAGES":
            case "AUTHOR":
            case "TITLE":
            case "SUBJECT":
            case "FILENAME":
            case "SECTION":
            case "SECTIONPAGES":
            {
                // BUG-R7A: these branches took no args of their own, so the
                // entire `rest` (general switches like `\* roman`,
                // `\* MERGEFORMAT`, `\p`, `\* arabic`) was dropped — replay
                // produced a bare ` PAGE `. Capture the whole residual as
                // trailing switches; AddField splices them back verbatim.
                if (!string.IsNullOrWhiteSpace(rest)) props["switches"] = rest.Trim();
                break;
            }
            case "DATE":
            case "TIME":
            case "CREATEDATE":
            case "SAVEDATE":
            case "PRINTDATE":
            {
                // Preserve the `\@ "MMMM d, yyyy"` format switch so dump
                // round-trips Word's locale-formatted date fields. Without
                // this, BuildFieldAddProps dropped `rest` and replay
                // produced a bare DATE field rendered in the default
                // locale (BUG-X6-3). AddField consumes the value via
                // --prop format=…
                var fmtMatch = System.Text.RegularExpressions.Regex.Match(
                    rest ?? "", "\\\\@\\s+\"([^\"]+)\"");
                if (fmtMatch.Success)
                    props["format"] = fmtMatch.Groups[1].Value;
                // BUG-R7A: the `\@ "..."` capture above kept only the date
                // picture; any remaining general switch (`\* MERGEFORMAT`,
                // `\* Upper`, …) was dropped. Strip the consumed `\@ "..."`
                // span from `rest` and carry whatever is left as trailing
                // switches so DATE keeps BOTH `\@ "yyyy"` AND `\* MERGEFORMAT`.
                var dateResidual = fmtMatch.Success
                    ? (rest ?? "").Remove(fmtMatch.Index, fmtMatch.Length)
                    : (rest ?? "");
                dateResidual = dateResidual.Trim();
                if (!string.IsNullOrEmpty(dateResidual)) props["switches"] = dateResidual;
                break;
            }
            case "REF":
            case "PAGEREF":
            case "NOTEREF":
            {
                // First arg is the bookmark name (may be quoted).
                var name = ExtractFirstArg(rest);
                if (string.IsNullOrEmpty(name)) return null;
                props["bookmarkName"] = name;
                // BUG-R7A: only the bookmark name was captured; trailing
                // switches (`\h`, `\p`, `\* MERGEFORMAT`, …) were dropped.
                // Capture the residual after the bookmark name and emit it
                // via the `switches` prop (AddField splices it back). `\h`
                // flows through `switches` here, not the legacy `hyperlink`
                // prop, so there is no double-emission.
                var refSw = ExtractTrailingSwitches(rest, name);
                if (!string.IsNullOrEmpty(refSw)) props["switches"] = refSw;
                break;
            }
            case "SEQ":
            {
                var ident = ExtractFirstArg(rest);
                if (string.IsNullOrEmpty(ident)) return null;
                props["identifier"] = ident;
                // BUG-DUMP17-01: preserve trailing switches (\* ARABIC, \r N,
                // \n, \c, \h, \s …). Without this, dump→batch round-trips
                // strip every SEQ formatting switch and replay produces a
                // bare " SEQ Figure ".
                var seqSw = ExtractTrailingSwitches(rest, ident);
                if (!string.IsNullOrEmpty(seqSw)) props["switches"] = seqSw;
                break;
            }
            case "MERGEFIELD":
            {
                var name = ExtractFirstArg(rest);
                if (string.IsNullOrEmpty(name)) return null;
                props["fieldName"] = name;
                // BUG-DUMP17-02: preserve trailing switches (\* MERGEFORMAT,
                // \b, \f, \v …). Same shape as the SEQ case above.
                var mfSw = ExtractTrailingSwitches(rest, name);
                if (!string.IsNullOrEmpty(mfSw)) props["switches"] = mfSw;
                break;
            }
            case "HYPERLINK":
            {
                // BUG-DUMP15-02: HYPERLINK may carry any combination of a base
                // URL, `\l "anchor"`, and `\o "tooltip"`. The previous code
                // checked `\l` first and returned only the anchor, dropping
                // the URL entirely; `\o` was never parsed. Parse all three
                // independently so dump→batch round-trips preserve them.
                // The first non-switch token (if any) is the base URL.
                var restStr = rest ?? "";
                if (!System.Text.RegularExpressions.Regex.IsMatch(restStr.TrimStart(), @"^\\"))
                {
                    var url = ExtractFirstArg(restStr);
                    if (!string.IsNullOrEmpty(url)) props["url"] = url;
                }
                var anchorMatch = System.Text.RegularExpressions.Regex.Match(restStr, "\\\\l\\s+\"([^\"]+)\"");
                if (anchorMatch.Success) props["anchor"] = anchorMatch.Groups[1].Value;
                var tooltipMatch = System.Text.RegularExpressions.Regex.Match(restStr, "\\\\o\\s+\"([^\"]+)\"");
                if (tooltipMatch.Success) props["tooltip"] = tooltipMatch.Groups[1].Value;
                if (!props.ContainsKey("url") && !props.ContainsKey("anchor"))
                    return null;
                break;
            }
            default:
                // BUG-DUMP7-05: AddField's switch has no case for `=`,
                // numeric expression fields like `= PAGE - 1`, or any other
                // unrecognised code. Emitting fieldType=<code> would make
                // replay throw `Unknown field type '<code>'`. Drop the
                // unhelpful fieldType and pass the full trimmed instruction
                // through `instr` instead — AddField's raw-instruction
                // fallback rebuilds the chain verbatim. Drops `fieldType`
                // entirely so the caller doesn't reject the row up-front.
                props.Remove("fieldType");
                props["instr"] = trimmed;
                break;
        }
        if (!string.IsNullOrEmpty(display))
            props["text"] = display;
        return props;
    }

    private static string ExtractFirstArg(string s)
    {
        if (string.IsNullOrEmpty(s)) return "";
        var t = s.TrimStart();
        if (t.StartsWith('"'))
        {
            var end = t.IndexOf('"', 1);
            return end > 0 ? t[1..end] : "";
        }
        var spc = t.IndexOfAny(new[] { ' ', '\t' });
        return spc < 0 ? t : t[..spc];
    }

    // Return the portion of `s` that follows the first arg (which
    // ExtractFirstArg already returned), trimmed. Used by SEQ /
    // MERGEFIELD field parsing to preserve trailing switches like
    // `\* ARABIC \r N` or `\* MERGEFORMAT` so AddField can replay them
    // verbatim. BUG-DUMP17-01 / BUG-DUMP17-02.
    private static string ExtractTrailingSwitches(string? s, string firstArg)
    {
        if (string.IsNullOrEmpty(s) || string.IsNullOrEmpty(firstArg)) return "";
        var t = s.TrimStart();
        int consumed;
        if (t.StartsWith('"'))
        {
            var end = t.IndexOf('"', 1);
            if (end < 0) return "";
            consumed = end + 1;
        }
        else
        {
            consumed = firstArg.Length;
        }
        return consumed >= t.Length ? "" : t[consumed..].Trim();
    }

    // Parse a TOC field instruction (` TOC \o "1-3" \h \u \z `) into the
    // prop bag AddToc accepts. AddToc emits the canonical instruction so
    // round-tripping the parsed props back through it lands at the same
    // OOXML even when the source instruction had extra whitespace or
    // switch ordering.
    private static Dictionary<string, string> ParseTocInstruction(string instruction)
    {
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var lvl = System.Text.RegularExpressions.Regex.Match(instruction, "\\\\o\\s+\"([^\"]+)\"");
        if (lvl.Success) props["levels"] = lvl.Groups[1].Value;
        // \h = hyperlinks (default true on AddToc, but emit explicitly for clarity)
        props["hyperlinks"] = System.Text.RegularExpressions.Regex.IsMatch(instruction, "\\\\h\\b")
            ? "true" : "false";
        // \z suppresses page numbers; absence means pageNumbers=true
        props["pageNumbers"] = System.Text.RegularExpressions.Regex.IsMatch(instruction, "\\\\z\\b")
            ? "false" : "true";
        // BUG-X5-03: \t = custom-style→level mapping ("Style;level,..."),
        // \b = bookmark scope. Capture the quoted argument so AddToc can
        // round-trip them; otherwise custom TOC switches were silently
        // dropped on dump.
        var ct = System.Text.RegularExpressions.Regex.Match(instruction, "\\\\t\\s+\"([^\"]+)\"");
        if (ct.Success) props["customStyles"] = ct.Groups[1].Value;
        var cb = System.Text.RegularExpressions.Regex.Match(instruction, "\\\\b\\s+\"([^\"]+)\"");
        if (cb.Success) props["bookmark"] = cb.Groups[1].Value;
        return props;
    }
}
