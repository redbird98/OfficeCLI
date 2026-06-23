// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // CJK line-break hooks — partial methods are eliminated by the compiler when no implementation exists
    partial void OnHtmlParagraphBegin(Paragraph para);
    partial void OnHtmlParagraphEnd(StringBuilder sb);
    partial void OnHtmlRenderText(StringBuilder sb, string text, RunProperties? rProps, string? runStyle, ref bool handled);
    partial void OnHtmlRenderBreak(string? runStyle, ref bool handled);
    // Notify overlay that a <w:tab/> was just emitted as a visible `widthPt`
    // wide spacer. Overlay must account for this width in its per-line budget
    // since the browser lays it out inline and pushes subsequent text right.
    partial void OnHtmlRenderTab(double widthPt);

    // ==================== Paragraph Content ====================

    /// <summary>
    /// True when the paragraph's tab stops include a leader (dot/hyphen/…)
    /// AND the paragraph contains a &lt;w:tab&gt;. Such paragraphs need a flex
    /// container (has-leader-tab) so the .dot-leader span's flex:1 can grow
    /// — the &lt;w:tab&gt; path otherwise renders a plain non-flex &lt;p&gt;.
    /// </summary>
    private bool ParagraphHasLeaderTab(Paragraph para)
    {
        if (!para.Descendants<TabChar>().Any()) return false;
        var tabs = para.ParagraphProperties?.Tabs?.Elements<TabStop>();
        if (tabs == null || !tabs.Any())
        {
            var tsId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (tsId != null) tabs = ResolveTabStopsFromStyle(tsId);
        }
        return tabs?.Any(t =>
            t.Leader?.InnerText is "dot" or "hyphen" or "underscore"
            or "middleDot" or "dash" or "heavy") == true;
    }

    /// <summary>
    /// True when the paragraph contains a &lt;w:tab&gt; AND its tab stops include
    /// a center- or right-aligned positional stop without a leader (the classic
    /// three-part header "Left \t Center \t Right" structure). Such paragraphs
    /// are rendered with a flex band model (has-aligned-tab) instead of the
    /// fixed-width left-aligned inline-block path: each band flex-grows and
    /// text-aligns per the upcoming stop's Val, so the segments stay on one
    /// line and land left/centre/right exactly like Word. Leader stops keep the
    /// existing has-leader-tab (TOC dot-leader) path; pure left tabs keep the
    /// inline-block path. Both are untouched by this detector.
    /// </summary>
    private bool ParagraphHasAlignedTab(Paragraph para)
    {
        if (!para.Descendants<TabChar>().Any()) return false;
        var tabs = para.ParagraphProperties?.Tabs?.Elements<TabStop>();
        if (tabs == null || !tabs.Any())
        {
            var tsId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (tsId != null) tabs = ResolveTabStopsFromStyle(tsId);
        }
        return tabs?.Any(t =>
            t.Val?.InnerText is "center" or "right"
            && t.Leader?.InnerText is null or "none") == true;
    }

    /// <summary>
    /// True when any run after <paramref name="run"/> in the paragraph carries
    /// visible content (text, symbol, drawing, or another tab). Used by the
    /// aligned-tab band renderer to recognize a trailing underlined tab (a
    /// full-width heading underline rule): when the underlined tab is the last
    /// content, the band gets a stretching bottom-border instead of a zero-width
    /// empty span. A trailing run that holds only the paragraph mark or empty
    /// rPr is not content.
    /// </summary>
    private static bool RunHasContentAfter(Run run, Paragraph para)
    {
        bool seenRun = false;
        foreach (var child in para.ChildElements)
        {
            if (!seenRun)
            {
                if (ReferenceEquals(child, run)) seenRun = true;
                continue;
            }
            switch (child)
            {
                case Run r:
                    if (r.Descendants<Text>().Any(t => !string.IsNullOrEmpty(t.Text))
                        || r.Descendants<TabChar>().Any()
                        || r.Descendants<SymbolChar>().Any()
                        || r.Descendants<Drawing>().Any()
                        || r.Descendants<CarriageReturn>().Any()
                        || r.Descendants<Break>().Any())
                        return true;
                    break;
                case Hyperlink:
                case DocumentFormat.OpenXml.Math.Paragraph:
                    return true;
            }
        }
        return false;
    }

    private void RenderParagraphHtml(StringBuilder sb, Paragraph para)
    {
        // Use <div> instead of <p> when paragraph contains block-level elements (text boxes, charts, shapes)
        var tag = HasBlockLevelDrawing(para) ? "div" : "p";
        sb.Append(BuildParagraphOpenTag(para, tag));
        RenderParagraphContentHtml(sb, para);
        sb.AppendLine($"</{tag}>");
    }

    // Builds the paragraph's opening tag (with class/style attributes). Shared
    // by RenderParagraphHtml and the mid-paragraph page-break handler, which
    // must reopen an identical <p> after closing the one interrupted by the
    // page-transition divs (otherwise </div> would close inside an open <p>).
    private string BuildParagraphOpenTag(Paragraph para, string tag)
    {
        var sb = new StringBuilder();
        sb.Append($"<{tag}");
        // Add CSS class for TOC paragraphs (suppress hyperlink styling).
        // Word does NOT apply the Hyperlink character style's color/underline to
        // the internal (\l bookmark) links a TOC field generates — entries render
        // in the toc-N paragraph style's own color (black/auto by default). The
        // styleId is unreliable as the TOC marker: Latin Word uses "TOC1"/"TOC2",
        // but WPS / Chinese Word assign numeric ids (e.g. "28") whose display NAME
        // is "toc 1". Match on the resolved style name too so both shapes get the
        // .toc suppression class. (Real body hyperlinks — <w:hyperlink r:id=…> in
        // non-TOC paragraphs — are unaffected and stay blue/underlined.)
        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        var classes = new List<string>();
        if (IsTocParagraphStyle(styleId, GetStyleName(para)))
            classes.Add("toc");
        // CONSISTENCY(run-special-content): paragraphs containing w:ptab
        // (header/footer left/center/right alignment) need a flex container
        // for the .ptab-spacer / .*-leader children to actually push their
        // siblings apart. The has-ptab class enables display:flex without
        // affecting paragraphs that don't need it.
        if (para.Descendants<PositionalTab>().Any())
            classes.Add("has-ptab");
        if (ParagraphHasLeaderTab(para))
            classes.Add("has-leader-tab");
        else if (ParagraphHasAlignedTab(para))
            classes.Add("has-aligned-tab");
        if (classes.Count > 0)
            sb.Append($" class=\"{string.Join(" ", classes)}\"");
        var pStyle = GetParagraphInlineCss(para);
        // A paragraph that anchors a wrapNone shape positioned relative to the
        // column/paragraph (e.g. checkbox rectangles floated over a label list)
        // becomes the position:relative containing block for those absolutely
        // positioned shapes (see ComputeParagraphAnchorAbsoluteCss). Without this
        // the shapes resolve against the .page box and the per-checkbox posOffset
        // is lost — they stack at the cell's left edge as a vertical ladder.
        if (ParagraphAnchorsSubParagraphShape(para))
            pStyle = string.IsNullOrEmpty(pStyle) ? "position:relative" : pStyle + ";position:relative";
        if (!string.IsNullOrEmpty(pStyle))
            sb.Append($" style=\"{pStyle}\"");
        sb.Append(">");
        return sb.ToString();
    }

    // A paragraph belongs to a TOC entry when its style is one of the toc-N
    // styles. Two authoring shapes exist: Latin Word styleId "TOC1".."TOC9"
    // (and the legacy "Contents"/"TOA"/"Index" families share the prefix idea
    // only for TOC), and WPS / localized Word where the styleId is opaque
    // (numeric) but the style display name is "toc 1".."toc 9" / "目录 1".
    // Matching either the styleId prefix or the normalized name catches both.
    private static bool IsTocParagraphStyle(string? styleId, string? styleName)
    {
        if (styleId != null && styleId.StartsWith("TOC", StringComparison.OrdinalIgnoreCase))
            return true;
        if (styleName != null)
        {
            // Normalize "toc 1" / "TOC 1" / "toc1" → compare prefix "toc".
            var trimmed = styleName.TrimStart();
            if (trimmed.StartsWith("toc", StringComparison.OrdinalIgnoreCase))
                return true;
        }
        return false;
    }

    private void RenderParagraphContentHtml(StringBuilder sb, Paragraph para)
    {
        OnHtmlParagraphBegin(para);
        _ctx.CurrentParagraphTabIndex = 0;
        _ctx.CurrentParagraphAlignedTab = ParagraphHasAlignedTab(para)
            && !ParagraphHasLeaderTab(para);
        _ctx.CurrentAlignedTabAlign = "left";
        // Mark where this paragraph's content begins so positional tabs can
        // retro-wrap the leading text into an absolute-width container.
        _ctx.CurrentParagraphTabSegmentStart = sb.Length;

        // Render bookmark anchors for internal hyperlink targets
        foreach (var bm in para.Elements<BookmarkStart>())
        {
            var bmName = bm.Name?.Value;
            if (!string.IsNullOrEmpty(bmName) && !bmName.StartsWith("_GoBack"))
                sb.Append($"<a id=\"{HtmlEncodeAttr(bmName)}\"></a>");
        }

        // Collect standalone images that precede a text box group (they overlay the group in Word)
        bool hasTextBoxGroup = HasTextBoxContent(para);
        var preGroupImages = hasTextBoxGroup ? new List<Drawing>() : null;
        bool textBoxSeen = false;
        // FORMCHECKBOX fields cache the glyph as a literal <w:t>☐/☑</w:t> between
        // fldChar.Separate and fldChar.End. Begin already emits the glyph, so
        // suppress the cached run to avoid rendering the checkbox twice.
        bool skipCachedCheckboxDisplay = false;

        foreach (var child in para.ChildElements)
        {
            if (child is Run run)
            {
                var runFldChar = run.GetFirstChild<FieldChar>()?.FieldCharType?.Value;
                if (runFldChar == FieldCharValues.Begin
                    && run.GetFirstChild<FieldChar>()!.GetFirstChild<FormFieldData>()?.GetFirstChild<CheckBox>() != null)
                {
                    RenderRunHtml(sb, run, para);
                    skipCachedCheckboxDisplay = true;
                    continue;
                }
                if (skipCachedCheckboxDisplay)
                {
                    if (runFldChar == FieldCharValues.End)
                        skipCachedCheckboxDisplay = false;
                    continue;
                }
                // Find drawing (direct child or inside mc:AlternateContent Choice)
                // SDK's Descendants<Drawing>() naturally skips mc:Fallback (VML w:pict)
                var drawing = run.GetFirstChild<Drawing>() ?? run.Descendants<Drawing>().FirstOrDefault();

                if (drawing != null && HasGroupOrShape(drawing))
                {
                    bool hasTextBox = HasTextBox(drawing);
                    if (hasTextBox && preGroupImages != null)
                    {
                        // Render group with preceding images overlaid into text box
                        RenderDrawingWithOverlaidImages(sb, drawing, preGroupImages);
                        preGroupImages.Clear();
                        textBoxSeen = true;
                    }
                    else
                    {
                        RenderDrawingHtml(sb, drawing);
                    }
                    continue;
                }

                // Collect standalone images before text box group for overlay
                if (hasTextBoxGroup && !textBoxSeen && drawing != null)
                {
                    preGroupImages!.Add(drawing);
                    continue;
                }

                RenderRunHtml(sb, run, para);
            }
            else if (child.LocalName is "ins" or "moveTo")
            {
                // Tracked insertions — underline to match Word's default revision mark style
                var author = child.GetAttributes().FirstOrDefault(a => a.LocalName == "author").Value;
                var authorAttr = string.IsNullOrEmpty(author) ? "" : $" title=\"Inserted by {HtmlEncodeAttr(author)}\"";
                sb.Append($"<span class=\"track-ins\" style=\"text-decoration:underline;color:#2E7D32\"{authorAttr}>");
                // Walk all nested runs so a <w:del> or <w:hyperlink> nested
                // inside <w:ins> doesn't drop its content (Descendants<Run>
                // picks up runs at any depth).
                foreach (var insRun in child.Descendants<Run>())
                    RenderRunHtml(sb, insRun, para);
                // Also render nested deletion text (ins-of-del revision) so
                // the reader sees what was removed within the insertion.
                var nestedDelText = string.Concat(child.Descendants()
                    .Where(e => e.LocalName is "del" or "moveFrom")
                    .SelectMany(d => d.Descendants())
                    .Where(e => e.LocalName is "delText" or "t")
                    .Select(e => e.InnerText));
                if (!string.IsNullOrEmpty(nestedDelText))
                    sb.Append($"<span class=\"track-del\" style=\"text-decoration:line-through;color:#C62828\">{HtmlEncode(nestedDelText)}</span>");
                sb.Append("</span>");
            }
            else if (child.LocalName is "del" or "moveFrom")
            {
                // Tracked deletions — strikethrough with color, preserving the deleted text
                // The delText inside del runs carries the actual deleted content; we render it so
                // a reader of the preview can see what was removed.
                var author = child.GetAttributes().FirstOrDefault(a => a.LocalName == "author").Value;
                var authorAttr = string.IsNullOrEmpty(author) ? "" : $" title=\"Deleted by {HtmlEncodeAttr(author)}\"";
                var delText = string.Concat(child.Descendants()
                    .Where(e => e.LocalName == "delText" || e.LocalName == "t")
                    .Select(e => e.InnerText));
                if (!string.IsNullOrEmpty(delText))
                    sb.Append($"<span class=\"track-del\" style=\"text-decoration:line-through;color:#C62828\"{authorAttr}>{HtmlEncode(delText)}</span>");
            }
            else if (child is Hyperlink hyperlink)
            {
                RenderHyperlinkHtml(sb, hyperlink, para);
            }
            else if (child.LocalName == "oMath" || child is M.OfficeMath)
            {
                var latex = FormulaParser.ToLatex(child);
                sb.Append($"<span class=\"katex-formula\" data-formula=\"{HtmlEncodeAttr(latex)}\"></span>");
            }
            else if (child.LocalName is "sdt" or "smartTag" or "customXml" or "fldSimple")
            {
                // Content controls, smart tags, custom XML, simple fields —
                // render hyperlinks with href + their own runs (TOC entries
                // are authored as <w:fldSimple> wrapping <w:hyperlink>),
                // then render bare runs. Runs nested inside a hyperlink are
                // emitted by the hyperlink branch so skip them at the
                // outer Run pass.
                var emittedRuns = new HashSet<OpenXmlElement>();
                foreach (var innerHyp in child.Descendants<Hyperlink>())
                {
                    RenderHyperlinkHtml(sb, innerHyp, para);
                    foreach (var r in innerHyp.Descendants<Run>())
                        emittedRuns.Add(r);
                }
                foreach (var innerRun in child.Descendants<Run>())
                {
                    if (emittedRuns.Contains(innerRun)) continue;
                    RenderRunHtml(sb, innerRun, para);
                }
            }
        }

        OnHtmlParagraphEnd(sb);
    }

    // ==================== Run Rendering ====================

    private void RenderRunHtml(StringBuilder sb, Run run, Paragraph para)
    {
        // Check for drawing (direct or inside mc:AlternateContent)
        var drawing = run.GetFirstChild<Drawing>()
            ?? run.Descendants<Drawing>().FirstOrDefault();
        if (drawing != null)
        {
            RenderDrawingHtml(sb, drawing);
            return;
        }

        // VML legacy picture (<w:pict>). The full geometry rendering is
        // deferred (see KNOWN_ISSUES #7e); as a safety net, extract any
        // text content so WordArt strings and textbox text don't vanish
        // from the preview entirely.
        var vmlPict = run.ChildElements.FirstOrDefault(c => c.LocalName == "pict");
        if (vmlPict != null)
        {
            // v:textbox → w:txbxContent → w:t
            var txbxTexts = vmlPict.Descendants().Where(e => e.LocalName == "t").Select(e => e.InnerText);
            // v:textpath string="..." (WordArt / classic watermark)
            var textpathStrings = vmlPict.Descendants()
                .Where(e => e.LocalName == "textpath")
                .Select(e => e.GetAttributes().FirstOrDefault(a => a.LocalName == "string").Value ?? "");
            var text = string.Join(" ", txbxTexts.Concat(textpathStrings).Where(s => !string.IsNullOrWhiteSpace(s)));
            if (!string.IsNullOrWhiteSpace(text))
                sb.Append($"<span class=\"vml-fallback\" style=\"color:#666;font-style:italic\">{HtmlEncode(text)}</span>");
            return;
        }

        // OLE embedded objects (Visio, Excel, etc.) carry a v:imagedata
        // preview image that we can render for a read-only snapshot.
        var oleObject = run.GetFirstChild<EmbeddedObject>();
        if (oleObject != null)
        {
            RenderOlePreviewHtml(sb, oleObject);
            return;
        }

        // Form field checkbox: fldChar begin with ffData/ffCheckBox — emit ☑ / ☐ glyph
        var fldChar = run.GetFirstChild<FieldChar>();
        if (fldChar?.FieldCharType?.Value == FieldCharValues.Begin)
        {
            var ffData = fldChar.GetFirstChild<FormFieldData>();
            var checkBox = ffData?.GetFirstChild<CheckBox>();
            if (checkBox != null)
            {
                var defaultChecked = checkBox.GetFirstChild<DefaultCheckBoxFormFieldState>()?.Val?.Value == true;
                var currentChecked = checkBox.GetFirstChild<Checked>()?.Val?.Value == true;
                var isChecked = currentChecked || defaultChecked;
                sb.Append(isChecked ? "☑" : "☐");
                return;
            }
        }

        // Footnote/endnote reference — render superscript number (don't return, run may also have text)
        var fnRef = run.GetFirstChild<FootnoteReference>();
        if (fnRef?.Id?.HasValue == true && fnRef.Id.Value > 0)
        {
            var fnId = (int)fnRef.Id.Value;
            _ctx.FootnoteRefs.Add(fnId);
            // #8a: when the current section has numRestart=eachSect, the
            // displayed number counts from 1 within that section; otherwise
            // it's the document-wide running total.
            int displayNum;
            if (_ctx.FnRestartEachSection)
            {
                _ctx.FnCountInSection++;
                displayNum = _ctx.FnCountInSection;
            }
            else
            {
                displayNum = _ctx.FootnoteRefs.Count;
            }
            var fnLabel = FormatNoteNumber(displayNum, GetFootnoteNumFmt());
            _ctx.FnLabels[fnId] = fnLabel;
            sb.Append($"<sup class=\"fn-ref\" style=\"{ResolveNoteRefSupStyle(run, para)}\"><a href=\"#fn{fnId}\" id=\"fnref{fnId}\" style=\"color:inherit;text-decoration:none\">{fnLabel}</a></sup>");
        }
        var enRef = run.GetFirstChild<EndnoteReference>();
        if (enRef?.Id?.HasValue == true && enRef.Id.Value > 0)
        {
            var enId = (int)enRef.Id.Value;
            _ctx.EndnoteRefs.Add(enId);
            var enNum = _ctx.EndnoteRefs.Count;
            var enLabel = FormatNoteNumber(enNum, GetEndnoteNumFmt());
            sb.Append($"<sup class=\"en-ref\" style=\"{ResolveNoteRefSupStyle(run, para)}\"><a href=\"#en{enId}\" id=\"enref{enId}\" style=\"color:inherit;text-decoration:none\">{enLabel}</a></sup>");
        }
        // FootnoteReferenceMark / EndnoteReferenceMark: don't skip the run, just ignore the mark element
        // (the run may also contain text that should be rendered)

        // Ruby (furigana) annotation — emit <ruby>base<rt>annotation</rt></ruby>
        var ruby = run.ChildElements.FirstOrDefault(c => c.LocalName == "ruby");
        if (ruby != null)
        {
            var rubyBase = ruby.ChildElements.FirstOrDefault(c => c.LocalName == "rubyBase");
            var rt = ruby.ChildElements.FirstOrDefault(c => c.LocalName == "rt");
            var baseText = string.Concat(rubyBase?.Descendants<Text>().Select(t => t.Text) ?? []);
            var rtText = string.Concat(rt?.Descendants<Text>().Select(t => t.Text) ?? []);
            if (!string.IsNullOrEmpty(baseText))
            {
                sb.Append($"<ruby>{HtmlEncode(baseText)}<rt>{HtmlEncode(rtText)}</rt></ruby>");
                return;
            }
        }

        var hasContent = run.ChildElements.Any(c =>
            c is Break || c is TabChar || c is SymbolChar || c is CarriageReturn
            // CONSISTENCY(run-special-content): PositionalTab is rendered as
            // a flex spacer (or leader span) by the ptab branch below — must
            // pass the hasContent gate or the run gets silently early-
            // returned, leaving header/footer left/center/right segments
            // collapsed in the html preview.
            || c is PositionalTab
            || c.LocalName is "noBreakHyphen" or "softHyphen"
            || (c is Text t && !string.IsNullOrEmpty(t.Text)));

        if (!hasContent) return;

        var rProps = ResolveEffectiveRunProperties(run, para);
        // w:vanish / w:specVanish — hidden text should be omitted from the
        // visual preview, matching native Word's default view behavior.
        if (rProps.Vanish != null && (rProps.Vanish.Val == null || rProps.Vanish.Val.Value))
            return;
        if (rProps.SpecVanish != null && (rProps.SpecVanish.Val == null || rProps.SpecVanish.Val.Value))
            return;
        var style = GetRunInlineCss(rProps, para);

        // Format revision (w:rPrChange) — a tracked formatting change. The
        // final format is already applied via GetRunInlineCss above; mirror
        // the ins/del revision marks by adding a restrained format-revision
        // indicator (dashed underline + author tooltip) so the reader sees
        // "this run's formatting was changed under track-changes" rather
        // than just the end result. The <w:rPrChange> element (the SDK's
        // RunPropertiesChange, which carries the author/date) lives on the
        // run's own direct rPr (not the style/docDefaults chain) — read it
        // there. PreviousRunProperties is only its inner snapshot.
        var rPrChange = run.RunProperties?.GetFirstChild<RunPropertiesChange>();
        string fmtRevClass = "";
        string fmtRevTitle = "";
        if (rPrChange != null)
        {
            // text-decoration:underline dashed doesn't collide with the
            // ins underline / del line-through (different color + dashed
            // style), and leaves the final font formatting untouched.
            style = string.IsNullOrEmpty(style)
                ? "text-decoration:underline dashed #6A1B9A;text-decoration-thickness:1px"
                : style + ";text-decoration:underline dashed #6A1B9A;text-decoration-thickness:1px";
            fmtRevClass = " class=\"track-fmt\"";
            var fmtAuthor = rPrChange.Author?.Value;
            fmtRevTitle = string.IsNullOrEmpty(fmtAuthor)
                ? " title=\"Formatting changed\""
                : $" title=\"Formatted by {HtmlEncodeAttr(fmtAuthor)}\"";
        }
        var needsSpan = !string.IsNullOrEmpty(style);

        // When line-break tracking is active, text is buffered and flushed later
        // with style spans — skip the outer span to avoid double-wrapping
        if (needsSpan && !_ctx.LineBreakEnabled)
            sb.Append($"<span{fmtRevClass}{fmtRevTitle} style=\"{style}\">");

        foreach (var child in run.ChildElements)
        {
            if (child is Break brk)
            {
                if (brk.Type?.Value == BreakValues.Page)
                {
                    // The PAGE_BREAK marker is later split on and replaced by
                    // page-transition </div>...<div> markup. If emitted while
                    // the run <span> and paragraph <p> are still open, those
                    // </div>s would close inside an open <p>/<span> (invalid
                    // nesting). Close the span + paragraph first, emit the
                    // marker, then reopen an identical <p> (and the span) for
                    // any remaining runs on the new page. Mirrors the column
                    // break branch below.
                    var pTag = HasBlockLevelDrawing(para) ? "div" : "p";
                    if (needsSpan) sb.Append("</span>");
                    sb.Append($"</{pTag}>");
                    sb.Append("<!--PAGE_BREAK-->");
                    sb.Append(BuildParagraphOpenTag(para, pTag));
                    if (needsSpan) sb.Append($"<span style=\"{style}\">");
                }
                else if (brk.Type?.Value == BreakValues.Column)
                {
                    // Close current span/paragraph, insert block-level column break, reopen
                    if (needsSpan) sb.Append("</span>");
                    sb.Append("</p><p style=\"break-before:column\">");
                    if (needsSpan) sb.Append($"<span style=\"{style}\">");
                }
                else
                {
                    // When CJK line-break tracking is active, text from <w:t>
                    // is buffered (via OnHtmlRenderText) for post-measurement
                    // flush; emitting <br> directly to sb here would land it
                    // BEFORE the buffered text in the output. Route through
                    // OnHtmlRenderBreak so the overlay can buffer the break
                    // in document order alongside text.
                    bool brkHandled = false;
                    OnHtmlRenderBreak(style, ref brkHandled);
                    if (!brkHandled) sb.Append("<br>");
                }
            }
            else if (child is TabChar)
            {
                // Resolve tab stops: direct on paragraph, or via its style
                var tabs = para.ParagraphProperties?.Tabs?.Elements<TabStop>();
                if (tabs == null || !tabs.Any())
                {
                    var tsId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                    if (tsId != null) tabs = ResolveTabStopsFromStyle(tsId);
                }
                // Aligned-tab band model (three-part "Left \t Center \t Right"
                // header): close the current leading text in a flex band whose
                // text-align matches the band's own alignment, then arm the next
                // band's alignment from THIS tab stop's Val. Each band flex:1, so
                // the segments share the line evenly and land left/centre/right
                // without overflowing onto a second line. Bypasses the
                // fixed-width inline-block path (which ignores Val and overflows).
                if (_ctx.CurrentParagraphAlignedTab)
                {
                    if (needsSpan) { sb.Append("</span>"); needsSpan = false; }
                    var bandStart = _ctx.CurrentParagraphTabSegmentStart;
                    var bandAlign = _ctx.CurrentAlignedTabAlign;
                    // Underlined (or otherwise decorated) tab to a stop with no
                    // trailing content: Word draws the run's text-decoration
                    // continuously across the tab gap, producing a full-width
                    // underlined heading separator ("Experience" + an underline
                    // rule out to the right tab stop). The band model otherwise
                    // renders the tab run as a zero-width empty span, so the
                    // underline stops at the word. Detect this case (decoration on
                    // the tab run AND nothing after this tab) and render the band
                    // with a stretching bottom-border that spans to the next stop.
                    bool underlineTab = !string.IsNullOrEmpty(style)
                        && style.Contains("text-decoration:underline", StringComparison.Ordinal);
                    bool tabIsLast = underlineTab && !RunHasContentAfter(run, para);
                    if (bandStart >= 0 && bandStart <= sb.Length)
                    {
                        var leading = sb.ToString(bandStart, sb.Length - bandStart);
                        sb.Length = bandStart;
                        if (tabIsLast)
                        {
                            // The band grows (flex:1) to fill the row; a solid
                            // bottom-border on it draws the underline continuously
                            // from the leading text out to the right edge. Suppress
                            // the trailing empty band so this band is the only flex
                            // child and spans the full content width.
                            sb.Append($"<span class=\"atab-band\" style=\"text-align:{bandAlign};border-bottom:1px solid currentColor\">{leading}</span>");
                        }
                        else
                        {
                            sb.Append($"<span class=\"atab-band\" style=\"text-align:{bandAlign}\">{leading}</span>");
                        }
                    }
                    if (tabIsLast)
                    {
                        // -1 makes the trailing-band logic skip (bandStart >= 0
                        // guard), so no empty right band steals flex space.
                        _ctx.CurrentParagraphTabIndex++;
                        _ctx.CurrentParagraphTabSegmentStart = -1;
                        continue;
                    }
                    // Arm the alignment for the band this tab opens.
                    var alignedStops = tabs?
                        .Where(t => t.Val?.InnerText != "clear" && t.Position?.HasValue == true)
                        .OrderBy(t => t.Position!.Value).ToList();
                    int aIdx = _ctx.CurrentParagraphTabIndex;
                    var nextVal = (alignedStops != null && aIdx < alignedStops.Count)
                        ? alignedStops[aIdx].Val?.InnerText : null;
                    _ctx.CurrentAlignedTabAlign = nextVal switch
                    {
                        "center" => "center",
                        "right" or "end" => "right",
                        _ => "left",
                    };
                    _ctx.CurrentParagraphTabIndex++;
                    _ctx.CurrentParagraphTabSegmentStart = sb.Length;
                    if (!string.IsNullOrEmpty(style) && !_ctx.LineBreakEnabled)
                    { sb.Append($"<span style=\"{style}\">"); needsSpan = true; }
                    continue;
                }
                // TOC-style special case: right-aligned tab with any leader.
                // Dot/hyphen/underscore/middleDot all fill the gap between
                // the current inline position and the right edge of the
                // content box via a flex-grow spacer.
                // BUG(multi-tab-leader): this must fire only when the CURRENT
                // tab stop (at CurrentParagraphTabIndex) is itself the
                // right+leader stop — not merely because the paragraph
                // contains one somewhere. Otherwise a leading center tab
                // (tabIdx 0) in a "center, right:dot" paragraph is wrongly
                // turned into a dot-leader and the center positioning is lost.
                var leaderOrderedStops = tabs?
                    .Where(t => t.Val?.InnerText != "clear" && t.Position?.HasValue == true)
                    .OrderBy(t => t.Position!.Value).ToList();
                int curTabIdx = _ctx.CurrentParagraphTabIndex;
                var curStop = (leaderOrderedStops != null && curTabIdx < leaderOrderedStops.Count)
                    ? leaderOrderedStops[curTabIdx] : null;
                var rightLeaderTab = (curStop?.Val?.InnerText == "right"
                    && curStop.Leader?.InnerText is "dot" or "hyphen" or "underscore" or "middleDot" or "dash" or "heavy")
                    ? curStop : null;
                if (rightLeaderTab != null)
                {
                    if (needsSpan) { sb.Append("</span>"); needsSpan = false; }
                    var leaderClass = rightLeaderTab.Leader?.InnerText switch
                    {
                        "hyphen" or "dash" => "hyphen-leader",
                        "underscore" or "heavy" => "underscore-leader",
                        "middleDot" => "middledot-leader",
                        _ => "dot-leader",
                    };
                    sb.Append($"<span class=\"{leaderClass}\"></span>");
                    // Advance segment tracking so a following tab (rare for a
                    // right-leader, but keeps multi-tab bookkeeping correct)
                    // measures from here, and reopen the run style span if any.
                    _ctx.CurrentParagraphTabSegmentStart = sb.Length;
                    _ctx.CurrentParagraphTabIndex++;
                    if (!string.IsNullOrEmpty(style) && !_ctx.LineBreakEnabled)
                    { sb.Append($"<span style=\"{style}\">"); needsSpan = true; }
                }
                else
                {
                    // General tab: emit inline-block with width = distance to Nth tab stop
                    // (or default 36pt = 0.5in fallback when no custom stops defined)
                    var orderedStops = tabs?
                        .Where(t => t.Val?.InnerText != "clear" && t.Position?.HasValue == true)
                        .OrderBy(t => t.Position!.Value).ToList();
                    double widthPt;
                    int tabIdx = _ctx.CurrentParagraphTabIndex;
                    if (orderedStops != null && tabIdx < orderedStops.Count)
                    {
                        var curPos = orderedStops[tabIdx].Position!.Value / 20.0; // twips → pt
                        var prevPos = tabIdx > 0 ? orderedStops[tabIdx - 1].Position!.Value / 20.0 : 0;
                        widthPt = curPos - prevPos;
                        // Tab stops are absolute positions from the page margin, but
                        // the paragraph content box is shifted right by the paragraph's
                        // left indent (rendered as padding-left on the <p>). For the
                        // FIRST tab segment (prevPos==0) the box starts at the indented
                        // origin, so to land its right edge on the absolute tab position
                        // we subtract the left indent from the box width. Later segments
                        // continue from the previous segment's absolute end, so they
                        // keep the plain (curPos - prevPos) width.
                        if (tabIdx == 0)
                        {
                            var leftIndentPt = GetParagraphLeftIndentPt(para);
                            widthPt = Math.Max(0, curPos - leftIndentPt);
                        }
                        // Handle tab leader for positional tabs. OOXML values:
                        //   none, dot, hyphen, underscore, heavy, middleDot (spec)
                        //   some authors also emit "dash" as a hyphen alias.
                        var leader = orderedStops[tabIdx].Leader?.InnerText;
                        var cssLeader = leader switch
                        {
                            "dot" => "border-bottom:1px dotted #000;",
                            // middleDot is centered dot between stops — best CSS equivalent is a
                            // thicker dotted border with larger spacing; browsers render dotted
                            // borders with square dots which read as middle dots at 2px width.
                            "middleDot" => "border-bottom:2px dotted #555;",
                            "hyphen" or "dash" => "border-bottom:1px dashed #000;",
                            "underscore" or "heavy" => "border-bottom:1px solid #000;",
                            _ => "",
                        };
                        // Tab absolute-alignment fix: instead of an EMPTY fixed-width
                        // spacer emitted AFTER the leading text (which makes the
                        // following text start at natural_text_width + gap — varying
                        // with text length), RETRO-WRAP the leading text since the
                        // last tab segment INSIDE a fixed-width inline-block. The
                        // following text then starts at the absolute tab position
                        // regardless of leading-text length (matches Word's tab stops).
                        if (needsSpan) { sb.Append("</span>"); needsSpan = false; }
                        var segStart = _ctx.CurrentParagraphTabSegmentStart;
                        if (segStart >= 0 && segStart <= sb.Length)
                        {
                            var leading = sb.ToString(segStart, sb.Length - segStart);
                            sb.Length = segStart;
                            // Use min-width (not fixed width): a tab advances the
                            // pen to AT LEAST the stop position, so when the leading
                            // text is shorter than the gap the box is exactly widthPt
                            // (following text lands on the absolute tab stop) and when
                            // it is longer the box grows to fit and pushes the
                            // following text past the stop — matching Word. A fixed
                            // width + overflow:hidden/nowrap instead CLIPPED any
                            // leading text wider than the gap, silently dropping
                            // visible body text (Word never clips at a tab stop).
                            sb.Append($"<span style=\"display:inline-block;min-width:{widthPt:0.##}pt;{cssLeader}vertical-align:bottom\">{leading}</span>");
                        }
                        else
                        {
                            // No tracked segment (shouldn't happen) — fall back to the
                            // original empty spacer to preserve the gap.
                            sb.Append($"<span style=\"display:inline-block;width:{widthPt:0.##}pt;{cssLeader}\"></span>");
                        }
                        // Next segment's leading text starts here (after this container).
                        _ctx.CurrentParagraphTabSegmentStart = sb.Length;
                        OnHtmlRenderTab(widthPt);
                        if (!string.IsNullOrEmpty(style) && !_ctx.LineBreakEnabled)
                        { sb.Append($"<span style=\"{style}\">"); needsSpan = true; }
                    }
                    else
                    {
                        // No explicit tab stop: use document-level defaultTabStop
                        // from settings.xml (twips → pt); fallback to 36pt (0.5in)
                        // when settings are missing.
                        var dts = _doc.MainDocumentPart?.DocumentSettingsPart?.Settings?.GetFirstChild<DefaultTabStop>();
                        double defTabPt = 36.0;
                        if (dts?.Val?.HasValue == true && dts.Val.Value > 0)
                            defTabPt = dts.Val.Value / 20.0;
                        sb.Append($"<span style=\"display:inline-block;width:{defTabPt:0.##}pt\"></span>");
                        OnHtmlRenderTab(defTabPt);
                    }
                    _ctx.CurrentParagraphTabIndex++;
                }
            }
            else if (child is PositionalTab ptabChild)
            {
                // CONSISTENCY(run-special-content): w:ptab is the OOXML
                // primitive Word emits in headers/footers to anchor
                // left/center/right alignment regions. Without a render
                // branch the html preview silently dropped these and the
                // three header segments collapsed into a single line.
                // Emit a flex-grow spacer (uses existing leader CSS classes
                // when a leader is set, otherwise a plain ptab-spacer with
                // fallback min-width so the gap is still visible inside
                // non-flex paragraphs). For paragraphs hosting ptabs the
                // outer container is already widened to flex via the
                // has-ptab class added in RenderParagraphHtml.
                if (needsSpan) { sb.Append("</span>"); needsSpan = false; }
                var ptabLeader = ptabChild.Leader?.HasValue == true
                    ? ptabChild.Leader.InnerText : null;
                var ptabClass = ptabLeader switch
                {
                    "dot" => "dot-leader",
                    "hyphen" or "dash" => "hyphen-leader",
                    "underscore" or "heavy" => "underscore-leader",
                    "middleDot" => "middledot-leader",
                    _ => "ptab-spacer",
                };
                sb.Append($"<span class=\"{ptabClass}\"></span>");
            }
            else if (child is CarriageReturn)
                sb.Append("<br>");
            else if (child.LocalName == "noBreakHyphen")
                sb.Append("\u2011"); // non-breaking hyphen
            else if (child.LocalName == "softHyphen")
                sb.Append("&shy;");
            else if (child is Text t && !string.IsNullOrEmpty(t.Text))
            {
                bool handled = false;
                OnHtmlRenderText(sb, t.Text, rProps, style, ref handled);
                if (!handled)
                    sb.Append(HtmlEncode(t.Text));
            }
            else if (child is SymbolChar sym)
            {
                // w:sym — render with correct font family for symbol fonts
                var charCode = sym.Char?.Value;
                var symFont = sym.Font?.Value;
                if (charCode != null && int.TryParse(charCode, System.Globalization.NumberStyles.HexNumber, null, out var code))
                {
                    if (symFont != null)
                        sb.Append($"<span style=\"font-family:'{CssSanitize(symFont)}'\">&#x{code:X};</span>");
                    else
                        sb.Append($"&#x{code:X};");
                }
                else
                    sb.Append("\u25A1"); // fallback: □
            }
        }

        if (needsSpan && !_ctx.LineBreakEnabled)
            sb.Append("</span>");

        // Close the trailing aligned-tab band (text after the last <w:tab/>),
        // so the final segment (e.g. the right-aligned "Right") is wrapped in a
        // flex band with the armed alignment. Only when at least one tab opened
        // a band (TabIndex > 0); a paragraph that ended up with no tab is left
        // alone.
        if (_ctx.CurrentParagraphAlignedTab && _ctx.CurrentParagraphTabIndex > 0)
        {
            var bandStart = _ctx.CurrentParagraphTabSegmentStart;
            var bandAlign = _ctx.CurrentAlignedTabAlign;
            if (bandStart >= 0 && bandStart <= sb.Length)
            {
                var leading = sb.ToString(bandStart, sb.Length - bandStart);
                sb.Length = bandStart;
                sb.Append($"<span class=\"atab-band\" style=\"text-align:{bandAlign}\">{leading}</span>");
            }
        }
    }

    // ==================== OLE Object Preview Rendering ====================

    /// <summary>
    /// Render the VML preview image that accompanies an embedded OLE object
    /// (e.g. a Visio diagram). Web-compatible formats (PNG/JPEG/GIF/SVG/WebP/BMP)
    /// render as a data-URI &lt;img&gt;; browser-unrenderable formats (EMF/WMF/TIFF)
    /// fall back to a sized placeholder &lt;div&gt;. Pure OpenXML — no GDI and no
    /// System.Drawing dependency.
    /// </summary>
    private void RenderOlePreviewHtml(StringBuilder sb, OpenXmlElement oleObj)
    {
        var imageData = oleObj.Descendants().FirstOrDefault(e => e.LocalName == "imagedata");
        if (imageData == null) return;

        // The r:id attribute lives in the relationships namespace.
        string? relId = null;
        foreach (var attr in imageData.GetAttributes())
        {
            if (attr.LocalName == "id" && (attr.NamespaceUri?.Contains("relationships") ?? false))
            {
                relId = attr.Value;
                break;
            }
        }
        if (string.IsNullOrEmpty(relId)) return;

        var dataUri = LoadImageAsDataUri(relId);
        if (dataUri == null) return;

        // Decide web-compatibility from the part's real content type, not the
        // returned data URI. PartToDataUri degrades undecodable WMF/EMF to an
        // SVG placeholder data URI; inferring from the URI string would
        // misclassify that placeholder as a renderable SVG and route the OLE
        // preview to the <img> branch instead of the sized placeholder block.
        var rawContentType = LoadImageContentType(relId);

        // Display size comes from the companion v:shape style
        // ("width:Xpt;height:Ypt"), falling back to the w:object
        // dxaOrig/dyaOrig twip attributes if the shape style is missing.
        double widthPt = 0, heightPt = 0;
        var shape = oleObj.Descendants().FirstOrDefault(e => e.LocalName == "shape");
        if (shape != null)
        {
            var styleAttr = shape.GetAttributes().FirstOrDefault(a => a.LocalName == "style").Value;
            if (!string.IsNullOrEmpty(styleAttr))
            {
                var wMatch = Regex.Match(styleAttr, @"width:([\d.]+)pt");
                var hMatch = Regex.Match(styleAttr, @"height:([\d.]+)pt");
                if (wMatch.Success)
                    double.TryParse(wMatch.Groups[1].Value,
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out widthPt);
                if (hMatch.Success)
                    double.TryParse(hMatch.Groups[1].Value,
                        System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out heightPt);
            }
        }
        if (widthPt == 0 || heightPt == 0)
        {
            foreach (var attr in oleObj.GetAttributes())
            {
                if (attr.LocalName == "dxaOrig" && int.TryParse(attr.Value, out var dxa))
                    widthPt = dxa / 20.0;
                if (attr.LocalName == "dyaOrig" && int.TryParse(attr.Value, out var dya))
                    heightPt = dya / 20.0;
            }
        }

        var widthPx = widthPt > 0 ? (long)(widthPt * 96 / 72) : 0;
        var heightPx = heightPt > 0 ? (long)(heightPt * 96 / 72) : 0;

        var ctForCompat = rawContentType ?? "";
        bool isWebCompatible = ctForCompat.Contains("image/png")
            || ctForCompat.Contains("image/jpeg")
            || ctForCompat.Contains("image/gif")
            || ctForCompat.Contains("image/svg")
            || ctForCompat.Contains("image/webp")
            || ctForCompat.Contains("image/bmp");

        if (isWebCompatible)
        {
            var widthAttr = widthPx > 0 ? $" width=\"{widthPx}\"" : "";
            var heightAttr = heightPx > 0 ? $" height=\"{heightPx}\"" : "";
            var sizeStyle = widthPx > 0 && heightPx > 0
                ? $"max-width:100%;width:{widthPx}px;height:{heightPx}px"
                : widthPx > 0
                    ? $"max-width:100%;width:{widthPx}px;height:auto"
                    : "max-width:100%";
            sb.Append($"<img src=\"{dataUri}\" alt=\"Embedded object\"{widthAttr}{heightAttr} style=\"{sizeStyle}\">");
        }
        else
        {
            // EMF / WMF / TIFF — browsers cannot render these natively.
            // Emit a sized placeholder so the layout keeps its footprint.
            var ph = widthPx > 0 && heightPx > 0
                ? $"width:{widthPx}px;height:{heightPx}px;max-width:100%"
                : "min-width:200px;min-height:100px";
            sb.Append($"<div class=\"ole-placeholder\" style=\"{ph};border:1px dashed #bbb;background:#f5f5f5;display:flex;align-items:center;justify-content:center;color:#888;font-size:13px;margin:8px 0\">");
            sb.Append("Embedded Object (preview not supported in browser)");
            sb.Append("</div>");
        }
    }

    // Footnote/endnote reference tracking is in _ctx.FootnoteRefs / _ctx.EndnoteRefs

    private void RenderFootnotesHtml(StringBuilder sb)
    {
        if (_ctx.FootnoteRefs.Count == 0) return;
        var fnPart = _doc.MainDocumentPart?.FootnotesPart;
        if (fnPart?.Footnotes == null) return;

        var fnSize = ResolveStyleFontSize("FootnoteText") ?? "10pt";
        var fnColor = ResolveStyleColor("FootnoteText");
        var fnColorCss = fnColor != null ? $";color:{fnColor}" : "";
        sb.AppendLine($"<div class=\"footnotes\" style=\"font-size:{fnSize}{fnColorCss}\">");
        sb.AppendLine("<hr style=\"margin-top:0;margin-bottom:0.5em;border:none;border-top:1px solid #757575;width:33%\">");

        var fnFmt = GetFootnoteNumFmt();
        int num = 0;
        // Snapshot: rendering a footnote body may itself reach a run carrying a
        // FootnoteReference, which appends to _ctx.FootnoteRefs. Iterating the
        // live List<int> mutates-during-enumerate → InvalidOperationException.
        foreach (var fnId in _ctx.FootnoteRefs.ToList())
        {
            num++;
            var fn = fnPart.Footnotes.Elements<Footnote>().FirstOrDefault(f => f.Id?.Value == fnId);
            if (fn == null) continue;

            // #8a: reuse the label that was stored at ref-emit time so the
            // bottom list matches the superscript. Falls back to the flat
            // running number when the ref emitter didn't cache a label
            // (e.g. footnote referenced from header/footer).
            var fnLabel = _ctx.FnLabels.TryGetValue(fnId, out var cached)
                ? cached
                : FormatNoteNumber(num, fnFmt);
            var fnSupStyle = ResolveNoteListSupStyle("FootnoteText");
            sb.Append($"<div id=\"fn{fnId}\" style=\"margin:0.3em 0\"><sup style=\"{fnSupStyle}\">{fnLabel}</sup> ");
            RenderFootnoteChildren(sb, fn);
            sb.AppendLine($" <a href=\"#fnref{fnId}\" style=\"color:inherit;text-decoration:none\">\u21A9</a></div>");
        }
        sb.AppendLine("</div>");
    }

    // Render paragraphs AND tables inside a footnote/endnote. The previous
    // implementation only iterated Elements<Paragraph>() so a footnote with
    // a nested table silently dropped the table (and when a footnote
    // contained only a table, the whole footnote rendered empty).
    private IEnumerable<OpenXmlPart> CollectHyperlinkHostParts()
    {
        var main = _doc.MainDocumentPart;
        if (main == null) yield break;
        yield return main;
        foreach (var hp in main.HeaderParts) yield return hp;
        foreach (var fp in main.FooterParts) yield return fp;
        if (main.FootnotesPart != null) yield return main.FootnotesPart;
        if (main.EndnotesPart != null) yield return main.EndnotesPart;
    }

    private void RenderHyperlinkHtml(StringBuilder sb, Hyperlink hyperlink, Paragraph para)
    {
        var relId = hyperlink.Id?.Value;
        string? url = null;
        if (relId != null)
        {
            // Hyperlink rels can live on the enclosing HeaderPart/FooterPart/
            // FootnotesPart/EndnotesPart, not just MainDocumentPart. Falling
            // back to a full-part sweep keeps header/footer links clickable.
            try
            {
                var parts = CollectHyperlinkHostParts();
                foreach (var part in parts)
                {
                    url = part.HyperlinkRelationships.FirstOrDefault(r => r.Id == relId)?.Uri?.ToString();
                    if (url != null) break;
                    url = part.ExternalRelationships.FirstOrDefault(r => r.Id == relId)?.Uri?.ToString();
                    if (url != null) break;
                }
            }
            catch { }
        }
        if (url == null && hyperlink.Anchor?.Value != null)
            url = $"#{hyperlink.Anchor.Value}";
        var urlSafe = url != null && IsSafeLinkUrl(url);
        if (urlSafe)
        {
            // CSS gotcha: an underline drawn by the ancestor <a> element cannot
            // be removed by a descendant span's text-decoration:none. When every
            // run in the hyperlink has explicit w:u val="none", suppress the <a>
            // default underline on the <a> element itself. (A default hyperlink
            // with no explicit underline state stays underlined.)
            var aStyle = HyperlinkUnderlineExplicitlyNone(hyperlink, para)
                ? " style=\"text-decoration:none\"" : "";
            var tooltip = hyperlink.Tooltip?.Value;
            var titleAttr = !string.IsNullOrEmpty(tooltip)
                ? $" title=\"{HtmlEncodeAttr(tooltip)}\"" : "";
            sb.Append($"<a href=\"{HtmlEncodeAttr(url!)}\"{(url!.StartsWith("#") ? "" : " target=\"_blank\"")}{titleAttr}{aStyle}>");
        }
        foreach (var descendant in hyperlink.Descendants<Run>())
            RenderRunHtml(sb, descendant, para);
        if (urlSafe)
            sb.Append("</a>");
    }

    /// <summary>
    /// True when the hyperlink has at least one text-bearing run and every such
    /// run resolves to an explicit w:u val="none". Used to suppress the ancestor
    /// &lt;a&gt; element's default underline (a descendant span cannot do it).
    /// </summary>
    private bool HyperlinkUnderlineExplicitlyNone(Hyperlink hyperlink, Paragraph para)
    {
        bool sawTextRun = false;
        foreach (var run in hyperlink.Descendants<Run>())
        {
            var hasText = run.ChildElements.Any(c => c is Text t && !string.IsNullOrEmpty(t.Text));
            if (!hasText) continue;
            sawTextRun = true;
            var rPr = ResolveEffectiveRunProperties(run, para);
            if (rPr.Underline?.Val == null || rPr.Underline.Val.InnerText != "none")
                return false;
        }
        return sawTextRun;
    }

    private void RenderFootnoteChildren(StringBuilder sb, OpenXmlElement note)
    {
        bool first = true;
        foreach (var child in note.ChildElements)
        {
            if (child is Paragraph p)
            {
                if (!first) sb.Append("<br>");
                RenderParagraphContentHtml(sb, p);
                first = false;
            }
            else if (child is Table tbl)
            {
                RenderTableHtml(sb, tbl);
                first = false;
            }
        }
    }

    private void RenderEndnotesHtml(StringBuilder sb)
    {
        if (_ctx.EndnoteRefs.Count == 0) return;
        var enPart = _doc.MainDocumentPart?.EndnotesPart;
        if (enPart?.Endnotes == null) return;

        var enSize = ResolveStyleFontSize("EndnoteText") ?? "10pt";
        sb.AppendLine($"<div class=\"endnotes\" style=\"font-size:{enSize}\">");
        sb.AppendLine("<hr style=\"margin-top:2em;margin-bottom:0.5em;border:none;border-top:1px solid #757575;width:33%\">");

        var enFmt = GetEndnoteNumFmt();
        int num = 0;
        foreach (var enId in _ctx.EndnoteRefs)
        {
            num++;
            var en = enPart.Endnotes.Elements<Endnote>().FirstOrDefault(e => e.Id?.Value == enId);
            if (en == null) continue;

            var enLabel = FormatNoteNumber(num, enFmt);
            var enIndent = ResolveStyleIndent("EndnoteText");
            var enIndentCss = enIndent != null ? $"text-indent:{enIndent}" : "";
            var enSupStyle = ResolveNoteListSupStyle("EndnoteText");
            sb.Append($"<div id=\"en{enId}\" style=\"margin:0.3em 0;{enIndentCss}\"><sup style=\"{enSupStyle}\">{enLabel}</sup> ");
            RenderFootnoteChildren(sb, en);
            // Back-reference link to the in-body endnote marker (id="enref{N}",
            // emitted at ref-render time). Mirrors the footnote back-link so the
            // two note lists are internally consistent.
            sb.AppendLine($" <a href=\"#enref{enId}\" style=\"color:inherit;text-decoration:none\">↩</a></div>");
        }
        sb.AppendLine("</div>");
    }

    /// <summary>Get the numbering format for footnotes (default: decimal per OOXML spec §17.11.11).</summary>
    private string GetFootnoteNumFmt()
    {
        // Priority: section properties > document settings > spec default
        var sectProps = _doc.MainDocumentPart?.Document?.Body
            ?.Descendants<SectionProperties>().LastOrDefault();
        var sectFmt = sectProps?.GetFirstChild<FootnoteProperties>()?.NumberingFormat?.Val?.InnerText;
        if (sectFmt != null) return sectFmt;

        var settingsFmt = _doc.MainDocumentPart?.DocumentSettingsPart?.Settings
            ?.GetFirstChild<FootnoteDocumentWideProperties>()?.NumberingFormat?.Val?.InnerText;
        if (settingsFmt != null) return settingsFmt;

        return "decimal";
    }

    /// <summary>Get the numbering format for endnotes (default: lowerRoman per OOXML spec §17.11.4).</summary>
    private string GetEndnoteNumFmt()
    {
        // Priority: section properties > document settings > spec default
        var sectProps = _doc.MainDocumentPart?.Document?.Body
            ?.Descendants<SectionProperties>().LastOrDefault();
        var sectFmt = sectProps?.GetFirstChild<EndnoteProperties>()?.NumberingFormat?.Val?.InnerText;
        if (sectFmt != null) return sectFmt;

        var settingsFmt = _doc.MainDocumentPart?.DocumentSettingsPart?.Settings
            ?.GetFirstChild<EndnoteDocumentWideProperties>()?.NumberingFormat?.Val?.InnerText;
        if (settingsFmt != null) return settingsFmt;

        return "lowerRoman";
    }

    /// <summary>Format a note number according to Word numbering format.</summary>
    private static string FormatNoteNumber(int num, string fmt)
    {
        return fmt switch
        {
            "lowerRoman" => ToLowerRoman(num),
            "upperRoman" => ToLowerRoman(num).ToUpperInvariant(),
            "lowerLetter" => num >= 1 && num <= 26 ? ((char)('a' + num - 1)).ToString() : num.ToString(),
            "upperLetter" => num >= 1 && num <= 26 ? ((char)('A' + num - 1)).ToString() : num.ToString(),
            _ => num.ToString(), // "decimal" and any other format
        };
    }

    private static string ToLowerRoman(int num)
    {
        if (num <= 0 || num > 3999) return num.ToString();
        var sb = new StringBuilder();
        ReadOnlySpan<(int value, string roman)> map =
        [
            (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"),
            (100, "c"), (90, "xc"), (50, "l"), (40, "xl"),
            (10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")
        ];
        foreach (var (value, roman) in map)
        {
            while (num >= value)
            {
                sb.Append(roman);
                num -= value;
            }
        }
        return sb.ToString();
    }
}
