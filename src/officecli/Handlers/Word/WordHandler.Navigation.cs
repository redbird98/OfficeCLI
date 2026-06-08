// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    // Perf: append-monotonic body caches for batch replay. Each body append
    // otherwise costs O(n) — navigation scans all paragraphs to resolve
    // /body/p[last()], and AddParagraph re-Count()s Elements<Paragraph>() to
    // name the result path — making an N-paragraph replay O(N²). Self-validating
    // against the live tree (see AppendedBodyParaCount), so no mutation site has
    // to remember to invalidate.
    private OpenXmlElement? _lastBodyParagraph;
    private int _bodyParaCount = -1;

    private void InvalidateBodyParaCache() { _lastBodyParagraph = null; _bodyParaCount = -1; ClearBodyChildIndex(); }

    // using-scope that invalidates the append/child caches when the enclosing
    // mutation method RETURNS. Remove/Move/Swap/CopyFrom navigate first (which
    // rebuilds the child-index cache to the PRE-mutation tree) and only then
    // mutate, so invalidating at entry would leave a stale cache behind. We must
    // invalidate on exit, after the structural change has happened.
    private readonly struct BodyCacheGuard : System.IDisposable
    {
        private readonly WordHandler _h;
        public BodyCacheGuard(WordHandler h) => _h = h;
        public void Dispose() => _h.InvalidateBodyParaCache();
    }

    // ==================== Navigation ====================

    /// <summary>
    /// OOXML toggle element (Bold, Italic, Strike, Caps, …) is "ON" when the
    /// element exists AND its <c>w:val</c> attribute is either absent or
    /// truthy. <c>&lt;w:b/&gt;</c> means ON; <c>&lt;w:b w:val="0"/&gt;</c>
    /// and <c>&lt;w:b w:val="false"/&gt;</c> mean explicitly OFF. Pure
    /// null-checks on the element flip the OFF case back to ON, corrupting
    /// canonical Get readback (BUG-R2-04). Use this helper at every
    /// toggle-readback site so the override is honored.
    /// </summary>
    private static bool IsToggleOn(Bold? t)   => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(Italic? t) => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(Strike? t) => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(DoubleStrike? t) => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(Caps? t) => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(SmallCaps? t) => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(Vanish? t) => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(Outline? t) => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(Shadow? t) => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(Emboss? t) => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(Imprint? t) => t != null && (t.Val == null || t.Val.Value);
    private static bool IsToggleOn(NoProof? t) => t != null && (t.Val == null || t.Val.Value);

    private DocumentNode GetRootNode(int depth)
    {
        var node = new DocumentNode { Path = "/", Type = "document" };
        var children = new List<DocumentNode>();

        var mainPart = _doc.MainDocumentPart;
        if (mainPart?.Document?.Body != null)
        {
            children.Add(new DocumentNode
            {
                Path = "/body",
                Type = "body",
                ChildCount = mainPart.Document.Body.ChildElements.Count
            });
        }

        if (mainPart?.StyleDefinitionsPart != null)
        {
            children.Add(new DocumentNode
            {
                Path = "/styles",
                Type = "styles",
                ChildCount = mainPart.StyleDefinitionsPart.Styles?.ChildElements.Count ?? 0
            });
        }

        int headerIdx = 0;
        if (mainPart?.HeaderParts != null)
        {
            foreach (var _ in mainPart.HeaderParts)
            {
                children.Add(new DocumentNode
                {
                    Path = $"/header[{headerIdx + 1}]",
                    Type = "header"
                });
                headerIdx++;
            }
        }

        int footerIdx = 0;
        if (mainPart?.FooterParts != null)
        {
            foreach (var _ in mainPart.FooterParts)
            {
                children.Add(new DocumentNode
                {
                    Path = $"/footer[{footerIdx + 1}]",
                    Type = "footer"
                });
                footerIdx++;
            }
        }

        if (mainPart?.NumberingDefinitionsPart != null)
        {
            children.Add(new DocumentNode { Path = "/numbering", Type = "numbering" });
        }

        // CONSISTENCY(footnotes-container): mirror /footnotes/footnote[N] enumeration
        // (Navigation.cs:785) — user entries only (id > 0), excluding separator/
        // continuation system rows so child counts match what `query footnote` returns.
        if (mainPart?.FootnotesPart?.Footnotes != null)
        {
            int fnCount = mainPart.FootnotesPart.Footnotes.Elements<Footnote>()
                .Count(f => f.Id?.Value > 0);
            if (fnCount > 0)
            {
                children.Add(new DocumentNode
                {
                    Path = "/footnotes",
                    Type = "footnotes",
                    ChildCount = fnCount
                });
            }
        }

        if (mainPart?.EndnotesPart?.Endnotes != null)
        {
            int enCount = mainPart.EndnotesPart.Endnotes.Elements<Endnote>()
                .Count(e => e.Id?.Value > 0);
            if (enCount > 0)
            {
                children.Add(new DocumentNode
                {
                    Path = "/endnotes",
                    Type = "endnotes",
                    ChildCount = enCount
                });
            }
        }

        if (mainPart?.WordprocessingCommentsPart?.Comments != null)
        {
            int cCount = mainPart.WordprocessingCommentsPart.Comments.Elements<Comment>().Count();
            if (cCount > 0)
            {
                children.Add(new DocumentNode
                {
                    Path = "/comments",
                    Type = "comments",
                    ChildCount = cCount
                });
            }
        }

        // Core document properties
        var props = _doc.PackageProperties;
        if (props.Title != null) node.Format["title"] = props.Title;
        if (props.Creator != null) node.Format["author"] = props.Creator;
        if (props.Subject != null) node.Format["subject"] = props.Subject;
        if (props.Keywords != null) node.Format["keywords"] = props.Keywords;
        if (props.Description != null) node.Format["description"] = props.Description;
        if (props.Category != null) node.Format["category"] = props.Category;
        if (props.LastModifiedBy != null) node.Format["lastModifiedBy"] = props.LastModifiedBy;
        if (props.Revision != null) node.Format["revisionNumber"] = props.Revision;
        if (props.Created != null) node.Format["created"] = props.Created.Value.ToString("o");
        if (props.Modified != null) node.Format["modified"] = props.Modified.Value.ToString("o");

        // BUG-DUMP10-03: surface the document-level page background color
        // (<w:document><w:background w:color="…"/>…). Without this, dump
        // dropped the page background entirely. Set side already accepts
        // the canonical `background` key (see WordHandler.Add.cs:565).
        if (mainPart?.Document?.GetFirstChild<DocumentBackground>() is { } bgEl
            && bgEl.Color?.Value is { Length: > 0 } bgColor)
        {
            node.Format["background"] = ParseHelpers.FormatHexColor(bgColor);
        }

        // Page size from last section properties (document default)
        var sectPr = mainPart?.Document?.Body?.GetFirstChild<SectionProperties>()
            ?? mainPart?.Document?.Body?.Descendants<SectionProperties>().LastOrDefault();
        if (sectPr != null)
        {
            var pageSize = sectPr.GetFirstChild<PageSize>();
            if (pageSize?.Width?.Value != null) node.Format["pageWidth"] = FormatTwipsToCm(pageSize.Width.Value);
            if (pageSize?.Height?.Value != null) node.Format["pageHeight"] = FormatTwipsToCm(pageSize.Height.Value);
            if (pageSize?.Orient?.Value != null) node.Format["orientation"] = pageSize.Orient.InnerText;
            var margins = sectPr.GetFirstChild<PageMargin>();
            if (margins != null)
            {
                if (margins.Top?.Value != null) node.Format["marginTop"] = FormatTwipsToCm((uint)Math.Abs(margins.Top.Value));
                if (margins.Bottom?.Value != null) node.Format["marginBottom"] = FormatTwipsToCm((uint)Math.Abs(margins.Bottom.Value));
                if (margins.Left?.Value != null) node.Format["marginLeft"] = FormatTwipsToCm(margins.Left.Value);
                if (margins.Right?.Value != null) node.Format["marginRight"] = FormatTwipsToCm(margins.Right.Value);
                // header/footer-from-edge distances + binding gutter (pgMar
                // @header/@footer/@gutter) — needed for dump→batch fidelity.
                if (margins.Header?.Value != null) node.Format["marginHeader"] = FormatTwipsToCm(margins.Header.Value);
                if (margins.Footer?.Value != null) node.Format["marginFooter"] = FormatTwipsToCm(margins.Footer.Value);
                if (margins.Gutter?.Value != null) node.Format["marginGutter"] = FormatTwipsToCm(margins.Gutter.Value);
            }

            // CONSISTENCY(root-vs-section-readback): the body-level sectPr surfaced at /
            // and at /section[N] (for the final section) must yield the same Format keys
            // so set/get round-trips at either path. Mirror BuildSectionNode in
            // WordHandler.Query.cs:786-863 — keep encoding identical (restart maps
            // "newPage"→"restartPage", "newSection"→"restartSection").
            var pgNumType = sectPr.GetFirstChild<PageNumberType>();
            if (pgNumType?.Start?.Value != null)
                node.Format["pageStart"] = pgNumType.Start.Value;
            if (pgNumType?.Format?.Value != null)
                node.Format["pageNumFmt"] = pgNumType.Format.InnerText;
            // BUG-DUMP11-01: w:pgNumType also carries chapStyle (heading style
            // index for chapter numbering) and chapSep (separator between
            // chapter and page numbers). Surfaced here so the body sectPr
            // round-trips chapter-numbering config.
            if (pgNumType?.ChapterStyle?.Value != null)
                node.Format["chapStyle"] = pgNumType.ChapterStyle.Value;
            if (pgNumType?.ChapterSeparator?.Value != null)
                node.Format["chapSep"] = pgNumType.ChapterSeparator.InnerText;

            if (sectPr.GetFirstChild<TitlePage>() != null)
                node.Format["titlePage"] = true;

            // Surface pgBorders presence as the same shorthand the Set side
            // accepts ('box' / 'none'). Without readback the value silently
            // disappeared on get and never surfaced in dump, so a `set /
            // --prop pgBorders=box` round-trip lost the borders. Match the
            // Set encoding: any PageBorders element with at least one side
            // surfaces as `box` — the set side only ever writes that exact
            // shape (four single-line sides), so a present element implies
            // it. Future encodings can refine this to expose width / style
            // / color individually once Set grows beyond the shorthand.
            if (sectPr.GetFirstChild<PageBorders>() != null)
                node.Format["pgBorders"] = "box";

            // Section-level RTL (Arabic / Hebrew page direction).
            if (sectPr.GetFirstChild<BiDi>() != null)
                node.Format["direction"] = "rtl";

            // <w:rtlGutter/> places the binding gutter on the right side.
            if (sectPr.GetFirstChild<GutterOnRight>() != null)
                node.Format["rtlGutter"] = true;

            // BUG-DUMP11-03: <w:noEndnote/> on a section suppresses endnote
            // collection at section end. Bare on/off toggle (no val attr).
            if (sectPr.GetFirstChild<NoEndnote>() != null)
                node.Format["noEndnote"] = true;

            var lnNum = sectPr.GetFirstChild<LineNumberType>();
            if (lnNum != null)
            {
                var countBy = lnNum.CountBy?.Value ?? 1;
                var restartVal = lnNum.Restart?.InnerText ?? "continuous";
                node.Format["lineNumbers"] = restartVal switch
                {
                    "newPage" => "restartPage",
                    "newSection" => "restartSection",
                    _ => "continuous"
                };
                if (countBy != 1) node.Format["lineNumberCountBy"] = countBy;
                // BUG-DUMP11-02: w:lnNumType/@w:start was silently dropped.
                // Surface as canonical lineNumberStart key.
                if (lnNum.Start?.Value is short lnStart)
                    node.Format["lineNumberStart"] = (int)lnStart;
            }

            // BUG-DUMP11-04: header / footer references (default / first /
            // even) — mirror BuildSectionNode in WordHandler.Query.cs so
            // Get('/') and /section[N] surface the same headerRef.<type> /
            // footerRef.<type> keys.
            if (mainPart != null)
            {
                string? primaryHeader = null;
                foreach (var href in sectPr.Elements<HeaderReference>())
                {
                    if (href.Id?.Value == null) continue;
                    var refType = href.Type?.InnerText ?? "default";
                    try
                    {
                        var part = mainPart.GetPartById(href.Id.Value) as DocumentFormat.OpenXml.Packaging.HeaderPart;
                        if (part != null)
                        {
                            var idx = mainPart.HeaderParts.ToList().IndexOf(part);
                            if (idx >= 0)
                            {
                                var pathRef = $"/header[{idx + 1}]";
                                node.Format[$"headerRef.{refType}"] = pathRef;
                                if (primaryHeader == null || refType == "default") primaryHeader = pathRef;
                            }
                        }
                    }
                    catch { /* dangling rel — skip */ }
                }
                if (primaryHeader != null) node.Format["headerRef"] = primaryHeader;

                string? primaryFooter = null;
                foreach (var fref in sectPr.Elements<FooterReference>())
                {
                    if (fref.Id?.Value == null) continue;
                    var refType = fref.Type?.InnerText ?? "default";
                    try
                    {
                        var part = mainPart.GetPartById(fref.Id.Value) as DocumentFormat.OpenXml.Packaging.FooterPart;
                        if (part != null)
                        {
                            var idx = mainPart.FooterParts.ToList().IndexOf(part);
                            if (idx >= 0)
                            {
                                var pathRef = $"/footer[{idx + 1}]";
                                node.Format[$"footerRef.{refType}"] = pathRef;
                                if (primaryFooter == null || refType == "default") primaryFooter = pathRef;
                            }
                        }
                    }
                    catch { /* dangling rel — skip */ }
                }
                if (primaryFooter != null) node.Format["footerRef"] = primaryFooter;
            }
        }

        // Document protection
        var settings = _doc.MainDocumentPart?.DocumentSettingsPart?.Settings;
        var docProtection = settings?.GetFirstChild<DocumentProtection>();
        if (docProtection != null)
        {
            var editText = docProtection.Edit?.InnerText;
            node.Format["protection"] = editText switch
            {
                "readOnly" => "readOnly",
                "comments" => "comments",
                "trackedChanges" => "trackedChanges",
                "forms" => "forms",
                _ => "none"
            };
            var enforced = docProtection.Enforcement?.Value;
            node.Format["protectionEnforced"] = enforced == true || enforced == null && docProtection.Edit != null;
        }
        else
        {
            node.Format["protection"] = "none";
            node.Format["protectionEnforced"] = false;
        }

        // Document-level settings (DocGrid, CJK, print/display, font embedding, layout flags, columns, etc.)
        PopulateDocSettings(node);
        PopulateCompatibility(node);
        PopulateDocDefaults(node);

        // Theme and Extended Properties
        Core.ThemeHandler.PopulateTheme(_doc.MainDocumentPart?.ThemePart, node);
        Core.ExtendedPropertiesHandler.PopulateExtendedProperties(_doc.ExtendedFilePropertiesPart, node);

        node.Children = children;
        node.ChildCount = children.Count;
        return node;
    }

    private record PathSegment(string Name, int? Index, string? StringIndex = null);

    /// <summary>
    /// Resolve InsertPosition (After/Before anchor path) to a 0-based int? index.
    /// Anchor path can be full (/body/p[@paraId=xxx]) or short (p[@paraId=xxx]).
    /// </summary>
    private int? ResolveAnchorPosition(OpenXmlElement parent, string parentPath, InsertPosition? position)
    {
        if (position == null) return null;
        if (position.Index.HasValue) return position.Index;

        var anchorPath = position.After ?? position.Before!;

        // Catch bare attribute selector without element wrapper, e.g. @paraId=XXX instead of p[@paraId=XXX]
        if (System.Text.RegularExpressions.Regex.IsMatch(anchorPath, @"^@(\w+)=(.+)$"))
            throw new ArgumentException($"Invalid anchor path \"{anchorPath}\". Did you mean: p[{anchorPath}]?");

        // Handle find: prefix — text-based anchoring within a paragraph
        if (anchorPath.StartsWith("find:", StringComparison.OrdinalIgnoreCase))
        {
            // Return a sentinel value; actual handling done in Add via AddAtFindPosition
            return FindAnchorIndex;
        }

        // Normalize: if short form (no leading /), prepend parentPath
        if (!anchorPath.StartsWith("/"))
            anchorPath = parentPath.TrimEnd('/') + "/" + anchorPath;

        // Top-level /watermark[N]? special case. Watermarks are stored in
        // the header parts, not the body — there is no body-level sibling
        // that represents the watermark. `add --type watermark` returns
        // "/watermark" as the new element's identity; to keep that path
        // round-trippable as --after/--before, treat it as a no-op
        // positional hint: --after /watermark appends to parent, --before
        // /watermark prepends. Callers needing a specific body position
        // should pass an explicit /body/p[N] anchor instead.
        {
            var wmMatch = System.Text.RegularExpressions.Regex.Match(anchorPath, @"^/watermark(?:\[(\d+)\])?$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (wmMatch.Success)
            {
                // Honour the positional-hint contract only when a watermark
                // actually exists in the doc. Otherwise fall through so the
                // standard "Anchor element not found" error fires — matching
                // /chart[1] and other absent-anchor behaviour. An explicit
                // index beyond the number of watermarks (there's at most one)
                // is out-of-range — error instead of silently appending.
                var wmExists = FindWatermark() != null;
                var wmCount = wmExists ? 1 : 0;
                if (wmMatch.Groups[1].Success)
                {
                    var wmIdx = int.Parse(wmMatch.Groups[1].Value);
                    if (wmIdx < 1 || wmIdx > wmCount)
                        throw new ArgumentException($"Anchor element not found: {anchorPath}");
                }
                else if (!wmExists)
                {
                    throw new ArgumentException($"Anchor element not found: {anchorPath}");
                }
                return position.After != null ? (int?)null : 0;
            }
        }

        // Virtual table column anchor: /body/tbl[N]/col[N]. ParsePath would
        // fail because <w:col> doesn't exist in OOXML. Used by `add column
        // --before/--after col[K]` and `add --from col[K] --before/--after col[J]`.
        // Validates that the anchor exists in the named table.
        {
            var colAnchorMatch = System.Text.RegularExpressions.Regex.Match(
                anchorPath, @"^/body/tbl\[(\d+)\]/col\[(\d+)\]$");
            if (colAnchorMatch.Success)
            {
                var anchorTableIdx = int.Parse(colAnchorMatch.Groups[1].Value);
                var anchorColIdx = int.Parse(colAnchorMatch.Groups[2].Value);
                var body = _doc.MainDocumentPart?.Document?.Body;
                var tables = body?.Elements<Table>().ToList() ?? new List<Table>();
                if (anchorTableIdx < 1 || anchorTableIdx > tables.Count)
                    throw new ArgumentException($"Anchor table not found: {anchorPath} (total tables at /body: {tables.Count})");
                var anchorGrid = tables[anchorTableIdx - 1].GetFirstChild<TableGrid>();
                var gridColCount = anchorGrid?.Elements<GridColumn>().Count() ?? 0;
                if (anchorColIdx < 1 || anchorColIdx > gridColCount)
                    throw new ArgumentException($"Anchor column not found: {anchorPath} (total columns: {gridColCount})");
                return position.After != null ? anchorColIdx : anchorColIdx - 1;
            }
        }

        var segments = ParsePath(anchorPath);
        var anchor = NavigateToElement(segments, out var ctx)
            ?? throw new ArgumentException($"Anchor element not found: {anchorPath}" + (ctx != null ? $". {ctx}" : ""));

        // Body-level <w:sectPr> (direct child of Body) must remain the last
        // child of body. `--after /body/sectPr` has no valid placement;
        // silently routing to "before sectPr" (the old behaviour) misleads
        // the caller. Reject with a clear error. Paragraph-level sectPr
        // (inside w:pPr) is unaffected — its carrier paragraph is the
        // anchor, not the sectPr itself.
        if (position.After != null && anchor is SectionProperties && anchor.Parent is Body)
        {
            throw new ArgumentException(
                "Cannot insert after body-level sectPr; it must remain the last child of body. " +
                "Use --before /body/sectPr (or omit the anchor to append before sectPr).");
        }

        // Find anchor's position among parent's children
        var siblings = parent.ChildElements.ToList();
        // /body/oMathPara[N] resolves to the inner M.Paragraph/oMathPara element;
        // when it lives inside a pure wrapper w:p, the wrapper is the actual
        // body child. Re-target the anchor to that wrapper so --after/--before
        // can find it among body siblings.
        if ((anchor is M.Paragraph || anchor.LocalName == "oMathPara")
            && anchor.Parent is Paragraph wrapAnchor
            && IsOMathParaWrapperParagraph(wrapAnchor)
            && parent.ChildElements.Contains(wrapAnchor))
        {
            anchor = wrapAnchor;
        }
        var anchorIdx = siblings.IndexOf(anchor);
        if (anchorIdx < 0)
            throw new ArgumentException($"Anchor element is not a child of {parentPath}: {anchorPath}");

        // CONSISTENCY(table-row-anchor): when inserting into a <w:tbl>, the
        // body's child list also contains tblPr / tblGrid / tblPrEx, but
        // AddRow indexes against parent.Elements<TableRow>() — using the
        // ChildElements offset there would push past the tail and silently
        // AppendChild. Translate the anchor's position into row-only space
        // so the AddRow contract (index = row-only index) holds.
        if (parent is Table tbl && anchor is TableRow trAnchor)
        {
            var rows = tbl.Elements<TableRow>().ToList();
            var rowIdx = rows.IndexOf(trAnchor);
            if (rowIdx < 0)
                throw new ArgumentException($"Anchor row is not a row of {parentPath}: {anchorPath}");
            if (position.After != null)
                return rowIdx + 1 >= rows.Count ? null : rowIdx + 1;
            return rowIdx;
        }

        if (position.After != null)
        {
            // Insert after anchor: if last child, return null (append)
            return anchorIdx + 1 >= siblings.Count ? null : anchorIdx + 1;
        }
        else
        {
            // Insert before anchor
            return anchorIdx;
        }
    }

    /// <summary>Sentinel value indicating find: anchor needs text-based resolution.</summary>
    private const int FindAnchorIndex = -99999;

    /// <summary>
    /// Build an SDT path segment using @sdtId= if available, otherwise positional index.
    /// </summary>
    private static string BuildSdtPathSegment(OpenXmlElement sdt, int positionalIndex)
    {
        var sdtProps = (sdt is SdtBlock sb ? sb.SdtProperties : (sdt as SdtRun)?.SdtProperties);
        var sdtIdVal = sdtProps?.GetFirstChild<SdtId>()?.Val?.Value;
        return sdtIdVal != null
            ? $"sdt[@sdtId={sdtIdVal}]"
            : $"sdt[{positionalIndex}]";
    }

    /// <summary>
    /// Build a paragraph path segment using @paraId= if available, otherwise positional index.
    /// E.g. "p[@paraId=1A2B3C4D]" or "p[3]".
    /// </summary>
    private static string BuildParaPathSegment(Paragraph para, int positionalIndex)
    {
        var paraId = para.ParagraphId?.Value;
        return !string.IsNullOrEmpty(paraId)
            ? $"p[@paraId={paraId}]"
            : $"p[{positionalIndex}]";
    }

    // 1-based position of a just-appended paragraph among its parent's Paragraph
    // children, for naming the result path. O(1) for the body via the
    // append-monotonic cache; other parents Count() directly (not the hot path).
    // Self-validating: the incremental count is trusted only when the paragraph
    // cached last time is the element immediately preceding this freshly-appended
    // one — any out-of-band insert/remove/non-paragraph sibling breaks the chain
    // and forces one O(n) reseed, so no mutation site must invalidate manually.
    // Also refreshes _lastBodyParagraph to keep /body/p[last()] navigation warm.
    private int AppendBodyParaFast(OpenXmlElement parent, Paragraph para)
    {
        // The Open XML SDK keeps children in a singly-linked list, so
        // InsertBefore(refChild) and PreviousSibling() scan for the predecessor
        // — O(N) each, making an N-paragraph append O(N²). InsertAfterSelf and
        // NextSibling are O(1). So append right after the cached last body
        // paragraph (whose NextSibling is the trailing sectPr) and bump the
        // cached count. Cold cache or an out-of-band mutation (NextSibling no
        // longer the sectPr) falls back to the O(N) append + recount, which then
        // reseeds the cache for the next run.
        if (parent is Body fastBody && _bodyParaCount >= 0
            && _lastBodyParagraph is Paragraph anchor
            && ReferenceEquals(anchor.Parent, fastBody)
            && anchor.NextSibling() is SectionProperties)
        {
            anchor.InsertAfterSelf(para);
            _lastBodyParagraph = para;
            return ++_bodyParaCount;
        }
        AppendToParent(parent, para);
        if (parent is Body coldBody)
        {
            _bodyParaCount = coldBody.Elements<Paragraph>().Count();
            _lastBodyParagraph = para;
            return _bodyParaCount;
        }
        return parent.Elements<Paragraph>().Count();
    }

    private static List<PathSegment> ParsePath(string path)
    {
        var segments = new List<PathSegment>();
        // Reject leading double-slash up front — the subsequent Trim('/') would
        // otherwise eat the second slash and silently resolve "//body" → /body,
        // "//header[1]" → /header[1], producing inconsistent behavior next to
        // "//section[1]" which already errors out as Path-not-found via the regex
        // dispatch. The earlier-dispatch regexes anchor on `^/` so they don't
        // match `^//…` either; failures fall through here and we now reject.
        if (path.StartsWith("//"))
            throw new ArgumentException(
                $"Malformed path '{path}'. Path must start with exactly one '/'.");
        // Reject trailing slash up front — the subsequent Trim('/') would
        // otherwise silently absorb it and produce a path that looks valid
        // (e.g. "/body/p[1]/" → "body/p[1]") while any callers
        // concatenating onto the raw input would end up with doubled
        // separators like "/body/p[1]//r[2]" in the returned path.
        if (path.Length > 1 && path.EndsWith("/"))
            throw new ArgumentException(
                $"Malformed path '{path}'. Trailing '/' is not allowed.");
        var parts = path.Trim('/').Split('/');

        foreach (var part in parts)
        {
            // Reject degenerate empty segments from trailing/duplicate slashes
            // (e.g. "/body/p[1]/" or "/body//p[1]"). Without this, ParsePath
            // would silently swallow the empty part and return a garbled
            // navigable path.
            if (part.Length == 0)
                throw new ArgumentException(
                    $"Malformed path '{path}'. Empty path segment (check for trailing or duplicate '/').");

            var bracketIdx = part.IndexOf('[');
            if (bracketIdx >= 0)
            {
                // Only single-predicate form is supported. Reject malformed
                // selectors like "p[1][2]" or "p[1]trailing" where content
                // follows the first closing ']'. Without this the trailing
                // junk is silently swallowed (e.g. "p[1][2]" would resolve
                // to "p[1]") which hides typos.
                if (!part.EndsWith("]"))
                    throw new ArgumentException(
                        $"Malformed path segment '{part}'. Expected 'name[index]' or 'name[@attr=value]'.");
                var firstClose = part.IndexOf(']');
                if (firstClose != part.Length - 1)
                    throw new ArgumentException(
                        $"Malformed path segment '{part}'. Multiple predicates are not supported — use a single 'name[...]' form.");

                var name = Core.PathAliases.Resolve(part[..bracketIdx]);
                var indexStr = part[(bracketIdx + 1)..^1];
                // Reject empty predicate "p[]" which Int32.TryParse silently
                // rejects but which then falls through as a StringIndex of "".
                if (indexStr.Length == 0)
                    throw new ArgumentException(
                        $"Malformed path segment '{part}'. Empty predicate — expected 'name[index]' or 'name[@attr=value]'.");
                if (int.TryParse(indexStr, out var idx))
                {
                    if (idx <= 0)
                        throw new ArgumentException(
                            $"Malformed path segment '{part}'. Index predicate must be a positive integer (1-based), got '{indexStr}'.");
                    segments.Add(new PathSegment(name, idx));
                }
                else
                {
                    // Only accept a tightly specified set of string predicates:
                    //   last()
                    //   @attr=value   where attr is a simple identifier
                    //                 ([A-Za-z_][A-Za-z0-9_]*) and value is
                    //                 either bare-word (no whitespace, not
                    //                 starting with '@' or quote) or
                    //                 double-quoted.
                    // Anything else (e.g. "XYZ", " 1", "@=X", "@paraId",
                    //   "@w:paraId=X", "@attr='X'") is rejected up front so
                    //   typos cannot silently hit the FirstOrDefault()
                    //   fallback in NavigateToElement.
                    var normalizedPredicate = ValidateAndNormalizePredicate(part, indexStr);
                    segments.Add(new PathSegment(name, null, normalizedPredicate));
                }
            }
            else
            {
                segments.Add(new PathSegment(Core.PathAliases.Resolve(part), null));
            }
        }

        return segments;
    }

    /// <summary>
    /// Validate a string predicate (the content inside [...] that isn't an
    /// integer) and return its normalized form. Accepted grammar:
    ///   last()
    ///   @ident=value            (bare value: no whitespace, no quotes, no '@')
    ///   @ident="quoted value"   (double-quoted value)
    /// Everything else throws ArgumentException so typos like "p[XYZ]",
    /// "p[ 1]", "p[@paraId]" (no =), "p[@=X]", "p[@w:paraId=X]" are rejected
    /// instead of silently falling through to childList.FirstOrDefault().
    /// </summary>
    private static string ValidateAndNormalizePredicate(string part, string predicate)
    {
        if (predicate == "last()")
            return predicate;

        if (predicate.Length > 0 && predicate[0] == '@')
        {
            // Must have '=' and a non-empty identifier before it.
            var eq = predicate.IndexOf('=');
            if (eq <= 1)
                throw new ArgumentException(
                    $"Malformed path segment '{part}'. Attribute predicate must be '[@name=value]' with a non-empty attribute name.");

            var attr = predicate[1..eq];
            // Simple identifier: [A-Za-z_][A-Za-z0-9_]*
            if (!System.Text.RegularExpressions.Regex.IsMatch(attr, "^[A-Za-z_][A-Za-z0-9_]*$"))
                throw new ArgumentException(
                    $"Malformed path segment '{part}'. Attribute name '{attr}' is not a simple identifier (no prefixes/colons).");

            var value = predicate[(eq + 1)..];
            if (value.Length == 0)
                throw new ArgumentException(
                    $"Malformed path segment '{part}'. Attribute predicate value is empty.");

            // Accept double-quoted value — strip quotes so downstream
            // comparisons (which use bare string equality) work uniformly.
            if (value.Length >= 2 && value[0] == '"' && value[^1] == '"')
            {
                var inner = value[1..^1];
                if (inner.Contains('"'))
                    throw new ArgumentException(
                        $"Malformed path segment '{part}'. Quoted attribute value must not contain embedded double quotes.");
                return $"@{attr}={inner}";
            }

            // Bare value: no whitespace, no quotes, no leading '@'.
            if (value[0] == '@' || value[0] == '\'' || value[0] == '"')
                throw new ArgumentException(
                    $"Malformed path segment '{part}'. Attribute value must be bare-word or double-quoted.");
            foreach (var c in value)
            {
                if (char.IsWhiteSpace(c))
                    throw new ArgumentException(
                        $"Malformed path segment '{part}'. Attribute value must not contain whitespace (use double quotes).");
            }
            return predicate;
        }

        throw new ArgumentException(
            $"Malformed path segment '{part}'. Predicate must be a positive integer, 'last()', or '[@attr=value]'.");
    }

    // PERF: cache the flattened+filtered body paragraph/table lists per Body
    // instance. /body/p[N] and /body/tbl[N] are resolved by index; without
    // the cache, dumping a 14k-paragraph doc made 14k Get calls × 14k walks
    // → O(n²). Invalidation is by explicit clear: ClearBodyChildIndex() is
    // called on every structural add (Add() entry) and on Remove/Move/Swap/
    // CopyFrom/raw-set (via InvalidateBodyParaCache / RawSet). Property-only
    // Set calls do NOT clear (correct — they don't change which paragraph sits
    // at index N), so a batch of set /body/p[N] keeps hitting the cache.
    //
    // The guard must NOT recompute body.ChildElements.Count: the SDK stores
    // children in a singly-linked list, so .Count is O(n), which made even a
    // cache *hit* O(n) and turned batch set/get of /body/p[N] back into O(n²).
    private readonly Dictionary<OpenXmlElement, (List<OpenXmlElement> paras, List<OpenXmlElement> tables)>
        _bodyChildIndexCache = new();

    // Per-body-child → owning section, built lazily by FindOwningSectionProperties.
    private Dictionary<OpenXmlElement, SectionProperties?>? _owningSectionCache;

    // Per-body paraId → paragraph, built lazily by GetBodyParaById. Resolving
    // /body/p[@paraId=X] via FirstOrDefault scan is O(n); dump emits one
    // Get(/body/p[@paraId]) per paragraph, so that scan made dump O(n²).
    private Dictionary<OpenXmlElement, Dictionary<string, Paragraph>>? _bodyParaByIdCache;

    // Drop the body child-index + owning-section + paraId caches after a
    // structural mutation. Called from Add() (body-level) and InvalidateBodyParaCache.
    private void ClearBodyChildIndex()
    {
        _bodyChildIndexCache.Clear();
        _owningSectionCache = null;
        _bodyParaByIdCache = null;
    }

    // O(1) /body/p[@paraId=X] over body-direct (incl. customXml) paragraphs.
    // Returns null for ids that live only inside tables/sdt — the caller falls
    // back to a Descendants scan for those (rare).
    private Paragraph? GetBodyParaById(Body body, string paraId)
    {
        _bodyParaByIdCache ??= new();
        if (!_bodyParaByIdCache.TryGetValue(body, out var map))
        {
            map = new(StringComparer.OrdinalIgnoreCase);
            foreach (var p in GetBodyParagraphIndex(body).OfType<Paragraph>())
                if (p.ParagraphId?.Value is { } id && !map.ContainsKey(id))
                    map[id] = p; // first-wins, matches the old FirstOrDefault
            _bodyParaByIdCache[body] = map;
        }
        return map.TryGetValue(paraId, out var found) ? found : null;
    }

    private List<OpenXmlElement> GetBodyParagraphIndex(Body body) => GetBodyChildIndex(body).paras;
    private List<OpenXmlElement> GetBodyTableIndex(Body body) => GetBodyChildIndex(body).tables;

    private (List<OpenXmlElement> paras, List<OpenXmlElement> tables) GetBodyChildIndex(Body body)
    {
        if (_bodyChildIndexCache.TryGetValue(body, out var entry))
            return (entry.paras, entry.tables);

        var flat = new List<OpenXmlElement>();
        void Collect(OpenXmlElement el)
        {
            foreach (var c in el.ChildElements)
            {
                if (c is CustomXmlBlock cx) Collect(cx);
                else flat.Add(c);
            }
        }
        Collect(body);

        var paras = new List<OpenXmlElement>();
        var tables = new List<OpenXmlElement>();
        foreach (var e in flat)
        {
            if (e is Paragraph p && !IsOMathParaWrapperParagraph(p)) paras.Add(p);
            else if (e is Table t) tables.Add(t);
        }
        _bodyChildIndexCache[body] = (paras, tables);
        return (paras, tables);
    }

    /// <summary>
    /// Return the raw OuterXml of the element at <paramref name="path"/>, or
    /// null if the path does not resolve. Unlike <see cref="Raw"/> (which only
    /// addresses whole package parts), this navigates the document DOM. Used by
    /// the dump emitter to round-trip rich block containers — a Table-of-Contents
    /// SDT, say — verbatim instead of flattening them to text.
    /// </summary>
    internal string? RawElementXml(string path)
    {
        try
        {
            var segments = ParsePath(path);
            return NavigateToElement(segments)?.OuterXml;
        }
        catch
        {
            return null;
        }
    }

    private OpenXmlElement? NavigateToElement(List<PathSegment> segments)
        => NavigateToElement(segments, out _, out _);

    private OpenXmlElement? NavigateToElement(List<PathSegment> segments, out string? availableContext)
        => NavigateToElement(segments, out availableContext, out _);

    private OpenXmlElement? NavigateToElement(List<PathSegment> segments, out string? availableContext, out string resolvedPath)
    {
        resolvedPath = "";
        availableContext = null;
        if (segments.Count == 0) return null;

        var first = segments[0];

        // Handle bookmark[@name=...] as top-level path
        if (first.Name.ToLowerInvariant() == "bookmark" && first.StringIndex != null
            && first.StringIndex.StartsWith("@name=", StringComparison.OrdinalIgnoreCase))
        {
            var targetName = first.StringIndex["@name=".Length..];
            var body = _doc.MainDocumentPart?.Document?.Body;
            return body?.Descendants<BookmarkStart>()
                .FirstOrDefault(b => b.Name?.Value == targetName);
        }

        // Handle /bookmark[N] (1-based positional, document order). Skips
        // _GoBack and other reserved bookmarks (names starting with '_') so
        // the index matches what `query bookmark` returns.
        if (first.Name.ToLowerInvariant() == "bookmark" && segments.Count == 1
            && first.Index.HasValue)
        {
            var body = _doc.MainDocumentPart?.Document?.Body;
            if (body != null)
            {
                var bks = body.Descendants<BookmarkStart>()
                    .Where(b => !(b.Name?.Value ?? "").StartsWith("_", StringComparison.Ordinal))
                    .ToList();
                var n = first.Index.Value;
                if (n >= 1 && n <= bks.Count) return bks[n - 1];
            }
        }

        // BUG-R36-B5: top-level /sdt[N] alias. The schema documents both
        // /sdt[N] and /body/p[N]/sdt[M], but only the body-anchored form
        // resolved. Resolve /sdt[N] positionally over body-level SdtBlock
        // elements (document order), mirroring the /bookmark[N] alias above.
        if (first.Name.ToLowerInvariant() == "sdt" && segments.Count == 1
            && first.Index.HasValue)
        {
            var body = _doc.MainDocumentPart?.Document?.Body;
            if (body != null)
            {
                var sdts = body.Descendants<SdtBlock>().Cast<OpenXmlElement>()
                    .Concat(body.Descendants<SdtRun>().Cast<OpenXmlElement>())
                    .ToList();
                var n = first.Index.Value;
                if (n >= 1 && n <= sdts.Count) return sdts[n - 1];
            }
        }
        if (first.Name.ToLowerInvariant() == "sdt" && segments.Count == 1
            && first.StringIndex != null
            && first.StringIndex.StartsWith("@sdtId=", StringComparison.OrdinalIgnoreCase))
        {
            var body = _doc.MainDocumentPart?.Document?.Body;
            if (body != null
                && int.TryParse(first.StringIndex["@sdtId=".Length..], out var targetId))
            {
                return body.Descendants<SdtBlock>().Cast<OpenXmlElement>()
                    .Concat(body.Descendants<SdtRun>().Cast<OpenXmlElement>())
                    .FirstOrDefault(s =>
                        (s as SdtBlock)?.SdtProperties?.GetFirstChild<SdtId>()?.Val?.Value == targetId
                        || (s as SdtRun)?.SdtProperties?.GetFirstChild<SdtId>()?.Val?.Value == targetId);
            }
        }

        // Top-level /section[N] anchor routing. `add --type section` returns
        // "/section[N]" as the new element's identity; resolving it to the
        // carrier paragraph (the one whose pPr holds the Nth sectPr) lets
        // callers use it directly as --after/--before. Body-level sectPr
        // (the final section) is intentionally NOT an anchor target here —
        // it must remain the last child of body; anchor use is rejected in
        // ResolveAnchorPosition.
        if (first.Name.ToLowerInvariant() == "section" && segments.Count == 1 && first.Index.HasValue)
        {
            var body = _doc.MainDocumentPart?.Document?.Body;
            if (body != null)
            {
                var n = first.Index.Value;
                var sectParas = body.Elements<Paragraph>()
                    .Where(p => p.ParagraphProperties?.GetFirstChild<SectionProperties>() != null)
                    .ToList();
                if (n >= 1 && n <= sectParas.Count)
                    return sectParas[n - 1];
            }
        }

        // Top-level /chart[N] anchor routing. `add --type chart` returns
        // "/chart[N]" as the new element's identity; resolve it to the
        // body-level paragraph containing the Nth chart drawing so callers
        // can use the returned path directly as --after/--before.
        if (first.Name.ToLowerInvariant() == "chart" && segments.Count == 1 && first.Index.HasValue)
        {
            var charts = GetAllWordCharts();
            var n = first.Index.Value;
            if (n >= 1 && n <= charts.Count)
            {
                OpenXmlElement? cur = charts[n - 1].Inline;
                while (cur != null && cur is not Paragraph) cur = cur.Parent;
                if (cur is Paragraph chartPara) return chartPara;
            }
        }

        // Top-level /toc[N] anchor routing. `add --type toc` returns
        // "/toc[N]" as the new element's identity; resolve it to the Nth
        // body paragraph whose descendants include a FieldCode starting
        // with "TOC" (mirrors AddToc's counting logic) so callers can use
        // the returned path directly as --after/--before.
        if (first.Name.ToLowerInvariant() == "toc" && segments.Count == 1 && first.Index.HasValue)
        {
            var body = _doc.MainDocumentPart?.Document?.Body;
            if (body != null)
            {
                var tocParas = body.Elements<Paragraph>()
                    .Where(p => p.Descendants<FieldCode>().Any(fc =>
                        fc.Text != null && fc.Text.TrimStart().StartsWith("TOC", StringComparison.OrdinalIgnoreCase)))
                    .ToList();
                var n = first.Index.Value;
                if (n >= 1 && n <= tocParas.Count)
                    return tocParas[n - 1];
            }
        }

        // Top-level /formfield[N] anchor routing. `add --type formfield`
        // returns "/formfield[N]" as the new element's identity; resolve it to
        // the body-level paragraph containing the Nth form field's begin-run
        // so callers can use the returned path directly as --after/--before.
        // R14-bug4: also accept /formfield[@name=NAME] (the schema-documented
        // stable form, signalled by first.StringIndex starting with "@name=").
        if (first.Name.ToLowerInvariant() == "formfield" && segments.Count == 1)
        {
            var allFf = FindFormFields();
            (FieldInfo Field, FormFieldData FfData) hit = default;
            if (first.Index.HasValue)
            {
                var n = first.Index.Value;
                if (n >= 1 && n <= allFf.Count) hit = allFf[n - 1];
            }
            else if (first.StringIndex != null
                && first.StringIndex.StartsWith("@name=", StringComparison.OrdinalIgnoreCase))
            {
                var target = first.StringIndex["@name=".Length..];
                hit = allFf.FirstOrDefault(ff =>
                    ff.FfData.GetFirstChild<FormFieldName>()?.Val?.Value == target);
            }
            if (hit.Field != null)
            {
                var beginRun = hit.Field.BeginRun;
                OpenXmlElement? cur = beginRun;
                while (cur != null && cur is not Paragraph) cur = cur.Parent;
                return cur ?? beginRun;
            }
        }

        OpenXmlElement? current = first.Name.ToLowerInvariant() switch
        {
            "body" => _doc.MainDocumentPart?.Document?.Body,
            "styles" => _doc.MainDocumentPart?.StyleDefinitionsPart?.Styles,
            "header" => _doc.MainDocumentPart?.HeaderParts.ElementAtOrDefault((first.Index ?? 1) - 1)?.Header,
            "footer" => _doc.MainDocumentPart?.FooterParts.ElementAtOrDefault((first.Index ?? 1) - 1)?.Footer,
            "numbering" => _doc.MainDocumentPart?.NumberingDefinitionsPart?.Numbering,
            "settings" => _doc.MainDocumentPart?.DocumentSettingsPart?.Settings,
            "comments" => _doc.MainDocumentPart?.WordprocessingCommentsPart?.Comments,
            // /footnotes and /endnotes are container aliases so that
            // /footnotes/footnote[N] and /endnotes/endnote[N] work as
            // documented in the help text. The Nth user note is also
            // selectable directly via /footnote[N] (positional) or
            // /footnote[@footnoteId=N] (id-based) — those paths bypass
            // this switch via the `current == null` block below.
            "footnotes" => _doc.MainDocumentPart?.FootnotesPart?.Footnotes,
            "endnotes" => _doc.MainDocumentPart?.EndnotesPart?.Endnotes,
            _ => null
        };

        string parentPath = "/" + first.Name + (first.Index.HasValue ? $"[{first.Index}]" : "");

        // Top-level /footnote[@footnoteId=N] / /footnote[N] routing. Mirrors
        // WordHandler.Add.cs's TryResolveFootnoteOrEndnoteBody so that paths
        // returned by `add` under a footnote/endnote are round-trippable via
        // `get` and usable as --after/--before anchors.
        if (current == null)
        {
            var fname = first.Name.ToLowerInvariant();
            if (fname == "footnote")
            {
                int? fnId = first.Index;
                if (fnId == null && first.StringIndex != null
                    && first.StringIndex.StartsWith("@footnoteId=", StringComparison.OrdinalIgnoreCase)
                    && int.TryParse(first.StringIndex["@footnoteId=".Length..], out var idv))
                {
                    fnId = idv;
                }
                if (fnId != null)
                {
                    current = _doc.MainDocumentPart?.FootnotesPart?.Footnotes?
                        .Elements<Footnote>().FirstOrDefault(f => f.Id?.Value == fnId.Value);
                    parentPath = $"/footnote[@footnoteId={fnId}]";
                }
            }
            else if (fname == "endnote")
            {
                int? enId = first.Index;
                if (enId == null && first.StringIndex != null
                    && first.StringIndex.StartsWith("@endnoteId=", StringComparison.OrdinalIgnoreCase)
                    && int.TryParse(first.StringIndex["@endnoteId=".Length..], out var idv))
                {
                    enId = idv;
                }
                if (enId != null)
                {
                    current = _doc.MainDocumentPart?.EndnotesPart?.Endnotes?
                        .Elements<Endnote>().FirstOrDefault(e => e.Id?.Value == enId.Value);
                    parentPath = $"/endnote[@endnoteId={enId}]";
                }
            }
        }

        for (int i = 1; i < segments.Count && current != null; i++)
        {
            var seg = segments[i];
            IEnumerable<OpenXmlElement> children;
            // When the current element is a block-level SDT, transparently
            // descend into its SdtContentBlock so paths like
            // /body/sdt[@sdtId=X]/p[N] resolve to paragraphs physically
            // nested inside the content wrapper. Mirrors GetBodyElements()
            // which already flattens SdtBlock when iterating body children.
            if (current is SdtBlock navSdtBlock)
            {
                var contentBlock = navSdtBlock.GetFirstChild<SdtContentBlock>();
                if (contentBlock != null) current = contentBlock;
            }
            else if (current is SdtRun navSdtRun)
            {
                var contentRun = navSdtRun.GetFirstChild<SdtContentRun>();
                if (contentRun != null) current = contentRun;
            }

            // Allow an explicit "/sdtContent" segment as a no-op selector: after
            // the transparent descend above, `current` is already the
            // SdtContent{Block,Run}. This keeps the ValidateParentChild hint
            // ("Add under <sdt>/sdtContent instead") literally navigable.
            if (seg.Name.Equals("sdtContent", StringComparison.OrdinalIgnoreCase)
                && (current is SdtContentBlock || current is SdtContentRun))
            {
                parentPath += "/sdtContent";
                continue;
            }

            // O(1) fast-path for /body/p[last()] — batch run-add resolves this
            // once per run (100k+ times in a large replay). Trusted when the
            // cached last paragraph is still parented to this body and its
            // NextSibling is the trailing sectPr, i.e. it is genuinely last
            // (NextSibling is O(1) on the SDK's singly-linked list, unlike a
            // full ToList()+LastOrDefault scan). Emit the same "/p[N]" segment
            // the scan below would, so resolvedPath is identical. A stale hit
            // falls through to the scan.
            if (seg.StringIndex == "last()" && seg.Name == "p" && current is Body
                && _bodyParaCount >= 0
                && _lastBodyParagraph is Paragraph lbpFast
                && ReferenceEquals(lbpFast.Parent, current)
                && lbpFast.NextSibling() is SectionProperties)
            {
                parentPath += $"/p[{_bodyParaCount}]";
                current = lbpFast;
                continue;
            }

            if (current is Body body2 && (seg.Name.ToLowerInvariant() == "p" || seg.Name.ToLowerInvariant() == "tbl"))
            {
                // Only count direct body-level paragraphs/tables, skip those inside SdtBlock containers.
                // #6: paragraphs whose sole content is m:oMathPara are
                // counted via the /body/oMathPara[N] path instead, so the
                // /body/p[N] enumeration skips them to match HTML-preview
                // data-path attribution (which also skips them).
                // BUG-DUMP8-01/02: w:customXml body wrappers are non-structural —
                // recursively flatten so paragraphs/tables nested inside one
                // (or several) levels of CustomXmlBlock surface in the same
                // /body/p[N] / /body/tbl[N] enumeration. Mirrors the listing
                // logic in WalkBodyChild for `get /body`; without this, path
                // resolution diverged from listing and `get /body/p[1]` threw
                // "Path not found" on customXml-wrapped paragraphs.
                // PERF: cache the filtered lists per Body instance + child count.
                // Without the cache, dumping a doc with N body paragraphs costs
                // O(N²) because the dump emitter calls Get("/body/p[K]") for
                // every K in 1..N, and each call re-walked body.ChildElements.
                // Real-world 14k-paragraph doc: 5+ minutes → seconds.
                children = seg.Name.ToLowerInvariant() == "p"
                    ? GetBodyParagraphIndex(body2)
                    : GetBodyTableIndex(body2);
            }
            else if (current is Body body3 && seg.Name == "oMathPara")
            {
                // oMathPara can be direct body children or wrapped inside w:p elements
                var mathParas = new List<OpenXmlElement>();
                foreach (var el in body3.ChildElements)
                {
                    if (el.LocalName == "oMathPara" || el is M.Paragraph)
                        mathParas.Add(el);
                    else if (el is Paragraph wp && IsOMathParaWrapperParagraph(wp))
                    {
                        // Only pure-wrapper paragraphs (pPr + single oMathPara child)
                        // — otherwise /body/p[N] and /body/oMathPara[M] would both
                        // address the same paragraph (mixed prose + inline math),
                        // causing Get/Set/Remove to diverge by callsite.
                        var inner = wp.ChildElements.FirstOrDefault(c => c.LocalName == "oMathPara" || c is M.Paragraph);
                        if (inner != null) mathParas.Add(inner);
                    }
                }
                children = mathParas;
            }
            else
            {
                children = seg.Name.ToLowerInvariant() switch
                {
                    "p" => current.Elements<Paragraph>().Cast<OpenXmlElement>(),
                    // BUG-D1-TXBX-RUN-INDEX: /<para>/r[N] must enumerate the
                    // SAME runs NodeBuilder surfaces as paragraph children —
                    // i.e. skip runs that live INSIDE a child textbox
                    // (their canonical path is /<para>/textbox[N]/p[M]/r[K]).
                    // Without this filter, a paragraph hosting 3 textbox
                    // drawings (Card A / B / C side-by-side) resolves r[2] to
                    // a text run INSIDE Card A's textbox content instead of
                    // Card B's drawing run, so WordBatchEmitter probes the
                    // wrong XML, IsTextboxDrawing returns false, and Cards
                    // B/C silently degrade to plain `add r` rows on dump
                    // round-trip — losing the textboxes entirely and
                    // misaligning every subsequent /body/textbox[K] index.
                    // Mirrors GetAllRuns in WordHandler.Helpers (also skips
                    // SimpleField/SdtRun-nested runs for the same path-
                    // stability reason).
                    "r" => current.Descendants<Run>()
                        .Where(r => r.GetFirstChild<CommentReference>() == null)
                        .Where(r => r.Ancestors<SdtRun>().FirstOrDefault() == null)
                        .Where(r => r.Ancestors<SimpleField>().FirstOrDefault() == null)
                        .Where(r =>
                        {
                            var tbc = r.Ancestors<TextBoxContent>().FirstOrDefault();
                            if (tbc == null) return true;
                            foreach (var anc in tbc.Ancestors())
                            {
                                if (ReferenceEquals(anc, current)) return false;
                            }
                            return true;
                        })
                        .Cast<OpenXmlElement>(),
                    "tbl" => current.Elements<Table>().Cast<OpenXmlElement>(),
                    "tr" => current.Elements<TableRow>().Cast<OpenXmlElement>(),
                    "tc" => current.Elements<TableCell>().Cast<OpenXmlElement>(),
                    "sdt" => current.ChildElements
                        .Where(e => e is SdtBlock || e is SdtRun).Cast<OpenXmlElement>(),
                    // v5.7-cont: /body/textbox[N] → walk descendant drawings,
                    // pick the Nth wps:txbx host, return its w:txbxContent
                    // so child p[M] resolves naturally via the next loop iter.
                    "textbox" => current.Descendants<Drawing>()
                        .Where(d => d.InnerXml.Contains("<wps:txbx") || d.InnerXml.Contains("txBox=\"1\""))
                        .Select(d => (OpenXmlElement?)d.Descendants().FirstOrDefault(e =>
                            e.LocalName == "txbxContent"
                            && e.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main"))
                        .Where(e => e != null)
                        .Cast<OpenXmlElement>(),
                    // v5.7-cont: /body/shape[N] → walk descendant drawings,
                    // pick the Nth wps:wsp that isn't itself a textbox. Returns
                    // the wps:wsp element so children resolve in its scope.
                    "shape" => current.Descendants<Drawing>()
                        .Where(d => d.InnerXml.Contains("<wps:wsp")
                                 && !d.InnerXml.Contains("<wps:txbx")
                                 && !d.InnerXml.Contains("txBox=\"1\""))
                        .Select(d => (OpenXmlElement?)d.Descendants().FirstOrDefault(e =>
                            e.LocalName == "wsp"
                            && e.NamespaceUri == "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"))
                        .Where(e => e != null)
                        .Cast<OpenXmlElement>(),
                    // /<para>/tab[N] and /styles/<id>/tab[N] descend
                    // transparently through pPr/tabs (or StyleParagraph-
                    // Properties/tabs) so the user-facing path stays flat
                    // instead of leaking the OOXML containers (.../pPr/tabs/tab).
                    // Symmetric with how AddTab returns the flat form.
                    "tab" when current is Paragraph navParaT
                        => navParaT.ParagraphProperties?.GetFirstChild<Tabs>()?.Elements<TabStop>().Cast<OpenXmlElement>()
                           ?? Enumerable.Empty<OpenXmlElement>(),
                    "tab" when current is Style navStyleT
                        => navStyleT.StyleParagraphProperties?.GetFirstChild<Tabs>()?.Elements<TabStop>().Cast<OpenXmlElement>()
                           ?? Enumerable.Empty<OpenXmlElement>(),
                    // /styles/<key> resolves <key> as a styleId or styleName
                    // (matches Set.Dispatch.cs's regex+OR matching), so paths
                    // like /styles/Heading1 are navigable for Add/Get/Set.
                    // The segment name here IS the key, not an OOXML local-
                    // name; downstream FirstOrDefault picks the (single) match.
                    _ when current is Styles navStylesContainer
                        => navStylesContainer.Elements<Style>().Where(s =>
                            string.Equals(s.StyleId?.Value, seg.Name, StringComparison.Ordinal)
                            || string.Equals(s.StyleName?.Val?.Value, seg.Name, StringComparison.Ordinal))
                           .Cast<OpenXmlElement>(),
                    // CONSISTENCY(footnotes-container): /footnotes/footnote[N]
                    // enumerates user footnotes only (id > 0), matching what
                    // `query footnote` returns and the positional /footnote[N]
                    // routing used by Add. The schema's separator/continuation
                    // entries (id=-1, id=0) are excluded so positional indexes
                    // line up across paths.
                    "footnote" when current is Footnotes fns
                        => fns.Elements<Footnote>().Where(f => f.Id?.Value > 0).Cast<OpenXmlElement>(),
                    "endnote" when current is Endnotes ens
                        => ens.Elements<Endnote>().Where(e => e.Id?.Value > 0).Cast<OpenXmlElement>(),
                    _ => current.ChildElements.Where(e => e.LocalName == seg.Name).Cast<OpenXmlElement>()
                };
            }

            // Reuse the list when the switch already produced one (e.g.
            // GetBodyParagraphIndex/GetBodyTableIndex return a cached
            // List<OpenXmlElement>). ToList() would copy it O(n) on EVERY
            // navigation, so resolving /body/p[N] for N=1..n — batch set/get of
            // paragraphs — was O(n²). The list is only read here (ElementAt /
            // LastOrDefault / IndexOf), so sharing the cached reference is safe.
            var childList = children as List<OpenXmlElement> ?? children.ToList();
            OpenXmlElement? next;
            if (seg.Index.HasValue)
                next = childList.ElementAtOrDefault(seg.Index.Value - 1);
            else if (seg.StringIndex == "last()")
                next = childList.LastOrDefault();
            else if (seg.StringIndex != null && seg.StringIndex.StartsWith("@paraId=", StringComparison.OrdinalIgnoreCase))
            {
                var targetId = seg.StringIndex["@paraId=".Length..];
                // CONSISTENCY(paraid-global-uniqueness): paraId is globally
                // unique across body/headers/footers/footnotes/endnotes/
                // comments (EnsureAllParaIds scans every part). Resolve by
                // descendants too — direct-child-only scan made cell paras
                // unreachable from the canonical /body/p[@paraId=...] form
                // that AddPtab/AddBreak/AddField return for cell parents.
                // Body case uses the O(1) paraId map (dump resolves one
                // /body/p[@paraId] per paragraph — the scan was O(n²)).
                next = current is Body navParaByIdBody
                    ? GetBodyParaById(navParaByIdBody, targetId)
                    : childList.OfType<Paragraph>()
                        .FirstOrDefault(p => string.Equals(p.ParagraphId?.Value, targetId, StringComparison.OrdinalIgnoreCase));
                if (next == null)
                {
                    next = (current as OpenXmlElement)?.Descendants<Paragraph>()
                        .FirstOrDefault(p => string.Equals(p.ParagraphId?.Value, targetId, StringComparison.OrdinalIgnoreCase));
                }
            }
            else if (seg.StringIndex != null && seg.StringIndex.StartsWith("@textId=", StringComparison.OrdinalIgnoreCase))
            {
                var targetId = seg.StringIndex["@textId=".Length..];
                next = childList.OfType<Paragraph>()
                    .FirstOrDefault(p => string.Equals(p.TextId?.Value, targetId, StringComparison.OrdinalIgnoreCase));
                if (next == null)
                {
                    next = (current as OpenXmlElement)?.Descendants<Paragraph>()
                        .FirstOrDefault(p => string.Equals(p.TextId?.Value, targetId, StringComparison.OrdinalIgnoreCase));
                }
            }
            else if (seg.StringIndex != null && seg.StringIndex.StartsWith("@commentId=", StringComparison.OrdinalIgnoreCase))
            {
                var targetId = seg.StringIndex["@commentId=".Length..];
                next = childList.OfType<Comment>()
                    .FirstOrDefault(c => c.Id?.Value == targetId);
            }
            else if (seg.StringIndex != null && seg.StringIndex.StartsWith("@name=", StringComparison.OrdinalIgnoreCase))
            {
                // Generic @name=... selector, used by bookmarkStart[@name=X]
                // so that the path returned by AddBookmark is navigable.
                var targetName = seg.StringIndex["@name=".Length..];
                next = childList.FirstOrDefault(e =>
                    e is BookmarkStart bs && string.Equals(bs.Name?.Value, targetName, StringComparison.Ordinal));
            }
            else if (seg.StringIndex != null && seg.StringIndex.StartsWith("@sdtId=", StringComparison.OrdinalIgnoreCase))
            {
                var targetId = seg.StringIndex["@sdtId=".Length..];
                next = childList.Where(e => e is SdtBlock or SdtRun)
                    .FirstOrDefault(e =>
                    {
                        var sdtId = (e is SdtBlock sb ? sb.SdtProperties : (e as SdtRun)?.SdtProperties)
                            ?.GetFirstChild<SdtId>()?.Val?.Value;
                        return sdtId?.ToString() == targetId;
                    });
            }
            // CONSISTENCY(id-selectors): mirror @paraId/@commentId/@sdtId — accept @id= for
            // numbering/abstractNum (w:abstractNumId@val) and numbering/num (w:num@numId).
            else if (seg.StringIndex != null && seg.StringIndex.StartsWith("@id=", StringComparison.OrdinalIgnoreCase))
            {
                var targetId = seg.StringIndex["@id=".Length..];
                next = childList.FirstOrDefault(e => e switch
                {
                    AbstractNum an => an.AbstractNumberId?.Value.ToString() == targetId,
                    NumberingInstance ni => ni.NumberID?.Value.ToString() == targetId,
                    _ => false,
                });
            }
            else if (seg.StringIndex != null && seg.StringIndex.StartsWith("@", StringComparison.Ordinal))
            {
                // Unrecognized attribute predicate — throw rather than silently returning
                // the first element. ValidateAndNormalizePredicate accepts any @ident=value
                // syntactically, but not every attribute maps to a Word OOXML concept.
                // Comment on the gap: expand the dispatch chain above when a new attribute
                // needs to be addressable (e.g. @bookmarkId=, @w14:paraId=).
                var eq = seg.StringIndex.IndexOf('=');
                var attrName = eq > 0 ? seg.StringIndex[1..eq] : seg.StringIndex[1..];
                throw new ArgumentException(
                    $"Attribute predicate '@{attrName}' is not a recognized Word path attribute. " +
                    $"Supported attributes: @paraId, @textId, @commentId, @sdtId, @id, @name.");
            }
            else
                next = childList.FirstOrDefault();

            if (next == null)
            {
                availableContext = BuildAvailableContext(current, parentPath, seg.Name, childList.Count);
                return null;
            }

            // Build path segment: prefer stable ID when available, fallback to positional.
            // Use the resolved element's LocalName (always canonical lowercase for OOXML)
            // rather than seg.Name (which echoes user capitalization like 'P'), so the
            // returned path round-trips cleanly and matches Query's canonical form.
            // Style is exempt — /styles/<id> uses the user-supplied styleId/Name as the key.
            var canonName = (next is Style) ? seg.Name : next.LocalName;
            // BUG-D1-TXBX-PATH: the textbox / shape segments resolve to
            // wps:wsp + w:txbxContent OOXML elements whose LocalName is
            // "wsp" / "txbxContent" — neither is navigable as a path
            // segment (the navigable name is the user-facing "textbox" /
            // "shape" form handled at the children-selector switch above).
            // Preserve the user-supplied segment name so resolvedPath stays
            // re-navigable; without this, descendant Get calls on children
            // like /<host>/textbox[N]/tbl[K]/tr[J] fail with
            // "No txbxContent found at /body".
            if (canonName == "txbxContent" || canonName == "wsp")
                canonName = seg.Name.ToLowerInvariant();
            if (next is Paragraph navPara && !string.IsNullOrEmpty(navPara.ParagraphId?.Value))
            {
                parentPath += "/" + canonName + $"[@paraId={navPara.ParagraphId.Value}]";
            }
            else if (next is Comment navComment && navComment.Id?.Value != null)
            {
                parentPath += "/" + canonName + $"[@commentId={navComment.Id.Value}]";
            }
            else if (next is Style navStyle)
            {
                // Style is keyed by styleId — emit /styles/<id> without a
                // positional [N] suffix to match Query's canonical form.
                parentPath += "/" + (navStyle.StyleId?.Value ?? seg.Name);
            }
            else if (next is SdtBlock or SdtRun)
            {
                var sdtProps = (next is SdtBlock sb2 ? sb2.SdtProperties : (next as SdtRun)?.SdtProperties);
                var sdtIdVal = sdtProps?.GetFirstChild<SdtId>()?.Val?.Value;
                if (sdtIdVal != null)
                    parentPath += "/" + canonName + $"[@sdtId={sdtIdVal}]";
                else
                {
                    var posIdx = childList.IndexOf(next) + 1;
                    parentPath += "/" + canonName + $"[{posIdx}]";
                }
            }
            else
            {
                var posIdx = childList.IndexOf(next) + 1;
                parentPath += "/" + canonName + $"[{posIdx}]";
            }
            current = next;
        }

        resolvedPath = parentPath;
        return current;
    }

    /// <summary>
    /// Build a context string describing available children when navigation fails.
    /// </summary>
    private static string BuildAvailableContext(OpenXmlElement parent, string parentPath, string requestedType, int matchCount)
    {
        if (matchCount > 0)
            return $"Available at {parentPath}: {requestedType}[1]..{requestedType}[{matchCount}]";

        // List distinct child types at this level
        var childTypes = parent.ChildElements
            .GroupBy(c => c.LocalName)
            .Select(g => $"{g.Key}({g.Count()})")
            .Take(10)
            .ToList();

        return childTypes.Count > 0
            ? $"No {requestedType} found at {parentPath}. Available children: {string.Join(", ", childTypes)}"
            : $"No children at {parentPath}";
    }

    private DocumentNode BookmarkStartToNode(BookmarkStart bkStart, DocumentNode node)
    {
        node.Type = "bookmark";
        node.Format["name"] = bkStart.Name?.Value ?? "";
        node.Format["id"] = bkStart.Id?.Value ?? "";
        // BUG-DUMP10-04: for cross-paragraph bookmark spans, walk
        // forward over sibling paragraphs in the same body and
        // surface the BookmarkEnd's paragraph offset (0-based).
        // 0 = same paragraph (default; AddBookmark places End next to
        // Start). >0 = the End sits N paragraphs after the Start.
        // Without this, dump emitted only the BookmarkStart and
        // AddBookmark always re-emitted the End in the same paragraph,
        // collapsing every multi-paragraph bookmark on round-trip.
        var bkStartId = bkStart.Id?.Value;
        if (!string.IsNullOrEmpty(bkStartId)
            && bkStart.Ancestors<Paragraph>().FirstOrDefault() is { } startPara
            && startPara.Parent is OpenXmlElement bodyParent)
        {
            var siblings = bodyParent.Elements<Paragraph>().ToList();
            int startIdx = siblings.IndexOf(startPara);
            if (startIdx >= 0)
            {
                for (int i = startIdx; i < siblings.Count; i++)
                {
                    var endHere = siblings[i].Descendants<BookmarkEnd>()
                        .FirstOrDefault(be => be.Id?.Value == bkStartId);
                    if (endHere != null)
                    {
                        int offset = i - startIdx;
                        if (offset > 0) node.Format["endPara"] = offset;
                        break;
                    }
                }
            }
        }
        var bkText = GetBookmarkText(bkStart);
        if (!string.IsNullOrEmpty(bkText))
            node.Text = bkText;
        return node;
    }

    private DocumentNode FootnoteToNode(Footnote fnEl, DocumentNode node, string path, int depth)
    {
        node.Type = "footnote";
        // Strip the reference-mark leading space (CONSISTENCY with Query
        // get-by-id and `query footnote`). Without this branch the
        // generic InnerText fallback below would return " fn-text".
        node.Text = GetFootnoteText(fnEl);
        if (fnEl.Id?.Value != null) node.Format["id"] = fnEl.Id.Value;
        if (fnEl.Type?.Value != null) node.Format["type"] = fnEl.Type.InnerText;
        // R44 minor-5: surface first-run formatting on footnote node so
        // bold/italic/size/color set via Add/Set round-trip through Get.
        // Mirrors the hyperlink firstRun pattern at line ~2746 above.
        var fnFirstRun = fnEl.Descendants<Run>().FirstOrDefault(r => r.GetFirstChild<Text>() != null);
        if (fnFirstRun?.RunProperties != null)
        {
            var rp = fnFirstRun.RunProperties;
            if (rp.RunFonts?.Ascii?.Value != null) node.Format["font"] = rp.RunFonts.Ascii.Value;
            if (rp.FontSize?.Val?.Value != null)
                node.Format["size"] = $"{int.Parse(rp.FontSize.Val.Value) / 2.0:0.##}pt";
            if (rp.Bold != null) node.Format["bold"] = IsToggleOn(rp.Bold);
            if (rp.Italic != null) node.Format["italic"] = IsToggleOn(rp.Italic);
            if (rp.Color?.ThemeColor?.HasValue == true) node.Format["color"] = rp.Color.ThemeColor.InnerText;
            else if (rp.Color?.Val?.Value != null) node.Format["color"] = ParseHelpers.FormatHexColor(rp.Color.Val.Value);
            if (rp.Underline?.Val != null) node.Format["underline"] = rp.Underline.Val.InnerText;
            if (rp.Strike != null) node.Format["strike"] = IsToggleOn(rp.Strike);
            if (rp.Highlight?.Val != null) node.Format["highlight"] = rp.Highlight.Val.InnerText;
        }
        // R20-wbt-1: surface direction from the first content paragraph's
        // pPr.BiDi so the cascade (already applied by ApplyFootnoteEndnoteFormatKeys)
        // round-trips through Get. Mirrors the paragraph readback below.
        var fnBidi = fnEl.Descendants<Paragraph>().FirstOrDefault()?.ParagraphProperties?.GetFirstChild<BiDi>();
        if (fnBidi != null)
            node.Format["direction"] = TryReadOnOff(fnBidi.Val) == true ? "rtl" : "ltr";
        // BUG-DUMP8-05/06: Paragraph branch surfaces inline w:sym (as
        // sym= run children) and m:oMath (as equation children) but the
        // Footnote branch returned early after flat text/format, so
        // sym and oMath inside footnote bodies were silently dropped.
        // Walk descendant runs/equations and surface them as children
        // on the footnote node, mirroring the paragraph walker's keys.
        if (depth > 0)
        {
            int fnSymIdx = 0;
            foreach (var symRun in fnEl.Descendants<Run>())
            {
                var symEl = symRun.GetFirstChild<SymbolChar>();
                if (symEl?.Char?.Value == null) continue;
                var symFontVal = symEl.Font?.Value ?? "";
                var symNode = new DocumentNode
                {
                    Type = "run",
                    Path = $"{path}/r[{fnSymIdx + 1}]",
                };
                symNode.Format["sym"] = $"{symFontVal}:{symEl.Char.Value}";
                node.Children.Add(symNode);
                fnSymIdx++;
            }
            int fnEqIdx = 0;
            foreach (var fnEq in fnEl.Descendants<M.OfficeMath>())
            {
                node.Children.Add(ElementToNode(fnEq, $"{path}/equation[{fnEqIdx + 1}]", depth - 1));
                fnEqIdx++;
            }
        }
        return node;
    }

    private DocumentNode EndnoteToNode(Endnote enEl, DocumentNode node, string path, int depth)
    {
        node.Type = "endnote";
        node.Text = GetFootnoteText(enEl);
        if (enEl.Id?.Value != null) node.Format["id"] = enEl.Id.Value;
        if (enEl.Type?.Value != null) node.Format["type"] = enEl.Type.InnerText;
        // R44 minor-5: mirror footnote firstRun readback for endnote.
        var enFirstRun = enEl.Descendants<Run>().FirstOrDefault(r => r.GetFirstChild<Text>() != null);
        if (enFirstRun?.RunProperties != null)
        {
            var rp = enFirstRun.RunProperties;
            if (rp.RunFonts?.Ascii?.Value != null) node.Format["font"] = rp.RunFonts.Ascii.Value;
            if (rp.FontSize?.Val?.Value != null)
                node.Format["size"] = $"{int.Parse(rp.FontSize.Val.Value) / 2.0:0.##}pt";
            if (rp.Bold != null) node.Format["bold"] = IsToggleOn(rp.Bold);
            if (rp.Italic != null) node.Format["italic"] = IsToggleOn(rp.Italic);
            if (rp.Color?.ThemeColor?.HasValue == true) node.Format["color"] = rp.Color.ThemeColor.InnerText;
            else if (rp.Color?.Val?.Value != null) node.Format["color"] = ParseHelpers.FormatHexColor(rp.Color.Val.Value);
            if (rp.Underline?.Val != null) node.Format["underline"] = rp.Underline.Val.InnerText;
            if (rp.Strike != null) node.Format["strike"] = IsToggleOn(rp.Strike);
            if (rp.Highlight?.Val != null) node.Format["highlight"] = rp.Highlight.Val.InnerText;
        }
        var enBidi = enEl.Descendants<Paragraph>().FirstOrDefault()?.ParagraphProperties?.GetFirstChild<BiDi>();
        if (enBidi != null)
            node.Format["direction"] = TryReadOnOff(enBidi.Val) == true ? "rtl" : "ltr";
        // CONSISTENCY with Footnote: surface inline w:sym / m:oMath
        // descendants so dump round-trips them through batch.
        if (depth > 0)
        {
            int enSymIdx = 0;
            foreach (var symRun in enEl.Descendants<Run>())
            {
                var symEl = symRun.GetFirstChild<SymbolChar>();
                if (symEl?.Char?.Value == null) continue;
                var symFontVal = symEl.Font?.Value ?? "";
                var symNode = new DocumentNode
                {
                    Type = "run",
                    Path = $"{path}/r[{enSymIdx + 1}]",
                };
                symNode.Format["sym"] = $"{symFontVal}:{symEl.Char.Value}";
                node.Children.Add(symNode);
                enSymIdx++;
            }
            int enEqIdx = 0;
            foreach (var enEq in enEl.Descendants<M.OfficeMath>())
            {
                node.Children.Add(ElementToNode(enEq, $"{path}/equation[{enEqIdx + 1}]", depth - 1));
                enEqIdx++;
            }
        }
        return node;
    }

    private DocumentNode CommentToNode(Comment comment, DocumentNode node)
    {
        node.Type = "comment";
        node.Text = string.Join("", comment.Descendants<Text>().Select(t => t.Text));
        if (comment.Author?.Value != null) node.Format["author"] = comment.Author.Value;
        if (comment.Initials?.Value != null) node.Format["initials"] = comment.Initials.Value;
        if (comment.Id?.Value != null) node.Format["id"] = comment.Id.Value;
        if (comment.Date?.Value != null) node.Format["date"] = comment.Date.Value.ToString("o");
        if (comment.Id?.Value != null)
        {
            var anchorPath = FindCommentAnchorPath(comment.Id.Value);
            if (anchorPath != null) node.Format["anchoredTo"] = anchorPath;
        }
        // R21-WB-1: surface direction from the first content paragraph's
        // pPr.BiDi so the cascade (already applied by ApplyCommentFormatKeys)
        // round-trips through Get. Mirrors footnote/endnote readback above.
        var cmtBidi = comment.Descendants<Paragraph>().FirstOrDefault()?.ParagraphProperties?.GetFirstChild<BiDi>();
        if (cmtBidi != null)
            node.Format["direction"] = TryReadOnOff(cmtBidi.Val) == true ? "rtl" : "ltr";
        return node;
    }

    private DocumentNode SectionPropertiesToNode(SectionProperties sectPrEl, string path)
    {
        // CONSISTENCY(section-readback): /body/sectPr[N] should surface
        // the same Format keys as /section[N] so direction, page size,
        // margins, etc. are visible regardless of which path the caller
        // used. Delegate to BuildSectionNode but preserve the original
        // path the caller asked for.
        return BuildSectionNode(sectPrEl, path);
    }

    private DocumentNode RunToNode(Run run, DocumentNode node, string path)
    {
        node.Type = "run";
        node.Text = GetRunText(run);
        // BUG-DUMP7-01: surface <w:sym w:font=… w:char=…/> as a `sym`
        // Format key (font:hex). GetRunText also surfaces the resolved
        // Unicode glyph as Text so the run looks non-empty, but Text
        // alone is lossy — Wingdings F0E0 ↦ U+F0E0 would replay as a
        // plain text run in a non-symbol font and the glyph would
        // disappear. AddRun consumes `sym=` to rebuild SymbolChar.
        var symEl = run.GetFirstChild<SymbolChar>();
        if (symEl?.Char?.Value != null)
        {
            var symFontVal = symEl.Font?.Value ?? "";
            node.Format["sym"] = $"{symFontVal}:{symEl.Char.Value}";
        }
        // BUG-DUMP4-02: surface track-change attribution from any
        // InsertedRun/DeletedRun ancestor wrapping this run. Descendants<Run>
        // unwraps the wrapper so the run looks plain on the curated
        // surface; without this the author/date attribution silently
        // disappears on dump round-trip even though the inner text
        // survives.
        var insAncestor = run.Ancestors<InsertedRun>().FirstOrDefault();
        var moveFromAncestor = insAncestor == null ? run.Ancestors<MoveFromRun>().FirstOrDefault() : null;
        var moveToAncestor = (insAncestor == null && moveFromAncestor == null)
            ? run.Ancestors<MoveToRun>().FirstOrDefault() : null;
        if (insAncestor != null)
        {
            node.Format["revision.type"] = "ins";
            if (!string.IsNullOrEmpty(insAncestor.Author?.Value))
                node.Format["revision.author"] = insAncestor.Author!.Value!;
            if (insAncestor.Date?.Value is DateTime insDate)
                node.Format["revision.date"] = insDate.ToString("o");
            if (insAncestor.Id?.Value is { } insId)
                node.Format["revision.id"] = insId.ToString();
        }
        else if (moveFromAncestor != null)
        {
            // CONSISTENCY(revision-wrapper-readback): w:moveFrom/w:moveTo wrap
            // a run identically to w:ins/w:del; without a Get-side branch the
            // wrapper attribution silently disappeared on dump round-trip and
            // the BatchEmitter re-emitted the run as a plain `add r` (losing
            // the move pairing entirely). Add accepts revision.type=moveFrom/
            // moveTo + author/date/id.
            node.Format["revision.type"] = "moveFrom";
            if (!string.IsNullOrEmpty(moveFromAncestor.Author?.Value))
                node.Format["revision.author"] = moveFromAncestor.Author!.Value!;
            if (moveFromAncestor.Date?.Value is DateTime mfDate)
                node.Format["revision.date"] = mfDate.ToString("o");
            if (moveFromAncestor.Id?.Value is { } mfId)
                node.Format["revision.id"] = mfId.ToString();
        }
        else if (moveToAncestor != null)
        {
            node.Format["revision.type"] = "moveTo";
            if (!string.IsNullOrEmpty(moveToAncestor.Author?.Value))
                node.Format["revision.author"] = moveToAncestor.Author!.Value!;
            if (moveToAncestor.Date?.Value is DateTime mtDate)
                node.Format["revision.date"] = mtDate.ToString("o");
            if (moveToAncestor.Id?.Value is { } mtId)
                node.Format["revision.id"] = mtId.ToString();
        }
        else
        {
            var delAncestor = run.Ancestors<DeletedRun>().FirstOrDefault();
            if (delAncestor != null)
            {
                node.Format["revision.type"] = "del";
                if (!string.IsNullOrEmpty(delAncestor.Author?.Value))
                    node.Format["revision.author"] = delAncestor.Author!.Value!;
                if (delAncestor.Date?.Value is DateTime delDate)
                    node.Format["revision.date"] = delDate.ToString("o");
                if (delAncestor.Id?.Value is { } delId)
                    node.Format["revision.id"] = delId.ToString();
            }
            else
            {
                // AddRun writes <w:rPrChange> for `trackChange=format`. The
                // rPrChange block carries the same author/date attribution
                // as the ins/del wrappers, but rides inside <w:rPr> rather
                // than wrapping the run.
                var rPrChange = run.RunProperties?.GetFirstChild<RunPropertiesChange>();
                if (rPrChange != null)
                {
                    node.Format["revision.type"] = "format";
                    if (!string.IsNullOrEmpty(rPrChange.Author?.Value))
                        node.Format["revision.author"] = rPrChange.Author!.Value!;
                    if (rPrChange.Date?.Value is DateTime rDate)
                        node.Format["revision.date"] = rDate.ToString("o");
                    if (rPrChange.Id?.Value is { } rId)
                        node.Format["revision.id"] = rId.ToString();
                    // BUG: rPrChange's inner snapshot (PreviousRunProperties
                    // when written by SetRevision; plain RunProperties when
                    // written by AddRun's empty-marker path; foreign producers
                    // may use either) carries the pre-change run properties
                    // (bold/italic/color/size/…). AddRun's rPrChange branch
                    // re-stamps the marker with an empty snapshot, losing
                    // what Word's Reject Change would restore. Surface
                    // beforeLost so BatchEmitter can warn. See
                    // WordHandler.Set.Revision.cs note on which strongly-typed
                    // child class round-trips through OpenXml 3.x.
                    var prevPrev = rPrChange.GetFirstChild<PreviousRunProperties>();
                    var prevPlain = rPrChange.GetFirstChild<RunProperties>();
                    bool prevHas = (prevPrev != null && prevPrev.HasChildren)
                                || (prevPlain != null && prevPlain.HasChildren);
                    if (prevHas)
                        node.Format["revision.beforeLost"] = true;
                }
            }
        }
        // CONSISTENCY(canonical-keys): mirror style Get (WordHandler.Query.cs:546-553) —
        // emit per-script font slots, no flat "font" alias. R6 BUG-1: previously
        // collapsed all 4 slots into a single "font" via GetRunFont (Ascii first).
        var rFonts = run.RunProperties?.RunFonts;
        if (rFonts != null)
        {
            // CONSISTENCY(canonical-keys): collapse Ascii+HighAnsi into
            // `font.latin` (canonical per schema docx/run.json) when they
            // match — the round-trip case for `font.latin=` Set. Differing
            // slots fall back to legacy `font.ascii` / `font.hAnsi` keys.
            var ascii = string.IsNullOrEmpty(rFonts.Ascii?.Value) ? null : rFonts.Ascii!.Value;
            var hAnsi = string.IsNullOrEmpty(rFonts.HighAnsi?.Value) ? null : rFonts.HighAnsi!.Value;
            if (ascii != null && hAnsi != null && ascii == hAnsi)
                node.Format["font.latin"] = ascii;
            else
            {
                if (ascii != null && hAnsi != null)
                {
                    node.Format["font.ascii"] = ascii;
                    node.Format["font.hAnsi"] = hAnsi;
                }
                // BUG-DUMP-FONT-LATIN: ascii-only (hAnsi absent) must emit
                // font.ascii, NOT font.latin. font.latin sets BOTH slots on
                // replay, so a source <w:rFonts ascii="X"/> round-tripped to
                // ascii="X" hAnsi="X" — silently rebinding extended-Latin
                // glyphs (which had fallen back to the style/docDefaults font)
                // to X. Same for hAnsi-only.
                else if (ascii != null) node.Format["font.ascii"] = ascii;
                else if (hAnsi != null) node.Format["font.hAnsi"] = hAnsi;
            }
            if (!string.IsNullOrEmpty(rFonts.EastAsia?.Value)) node.Format["font.ea"] = rFonts.EastAsia!.Value!;
            // BUG-DUMP14-03: theme-font slots (asciiTheme/hAnsiTheme/
            // eastAsiaTheme/cstheme) bind a run to a theme major/minor
            // font instead of a literal face name. Without surfacing
            // them, documents using theme fonts lose all font bindings
            // on round-trip (only literal Ascii/HighAnsi were read).
            if (rFonts.AsciiTheme?.HasValue == true)
                node.Format["font.asciiTheme"] = rFonts.AsciiTheme.InnerText;
            if (rFonts.HighAnsiTheme?.HasValue == true)
                node.Format["font.hAnsiTheme"] = rFonts.HighAnsiTheme.InnerText;
            if (rFonts.EastAsiaTheme?.HasValue == true)
                node.Format["font.eaTheme"] = rFonts.EastAsiaTheme.InnerText;
            if (rFonts.ComplexScriptTheme?.HasValue == true)
                node.Format["font.csTheme"] = rFonts.ComplexScriptTheme.InnerText;
            // <w:rFonts w:hint="eastAsia|cs|default"/> selects which font slot
            // renders ambiguous characters (CJK digits/punctuation). Dropping it
            // changes glyph widths on round-trip — a tight CJK line can rewrap.
            // Round-trips via the font.hint key (ApplyRunFormatting writes it).
            if (rFonts.Hint?.HasValue == true)
                node.Format["font.hint"] = rFonts.Hint.InnerText;
        }
        // <w:lang/> three slots: val (latin) / eastAsia / bidi (cs).
        // CONSISTENCY(canonical-keys): mirror font.latin/font.ea/font.cs vocabulary.
        var rLang = run.RunProperties?.GetFirstChild<Languages>();
        if (rLang != null)
        {
            if (rLang.Val?.Value != null) node.Format["lang.latin"] = rLang.Val.Value;
            if (rLang.EastAsia?.Value != null) node.Format["lang.ea"] = rLang.EastAsia.Value;
            if (rLang.Bidi?.Value != null) node.Format["lang.cs"] = rLang.Bidi.Value;
        }
        var size = GetRunFontSize(run);
        if (size != null) node.Format["size"] = size;
        if (run.RunProperties?.Bold != null) node.Format["bold"] = IsToggleOn(run.RunProperties.Bold);
        if (run.RunProperties?.Italic != null) node.Format["italic"] = IsToggleOn(run.RunProperties.Italic);
        // Complex-script readback (font.cs / size.cs / bold.cs / italic.cs).
        // See WordHandler.I18n.cs.
        ReadComplexScriptRunFormatting(run.RunProperties, null, node.Format);
        if (run.RunProperties?.Color?.ThemeColor?.HasValue == true) node.Format["color"] = run.RunProperties.Color.ThemeColor.InnerText;
        else if (run.RunProperties?.Color?.Val?.Value != null) node.Format["color"] = ParseHelpers.FormatHexColor(run.RunProperties.Color.Val.Value);
        if (run.RunProperties?.Underline?.Val != null) node.Format["underline"] = run.RunProperties.Underline.Val.InnerText;
        // CONSISTENCY(underline-color): backfilled from style Get edc8f884.
        if (run.RunProperties?.Underline?.Color?.Value != null)
            node.Format["underline.color"] = ParseHelpers.FormatHexColor(run.RunProperties.Underline.Color.Value);
        if (run.RunProperties?.Strike != null) node.Format["strike"] = IsToggleOn(run.RunProperties.Strike);
        if (run.RunProperties?.Highlight?.Val != null) node.Format["highlight"] = run.RunProperties.Highlight.Val.InnerText;
        if (run.RunProperties?.Caps != null) node.Format["caps"] = IsToggleOn(run.RunProperties.Caps);
        if (run.RunProperties?.SmallCaps != null) node.Format["smallcaps"] = IsToggleOn(run.RunProperties.SmallCaps);
        if (run.RunProperties?.DoubleStrike != null) node.Format["dstrike"] = IsToggleOn(run.RunProperties.DoubleStrike);
        if (run.RunProperties?.Vanish != null) node.Format["vanish"] = IsToggleOn(run.RunProperties.Vanish);
        if (run.RunProperties?.Outline != null) node.Format["outline"] = IsToggleOn(run.RunProperties.Outline);
        if (run.RunProperties?.Shadow != null) node.Format["shadow"] = IsToggleOn(run.RunProperties.Shadow);
        if (run.RunProperties?.Emboss != null) node.Format["emboss"] = IsToggleOn(run.RunProperties.Emboss);
        if (run.RunProperties?.Imprint != null) node.Format["imprint"] = IsToggleOn(run.RunProperties.Imprint);
        if (run.RunProperties?.NoProof != null) node.Format["noproof"] = IsToggleOn(run.RunProperties.NoProof);
        if (run.RunProperties?.RightToLeftText != null)
        {
            // <w:rtl/> with no Val attribute implies true; <w:rtl w:val="0"/>
            // is an explicit off-override (overrides inherited docDefaults).
            // CONSISTENCY(canonical-key): paragraphs and sections surface
            // this property as Format["direction"]="rtl"|"ltr"; runs must
            // match so users see one canonical key across scopes (R16-bt-1).
            var rtlVal = run.RunProperties.RightToLeftText.Val;
            var on = rtlVal == null ? true : rtlVal.Value;
            node.Format["direction"] = on ? "rtl" : "ltr";
        }
        if (run.RunProperties?.VerticalTextAlignment?.Val?.Value == VerticalPositionValues.Superscript)
            node.Format["superscript"] = true;
        if (run.RunProperties?.VerticalTextAlignment?.Val?.Value == VerticalPositionValues.Subscript)
            node.Format["subscript"] = true;
        // ApplyRunFormatting writes <w:position> for `position` (raised /
        // lowered baseline offset in half-points). Mirror it on the Get
        // side so the round-trip key survives.
        var posVal = run.RunProperties?.GetFirstChild<Position>()?.Val?.Value;
        if (!string.IsNullOrEmpty(posVal))
            node.Format["position"] = posVal;
        if (run.RunProperties?.Spacing?.Val?.HasValue == true)
            node.Format["charSpacing"] = $"{run.RunProperties.Spacing.Val.Value / 20.0:0.##}pt";
        // BUG-DUMP22-08: <w:bdr/> (character border) is multi-attribute
        // (val + sz + color + space) so the long-tail FillUnknownChildProps
        // skipped it (attrCount > 1), leaving only the surface bare key
        // with no sub-attrs. Emit the colon-encoded compound form that
        // ApplyRunFormatting consumes on replay so dump round-trips
        // preserve size and color.
        var rBdr = run.RunProperties?.GetFirstChild<Border>();
        if (rBdr?.Val?.HasValue == true)
        {
            var bdrStyle = rBdr.Val!.InnerText;
            var bdrSize = rBdr.Size?.Value;
            var bdrColor = rBdr.Color?.Value;
            var bdrSpace = rBdr.Space?.Value;
            node.Format["bdr"] = string.Join(';', new[]
            {
                bdrStyle,
                bdrSize?.ToString() ?? "",
                string.IsNullOrEmpty(bdrColor) ? "" : ParseHelpers.FormatHexColor(bdrColor),
                bdrSpace?.ToString() ?? "0"
            });
        }
        if (run.RunProperties?.Shading != null)
        {
            // BUG-DUMP22-01/02: surface val/fill/color sub-keys instead of
            // a bare `shading=fill` value. The bare form silently coerced
            // val to "clear" and dropped color on dump round-trip. Mirrors
            // the paragraph/table/cell shading reader (round-21 fix).
            var rShdVal = run.RunProperties.Shading.Val?.InnerText;
            var rShdFill = run.RunProperties.Shading.Fill?.Value;
            var rShdColor = run.RunProperties.Shading.Color?.Value;
            if (!string.IsNullOrEmpty(rShdVal)) node.Format["shading.val"] = rShdVal;
            if (!string.IsNullOrEmpty(rShdFill)) node.Format["shading.fill"] = ParseHelpers.FormatHexColor(rShdFill);
            if (!string.IsNullOrEmpty(rShdColor)) node.Format["shading.color"] = ParseHelpers.FormatHexColor(rShdColor);
        }
        // w14 text effects
        ReadW14TextEffects(run.RunProperties, node);
        // BUG-DUMP10-01: w:eastAsianLayout (vert/combine/vertCompress)
        // is a multi-attribute child the long-tail FillUnknownChildProps
        // skips (it only handles single-val/no-attr leaves). Without an
        // explicit reader, vertical-text and two-lines-in-one CJK layout
        // was silently dropped on dump→batch round-trip. Set side is
        // covered by TypedAttributeFallback.TrySet which creates the
        // dotted child + attr automatically.
        if (run.RunProperties?.GetFirstChild<EastAsianLayout>() is { } eal)
        {
            if (eal.Vertical?.Value == true) node.Format["eastAsianLayout.vert"] = "1";
            if (eal.Combine?.Value == true) node.Format["eastAsianLayout.combine"] = "1";
            if (eal.VerticalCompress?.HasValue == true)
                node.Format["eastAsianLayout.vertCompress"] = eal.VerticalCompress.InnerText;
            if (eal.CombineBrackets?.HasValue == true)
                node.Format["eastAsianLayout.combineBrackets"] = eal.CombineBrackets.InnerText;
        }
        // Long-tail fallback: surface every rPr child the curated reader
        // didn't consume. Symmetric with the Set-side TryCreateTypedChild
        // fallback in SetElementRun (WordHandler.Set.Element.cs).
        FillUnknownChildProps(run.RunProperties, node);
        // Image properties if run contains a Drawing.
        // BUG-R5-T3: previously this branch wrote only id/name/alt/width/
        // height/relId — wrap/hPosition/vPosition/hRelative/vRelative/
        // behindText for floating pictures were silently dropped, which
        // also broke dump→batch round-trip (WordBatchEmitter relies on Get).
        // Reuse CreateImageNode (the canonical picture-node builder) and
        // merge its Format bag into the run node.
        var runDrawing = run.GetFirstChild<Drawing>();
        if (runDrawing != null)
        {
            var picNode = CreateImageNode(runDrawing, run, path);
            node.Type = picNode.Type;
            if (!string.IsNullOrEmpty(picNode.Text)) node.Text = picNode.Text;
            foreach (var kv in picNode.Format)
                node.Format[kv.Key] = kv.Value;
        }
        // OLE object if run contains an EmbeddedObject. The underlying
        // logic is the same as CreateOleNode — reuse it so Get/Query
        // return identical shapes.
        var runOle = run.GetFirstChild<EmbeddedObject>();
        if (runOle != null)
        {
            // CONSISTENCY(ole-host-part): mirror Query.cs's header/footer
            // OLE handling — the EmbeddedObjectPart relationship lives on
            // the owning Header/Footer part, not the MainDocumentPart.
            // Walk ancestors to find the host part so CreateOleNode can
            // populate contentType/fileSize instead of returning orphan.
            OpenXmlPart? hostPart = _doc.MainDocumentPart;
            var headerAncestor = run.Ancestors<Header>().FirstOrDefault();
            if (headerAncestor != null && _doc.MainDocumentPart != null)
            {
                var hp = _doc.MainDocumentPart.HeaderParts
                    .FirstOrDefault(p => ReferenceEquals(p.Header, headerAncestor));
                if (hp != null) hostPart = hp;
            }
            else
            {
                var footerAncestor = run.Ancestors<Footer>().FirstOrDefault();
                if (footerAncestor != null && _doc.MainDocumentPart != null)
                {
                    var fp = _doc.MainDocumentPart.FooterParts
                        .FirstOrDefault(p => ReferenceEquals(p.Footer, footerAncestor));
                    if (fp != null) hostPart = fp;
                }
            }
            var oleNode = CreateOleNode(runOle, run, path, hostPart);
            // Keep the node's path as-is, but swap in the OLE-sourced
            // type/format bag.
            node.Type = oleNode.Type;
            foreach (var kv in oleNode.Format)
                node.Format[kv.Key] = kv.Value;
            if (!string.IsNullOrEmpty(oleNode.Text))
                node.Text = oleNode.Text;
        }
        // CONSISTENCY(run-special-content): runs that primarily carry inline
        // structure (ptab, fldChar, instrText, tab, break) instead of a
        // <w:t> payload were previously surfaced as opaque
        // {type:"run", text:""} placeholders — six of these in a row in
        // header/footer paragraphs (PAGE field begin/instr/separate/end +
        // ptab anchors), all indistinguishable. Upgrade the node.Type so
        // callers walking paragraph.children can rebuild left/center/right
        // alignment regions and detect field markers without reparsing the
        // raw OOXML themselves. Mirrors the type=picture / type=ole
        // pattern above.
        //
        // Each block is gated on `node.Type == "run"` so that:
        //   (a) Drawing/EmbeddedObject (already upgraded above to
        //       picture/ole) wins over a co-residing <w:br>/<w:tab> —
        //       picture+break is a real Word emission and the picture
        //       identity must not be silently overwritten;
        //   (b) the first matching structural element wins when several
        //       coexist in one run (rare but possible), keeping node.Type
        //       single-valued and deterministic. ptab is checked first
        //       (most semantically distinctive), then fieldChar, then
        //       instrText, then tab, then break.
        if (node.Type == "run")
        {
            var ptabEl = run.GetFirstChild<PositionalTab>();
            if (ptabEl != null)
            {
                node.Type = "ptab";
                // Open XML SDK v3 enum .ToString() returns "FooValues { }"
                // — use .InnerText to get the actual XML attribute value
                // ("center", "right", "begin", etc.). Same trap as the
                // LineSpacingRuleValues note in WordHandler CLAUDE.md.
                if (ptabEl.Alignment?.HasValue == true)
                    node.Format["align"] = ptabEl.Alignment.InnerText;
                if (ptabEl.RelativeTo?.HasValue == true)
                    node.Format["relativeTo"] = ptabEl.RelativeTo.InnerText;
                if (ptabEl.Leader?.HasValue == true)
                    node.Format["leader"] = ptabEl.Leader.InnerText;
            }
        }
        if (node.Type == "run")
        {
            var fldCharEl = run.GetFirstChild<FieldChar>();
            if (fldCharEl != null)
            {
                node.Type = "fieldChar";
                if (fldCharEl.FieldCharType?.HasValue == true)
                    node.Format["fieldCharType"] = fldCharEl.FieldCharType.InnerText;
                // CONSISTENCY(field-cache-stale): expose dirty so audit
                // tools can verify whether Set instr / Set cached
                // properly flagged the owning field for recompute. The
                // attribute persists in OOXML; surfacing it via Get
                // closes the loop the Round 3 dirty fix opened.
                if (fldCharEl.Dirty?.Value == true)
                    node.Format["dirty"] = true;
                if (fldCharEl.FormFieldData != null)
                {
                    node.Format["hasFormFieldData"] = true;
                    // R14-bug1+2: surface the ffData payload on the fieldChar
                    // begin node so WordBatchEmitter.CollapseFieldChains can
                    // emit a `type=formfield` synth carrying the full
                    // ffData props (name, default, maxLength, checkBox.size,
                    // checkBox.default, ddList items, helpText, statusText,
                    // entryMacro, exitMacro, calcOnExit, textInput.type,
                    // textInput.format). Without these keys, the dump
                    // emitted a bare `add field instr=FORMTEXT` row that
                    // walked the BuildFieldAddProps default arm and lost
                    // every ffData wrapper on replay.
                    var ffd = fldCharEl.FormFieldData;
                    var ffName = ffd.GetFirstChild<FormFieldName>()?.Val?.Value;
                    if (!string.IsNullOrEmpty(ffName))
                        node.Format["ffName"] = ffName;
                    var ffEnabled = ffd.GetFirstChild<Enabled>();
                    if (ffEnabled != null)
                        node.Format["ffEnabled"] = ffEnabled.Val?.Value ?? true;
                    var ffHelp = ffd.GetFirstChild<HelpText>()?.Val?.Value;
                    if (!string.IsNullOrEmpty(ffHelp))
                        node.Format["ffHelpText"] = ffHelp;
                    var ffStatus = ffd.GetFirstChild<StatusText>()?.Val?.Value;
                    if (!string.IsNullOrEmpty(ffStatus))
                        node.Format["ffStatusText"] = ffStatus;
                    var ffEntry = ffd.GetFirstChild<EntryMacro>()?.Val?.Value;
                    if (!string.IsNullOrEmpty(ffEntry))
                        node.Format["ffEntryMacro"] = ffEntry;
                    var ffExit = ffd.GetFirstChild<ExitMacro>()?.Val?.Value;
                    if (!string.IsNullOrEmpty(ffExit))
                        node.Format["ffExitMacro"] = ffExit;
                    var ffCalc = ffd.GetFirstChild<CalculateOnExit>();
                    if (ffCalc != null)
                        node.Format["ffCalcOnExit"] = ffCalc.Val?.Value ?? true;

                    var ti = ffd.GetFirstChild<TextInput>();
                    var cb = ffd.GetFirstChild<CheckBox>();
                    var dd = ffd.GetFirstChild<DropDownListFormField>();
                    if (ti != null)
                    {
                        node.Format["ffType"] = "text";
                        var tDef = ti.GetFirstChild<DefaultTextBoxFormFieldString>()?.Val?.Value;
                        if (!string.IsNullOrEmpty(tDef))
                            node.Format["ffDefault"] = tDef;
                        var tMax = ti.GetFirstChild<MaxLength>()?.Val?.Value;
                        if (tMax != null && tMax.Value != 0)
                            node.Format["ffMaxLength"] = (int)tMax;
                        var tTyp = ti.GetFirstChild<TextBoxFormFieldType>()?.Val?.InnerText;
                        if (!string.IsNullOrEmpty(tTyp))
                            node.Format["ffTextType"] = tTyp;
                        var tFmt = ti.GetFirstChild<Format>()?.Val?.Value;
                        if (!string.IsNullOrEmpty(tFmt))
                            node.Format["ffTextFormat"] = tFmt;
                    }
                    else if (cb != null)
                    {
                        node.Format["ffType"] = "checkbox";
                        var cChecked = cb.GetFirstChild<Checked>();
                        var cDefault = cb.GetFirstChild<DefaultCheckBoxFormFieldState>();
                        var isChk = cChecked?.Val?.Value ?? cDefault?.Val?.Value ?? false;
                        node.Format["ffChecked"] = isChk;
                        var cSize = cb.GetFirstChild<FormFieldSize>()?.Val?.Value;
                        if (!string.IsNullOrEmpty(cSize))
                            node.Format["ffCheckBoxSize"] = cSize;
                    }
                    else if (dd != null)
                    {
                        node.Format["ffType"] = "dropdown";
                        var items = dd.Elements<ListEntryFormField>()
                            .Select(li => li.Val?.Value ?? "").ToList();
                        if (items.Count > 0)
                            node.Format["ffItems"] = string.Join(",", items);
                        var dSel = dd.GetFirstChild<DropDownListSelection>()?.Val?.Value;
                        if (dSel != null)
                            node.Format["ffDefault"] = (int)dSel;
                    }
                }
            }
        }
        if (node.Type == "run")
        {
            var instrEl = run.GetFirstChild<FieldCode>();
            if (instrEl != null)
            {
                node.Type = "instrText";
                node.Format["instruction"] = instrEl.Text ?? "";
                // CONSISTENCY(canonical-keys): also surface the
                // instruction as node.Text so selector text-contains
                // searches (`instrText[text~=PAGE]`) and Get readback
                // agree. Without this, MatchesRunSelector's
                // GetRunText fallback hits the <w:instrText> content
                // while Navigation hands callers an empty Text — the
                // two surfaces disagreed on what the run "says".
                node.Text = instrEl.Text ?? "";
            }
        }
        // CONSISTENCY(run-text-tab): the type-upgrade for tab/break runs
        // checks "no Text element" (not "node.Text empty") because
        // GetRunText now surfaces TabChar as \t in node.Text. A pure
        // <w:r><w:tab/></w:r> run has no <w:t> child but node.Text="\t".
        if (node.Type == "run" && !run.Elements<Text>().Any())
        {
            var tabEl = run.GetFirstChild<TabChar>();
            if (tabEl != null)
            {
                node.Type = "tab";
                node.Text = "";
            }
        }
        // CONSISTENCY(run-text-break): gate on "no Text element" (not
        // "node.Text empty"), same as the tab upgrade above. GetRunText
        // surfaces a soft <w:br/> (textWrapping) as \n in node.Text, so a
        // pure <w:r><w:br/></w:r> break run is no longer Text-empty and the
        // old IsNullOrEmpty gate skipped the upgrade — leaving it as a run
        // with text="\n" that the emitter mis-rendered. A mixed run
        // <w:t>foo</w:t><w:br/> still has a <w:t> child, so it stays a run
        // (text="foo\n") and the inline break is preserved as \n.
        if (node.Type == "run" && !run.Elements<Text>().Any())
        {
            var breakEl = run.GetFirstChild<Break>();
            if (breakEl != null)
            {
                node.Type = "break";
                node.Text = "";
                // Normalize "textWrapping" → "line" on emit. OOXML treats
                // a typeless <w:br/> as textWrapping (the default), but
                // AddBreak's user-facing vocab uses "line"; without
                // normalisation, dump round-trip emits `type=line` from
                // typeless source and `type=textWrapping` from the
                // explicitly-stamped replay target — semantically
                // identical, byte-different.
                if (breakEl.Type?.HasValue == true)
                {
                    var bt = breakEl.Type.InnerText;
                    node.Format["breakType"] = string.Equals(bt, "textWrapping", StringComparison.OrdinalIgnoreCase)
                        ? "line"
                        : bt;
                }
            }
        }

        if (run.Parent is Hyperlink hlParent)
        {
            // BUG-DUMP10-05: a hyperlink wrapper with neither r:id nor
            // anchor (tooltip-only / history-only) used to fall through
            // both branches below, leaving the run with no Format keys
            // that would trigger the WordBatchEmitter hyperlink-emit guard.
            // Surface a sentinel so the wrapper survives even when there
            // is no destination — required for w:hyperlink[@w:tooltip]
            // bookmarks-style hover popups.
            node.Format["isHyperlink"] = true;
            if (hlParent.Id?.Value != null)
            {
                try
                {
                    var rel = ResolveHyperlinkRelationship(hlParent, hlParent.Id.Value);
                    // CONSISTENCY(docx-hyperlink-canonical-url): schema docx/hyperlink.json
                    // declares `url` as the canonical key; `link` is accepted as an input
                    // alias by Add/Set but Get normalizes output to `url`.
                    // Use OriginalString rather than ToString() — System.Uri normalises
                    // a bare authority by appending a trailing `/` (e.g.
                    // `https://example.com` → `https://example.com/`), which would
                    // surface as a get-side drift the on-disk .rels Target lacks.
                    if (rel != null) node.Format["url"] = rel.Uri.OriginalString;
                }
                catch { }
            }
            // CONSISTENCY(internal-anchor-hyperlink): runs inside an
            // internal anchor hyperlink (w:hyperlink[@w:anchor]) had no
            // r:id, so `anchor` was never surfaced on the run. The
            // WordBatchEmitter hyperlink branch keys off Format["anchor"]/
            // ["url"] to emit `add hyperlink`; without anchor the run
            // was demoted to a plain `add r` and the link was lost on
            // dump→batch round-trip.
            if (hlParent.Anchor?.Value != null)
                node.Format["anchor"] = hlParent.Anchor.Value;
            // BUG-DUMP24-02: w:docLocation is a separate "location in
            // target document" attribute, distinct from w:anchor. Surface
            // it so dump→batch round-trips the wrapping hyperlink fully.
            if (hlParent.DocLocation?.Value != null)
                node.Format["docLocation"] = hlParent.DocLocation.Value;
            // BUG-DUMP10-02: surface the tooltip / tgtFrame / history
            // attributes from the wrapping hyperlink so dump→batch
            // round-trip preserves them. Same canonical keys as the
            // standalone Hyperlink branch below.
            if (hlParent.Tooltip?.Value != null)
                node.Format["tooltip"] = hlParent.Tooltip.Value;
            if (hlParent.TargetFrame?.Value != null)
                node.Format["tgtFrame"] = hlParent.TargetFrame.Value;
            if (hlParent.History?.Value == true)
                node.Format["history"] = true;
        }

        // Populate effective.* properties from style inheritance.
        // CONSISTENCY(run-special-content): runs whose primary payload
        // is a structural inline element (ptab/fieldChar/instrText/tab/
        // break) carry no glyph for font/size/color to apply to;
        // emitting effective.size / effective.font.* on them only
        // floods output with noise and primes audit tools to misread
        // cosmetic styles on a "fldChar end" marker as meaningful.
        // Picture/ole runs are gated for the same reason — their
        // typography is irrelevant to the embedded media.
        var parentPara = run.Ancestors<Paragraph>().FirstOrDefault();
        if (parentPara != null && node.Type == "run")
            PopulateEffectiveRunProperties(node, run, parentPara);

        // Same noise-suppression for direct rPr-level keys read before
        // the type upgrade above (font.*/size/bold/...): on a pure MARKER
        // run (fieldChar / instrText / break) the rPr has no glyph to paint,
        // so surfacing font/size/color is noise that primes audit tools to
        // misread cosmetic styling on a structural marker as meaningful —
        // strip it so the bag shows only the role-defining keys
        // (fieldCharType, instr, breakType, …).
        //
        // BUG-DUMP-TABRPR: tab and ptab are deliberately NOT in this list
        // (they were, historically, under the same "no glyph" rationale —
        // but that was wrong for them). A tab / positional tab is a SIZED
        // inline element: Word routinely stamps rPr on these runs (font/size/
        // szCs drive the line height and the leader-dot glyphs the tab
        // paints), and stripping it dropped that rPr on every dump→batch
        // round-trip, leaving an empty <w:rPr/>. The earlier strip never
        // needed them anyway — the tab/ptab assertions
        // (Req5Round12FuzzTabSpecialTests) only check a run shows no
        // typography AFTER Set rejected/never-applied any (so the run
        // genuinely has none to read); they don't require hiding typography
        // that a source document already authored.
        if (node.Type is "fieldChar" or "instrText" or "break")
        {
            foreach (var noiseKey in TypographyOnlyKeys)
                node.Format.Remove(noiseKey);
        }
        return node;
    }

    private DocumentNode HyperlinkToNode(Hyperlink hyperlink, DocumentNode node)
    {
        node.Type = "hyperlink";
        node.Text = string.Concat(hyperlink.Descendants<Text>().Select(t => t.Text));
        var relId = hyperlink.Id?.Value;
        if (relId != null)
        {
            try
            {
                var rel = ResolveHyperlinkRelationship(hyperlink, relId);
                // CONSISTENCY(docx-hyperlink-canonical-url): see note above.
                if (rel != null) node.Format["url"] = rel.Uri.OriginalString;
            }
            catch { }
        }
        // Internal-anchor hyperlink (`add --type hyperlink --prop anchor=Foo`)
        // sets w:hyperlink/@w:anchor instead of @r:id. Surface it so set/get
        // round-trips and users can debug why a link points where it does.
        if (hyperlink.Anchor?.Value != null)
            node.Format["anchor"] = hyperlink.Anchor.Value;
        // BUG-DUMP24-02: w:docLocation is a separate "location in target
        // document" attribute, distinct from w:anchor. Surface it so
        // dump→batch round-trips it.
        if (hyperlink.DocLocation?.Value != null)
            node.Format["docLocation"] = hyperlink.DocLocation.Value;
        // BUG-DUMP10-02: tooltip / tgtFrame / history attributes are
        // independent of url/anchor — surface them so dump→batch
        // preserves the hover popup, target window, and history flag.
        if (hyperlink.Tooltip?.Value != null)
            node.Format["tooltip"] = hyperlink.Tooltip.Value;
        if (hyperlink.TargetFrame?.Value != null)
            node.Format["tgtFrame"] = hyperlink.TargetFrame.Value;
        if (hyperlink.History?.Value == true)
            node.Format["history"] = true;
        // Read run formatting from the first run inside the hyperlink
        var hlRun = hyperlink.Elements<Run>().FirstOrDefault(r => r.GetFirstChild<Text>() != null);
        if (hlRun?.RunProperties != null)
        {
            var rp = hlRun.RunProperties;
            if (rp.RunFonts?.Ascii?.Value != null) node.Format["font"] = rp.RunFonts.Ascii.Value;
            // BUG-DUMP17-07: surface per-script font slot so dump→batch
            // round-trip preserves font.cs on hyperlink runs.
            if (rp.RunFonts?.ComplexScript?.Value != null) node.Format["font.cs"] = rp.RunFonts.ComplexScript.Value;
            if (rp.FontSize?.Val?.Value != null)
                node.Format["size"] = $"{int.Parse(rp.FontSize.Val.Value) / 2.0:0.##}pt";
            if (rp.Bold != null) node.Format["bold"] = IsToggleOn(rp.Bold);
            if (rp.Italic != null) node.Format["italic"] = IsToggleOn(rp.Italic);
            if (rp.Color?.ThemeColor?.HasValue == true) node.Format["color"] = rp.Color.ThemeColor.InnerText;
            else if (rp.Color?.Val?.Value != null) node.Format["color"] = ParseHelpers.FormatHexColor(rp.Color.Val.Value);
            if (rp.Underline?.Val != null) node.Format["underline"] = rp.Underline.Val.InnerText;
            // CONSISTENCY(underline-color): backfilled from style Get edc8f884.
            if (rp.Underline?.Color?.Value != null)
                node.Format["underline.color"] = ParseHelpers.FormatHexColor(rp.Underline.Color.Value);
            if (rp.Strike != null) node.Format["strike"] = IsToggleOn(rp.Strike);
            if (rp.Highlight?.Val != null) node.Format["highlight"] = rp.Highlight.Val.InnerText;
            var rStyle = rp.GetFirstChild<RunStyle>();
            if (rStyle?.Val?.Value != null)
                node.Format["rStyle"] = rStyle.Val.Value;
        }
        return node;
    }

    private DocumentNode TableToNode(Table table, DocumentNode node, string path, int depth)
    {
        node.Type = "table";
        node.ChildCount = table.Elements<TableRow>().Count();
        var firstRow = table.Elements<TableRow>().FirstOrDefault();
        // Use grid column count (from TableGrid) instead of cell count for accurate column reporting
        var gridColCount = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().Count();
        // CONSISTENCY(format-stringy): user-facing numeric counts are
        // stored as strings to match other Word format keys (size "14pt",
        // spacing "12pt"). Avoids object-vs-int comparison surprises.
        node.Format["cols"] = (gridColCount ?? firstRow?.Elements<TableCell>().Count() ?? 0).ToString();
        node.Format["rows"] = node.ChildCount.ToString();
        // _gridCols: actual <w:gridCol> count (0 when TableGrid is missing
        // or empty), unbiased by the row-cell fallback that `cols` uses for
        // backward-compat. EmitTable reads this to decide whether to emit
        // `gridCols=0` on the dumped `add table` so AddTable leaves the
        // <w:tblGrid/> empty — preserving sources whose cells encode width
        // via tcW (or auto-fit). Underscore-prefixed to mark it as
        // internal-only (not a user-facing Set/Add key).
        node.Format["_gridCols"] = (gridColCount ?? 0).ToString();

        var tp = table.GetFirstChild<TableProperties>();
        if (tp != null)
        {
            // tblPrChange: `set table + trackChange.author` snapshots the
            // prior tblPr and stamps author/date. Surfaced under
            // tblPrChange.* so EmitTable can emit a follow-up
            // `set /body/tbl[N]` step carrying trackChange.author/date
            // (re-runs the snapshot+stamp on replay). Distinct namespace
            // from any future bare `trackChange=` on tables.
            var tblPrChange = tp.GetFirstChild<TablePropertiesChange>();
            if (tblPrChange != null)
            {
                if (!string.IsNullOrEmpty(tblPrChange.Author?.Value))
                    node.Format["tblPrChange.author"] = tblPrChange.Author!.Value!;
                if (tblPrChange.Date?.Value is DateTime tDate)
                    node.Format["tblPrChange.date"] = tDate.ToString("o");
            }
            // Table style
            // BUG-R3-05: empty Val (set via legacy code that wrote tblStyle
            // with empty string) must NOT surface as a "style" key.
            if (!string.IsNullOrEmpty(tp.TableStyle?.Val?.Value))
                node.Format["style"] = tp.TableStyle.Val.Value!;
            // Table borders. `LeftBorder`/`RightBorder` only catch
            // <w:left>/<w:right>; bidi-aware sources use <w:start>/<w:end>
            // which the SDK does NOT alias onto Left/Right (the typed
            // properties stay null). Walk all border children by local
            // name and map both forms onto the same canonical key — the
            // alternative is dropping borders for any doc whose tblBorders
            // uses the start/end naming (three-line-table2.docx).
            var tblBorders = tp.TableBorders;
            if (tblBorders != null)
            {
                ReadBorder(tblBorders.TopBorder, "border.top", node);
                ReadBorder(tblBorders.BottomBorder, "border.bottom", node);
                ReadBorder(tblBorders.InsideHorizontalBorder, "border.insideH", node);
                ReadBorder(tblBorders.InsideVerticalBorder, "border.insideV", node);
                foreach (var bChild in tblBorders.ChildElements)
                {
                    if (bChild is BorderType bt)
                    {
                        var ln = bChild.LocalName;
                        if (ln.Equals("left", StringComparison.OrdinalIgnoreCase)
                            || ln.Equals("start", StringComparison.OrdinalIgnoreCase))
                            ReadBorder(bt, "border.left", node);
                        else if (ln.Equals("right", StringComparison.OrdinalIgnoreCase)
                                 || ln.Equals("end", StringComparison.OrdinalIgnoreCase))
                            ReadBorder(bt, "border.right", node);
                    }
                }
            }
            // Table width
            if (tp.TableWidth?.Width?.Value != null)
            {
                var wType = tp.TableWidth.Type?.Value;
                // BUG-DUMP19-03: type=auto must round-trip as "auto", not
                // collapse to a bare dxa integer (Width="0").
                node.Format["width"] = wType == TableWidthUnitValues.Pct
                    ? (int.Parse(tp.TableWidth.Width.Value) / 50) + "%"
                    : wType == TableWidthUnitValues.Auto
                        ? "auto"
                        : tp.TableWidth.Width.Value;
            }
            else if (tp.TableWidth?.Type?.Value == TableWidthUnitValues.Auto)
            {
                // Some producers emit <w:tblW w:type="auto"/> without w:w.
                node.Format["width"] = "auto";
            }
            else
            {
                // Internal-only marker: source had no <w:tblW> element at
                // all. EmitTable reads this to tell AddTable to skip the
                // default-tblW stamp; without it, replay grows
                // a <w:tblW w:w="<sum-of-gridCol>" w:type="dxa"/> that
                // the source never had, and the next dump surfaces a
                // phantom `width=…` key.
                node.Format["_noTblW"] = true;
            }
            // Alignment
            if (tp.TableJustification?.Val?.Value != null)
                node.Format["align"] = tp.TableJustification.Val.InnerText;
            // Indent
            if (tp.TableIndentation?.Width?.Value != null)
                node.Format["indent"] = tp.TableIndentation.Width.Value;
            // Cell spacing
            if (tp.TableCellSpacing?.Width?.Value != null)
                node.Format["cellSpacing"] = tp.TableCellSpacing.Width.Value;
            // Layout — emit "autofit" (not "auto") so the readback token
            // matches the canonical input vocabulary documented in the
            // table add/set help. Set accepts both "auto" and "autofit"
            // (anything not "fixed" maps to Autofit), so this only affects
            // get and is round-trip safe with the dump/replay pipeline.
            if (tp.TableLayout?.Type?.Value != null)
                node.Format["layout"] = tp.TableLayout.Type.Value == TableLayoutValues.Fixed ? "fixed" : "autofit";
            // Direction (CT_TblPrBase / w:bidiVisual). Mirrors paragraph
            // direction vocabulary; presence-only readback (no bidiVisual
            // means no key — LTR is the default).
            if (tp.GetFirstChild<BiDiVisual>() != null)
                node.Format["direction"] = "rtl";
            // Default cell margin (padding)
            var dcm = tp.TableCellMarginDefault;
            if (dcm?.TopMargin?.Width?.Value != null)
                node.Format["padding.top"] = dcm.TopMargin.Width.Value;
            if (dcm?.BottomMargin?.Width?.Value != null)
                node.Format["padding.bottom"] = dcm.BottomMargin.Width.Value;
            if (dcm?.TableCellLeftMargin?.Width?.Value != null)
                node.Format["padding.left"] = dcm.TableCellLeftMargin.Width.Value;
            if (dcm?.TableCellRightMargin?.Width?.Value != null)
                node.Format["padding.right"] = dcm.TableCellRightMargin.Width.Value;
            // Table-level shading (w:tblPr/w:shd). Mirror paragraph shading
            // pattern: split into shading.val/.fill/.color sub-keys.
            // WordBatchEmitter's shading-fold collapses these into a single
            // semicolon-encoded `shading=VAL;FILL[;COLOR]` value, which
            // AddTable consumes via the existing "shading" case.
            // BUG-DUMP22-09: floating-table position (<w:tblpPr/>) and
            // overlap (<w:tblOverlap/>) — both were silently dropped on
            // dump, leaving floating tables stuck inline on round-trip.
            // Surface tblpPr's six attrs as tblp.* dotted keys (using the
            // OOXML attribute local names verbatim) plus tblOverlap as a
            // dotted sibling so AddTable's TypedAttributeFallback can
            // re-create the elements verbatim. CONSISTENCY(canonical-keys):
            // dotted-segment-as-element-prefix matches ind.firstLine and
            // pBdr.top patterns.
            var tblpPr = tp.GetFirstChild<TablePositionProperties>();
            if (tblpPr != null)
            {
                if (tblpPr.HorizontalAnchor?.HasValue == true)
                    node.Format["tblp.horzAnchor"] = tblpPr.HorizontalAnchor.InnerText;
                if (tblpPr.VerticalAnchor?.HasValue == true)
                    node.Format["tblp.vertAnchor"] = tblpPr.VerticalAnchor.InnerText;
                if (tblpPr.TablePositionX?.HasValue == true)
                    node.Format["tblp.tblpX"] = tblpPr.TablePositionX.Value!;
                if (tblpPr.TablePositionY?.HasValue == true)
                    node.Format["tblp.tblpY"] = tblpPr.TablePositionY.Value!;
                if (tblpPr.TablePositionXAlignment?.HasValue == true)
                    node.Format["tblp.tblpXSpec"] = tblpPr.TablePositionXAlignment.InnerText;
                if (tblpPr.TablePositionYAlignment?.HasValue == true)
                    node.Format["tblp.tblpYSpec"] = tblpPr.TablePositionYAlignment.InnerText;
                if (tblpPr.LeftFromText?.HasValue == true)
                    node.Format["tblp.leftFromText"] = tblpPr.LeftFromText.Value!;
                if (tblpPr.RightFromText?.HasValue == true)
                    node.Format["tblp.rightFromText"] = tblpPr.RightFromText.Value!;
                if (tblpPr.TopFromText?.HasValue == true)
                    node.Format["tblp.topFromText"] = tblpPr.TopFromText.Value!;
                if (tblpPr.BottomFromText?.HasValue == true)
                    node.Format["tblp.bottomFromText"] = tblpPr.BottomFromText.Value!;
            }
            var tblOverlap = tp.GetFirstChild<TableOverlap>();
            if (tblOverlap?.Val?.HasValue == true)
                node.Format["tblOverlap.val"] = tblOverlap.Val.InnerText;
            if (tp.Shading != null)
            {
                var tShdVal = tp.Shading.Val?.InnerText;
                var tShdFill = tp.Shading.Fill?.Value;
                var tShdColor = tp.Shading.Color?.Value;
                if (!string.IsNullOrEmpty(tShdVal)) node.Format["shading.val"] = tShdVal;
                if (!string.IsNullOrEmpty(tShdFill)) node.Format["shading.fill"] = ParseHelpers.FormatHexColor(tShdFill);
                if (!string.IsNullOrEmpty(tShdColor)) node.Format["shading.color"] = ParseHelpers.FormatHexColor(tShdColor);
            }

            // BUG-R3-01: tblLook readback — Set wrote the XML correctly, but
            // Get never read it back (Set/Get round-trip gap). Emit both the
            // short-form lowercase keys (firstrow/lastrow/bandrow — match
            // Set's case-insensitive vocabulary and project canonical
            // pattern: vmerge/colspan) AND OOXML-attribute-name camelCase
            // keys (firstRow/bandedRows — verbatim attribute names) so
            // batch round-trip works either way. The two forms exist for
            // historical-vocabulary parity; values are kept consistent
            // across both keys (lowercase stores "true"/"false" string,
            // camelCase stores bool).
            // BUG-R4-01/06: Get emits ONLY canonical camelCase keys
            // (firstRow/lastRow/firstCol/lastCol/bandedRows/bandedCols).
            // Set still accepts lowercase aliases (firstrow/bandrow/etc)
            // as input — see Set.Element.cs. Internal hex `tblLook.val`
            // is NOT surfaced (was a dump-poisoning impl detail).
            var tblLookRead = tp.GetFirstChild<TableLook>();
            if (tblLookRead != null)
            {
                if (tblLookRead.FirstRow?.HasValue == true && tblLookRead.FirstRow.Value)
                    node.Format["firstRow"] = true;
                if (tblLookRead.LastRow?.HasValue == true && tblLookRead.LastRow.Value)
                    node.Format["lastRow"] = true;
                if (tblLookRead.FirstColumn?.HasValue == true && tblLookRead.FirstColumn.Value)
                    node.Format["firstCol"] = true;
                if (tblLookRead.LastColumn?.HasValue == true && tblLookRead.LastColumn.Value)
                    node.Format["lastCol"] = true;
                // banding semantics are inverted: noHBand=true means NO banding.
                // Emit only when banding IS active (noHBand=false explicitly set).
                if (tblLookRead.NoHorizontalBand?.HasValue == true && !tblLookRead.NoHorizontalBand.Value)
                    node.Format["bandedRows"] = true;
                if (tblLookRead.NoVerticalBand?.HasValue == true && !tblLookRead.NoVerticalBand.Value)
                    node.Format["bandedCols"] = true;
            }

            // Accessibility: table caption / description. Set writes
            // <w:tblCaption w:val="…"/> and <w:tblDescription w:val="…"/>
            // (see Set.Element.cs table branch). Without the readback,
            // get/dump silently drops these on round-trip.
            var tblCaption = tp.GetFirstChild<TableCaption>();
            if (!string.IsNullOrEmpty(tblCaption?.Val?.Value))
                node.Format["caption"] = tblCaption.Val.Value!;
            var tblDescription = tp.GetFirstChild<TableDescription>();
            if (!string.IsNullOrEmpty(tblDescription?.Val?.Value))
                node.Format["description"] = tblDescription.Val.Value!;
        }

        // Column widths from grid
        var gridCols = table.GetFirstChild<TableGrid>()?.Elements<GridColumn>().ToList();
        if (gridCols != null && gridCols.Count > 0)
            node.Format["colWidths"] = string.Join(",", gridCols.Select(g => (g.Width?.Value ?? "0") + "dxa"));

        if (depth > 0)
        {
            int rowIdx = 0;
            foreach (var row in table.Elements<TableRow>())
            {
                var rowNode = new DocumentNode
                {
                    Path = $"{path}/tr[{rowIdx + 1}]",
                    Type = "row",
                    ChildCount = row.Elements<TableCell>().Count()
                };
                ReadRowProps(row, rowNode);
                if (depth > 1)
                {
                    int cellIdx = 0;
                    foreach (var cell in row.Elements<TableCell>())
                    {
                        var cellNode = new DocumentNode
                        {
                            Path = $"{path}/tr[{rowIdx + 1}]/tc[{cellIdx + 1}]",
                            Type = "cell",
                            Text = string.Join("", cell.Descendants<Text>().Select(t => t.Text)),
                            // CONSISTENCY(cell-children): include nested Table children alongside Paragraphs.
                            ChildCount = cell.Elements<OpenXmlElement>().Count(e => e is Paragraph || e is Table)
                        };
                        ReadCellProps(cell, cellNode);
                        if (depth > 2)
                        {
                            int cellPIdx = 0, cellTblIdx = 0;
                            foreach (var cellChild in cell.Elements<OpenXmlElement>())
                            {
                                if (cellChild is Paragraph cellPara)
                                {
                                    cellPIdx++;
                                    var cParaSegment = BuildParaPathSegment(cellPara, cellPIdx);
                                    cellNode.Children.Add(ElementToNode(cellPara, $"{path}/tr[{rowIdx + 1}]/tc[{cellIdx + 1}]/{cParaSegment}", depth - 3));
                                }
                                else if (cellChild is Table cellTbl)
                                {
                                    cellTblIdx++;
                                    cellNode.Children.Add(ElementToNode(cellTbl, $"{path}/tr[{rowIdx + 1}]/tc[{cellIdx + 1}]/tbl[{cellTblIdx}]", depth - 3));
                                }
                            }
                        }
                        rowNode.Children.Add(cellNode);
                        cellIdx++;
                    }
                }
                node.Children.Add(rowNode);
                rowIdx++;
            }
        }
        return node;
    }

    private DocumentNode TableCellToNode(TableCell directCell, DocumentNode node, string path, int depth)
    {
        node.Type = "cell";
        node.Text = string.Join("", directCell.Descendants<Text>().Select(t => t.Text));
        // CONSISTENCY(cell-children): include nested Table children alongside Paragraphs.
        node.ChildCount = directCell.Elements<OpenXmlElement>().Count(e => e is Paragraph || e is Table);
        ReadCellProps(directCell, node);
        if (depth > 0)
        {
            int dcPIdx = 0, dcTblIdx = 0;
            foreach (var dcChild in directCell.Elements<OpenXmlElement>())
            {
                if (dcChild is Paragraph cellPara)
                {
                    dcPIdx++;
                    var dcParaSegment = BuildParaPathSegment(cellPara, dcPIdx);
                    node.Children.Add(ElementToNode(cellPara, $"{path}/{dcParaSegment}", depth - 1));
                }
                else if (dcChild is Table dcTbl)
                {
                    dcTblIdx++;
                    node.Children.Add(ElementToNode(dcTbl, $"{path}/tbl[{dcTblIdx}]", depth - 1));
                }
            }
        }
        return node;
    }

    private DocumentNode TableRowToNode(TableRow directRow, DocumentNode node, string path, int depth)
    {
        node.Type = "row";
        node.ChildCount = directRow.Elements<TableCell>().Count();
        ReadRowProps(directRow, node);
        if (depth > 0)
        {
            int cellIdx = 0;
            foreach (var cell in directRow.Elements<TableCell>())
            {
                var cellNode = new DocumentNode
                {
                    Path = $"{path}/tc[{cellIdx + 1}]",
                    Type = "cell",
                    Text = string.Join("", cell.Descendants<Text>().Select(t => t.Text)),
                    // CONSISTENCY(cell-children): include nested Table children alongside Paragraphs.
                    ChildCount = cell.Elements<OpenXmlElement>().Count(e => e is Paragraph || e is Table)
                };
                ReadCellProps(cell, cellNode);
                if (depth > 1)
                {
                    int drPIdx = 0, drTblIdx = 0;
                    foreach (var drChild in cell.Elements<OpenXmlElement>())
                    {
                        if (drChild is Paragraph cellPara)
                        {
                            drPIdx++;
                            var drParaSegment = BuildParaPathSegment(cellPara, drPIdx);
                            cellNode.Children.Add(ElementToNode(cellPara, $"{path}/tc[{cellIdx + 1}]/{drParaSegment}", depth - 2));
                        }
                        else if (drChild is Table drTbl)
                        {
                            drTblIdx++;
                            cellNode.Children.Add(ElementToNode(drTbl, $"{path}/tc[{cellIdx + 1}]/tbl[{drTblIdx}]", depth - 2));
                        }
                    }
                }
                node.Children.Add(cellNode);
                cellIdx++;
            }
        }
        return node;
    }

    private DocumentNode SdtBlockToNode(SdtBlock sdtBlockNode, DocumentNode node)
    {
        node.Type = "sdt";
        var sdtProps = sdtBlockNode.SdtProperties;
        if (sdtProps != null)
        {
            var alias = sdtProps.GetFirstChild<SdtAlias>();
            if (alias?.Val?.Value != null) node.Format["alias"] = alias.Val.Value;
            var tagEl = sdtProps.GetFirstChild<Tag>();
            if (tagEl?.Val?.Value != null) node.Format["tag"] = tagEl.Val.Value;
            var lockEl = sdtProps.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Lock>();
            if (lockEl?.Val?.Value != null) node.Format["lock"] = lockEl.Val.InnerText;
            var sdtId = sdtProps.GetFirstChild<SdtId>();
            if (sdtId?.Val?.Value != null) node.Format["id"] = sdtId.Val.Value;

            // Determine SDT type (check specific types first, text last as fallback)
            if (sdtProps.GetFirstChild<SdtContentDropDownList>() != null) node.Format["type"] = "dropdown";
            else if (sdtProps.GetFirstChild<SdtContentComboBox>() != null) node.Format["type"] = "combobox";
            else if (sdtProps.GetFirstChild<SdtContentDate>() != null) node.Format["type"] = "date";
            else if (sdtProps.GetFirstChild<SdtContentText>() != null) node.Format["type"] = "text";
            else node.Format["type"] = "richtext";

            // Read date format for date controls
            var dateContent = sdtProps.GetFirstChild<SdtContentDate>();
            if (dateContent?.DateFormat?.Val?.Value != null)
                node.Format["format"] = dateContent.DateFormat.Val.Value;

            // Editable status
            node.Format["editable"] = IsSdtEditable(sdtProps);

            // Placeholder detection
            var showingPlcHdr = sdtProps.GetFirstChild<ShowingPlaceholder>();
            if (showingPlcHdr != null)
            {
                node.Format["placeholder"] = true;
                var plcHdrText = sdtProps.GetFirstChild<SdtPlaceholder>()?.DocPartReference?.Val?.Value;
                if (plcHdrText != null) node.Format["placeholderText"] = plcHdrText;
            }

            // Read dropdown/combobox items
            var ddl = sdtProps.GetFirstChild<SdtContentDropDownList>();
            var combo = sdtProps.GetFirstChild<SdtContentComboBox>();
            var listItems = ddl?.Elements<ListItem>() ?? combo?.Elements<ListItem>();
            if (listItems != null)
            {
                // BUG-R5-07: SDT ListItems carry distinct DisplayText and
                // Value attrs. Real Word docs commonly differ (e.g.
                // "Draft|DRAFT"). Emit the pipe form when value !=
                // displayText so dump→add round-trips. ParseSdtItems on
                // the Add side accepts both bare and piped forms.
                var items = listItems.Select(li =>
                {
                    var disp = li.DisplayText?.Value ?? li.Value?.Value ?? "";
                    var val = li.Value?.Value ?? li.DisplayText?.Value ?? "";
                    return disp == val ? disp : $"{disp}|{val}";
                }).ToList();
                if (items.Count > 0) node.Format["items"] = string.Join(",", items);
            }
        }
        node.Text = string.Concat(sdtBlockNode.Descendants<Text>().Select(t => t.Text));
        var sdtContent = sdtBlockNode.SdtContentBlock;
        node.ChildCount = sdtContent?.ChildElements.Count ?? 0;
        return node;
    }

    private DocumentNode SdtRunToNode(SdtRun sdtRunNode, DocumentNode node)
    {
        node.Type = "sdt";
        var sdtProps = sdtRunNode.SdtProperties;
        if (sdtProps != null)
        {
            var alias = sdtProps.GetFirstChild<SdtAlias>();
            if (alias?.Val?.Value != null) node.Format["alias"] = alias.Val.Value;
            var tagEl = sdtProps.GetFirstChild<Tag>();
            if (tagEl?.Val?.Value != null) node.Format["tag"] = tagEl.Val.Value;
            var lockEl = sdtProps.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Lock>();
            if (lockEl?.Val?.Value != null) node.Format["lock"] = lockEl.Val.InnerText;
            var sdtId = sdtProps.GetFirstChild<SdtId>();
            if (sdtId?.Val?.Value != null) node.Format["id"] = sdtId.Val.Value;

            if (sdtProps.GetFirstChild<SdtContentDropDownList>() != null) node.Format["type"] = "dropdown";
            else if (sdtProps.GetFirstChild<SdtContentComboBox>() != null) node.Format["type"] = "combobox";
            else if (sdtProps.GetFirstChild<SdtContentDate>() != null) node.Format["type"] = "date";
            else if (sdtProps.GetFirstChild<SdtContentText>() != null) node.Format["type"] = "text";
            else node.Format["type"] = "richtext";

            // Editable status
            node.Format["editable"] = IsSdtEditable(sdtProps);

            // Placeholder detection
            var showingPlcHdrRun = sdtProps.GetFirstChild<ShowingPlaceholder>();
            if (showingPlcHdrRun != null)
            {
                node.Format["placeholder"] = true;
                var plcHdrTextRun = sdtProps.GetFirstChild<SdtPlaceholder>()?.DocPartReference?.Val?.Value;
                if (plcHdrTextRun != null) node.Format["placeholderText"] = plcHdrTextRun;
            }

            var ddl = sdtProps.GetFirstChild<SdtContentDropDownList>();
            var combo = sdtProps.GetFirstChild<SdtContentComboBox>();
            var listItems = ddl?.Elements<ListItem>() ?? combo?.Elements<ListItem>();
            if (listItems != null)
            {
                // BUG-R5-07: SDT ListItems carry distinct DisplayText and
                // Value attrs. Real Word docs commonly differ (e.g.
                // "Draft|DRAFT"). Emit the pipe form when value !=
                // displayText so dump→add round-trips. ParseSdtItems on
                // the Add side accepts both bare and piped forms.
                var items = listItems.Select(li =>
                {
                    var disp = li.DisplayText?.Value ?? li.Value?.Value ?? "";
                    var val = li.Value?.Value ?? li.DisplayText?.Value ?? "";
                    return disp == val ? disp : $"{disp}|{val}";
                }).ToList();
                if (items.Count > 0) node.Format["items"] = string.Join(",", items);
            }
        }
        node.Text = string.Concat(sdtRunNode.Descendants<Text>().Select(t => t.Text));
        return node;
    }

    private DocumentNode OfficeMathToNode(M.OfficeMath inlineMath, DocumentNode node)
    {
        node.Type = "equation";
        node.Format["mode"] = "inline";
        try { node.Text = Core.FormulaParser.ToLatex(inlineMath); }
        catch { node.Text = inlineMath.InnerText; }
        if (string.IsNullOrEmpty(node.Text))
            node.Text = inlineMath.InnerText;
        return node;
    }

    private DocumentNode HeaderFooterToNode(OpenXmlElement element, DocumentNode node, string path, int depth)
    {
        // Header/Footer: enumerate block-level children. Tables are valid
        // block-level OOXML inside hdr/ftr (same schema as body), so list
        // them alongside paragraphs. Mirrors body-listing logic above.
        node.Type = element is Header ? "header" : "footer";
        node.Text = string.Concat(element.Descendants<Text>().Select(t => t.Text));
        node.ChildCount = element.Elements<Paragraph>().Count() + element.Elements<Table>().Count();
        if (depth > 0)
        {
            int pIdx = 0, tblIdx = 0;
            foreach (var child in element.ChildElements)
            {
                if (child is Paragraph hfPara)
                {
                    pIdx++;
                    var paraSegment = BuildParaPathSegment(hfPara, pIdx);
                    node.Children.Add(ElementToNode(hfPara, $"{path}/{paraSegment}", depth - 1));
                }
                else if (child is Table)
                {
                    tblIdx++;
                    node.Children.Add(ElementToNode(child, $"{path}/tbl[{tblIdx}]", depth - 1));
                }
            }
        }
        return node;
    }

    private DocumentNode BodyToNode(Body bodyNode, DocumentNode node, string path, int depth)
    {
        // CONSISTENCY(body-listing): enumerate body children using the
        // same p[N]/oMathPara[M] counting rules as NavigateToElement so
        // `get /body` emits paths that `get <path>` can resolve. The
        // generic fallback would count every LocalName, listing wrapper
        // <w:p> (pure oMathPara) as p[2] even though the resolver skips
        // them. Mirrors the logic in WordHandler.View.ViewAsText.
        node.ChildCount = bodyNode.ChildElements.Count;
        if (depth > 0)
        {
            int pIdx = 0, tblIdx = 0, mathParaIdx = 0, sdtIdx = 0;
            // BUG-DUMP7-04: w:customXml body wrappers are non-structural —
            // their inner paragraphs and tables should appear as direct
            // body children (with shared p/tbl/sdt counters) so the
            // wrapper itself is invisible to dump but its content
            // round-trips. Recursively flatten any depth of customXml
            // nesting. Without this, the wrapper fell to the generic
            // else and its children were never enumerated.
            void WalkBodyChild(OpenXmlElement child)
            {
                if (child.LocalName == "oMathPara" || child is M.Paragraph)
                {
                    mathParaIdx++;
                    node.Children.Add(ElementToNode(child, $"{path}/oMathPara[{mathParaIdx}]", depth - 1));
                }
                else if (child is Paragraph bPara)
                {
                    if (IsOMathParaWrapperParagraph(bPara))
                    {
                        mathParaIdx++;
                        node.Children.Add(ElementToNode(bPara, $"{path}/oMathPara[{mathParaIdx}]", depth - 1));
                    }
                    else
                    {
                        pIdx++;
                        var bSeg = BuildParaPathSegment(bPara, pIdx);
                        node.Children.Add(ElementToNode(bPara, $"{path}/{bSeg}", depth - 1));
                    }
                }
                else if (child is Table)
                {
                    tblIdx++;
                    node.Children.Add(ElementToNode(child, $"{path}/tbl[{tblIdx}]", depth - 1));
                }
                else if (child is SdtBlock)
                {
                    sdtIdx++;
                    node.Children.Add(ElementToNode(child, $"{path}/sdt[{sdtIdx}]", depth - 1));
                }
                else if (child is CustomXmlBlock cxBlock)
                {
                    foreach (var inner in cxBlock.ChildElements)
                        WalkBodyChild(inner);
                }
                else
                {
                    // Non-structural (sectPr etc.) — keep localName naming
                    node.Children.Add(ElementToNode(child, $"{path}/{child.LocalName}[1]", depth - 1));
                }
            }
            foreach (var child in bodyNode.ChildElements)
                WalkBodyChild(child);
        }
        return node;
    }

    private DocumentNode ParagraphToNode(Paragraph para, DocumentNode node, string path, int depth)
    {
        node.Type = "paragraph";
        node.Text = GetParagraphText(para);
        node.Style = GetStyleName(para);
        node.Preview = node.Text?.Length > 50 ? node.Text[..50] + "..." : node.Text;
        node.ChildCount = GetAllRuns(para).Count();

        if (!string.IsNullOrEmpty(para.ParagraphId?.Value))
            node.Format["paraId"] = para.ParagraphId.Value;
        // textId intentionally NOT exposed in Format: Set() rewrites it on
        // every mutation (see WordHandler.Set.cs "para.TextId = GenerateParaId()"),
        // which would let an AI agent comparing consecutive Get snapshots see
        // spurious diffs and mistake idempotent edits for real changes. paraId
        // is stable and sufficient for identity. The underlying w14:textId
        // attribute is still present in the OOXML; only the user-facing
        // DocumentNode.Format projection hides it.

        var pProps = para.ParagraphProperties;
        // AddParagraph writes <w:pPrChange> for `trackChange=format`. The
        // pPrChange block carries author/date attribution alongside a
        // baseline snapshot of the pre-format pPr — mirror what the run
        // side does for <w:rPrChange>.
        var pPrChange = pProps?.GetFirstChild<ParagraphPropertiesChange>();
        if (pPrChange != null)
        {
            node.Format["revision.type"] = "format";
            if (!string.IsNullOrEmpty(pPrChange.Author?.Value))
                node.Format["revision.author"] = pPrChange.Author!.Value!;
            if (pPrChange.Date?.Value is DateTime pDate)
                node.Format["revision.date"] = pDate.ToString("o");
            if (pPrChange.Id?.Value is { } pcId)
                node.Format["revision.id"] = pcId.ToString();
            // BUG: pPrChange's previous-pPr snapshot (alignment, spacing,
            // indent, …) is what Word's Reject Change restores. Add v1
            // re-stamps the marker with empty previous-pPr, losing the
            // snapshot on dump→batch round-trip. SDK 3.x writes the snapshot
            // as ParagraphPropertiesExtended (SetRevision path) or
            // PreviousParagraphProperties (AddParagraph's empty-marker path /
            // foreign producers). Probe both.
            var prevExt = pPrChange.GetFirstChild<ParagraphPropertiesExtended>();
            var prevPpr = pPrChange.GetFirstChild<PreviousParagraphProperties>();
            bool prevPHas = (prevExt != null && prevExt.HasChildren)
                         || (prevPpr != null && prevPpr.HasChildren);
            if (prevPHas)
                node.Format["revision.beforeLost"] = true;
        }
        // paraMarkIns: `<w:pPr><w:rPr><w:ins .../></w:rPr></w:pPr>` records
        // that the paragraph mark itself was inserted as a tracked change —
        // distinct from pPrChange (format-change snapshot) and from any
        // content-run wrappers. Surfaced under a paraMarkIns.* namespace so
        // dump can translate it back into AddParagraph's bare-trackChange.*
        // form on replay (which re-creates both the ¶ mark and any
        // accompanying content wrapping in one step). Kept distinct from
        // `trackChange=format` to avoid clobbering pPrChange attribution
        // when both are present on the same paragraph.
        var pmrpRev = pProps?.ParagraphMarkRunProperties;
        if (pmrpRev != null)
        {
            var pMarkIns = pmrpRev.GetFirstChild<Inserted>();
            if (pMarkIns != null)
            {
                if (!string.IsNullOrEmpty(pMarkIns.Author?.Value))
                    node.Format["paraMarkIns.author"] = pMarkIns.Author!.Value!;
                if (pMarkIns.Date?.Value is DateTime piDate)
                    node.Format["paraMarkIns.date"] = piDate.ToString("o");
                if (pMarkIns.Id?.Value is { } piId)
                    node.Format["paraMarkIns.id"] = piId.ToString();
            }
            // paraMarkDel: <w:pPr><w:rPr><w:del .../></w:rPr></w:pPr> records
            // that the paragraph mark was *deleted* as a tracked change
            // (paragraph-join revision). Mirror paraMarkIns so dump round-trip
            // preserves it.
            var pMarkDel = pmrpRev.GetFirstChild<Deleted>();
            if (pMarkDel != null)
            {
                if (!string.IsNullOrEmpty(pMarkDel.Author?.Value))
                    node.Format["paraMarkDel.author"] = pMarkDel.Author!.Value!;
                if (pMarkDel.Date?.Value is DateTime pdDate)
                    node.Format["paraMarkDel.date"] = pdDate.ToString("o");
                if (pMarkDel.Id?.Value is { } pdId)
                    node.Format["paraMarkDel.id"] = pdId.ToString();
            }
        }
        if (pProps != null)
        {
            if (pProps.ParagraphStyleId?.Val?.Value != null)
            {
                // CONSISTENCY(style-dual-key): `style` carries the OOXML
                // styleId (canonical handle used by basedOn/pStyle/rStyle).
                // `styleName` carries the user-facing display name. Both
                // are emitted so query selectors can pick precision
                // (styleId=/styleName=) or convenience (style=, lenient).
                node.Format["style"] = pProps.ParagraphStyleId.Val.Value;
                node.Format["styleId"] = pProps.ParagraphStyleId.Val.Value;
                var displayName = GetStyleName(para);
                if (!string.IsNullOrEmpty(displayName))
                    node.Format["styleName"] = displayName;
            }
            if (pProps.Justification?.Val != null)
            {
                var alignText = pProps.Justification.Val.InnerText;
                var alignValue = alignText == "both" ? "justify" : alignText;
                node.Format["align"] = alignValue;
            }
            if (pProps.SpacingBetweenLines != null)
            {
                if (pProps.SpacingBetweenLines.Before?.Value != null)
                {
                    node.Format["spaceBefore"] = SpacingConverter.FormatWordSpacing(pProps.SpacingBetweenLines.Before.Value);
                }
                if (pProps.SpacingBetweenLines.After?.Value != null)
                {
                    node.Format["spaceAfter"] = SpacingConverter.FormatWordSpacing(pProps.SpacingBetweenLines.After.Value);
                }
                if (pProps.SpacingBetweenLines.Line?.Value != null)
                {
                    node.Format["lineSpacing"] = SpacingConverter.FormatWordLineSpacing(
                        pProps.SpacingBetweenLines.Line.Value,
                        pProps.SpacingBetweenLines.LineRule?.InnerText);
                }
                if (pProps.SpacingBetweenLines.LineRule?.HasValue == true)
                {
                    node.Format["lineRule"] = pProps.SpacingBetweenLines.LineRule.InnerText;
                }
                // CONSISTENCY(ind-chars): mirror style-level Get (Query.cs)
                // for the chars-unit space-before/after slots so P1-7
                // round-trip works on paragraphs as well as styles.
                if (pProps.SpacingBetweenLines.BeforeLines?.Value != null)
                {
                    node.Format["spaceBeforeLines"] = pProps.SpacingBetweenLines.BeforeLines.Value;
                }
                if (pProps.SpacingBetweenLines.AfterLines?.Value != null)
                {
                    node.Format["spaceAfterLines"] = pProps.SpacingBetweenLines.AfterLines.Value;
                }
            }
            if (pProps.Indentation != null)
            {
                var ind = pProps.Indentation;
                // CONSISTENCY(unit-qualified-spacing): indents return "Xpt" via SpacingConverter,
                // matching spaceBefore/spaceAfter (Canonical DocumentNode.Format Rules).
                if (ind.FirstLine?.Value != null) node.Format["firstLineIndent"] = SpacingConverter.FormatWordSpacing(ind.FirstLine.Value);
                if (ind.Hanging?.Value != null) node.Format["hangingIndent"] = SpacingConverter.FormatWordSpacing(ind.Hanging.Value);
                // CONSISTENCY(ind-start-end): modern Word writes <w:ind w:start>/<w:end> instead of left/right.
                var leftTwips = ind.Left?.Value ?? ind.Start?.Value;
                if (leftTwips != null) node.Format["indent"] = SpacingConverter.FormatWordSpacing(leftTwips);
                var rightTwips = ind.Right?.Value ?? ind.End?.Value;
                if (rightTwips != null) node.Format["rightIndent"] = SpacingConverter.FormatWordSpacing(rightTwips);
                // CONSISTENCY(ind-chars): chars-unit indents (Chinese typography) — backfilled from style Get edc8f884.
                if (ind.FirstLineChars?.Value != null) node.Format["firstLineChars"] = ind.FirstLineChars.Value;
                if (ind.HangingChars?.Value != null) node.Format["hangingChars"] = ind.HangingChars.Value;
                var leftChars = ind.LeftChars?.Value ?? ind.StartCharacters?.Value;
                if (leftChars != null) node.Format["leftChars"] = leftChars;
                var rightChars = ind.RightChars?.Value ?? ind.EndCharacters?.Value;
                if (rightChars != null) node.Format["rightChars"] = rightChars;
            }
            if (pProps.KeepNext != null)
            {
                var v = pProps.KeepNext.Val;
                node.Format["keepNext"] = v == null || v.Value;
            }
            if (pProps.KeepLines != null)
            {
                var v = pProps.KeepLines.Val;
                node.Format["keepLines"] = v == null || v.Value;
            }
            if (pProps.PageBreakBefore != null)
            {
                var v = pProps.PageBreakBefore.Val;
                node.Format["pageBreakBefore"] = v == null || v.Value;
            }
            if (pProps.WidowControl != null)
            {
                // Val == null or Val == true means enabled; Val == false means explicitly disabled
                var wcVal = pProps.WidowControl.Val;
                node.Format["widowControl"] = wcVal == null || wcVal.Value;
            }
            if (pProps.BiDi != null)
            {
                // <w:bidi/> default Val is true; explicit Val=false toggles
                // it off. Emit canonical 'direction' so writers can clone
                // the paragraph with the same key they used to set it.
                // R8-fuzz-5: pProps.BiDi.Val.Value invokes OnOffValue.Parse
                // and throws FormatException on garbage attribute text
                // (e.g. <w:bidi w:val="garbage"/>). Skip the key on
                // unparseable input — Get must never crash on a doc that
                // disk-loaded fine, even when validate would flag the same
                // attribute as schema-invalid.
                bool? bidiOn = TryReadOnOff(pProps.BiDi.Val);
                if (bidiOn.HasValue)
                    node.Format["direction"] = bidiOn.Value ? "rtl" : "ltr";
            }
            if (pProps.ContextualSpacing != null)
            {
                var csVal = pProps.ContextualSpacing.Val;
                node.Format["contextualSpacing"] = csVal == null || csVal.Value;
            }
            if (pProps.Shading != null)
            {
                // CONSISTENCY(canonical-keys): split shading into shading.val/.fill/.color sub-keys
                // matching the OOXML attribute structure. No compound semicolon string.
                var shdVal = pProps.Shading.Val?.InnerText;
                var shdFill = pProps.Shading.Fill?.Value;
                var shdColor = pProps.Shading.Color?.Value;
                if (!string.IsNullOrEmpty(shdVal)) node.Format["shading.val"] = shdVal;
                if (!string.IsNullOrEmpty(shdFill)) node.Format["shading.fill"] = ParseHelpers.FormatHexColor(shdFill);
                if (!string.IsNullOrEmpty(shdColor)) node.Format["shading.color"] = ParseHelpers.FormatHexColor(shdColor);
            }

            var pBdr = pProps.ParagraphBorders;
            if (pBdr != null)
            {
                ReadBorder(pBdr.TopBorder, "pbdr.top", node);
                ReadBorder(pBdr.BottomBorder, "pbdr.bottom", node);
                ReadBorder(pBdr.LeftBorder, "pbdr.left", node);
                ReadBorder(pBdr.RightBorder, "pbdr.right", node);
                ReadBorder(pBdr.BetweenBorder, "pbdr.between", node);
                ReadBorder(pBdr.BarBorder, "pbdr.bar", node);
            }

            var numProps = pProps.NumberingProperties;
            if (numProps != null && numProps.NumberingId?.Val?.Value != null)
            {
                var numIdVal = numProps.NumberingId.Val.Value;
                node.Format["numId"] = numIdVal.ToString();
                var ilvlVal = numProps.NumberingLevelReference?.Val?.Value ?? 0;
                // R29-3: surface under the canonical key 'numLevel' (paragraph.json
                // declares numLevel canonical with ilvl as an input alias; Get
                // normalizes to the single canonical key). Style/abstractNum-level
                // contexts keep 'ilvl' as their own canonical key.
                node.Format["numLevel"] = ilvlVal.ToString();
                // numId=0 is the OOXML "remove numbering" sentinel — the paragraph
                // explicitly opts out of any inherited list style. Skip numFmt /
                // listStyle / start lookup so Get does not falsely advertise a list.
                if (numIdVal != 0)
                {
                    var numFmt = GetNumberingFormat(numIdVal, ilvlVal);
                    node.Format["numFmt"] = numFmt;
                    node.Format["listStyle"] = numFmt.ToLowerInvariant() == "bullet" ? "bullet" : "ordered";
                    var start = GetStartValue(numIdVal, ilvlVal);
                    if (start != null)
                        node.Format["start"] = start.Value;
                }
            }
            else
            {
                // Fall back to the style chain — paragraphs that inherit numbering
                // from styles like ListBullet / ListNumber don't have a direct numPr,
                // but Get should still surface the effective list metadata.
                var inherited = ResolveNumPrFromStyle(para);
                if (inherited.HasValue)
                {
                    var (inhId, inhLvl) = inherited.Value;
                    node.Format["numId"] = inhId.ToString();
                    // R29-3: canonical key 'numLevel' (see direct-numPr branch above).
                    node.Format["numLevel"] = inhLvl.ToString();
                    // BUG-DUMP26-01: flag style-inherited values so WordBatchEmitter
                    // can suppress them on `add p` — they're already covered by
                    // the paragraph's style and emitting them would semantically
                    // promote inherited→explicit on round-trip. Mirrors the
                    // round-1 first-run hoist precedent.
                    node.Format["numInherited"] = "true";
                    var numFmt = GetNumberingFormat(inhId, inhLvl);
                    node.Format["numFmt"] = numFmt;
                    node.Format["listStyle"] = numFmt.ToLowerInvariant() == "bullet" ? "bullet" : "ordered";
                    var start = GetStartValue(inhId, inhLvl);
                    if (start != null)
                        node.Format["start"] = start.Value;
                }
            }

            // CONSISTENCY(outline-lvl): backfilled from style Get edc8f884. Paragraph-level outlineLvl overrides style.
            if (pProps.OutlineLevel?.Val?.Value != null)
                node.Format["outlineLvl"] = (int)pProps.OutlineLevel.Val.Value;

            // CONSISTENCY(tabs): backfilled from style Get edc8f884.
            if (pProps.Tabs != null)
            {
                var tabList = new List<Dictionary<string, object?>>();
                foreach (var tab in pProps.Tabs.Elements<TabStop>())
                {
                    var t = new Dictionary<string, object?>();
                    if (tab.Position?.Value != null) t["pos"] = tab.Position.Value;
                    if (tab.Val?.HasValue == true) t["val"] = tab.Val.InnerText;
                    if (tab.Leader?.HasValue == true) t["leader"] = tab.Leader.InnerText;
                    if (t.Count > 0) tabList.Add(t);
                }
                if (tabList.Count > 0) node.Format["tabs"] = tabList;
            }

            // Long-tail fallback: surface every pPr child the curated reader
            // didn't consume. Symmetric with the Set-side TryCreateTypedChild
            // fallback in SetElementParagraph (WordHandler.Set.Element.cs).
            FillUnknownChildProps(pProps, node);

            // CONSISTENCY(add-set-symmetry): inline section break.
            // A paragraph carrying <w:sectPr> inside its <w:pPr> is the
            // OOXML representation of a mid-document section break (the
            // last paragraph before the break holds the section's
            // properties). AddSection on /body produces exactly this
            // shape, but Get used to expose nothing — leaving the
            // paragraph indistinguishable from a regular empty para.
            // Surface it as `sectionBreak` (Add prop name match) plus
            // companion section-property keys readers expect.
            var inlineSectPr = pProps.GetFirstChild<SectionProperties>();
            if (inlineSectPr != null)
            {
                var sectMark = inlineSectPr.GetFirstChild<SectionType>()?.Val?.InnerText;
                node.Format["sectionBreak"] = sectMark ?? "nextPage";

                // Per-section page layout when overridden on this break.
                var pgSz = inlineSectPr.GetFirstChild<PageSize>();
                if (pgSz?.Width?.Value != null)
                    node.Format["sectionBreak.pageWidth"] = FormatTwipsToCm(pgSz.Width.Value);
                if (pgSz?.Height?.Value != null)
                    node.Format["sectionBreak.pageHeight"] = FormatTwipsToCm(pgSz.Height.Value);
                if (pgSz?.Orient?.Value != null)
                    node.Format["sectionBreak.orientation"] = pgSz.Orient.InnerText;

                var pgMar = inlineSectPr.GetFirstChild<PageMargin>();
                if (pgMar != null)
                {
                    if (pgMar.Top?.Value != null)
                        node.Format["sectionBreak.marginTop"] = FormatTwipsToCm((uint)Math.Abs(pgMar.Top.Value));
                    if (pgMar.Bottom?.Value != null)
                        node.Format["sectionBreak.marginBottom"] = FormatTwipsToCm((uint)Math.Abs(pgMar.Bottom.Value));
                    if (pgMar.Left?.Value != null)
                        node.Format["sectionBreak.marginLeft"] = FormatTwipsToCm(pgMar.Left.Value);
                    if (pgMar.Right?.Value != null)
                        node.Format["sectionBreak.marginRight"] = FormatTwipsToCm(pgMar.Right.Value);
                    // header/footer-from-edge + binding gutter (mirror the root
                    // and /section[N] readbacks) so a mid-document section
                    // break round-trips its full pgMar, not just the 4 edges.
                    if (pgMar.Header?.Value != null)
                        node.Format["sectionBreak.marginHeader"] = FormatTwipsToCm(pgMar.Header.Value);
                    if (pgMar.Footer?.Value != null)
                        node.Format["sectionBreak.marginFooter"] = FormatTwipsToCm(pgMar.Footer.Value);
                    if (pgMar.Gutter?.Value != null)
                        node.Format["sectionBreak.marginGutter"] = FormatTwipsToCm(pgMar.Gutter.Value);
                }

                var pgNum = inlineSectPr.GetFirstChild<PageNumberType>();
                if (pgNum?.Start?.Value != null)
                    node.Format["sectionBreak.pageStart"] = pgNum.Start.Value;
                if (pgNum?.Format?.Value != null)
                    node.Format["sectionBreak.pageNumFmt"] = pgNum.Format.InnerText;

                if (inlineSectPr.GetFirstChild<TitlePage>() != null)
                    node.Format["sectionBreak.titlePage"] = true;

                // BUG-DUMP9-06: Columns / VerticalTextAlignmentOnPage on
                // an inline sectPr carrier were silently dropped — only
                // the root sectPr reader handled them. Surface as
                // sectionBreak.columns / sectionBreak.vAlign so dump
                // round-trips the carrier sectPr.
                var sbCols = inlineSectPr.GetFirstChild<Columns>();
                if (sbCols != null)
                {
                    if (sbCols.ColumnCount?.Value != null)
                        node.Format["sectionBreak.columns"] = (int)sbCols.ColumnCount.Value;
                    if (sbCols.Space?.Value != null && uint.TryParse(sbCols.Space.Value, out var sbColSpaceTwips))
                        node.Format["sectionBreak.columnSpace"] = FormatTwipsToCm(sbColSpaceTwips);
                    if (sbCols.EqualWidth?.Value != null)
                        node.Format["sectionBreak.columns.equalWidth"] = sbCols.EqualWidth.Value;
                    if (sbCols.Separator?.Value == true)
                        node.Format["sectionBreak.columns.separator"] = true;
                }

                var sbVAlign = inlineSectPr.GetFirstChild<VerticalTextAlignmentOnPage>();
                if (sbVAlign?.Val != null)
                    node.Format["sectionBreak.vAlign"] = sbVAlign.Val.InnerText;

                var lnNum = inlineSectPr.GetFirstChild<LineNumberType>();
                if (lnNum != null)
                {
                    node.Format["sectionBreak.lineNumbers"] = lnNum.Restart?.InnerText switch
                    {
                        "newPage" => "restartPage",
                        "newSection" => "restartSection",
                        _ => "continuous"
                    };
                    if (lnNum.CountBy?.Value is short cb && cb > 1)
                        node.Format["sectionBreak.lineNumberCountBy"] = cb;
                }

                // BUG-DUMP-SECGRID: the document grid (<w:docGrid>) on a
                // mid-document section break governs CJK line pitch / lines-
                // per-page; dropping it on dump→batch reflowed every page of
                // that section (a major pagination-drift source). Mirror the
                // root / section readback (docGrid.type / .linePitch / .charSpace);
                // AddSection replays them via its docGrid.* typed fallback.
                var sbGrid = inlineSectPr.GetFirstChild<DocGrid>();
                if (sbGrid != null)
                {
                    if (sbGrid.Type?.HasValue == true)
                        node.Format["sectionBreak.docGrid.type"] = sbGrid.Type.InnerText;
                    if (sbGrid.LinePitch?.Value != null)
                        node.Format["sectionBreak.docGrid.linePitch"] = sbGrid.LinePitch.Value;
                    if (sbGrid.CharacterSpace?.Value != null)
                        node.Format["sectionBreak.docGrid.charSpace"] = sbGrid.CharacterSpace.Value;
                }
            }
        }

        // BUG-DUMP9-02: surface paragraph-mark-only run formatting under
        // the `markRPr.*` namespace whenever pPr/rPr exists. The
        // run-fallback path below promotes mark rPr to bare keys only
        // when there are no runs (round-1 hoisting fix); when runs are
        // present, mark-only formatting on the ¶ glyph used to be
        // silently dropped on dump round-trip. Emit dedicated keys so
        // replay can target ParagraphMarkRunProperties without conflating
        // with run-level formatting.
        var pmrpForDump = para.ParagraphProperties?.ParagraphMarkRunProperties;
        // Suppress markRPr.* dotted keys when the paragraph has no
        // text-bearing runs — the bare keys below (size, font.latin, …)
        // already cover markRPr via the firstRun-fallback path. Emitting
        // both forms on an empty paragraph means dump→batch→dump
        // surfaces phantom markRPr.* keys even after AddParagraph
        // routed the formatting correctly (BUG-DUMP-MARKRPR-DOUBLE).
        // The dotted form's purpose is to distinguish the ¶ glyph's
        // formatting from the visible text — only meaningful when text
        // runs exist.
        var hasTextRun = para.Elements<Run>()
            .Any(r => r.GetFirstChild<Text>() != null
                      && !string.IsNullOrEmpty(r.GetFirstChild<Text>()?.Text));
        if (pmrpForDump != null && hasTextRun)
        {
            var b = pmrpForDump.GetFirstChild<Bold>();
            if (b != null) node.Format["markRPr.bold"] = IsToggleOn(b);
            var i = pmrpForDump.GetFirstChild<Italic>();
            if (i != null) node.Format["markRPr.italic"] = IsToggleOn(i);
            var s = pmrpForDump.GetFirstChild<Strike>();
            if (s != null) node.Format["markRPr.strike"] = IsToggleOn(s);
            var u = pmrpForDump.GetFirstChild<Underline>();
            if (u?.Val?.HasValue == true) node.Format["markRPr.underline"] = u.Val.InnerText;
            var fs = pmrpForDump.GetFirstChild<FontSize>();
            if (fs?.Val?.Value != null)
                node.Format["markRPr.size"] = $"{int.Parse(fs.Val.Value) / 2.0:0.##}pt";
            // BUG-DUMP-MARKRPR-CS: the ¶-mark's complex-script + kern slots
            // (szCs / bCs / iCs / kern) were never emitted under markRPr.*, so
            // dump→batch silently dropped them on every paragraph that has text
            // runs (the empty-paragraph fallback path already reads them via
            // ReadComplexScriptRunFormatting). Mirror that readback here so the
            // ¶ glyph's CJK metrics survive the round-trip. ApplyRunFormatting
            // consumes size.cs / bold.cs / italic.cs / kern on replay.
            var fsCs = pmrpForDump.GetFirstChild<FontSizeComplexScript>();
            if (fsCs?.Val?.Value is string fsCsVal && int.TryParse(fsCsVal, out var fsCsHp))
                node.Format["markRPr.size.cs"] = $"{fsCsHp / 2.0:0.##}pt";
            var bCsMark = pmrpForDump.GetFirstChild<BoldComplexScript>();
            if (bCsMark != null && (bCsMark.Val == null || bCsMark.Val.Value))
                node.Format["markRPr.bold.cs"] = true;
            var iCsMark = pmrpForDump.GetFirstChild<ItalicComplexScript>();
            if (iCsMark != null && (iCsMark.Val == null || iCsMark.Val.Value))
                node.Format["markRPr.italic.cs"] = true;
            var kernMark = pmrpForDump.GetFirstChild<Kern>();
            if (kernMark?.Val?.HasValue == true)
                node.Format["markRPr.kern"] = kernMark.Val.Value.ToString();
            var clr = pmrpForDump.GetFirstChild<Color>();
            if (clr != null)
            {
                if (clr.ThemeColor?.HasValue == true)
                    node.Format["markRPr.color"] = clr.ThemeColor.InnerText;
                else if (clr.Val?.Value != null)
                    node.Format["markRPr.color"] = ParseHelpers.FormatHexColor(clr.Val.Value);
            }
            var rf = pmrpForDump.GetFirstChild<RunFonts>();
            // BUG-DUMP-MARKRPR-HANSI / BUG-DUMP-FONT-LATIN: collapse to
            // font.latin only when BOTH slots are present and equal. Divergent
            // slots emit both; ascii-only emits font.ascii and hAnsi-only
            // emits font.hAnsi — NEVER font.latin for a single slot, since
            // font.latin sets both on replay and would silently rebind the ¶
            // mark's missing slot (extended-Latin falls back to the
            // style/docDefaults font when hAnsi is absent).
            var markAscii = rf?.Ascii?.Value;
            var markHAnsi = rf?.HighAnsi?.Value;
            if (markAscii != null && markHAnsi != null)
            {
                if (markAscii == markHAnsi)
                    node.Format["markRPr.font.latin"] = markAscii;
                else
                {
                    node.Format["markRPr.font.ascii"] = markAscii;
                    node.Format["markRPr.font.hAnsi"] = markHAnsi;
                }
            }
            else if (markAscii != null)
                node.Format["markRPr.font.ascii"] = markAscii;
            else if (markHAnsi != null)
                node.Format["markRPr.font.hAnsi"] = markHAnsi;
            if (rf?.EastAsia?.Value != null)
                node.Format["markRPr.font.ea"] = rf.EastAsia.Value;
            if (rf?.ComplexScript?.Value != null)
                node.Format["markRPr.font.cs"] = rf.ComplexScript.Value;
            // ¶-mark font hint + character spacing (mirror the run-level
            // font.hint / charSpacing readback so the paragraph mark's glyph
            // properties round-trip too — see RunToNode).
            if (rf?.Hint?.HasValue == true)
                node.Format["markRPr.font.hint"] = rf.Hint.InnerText;
            var pmSpacing = pmrpForDump.GetFirstChild<Spacing>();
            if (pmSpacing?.Val?.HasValue == true)
                node.Format["markRPr.charSpacing"] = $"{pmSpacing.Val.Value / 20.0:0.##}pt";
            var hl = pmrpForDump.GetFirstChild<Highlight>();
            if (hl?.Val?.HasValue == true) node.Format["markRPr.highlight"] = hl.Val.InnerText;
            // schemas/help/docx/paragraph.json declares rStyle add+set+get;
            // Add.Text.cs:437 writes <w:rStyle> into ParagraphMarkRunProperties,
            // but Get used to drop it. Emit at the paragraph-level canonical
            // key (no markRPr prefix) to match the schema's declaration.
            var rs = pmrpForDump.GetFirstChild<RunStyle>();
            if (rs?.Val?.Value != null) node.Format["rStyle"] = rs.Val.Value;
        }

        // First-run formatting on the paragraph node (like PPTX does for shapes).
        // Fall back to ParagraphMarkRunProperties when no runs exist (e.g. empty paragraph
        // that had formatting applied via Set before any text was added).
        var firstRun = para.Elements<Run>().FirstOrDefault(r => r.GetFirstChild<Text>() != null);
        var paraRp = firstRun?.RunProperties
            ?? (firstRun == null ? para.ParagraphProperties?.ParagraphMarkRunProperties as OpenXmlCompositeElement : null);
        if (paraRp != null)
        {
            RunProperties? rp = paraRp as RunProperties ?? null;
            ParagraphMarkRunProperties? markRp = paraRp as ParagraphMarkRunProperties ?? null;

            // CONSISTENCY(canonical-keys): mirror style Get (WordHandler.Query.cs:546-553) —
            // emit per-script font slots, no flat "font" alias. R6 BUG-1: previously only
            // emitted Ascii under "font" key, dropping eastAsia/hAnsi/cs slots.
            var pRunFonts = rp?.RunFonts ?? markRp?.GetFirstChild<RunFonts>();
            if (pRunFonts != null)
            {
                // CONSISTENCY(canonical-keys): schema (docx/run.json,
                // docx/paragraph.json) declares `font.latin` and `font.ea`
                // as canonical. Collapse Ascii+HighAnsi to `font.latin`
                // when they match (the round-trip case for `font.latin=`
                // Set). When they differ, emit both legacy slots so no
                // information is lost.
                var ascii = pRunFonts.Ascii?.Value;
                var hAnsi = pRunFonts.HighAnsi?.Value;
                if (ascii != null && hAnsi != null && ascii == hAnsi)
                {
                    if (!node.Format.ContainsKey("font.latin"))
                        node.Format["font.latin"] = ascii;
                }
                else if (ascii != null && hAnsi != null)
                {
                    // Two slots, divergent values — fall back to legacy keys.
                    if (!node.Format.ContainsKey("font.ascii"))
                        node.Format["font.ascii"] = ascii;
                    if (!node.Format.ContainsKey("font.hAnsi"))
                        node.Format["font.hAnsi"] = hAnsi;
                }
                // BUG-DUMP-FONT-LATIN: ascii-only → font.ascii, hAnsi-only →
                // font.hAnsi (NOT font.latin, which sets both on replay and
                // would rebind the absent slot).
                else if (ascii != null)
                {
                    if (!node.Format.ContainsKey("font.ascii"))
                        node.Format["font.ascii"] = ascii;
                }
                else if (hAnsi != null)
                {
                    if (!node.Format.ContainsKey("font.hAnsi"))
                        node.Format["font.hAnsi"] = hAnsi;
                }
                if (!string.IsNullOrEmpty(pRunFonts.EastAsia?.Value) && !node.Format.ContainsKey("font.ea"))
                    node.Format["font.ea"] = pRunFonts.EastAsia!.Value!;
                // BUG-DUMP15-03: surface theme-font slots on the paragraph
                // node (leaked from first run rPr) so dump→batch round-trip
                // preserves theme bindings. Mirrors the run-level readback
                // at the typed-Run branch below.
                if (pRunFonts.AsciiTheme?.HasValue == true && !node.Format.ContainsKey("font.asciiTheme"))
                    node.Format["font.asciiTheme"] = pRunFonts.AsciiTheme.InnerText;
                if (pRunFonts.HighAnsiTheme?.HasValue == true && !node.Format.ContainsKey("font.hAnsiTheme"))
                    node.Format["font.hAnsiTheme"] = pRunFonts.HighAnsiTheme.InnerText;
                if (pRunFonts.EastAsiaTheme?.HasValue == true && !node.Format.ContainsKey("font.eaTheme"))
                    node.Format["font.eaTheme"] = pRunFonts.EastAsiaTheme.InnerText;
                if (pRunFonts.ComplexScriptTheme?.HasValue == true && !node.Format.ContainsKey("font.csTheme"))
                    node.Format["font.csTheme"] = pRunFonts.ComplexScriptTheme.InnerText;
            }

            var fsVal = rp?.FontSize?.Val?.Value ?? markRp?.GetFirstChild<FontSize>()?.Val?.Value;
            if (fsVal != null && !node.Format.ContainsKey("size"))
                node.Format["size"] = $"{int.Parse(fsVal) / 2.0:0.##}pt";

            var boldEl = rp?.Bold ?? markRp?.GetFirstChild<Bold>();
            if (boldEl != null && !node.Format.ContainsKey("bold")) node.Format["bold"] = IsToggleOn(boldEl);

            var italicEl = rp?.Italic ?? markRp?.GetFirstChild<Italic>();
            if (italicEl != null && !node.Format.ContainsKey("italic")) node.Format["italic"] = IsToggleOn(italicEl);

            // Complex-script readback (font.cs / size.cs / bold.cs / italic.cs).
            // See WordHandler.I18n.cs.
            ReadComplexScriptRunFormatting(rp, markRp, node.Format);

            var colorEl = rp?.Color ?? markRp?.GetFirstChild<Color>();
            if (colorEl != null && !node.Format.ContainsKey("color"))
            {
                // Prefer theme color over Val when both set (Val often
                // "auto" when ThemeColor is the authoritative source).
                if (colorEl.ThemeColor?.HasValue == true)
                    node.Format["color"] = colorEl.ThemeColor.InnerText;
                else if (colorEl.Val?.Value != null)
                    node.Format["color"] = ParseHelpers.FormatHexColor(colorEl.Val.Value);
            }

            var ulEl = rp?.Underline ?? markRp?.GetFirstChild<Underline>();
            if (ulEl?.Val != null && !node.Format.ContainsKey("underline"))
                node.Format["underline"] = ulEl.Val.InnerText;
            // CONSISTENCY(underline-color): backfilled from style Get edc8f884.
            if (ulEl?.Color?.Value != null && !node.Format.ContainsKey("underline.color"))
                node.Format["underline.color"] = ParseHelpers.FormatHexColor(ulEl.Color.Value);

            var strikeEl = rp?.Strike ?? (OpenXmlLeafElement?)markRp?.GetFirstChild<Strike>();
            if (strikeEl != null && !node.Format.ContainsKey("strike")) node.Format["strike"] = true;

            var hlEl = rp?.Highlight ?? markRp?.GetFirstChild<Highlight>();
            if (hlEl?.Val != null && !node.Format.ContainsKey("highlight"))
                node.Format["highlight"] = hlEl.Val.InnerText;

            // BUG-DUMP-MARKRPR-CS: empty-paragraph ¶-mark kern. Text
            // paragraphs route the mark's kern through markRPr.kern above;
            // here we only cover the empty-paragraph case (markRp != null
            // ⇒ no text run) so we don't hoist a first run's kern onto the
            // paragraph node. AddParagraph's bare-key path applies it back
            // to the ¶ mark rPr on replay.
            if (markRp != null && !node.Format.ContainsKey("kern"))
            {
                var kEmpty = markRp.GetFirstChild<Kern>();
                if (kEmpty?.Val?.HasValue == true)
                    node.Format["kern"] = kEmpty.Val.Value.ToString();
            }
        }

        // Populate effective.* properties from style inheritance
        PopulateEffectiveParagraphProperties(node, para);

        if (depth > 0)
        {
            // BUG-DUMP13-02: interleave typed Runs and inline M.OfficeMath
            // equations in DOM order so paragraphs like `r1 / m:oMath / r2`
            // emit r1, equation, r2 (not r1, r2, equation). Previously
            // GetAllRuns appended every run first and the inline-equation
            // loop below appended all equations afterwards as a separate
            // group, so DOM order was lost on dump round-trip.
            //
            // We compute a DOM-position index per element via a single
            // descendant walk (Descendants() yields document order) and
            // use it to sort only the run+equation slice, leaving other
            // categories (sdt/bookmark/field/etc.) in their original
            // append order.
            int runIdx = 0;
            int inlineEqIdx = 0;
            var descendantPos = new Dictionary<OpenXmlElement, int>(ReferenceEqualityComparer.Instance);
            int dpi = 0;
            foreach (var d in para.Descendants())
                descendantPos[d] = dpi++;

            var runs = GetAllRuns(para);
            // BUG-DUMP9-04: m:oMath nested inside w:hyperlink is a
            // grandchild of the paragraph and was silently dropped.
            // BUG-DUMP8-03: include m:oMath nested inside w:ins/w:del
            // change-track wrappers — they are paragraph grandchildren,
            // not direct children, and were silently dropped on dump.
            var inlineEqsAll = para.Elements<M.OfficeMath>()
                .Concat(para.Elements<InsertedRun>().SelectMany(ins => ins.Elements<M.OfficeMath>()))
                .Concat(para.Elements<DeletedRun>().SelectMany(del => del.Elements<M.OfficeMath>()))
                .Concat(para.Elements<Hyperlink>().SelectMany(hl => hl.Elements<M.OfficeMath>()))
                .ToList();
            // BUG-DUMP15-04: paragraph hyperlink children for hyperlink-
            // scoped equation paths. m:oMath inside w:hyperlink must
            // surface as /…/p[N]/hyperlink[K]/equation[M] so dump→batch
            // replays the equation INSIDE the hyperlink rather than
            // alongside it. Index hyperlinks by their position among
            // the paragraph's direct Hyperlink children.
            var paraHyperlinks = para.Elements<Hyperlink>().ToList();

            // Merge runs and inline equations by DOM position, then emit
            // in that interleaved order.
            // BUG-DUMP15-02: bare <w:fldChar>/<w:instrText> direct children
            // of <w:p> (not wrapped in a <w:r>) are parsed as
            // OpenXmlUnknownElement and silently dropped from the children
            // list, which left CollapseFieldChains nothing to stitch and
            // dump→batch round-trips lost the entire HYPERLINK chain.
            // Surface them as synthetic fieldChar/instrText nodes so the
            // emitter can collapse them into a `field` row.
            const string wNs2 = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var bareFieldUnknowns = para.Elements<DocumentFormat.OpenXml.OpenXmlUnknownElement>()
                .Where(u => u.NamespaceUri == wNs2
                    && (u.LocalName == "fldChar" || u.LocalName == "instrText"))
                .ToList();
            // BUG-DUMP25-01: include direct-child BookmarkStart elements in
            // the DOM-ordered merge so a bookmark sitting between two runs
            // surfaces as `r, bookmark, r` rather than the legacy
            // `r, r, bookmark` (every bookmark hoisted to the tail of
            // node.Children). The trailing standalone bookmark loop below
            // is now skipped when this branch surfaces them.
            var paraBookmarks = para.Elements<BookmarkStart>().ToList();
            var ordered = runs.Select(r => (pos: descendantPos.TryGetValue(r, out var p) ? p : int.MaxValue, kind: "run", el: (OpenXmlElement)r))
                .Concat(inlineEqsAll.Select(e => (pos: descendantPos.TryGetValue(e, out var p) ? p : int.MaxValue, kind: "eq", el: (OpenXmlElement)e)))
                .Concat(bareFieldUnknowns.Select(u => (pos: descendantPos.TryGetValue(u, out var p) ? p : int.MaxValue, kind: u.LocalName == "fldChar" ? "fieldChar" : "instrText", el: (OpenXmlElement)u)))
                .Concat(paraBookmarks.Select(b => (pos: descendantPos.TryGetValue(b, out var p) ? p : int.MaxValue, kind: "bookmark", el: (OpenXmlElement)b)))
                .OrderBy(t => t.pos)
                .ToList();
            int bareFieldIdx = 0;
            foreach (var entry in ordered)
            {
                if (entry.kind == "run")
                {
                    var runNode = ElementToNode(entry.el, $"{path}/r[{runIdx + 1}]", depth - 1);
                    // BUG-DUMP18-02: surface a hyperlink-scoped subpath on
                    // runs that are direct children of <w:hyperlink>. The
                    // canonical Path stays flat (/…/r[N]) for back-compat
                    // with every existing caller; WordBatchEmitter's
                    // CollapseFieldChains carries this hint to the synth
                    // field-add row so a fldChar-chain field inside a
                    // hyperlink replays INSIDE the hyperlink instead of
                    // alongside it. Mirrors the SimpleField hyperlink-
                    // scope path emitted below.
                    if (entry.el.Parent is Hyperlink runHl)
                    {
                        int hlIdxRun = paraHyperlinks.IndexOf(runHl);
                        if (hlIdxRun >= 0)
                            runNode.Format["_hyperlinkParent"] = $"{path}/hyperlink[{hlIdxRun + 1}]";
                    }
                    node.Children.Add(runNode);
                    runIdx++;
                }
                else if (entry.kind == "eq")
                {
                    // BUG-DUMP15-04: equations whose immediate parent is
                    // <w:hyperlink> get a hyperlink-scoped path so the
                    // emitter can place the equation INSIDE the hyperlink
                    // on replay.
                    string eqPath;
                    if (entry.el.Parent is Hyperlink eqHl)
                    {
                        int hlIdx = paraHyperlinks.IndexOf(eqHl);
                        int hlEqIdx = eqHl.Elements<M.OfficeMath>()
                            .ToList().IndexOf((M.OfficeMath)entry.el);
                        eqPath = $"{path}/hyperlink[{hlIdx + 1}]/equation[{hlEqIdx + 1}]";
                    }
                    else
                    {
                        eqPath = $"{path}/equation[{inlineEqIdx + 1}]";
                        inlineEqIdx++;
                    }
                    node.Children.Add(ElementToNode(entry.el, eqPath, depth - 1));
                }
                else if (entry.kind == "bookmark")
                {
                    // BUG-DUMP25-01: emit BookmarkStart at its DOM position
                    // (sandwiched between sibling runs/equations) so dump→
                    // batch round-trips preserve mid-paragraph bookmark
                    // offsets like Word's _GoBack resume-cursor mark.
                    // Path index counts bookmarks among themselves to
                    // stay 1-based, mirroring the legacy bmIdx counter.
                    int bmPathIdx = paraBookmarks.IndexOf((BookmarkStart)entry.el);
                    node.Children.Add(ElementToNode(entry.el, $"{path}/bookmark[{bmPathIdx + 1}]", depth - 1));
                }
                else
                {
                    // BUG-DUMP15-02: synthesize fieldChar/instrText nodes
                    // for bare unknown elements so CollapseFieldChains can
                    // stitch the field. Mirrors the Run-based shape.
                    var u = (DocumentFormat.OpenXml.OpenXmlUnknownElement)entry.el;
                    var bn = new DocumentNode
                    {
                        Type = entry.kind,
                        Path = $"{path}/r[{runIdx + 1}]",
                    };
                    runIdx++;
                    if (entry.kind == "fieldChar")
                    {
                        var fct = u.GetAttribute("fldCharType", wNs2).Value;
                        if (!string.IsNullOrEmpty(fct))
                            bn.Format["fieldCharType"] = fct;
                    }
                    else // instrText
                    {
                        bn.Format["instruction"] = u.InnerText;
                        bn.Text = u.InnerText;
                    }
                    node.Children.Add(bn);
                    bareFieldIdx++;
                }
            }
            // BUG-DUMP5-06/07: <w:ruby> and <w:smartTag> aren't registered
            // as typed paragraph children in the OpenXml SDK schema set we
            // load — RawSet-injected fragments and SDK-untracked content
            // from real-world docx files surface them as
            // OpenXmlUnknownElement, so Descendants<Run>() inside
            // GetAllRuns skips every nested run (the inner <w:r> is also
            // an unknown element, not a typed Run). Walk the unknown
            // subtrees and synthesize plain `run` DocumentNodes from any
            // <w:r>/<w:t> children we find so the inner text round-trips
            // through dump→batch instead of vanishing.
            const string wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            foreach (var unkRun in para.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>())
            {
                if (unkRun.LocalName != "r" || unkRun.NamespaceUri != wNs) continue;
                // Only surface runs whose direct parent is an unknown
                // wrapper (ruby/rt/rubyBase/smartTag/customXml). Runs
                // whose parent is a typed Paragraph would already be
                // typed Runs and reached via GetAllRuns above; if they
                // somehow surface as unknown here it's because the
                // entire paragraph is malformed and we'd duplicate.
                // BUG-DUMP7-10: also accept InsertedRun/DeletedRun
                // ancestors — w:del>w:ruby in a malformed doc parses
                // ruby as unknown but the typed w:del wrapper still
                // sits between para and the unknown subtree, so the
                // ancestor (not just direct parent) needs the typed
                // change-track wrapper allowance.
                if (unkRun.Parent is not DocumentFormat.OpenXml.OpenXmlUnknownElement
                    && unkRun.Ancestors<InsertedRun>().FirstOrDefault() == null
                    && unkRun.Ancestors<DeletedRun>().FirstOrDefault() == null)
                    continue;
                var sbInner = new System.Text.StringBuilder();
                foreach (var tEl in unkRun.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>())
                {
                    if (tEl.NamespaceUri != wNs) continue;
                    // BUG-DUMP7-10: a w:del-wrapped ruby's inner runs
                    // carry their text in <w:delText>, not <w:t>.
                    // Without delText/instrText the "base"/"rt" text
                    // dropped silently and the paragraph surfaced empty.
                    if (tEl.LocalName == "t"
                        || tEl.LocalName == "delText"
                        || tEl.LocalName == "instrText")
                        sbInner.Append(tEl.InnerText);
                }
                if (sbInner.Length == 0) continue;
                var synthNode = new DocumentNode
                {
                    Type = "run",
                    Text = sbInner.ToString(),
                    Path = $"{path}/r[{runIdx + 1}]",
                };
                // BUG-DUMP7-10: preserve trackChange attribution from
                // the typed w:ins/w:del ancestor so the round-trip
                // re-emits the wrapper (mirrors the typed-Run branch
                // at the top of this method).
                var insAnc = unkRun.Ancestors<InsertedRun>().FirstOrDefault();
                if (insAnc != null)
                {
                    synthNode.Format["revision.type"] = "ins";
                    if (!string.IsNullOrEmpty(insAnc.Author?.Value))
                        synthNode.Format["revision.author"] = insAnc.Author!.Value!;
                    if (insAnc.Date?.Value is DateTime insAncDate)
                        synthNode.Format["revision.date"] = insAncDate.ToString("o");
                }
                else
                {
                    var delAnc = unkRun.Ancestors<DeletedRun>().FirstOrDefault();
                    if (delAnc != null)
                    {
                        synthNode.Format["revision.type"] = "del";
                        if (!string.IsNullOrEmpty(delAnc.Author?.Value))
                            synthNode.Format["revision.author"] = delAnc.Author!.Value!;
                        if (delAnc.Date?.Value is DateTime delAncDate)
                            synthNode.Format["revision.date"] = delAncDate.ToString("o");
                    }
                }
                node.Children.Add(synthNode);
                runIdx++;
            }
            // BUG-DUMP25-01: BookmarkStart children are now surfaced
            // inside the DOM-ordered `ordered` merge above, so a
            // bookmark between two runs round-trips at its original
            // intra-paragraph offset. The legacy standalone loop here
            // (which appended every bookmark at the tail of
            // node.Children) is intentionally left empty.
            // BUG-DUMP4-06: surface inline SdtRun (content control) children
            // so WordBatchEmitter can re-emit a typed `add sdt` row carrying
            // alias/tag/type metadata. Without this, GetAllRuns unwrapped
            // the SdtRun's inner Run as a plain `add r` and the metadata
            // was silently dropped on dump round-trip.
            int sdtRunIdx = 0;
            foreach (var sdtR in para.Elements<SdtRun>())
            {
                node.Children.Add(ElementToNode(sdtR, $"{path}/sdt[{sdtRunIdx + 1}]", depth - 1));
                sdtRunIdx++;
            }
            // BUG-DUMP7-03 / BUG-DUMP8-03 / BUG-DUMP9-04: inline <m:oMath>
            // children (including those nested inside w:ins/w:del/w:hyperlink
            // wrappers) are now interleaved with runs at the top of this
            // block (BUG-DUMP13-02) so DOM order is preserved. The
            // `inlineEqIdx` counter declared there carries forward into the
            // block-level oMathPara branch below.
            // BUG-DUMP12-02: surface block-level <m:oMathPara> children of a
            // mixed-content paragraph (paragraph that ALSO has ordinary
            // runs/hyperlinks/etc) as display equation nodes. The pure-wrapper
            // case is handled at the body level via the LocalName=="oMathPara"
            // branch in WalkBodyChild + IsOMathParaWrapperParagraph; the
            // mixed-content case falls through to plain p[N] and was silently
            // dropping the equation. We only emit when the para is NOT a pure
            // oMathPara wrapper, to avoid double-counting against the body
            // /oMathPara[M] addressing.
            if (!IsOMathParaWrapperParagraph(para))
            {
                foreach (var blockEq in para.Elements<M.Paragraph>())
                {
                    node.Children.Add(ElementToNode(blockEq, $"{path}/equation[{inlineEqIdx + 1}]", depth - 1));
                    inlineEqIdx++;
                }
            }
            // BUG-DUMP6-01: surface <w:fldSimple> children as typed `field`
            // nodes so WordBatchEmitter can re-emit `add field` with the
            // instruction preserved. Without this, GetAllRuns descended
            // into SimpleField and surfaced the inner display run as a
            // plain run, silently dropping the w:instr attribute.
            // BUG-DUMP9-03: w:fldSimple nested inside w:hyperlink is a
            // grandchild of the paragraph and was silently dropped.
            int fldSimpleIdx = 0;
            // BUG-DUMP18-02: w:fldSimple inside w:hyperlink must surface
            // as /…/p[N]/hyperlink[K]/field[M] so dump→batch replays the
            // field INSIDE the hyperlink rather than alongside it. Mirrors
            // BUG-DUMP15-04 hyperlink-scoped equation paths above.
            foreach (var fld in para.Elements<SimpleField>())
            {
                var instr = fld.Instruction?.Value ?? "";
                var displayText = string.Join("",
                    fld.Descendants<Text>().Select(t => t.Text));
                var fldNode = new DocumentNode
                {
                    Type = "field",
                    Text = displayText,
                    Path = $"{path}/field[{fldSimpleIdx + 1}]",
                };
                fldNode.Format["instruction"] = instr.Trim();
                var instrUpper = instr.Trim().Split(' ', 2)[0].ToUpperInvariant();
                if (!string.IsNullOrEmpty(instrUpper))
                    fldNode.Format["fieldType"] = instrUpper.ToLowerInvariant();
                // Cross-handler `evaluated` protocol — true whenever
                // there's a readable cached result. dirty=true (Word
                // re-renders on open) keeps evaluated=true since a value
                // is still available; view issues field_cache_stale
                // surfaces the dirty + cached combination separately.
                var fldDirty = fld.Dirty?.Value == true;
                if (fldDirty) fldNode.Format["dirty"] = true;
                fldNode.Format["evaluated"] = displayText.Length > 0;
                node.Children.Add(fldNode);
                fldSimpleIdx++;
            }
            for (int hlI = 0; hlI < paraHyperlinks.Count; hlI++)
            {
                var hl = paraHyperlinks[hlI];
                int perHlFldIdx = 0;
                foreach (var fld in hl.Elements<SimpleField>())
                {
                    var instr = fld.Instruction?.Value ?? "";
                    var displayText = string.Join("",
                        fld.Descendants<Text>().Select(t => t.Text));
                    var fldNode = new DocumentNode
                    {
                        Type = "field",
                        Text = displayText,
                        Path = $"{path}/hyperlink[{hlI + 1}]/field[{perHlFldIdx + 1}]",
                    };
                    fldNode.Format["instruction"] = instr.Trim();
                    var instrUpper = instr.Trim().Split(' ', 2)[0].ToUpperInvariant();
                    if (!string.IsNullOrEmpty(instrUpper))
                        fldNode.Format["fieldType"] = instrUpper.ToLowerInvariant();
                    var fldDirtyHl = fld.Dirty?.Value == true;
                    if (fldDirtyHl) fldNode.Format["dirty"] = true;
                    fldNode.Format["evaluated"] = displayText.Length > 0;
                    node.Children.Add(fldNode);
                    perHlFldIdx++;
                }
            }
        }
        return node;
    }

    private DocumentNode ElementToNode(OpenXmlElement element, string path, int depth)
    {
        var node = new DocumentNode { Path = path, Type = element.LocalName };

        if (element is BookmarkStart bkStart)
            return BookmarkStartToNode(bkStart, node);

        if (element is Footnote fnEl)
            return FootnoteToNode(fnEl, node, path, depth);

        if (element is Endnote enEl)
            return EndnoteToNode(enEl, node, path, depth);

        if (element is Comment comment)
            return CommentToNode(comment, node);

        if (element is SectionProperties sectPrEl)
            return SectionPropertiesToNode(sectPrEl, path);

        if (element is Paragraph para)
            return ParagraphToNode(para, node, path, depth);
        else if (element is Run run)
            return RunToNode(run, node, path);
        else if (element is Hyperlink hyperlink)
            return HyperlinkToNode(hyperlink, node);
        else if (element is Table table)
            return TableToNode(table, node, path, depth);
        else if (element is TableCell directCell)
            return TableCellToNode(directCell, node, path, depth);
        else if (element is TableRow directRow)
            return TableRowToNode(directRow, node, path, depth);
        else if (element is SdtBlock sdtBlockNode)
            return SdtBlockToNode(sdtBlockNode, node);
        else if (element is SdtRun sdtRunNode)
            return SdtRunToNode(sdtRunNode, node);
        else if (element.LocalName == "oMathPara" || element is M.Paragraph)
        {
            node.Type = "equation";
            node.Format["mode"] = "display";
            // BUG-DUMP19-02: surface m:oMathParaPr/m:jc as Format["align"] so
            // block-equation alignment round-trips. Without this the value is
            // silently dropped on read-back.
            var mathPPr = element.GetFirstChild<M.ParagraphProperties>();
            var jcVal = mathPPr?.Justification?.Val?.InnerText;
            if (!string.IsNullOrEmpty(jcVal))
            {
                node.Format["align"] = jcVal switch
                {
                    "centerGroup" => "centerGroup",
                    _ => jcVal // "left" | "center" | "right"
                };
            }
            // Extract LaTeX via FormulaParser
            var oMath = element.Descendants<M.OfficeMath>().FirstOrDefault();
            if (oMath != null)
            {
                try { node.Text = Core.FormulaParser.ToLatex(oMath); }
                catch { node.Text = element.InnerText; }
            }
            else
            {
                node.Text = element.InnerText;
            }
        }
        else if (element is M.OfficeMath inlineMath)
            return OfficeMathToNode(inlineMath, node);
        else if (element is Header or Footer)
            return HeaderFooterToNode(element, node, path, depth);
        else if (element is Body bodyNode)
            return BodyToNode(bodyNode, node, path, depth);
        else
        {
            // Generic fallback: collect XML attributes and child val patterns
            foreach (var attr in element.GetAttributes())
                node.Format[attr.LocalName] = attr.Value;
            foreach (var child in element.ChildElements)
            {
                if (child.ChildElements.Count == 0)
                {
                    foreach (var attr in child.GetAttributes())
                    {
                        if (attr.LocalName.Equals("val", StringComparison.OrdinalIgnoreCase))
                        {
                            node.Format[child.LocalName] = attr.Value;
                            break;
                        }
                    }
                }
            }

            var innerText = element.InnerText;
            if (!string.IsNullOrEmpty(innerText))
                node.Text = innerText.Length > 200 ? innerText[..200] + "..." : innerText;
            if (string.IsNullOrEmpty(innerText))
            {
                var outerXml = element.OuterXml;
                node.Preview = outerXml.Length > 200 ? outerXml[..200] + "..." : outerXml;
            }

            node.ChildCount = element.ChildElements.Count;
            if (depth > 0)
            {
                var typeCounters = new Dictionary<string, int>();
                foreach (var child in element.ChildElements)
                {
                    var name = child.LocalName;
                    typeCounters.TryGetValue(name, out int idx);
                    node.Children.Add(ElementToNode(child, $"{path}/{name}[{idx + 1}]", depth - 1));
                    typeCounters[name] = idx + 1;
                }
            }
        }

        return node;
    }

    private static void ReadRowProps(TableRow row, DocumentNode node)
    {
        var trPr = row.TableRowProperties;
        if (trPr == null) return;
        // trPrChange: `set row + trackChange.author` snapshots prior trPr.
        // Surfaced under trPrChange.* for round-trip; EmitTable emits a
        // follow-up `set tr[N]` with trackChange.author/date to reproduce.
        var trPrChange = trPr.GetFirstChild<TableRowPropertiesChange>();
        if (trPrChange != null)
        {
            if (!string.IsNullOrEmpty(trPrChange.Author?.Value))
                node.Format["trPrChange.author"] = trPrChange.Author!.Value!;
            if (trPrChange.Date?.Value is DateTime trDate)
                node.Format["trPrChange.date"] = trDate.ToString("o");
        }
        var rh = trPr.GetFirstChild<TableRowHeight>();
        if (rh?.Val?.Value != null)
        {
            // CONSISTENCY(unit-qualified-readback): docx stores row height
            // in twips (1440 twips = 1 inch = 2.54 cm); emit as "{n}cm" to
            // match the input vocabulary (Set accepts "2cm") and the other
            // docx length readbacks (pageWidth/pageHeight). Prior code
            // emitted pt and broke `set height=2cm` → `get height=56.7pt`
            // round-trip.
            var heightCm = rh.Val.Value * 2.54 / 1440.0;
            node.Format["height"] = $"{heightCm.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}cm";
            if (rh.HeightType?.Value == HeightRuleValues.Exact)
                node.Format["height.rule"] = "exact";
        }
        if (trPr.GetFirstChild<TableHeader>() != null)
            node.Format["header"] = true;
        if (trPr.GetFirstChild<CantSplit>() != null)
            node.Format["cantSplit"] = true;
    }

    private static void ReadCellProps(TableCell cell, DocumentNode node)
    {
        var tcPr = cell.TableCellProperties;
        if (tcPr != null)
        {
            // tcPrChange: `set cell + trackChange.author` snapshots prior tcPr.
            // Surfaced under tcPrChange.* for round-trip; EmitTable emits a
            // follow-up `set tr[N]/tc[M]` with trackChange.author/date.
            var tcPrChange = tcPr.GetFirstChild<TableCellPropertiesChange>();
            if (tcPrChange != null)
            {
                if (!string.IsNullOrEmpty(tcPrChange.Author?.Value))
                    node.Format["tcPrChange.author"] = tcPrChange.Author!.Value!;
                if (tcPrChange.Date?.Value is DateTime tcDate)
                    node.Format["tcPrChange.date"] = tcDate.ToString("o");
            }
            // Borders (including diagonal)
            var cb = tcPr.TableCellBorders;
            if (cb != null)
            {
                ReadBorder(cb.TopBorder, "border.top", node);
                ReadBorder(cb.BottomBorder, "border.bottom", node);
                ReadBorder(cb.LeftBorder, "border.left", node);
                ReadBorder(cb.RightBorder, "border.right", node);
                ReadBorder(cb.TopLeftToBottomRightCellBorder, "border.tl2br", node);
                ReadBorder(cb.TopRightToBottomLeftCellBorder, "border.tr2bl", node);
            }
            // Shading — check for gradient (w14:gradFill in mc:AlternateContent) first
            var mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
            var gradAc = tcPr.ChildElements
                .FirstOrDefault(e => e.LocalName == "AlternateContent" && e.NamespaceUri == mcNs);
            if (gradAc != null && gradAc.InnerXml.Contains("gradFill"))
            {
                // Parse gradient colors and angle from w14:gradFill XML
                var colors = new List<string>();
                foreach (var match in System.Text.RegularExpressions.Regex.Matches(
                    gradAc.InnerXml, @"val=""([0-9A-Fa-f]{6})"""))
                {
                    colors.Add(((System.Text.RegularExpressions.Match)match).Groups[1].Value);
                }
                var angleMatch = System.Text.RegularExpressions.Regex.Match(
                    gradAc.InnerXml, @"ang=""(\d+)""");
                var angle = angleMatch.Success ? int.Parse(angleMatch.Groups[1].Value) / 60000.0 : 0.0;
                var angleStr = angle % 1 == 0 ? $"{(int)angle}" : $"{angle:0.##}";
                if (colors.Count >= 2)
                {
                    node.Format["fill"] = $"gradient;{ParseHelpers.FormatHexColor(colors[0])};{ParseHelpers.FormatHexColor(colors[1])};{angleStr}";
                }
                else if (colors.Count == 1)
                {
                    node.Format["fill"] = ParseHelpers.FormatHexColor(colors[0]);
                }
            }
            else
            {
                var shd = tcPr.Shading;
                if (shd != null)
                {
                    // BUG-DUMP21-02 / BUG-R2-P3-11: emit only the canonical
                    // shading.val/.fill/.color sub-keys. Previously also
                    // emitted a legacy `fill` alias carrying the same value,
                    // which violated the root CLAUDE.md "one canonical key per
                    // semantic value" rule and showed up as duplicate output
                    // for every shaded cell. shading.fill is the canonical key
                    // (matches the OOXML attribute name).
                    var cShdVal = shd.Val?.InnerText;
                    var cShdFill = shd.Fill?.Value;
                    var cShdColor = shd.Color?.Value;
                    if (!string.IsNullOrEmpty(cShdVal)) node.Format["shading.val"] = cShdVal;
                    if (!string.IsNullOrEmpty(cShdFill)) node.Format["shading.fill"] = ParseHelpers.FormatHexColor(cShdFill);
                    if (!string.IsNullOrEmpty(cShdColor)) node.Format["shading.color"] = ParseHelpers.FormatHexColor(cShdColor);
                }
            }
            // Width
            // BUG-DUMP6-04: preserve w:tcW @type semantics. Mirror the table-level
            // width readback above (line ~1930) — pct widths are stored as
            // fifths-of-percent, so divide by 50 and append '%' so dump→batch
            // can recognize and re-emit pct cell widths.
            // BUG-R4-05: emit width with explicit unit suffix (dxa/%) — root
            // CLAUDE.md mandates unit-qualified width readback. Bare integer
            // ("3000") is the historic bug.
            if (tcPr.TableCellWidth?.Width?.Value != null)
            {
                var cwType = tcPr.TableCellWidth.Type?.Value;
                if (cwType == TableWidthUnitValues.Pct
                    && int.TryParse(tcPr.TableCellWidth.Width.Value, out var pctRaw))
                    node.Format["width"] = (pctRaw / 50) + "%";
                else if (cwType == TableWidthUnitValues.Auto)
                    node.Format["width"] = "auto";
                else if (cwType == TableWidthUnitValues.Nil
                    || tcPr.TableCellWidth.Width.Value == "0")
                    node.Format["width"] = "0dxa";
                else
                    node.Format["width"] = tcPr.TableCellWidth.Width.Value + "dxa";
            }
            // Vertical alignment
            if (tcPr.TableCellVerticalAlignment?.Val?.Value != null)
                node.Format["valign"] = tcPr.TableCellVerticalAlignment.Val.InnerText;
            // Vertical merge
            if (tcPr.VerticalMerge != null)
                node.Format["vmerge"] = tcPr.VerticalMerge.Val?.Value == MergedCellValues.Restart ? "restart" : "continue";
            // Horizontal merge — same toggle pattern as vmerge: ST_Merge val=restart
            // marks the leading cell of a horizontal span, bare <w:hMerge/> marks the
            // continuation cells. Without this read block dump→batch silently dropped
            // every horizontal span on round-trip.
            if (tcPr.HorizontalMerge != null)
                node.Format["hmerge"] = tcPr.HorizontalMerge.Val?.Value == MergedCellValues.Restart ? "restart" : "continue";
            // Grid span
            if (tcPr.GridSpan?.Val?.Value != null && tcPr.GridSpan.Val.Value > 1)
                node.Format["colspan"] = tcPr.GridSpan.Val.Value;
            // Cell padding/margins
            var mar = tcPr.TableCellMargin;
            if (mar != null)
            {
                if (mar.TopMargin?.Width?.Value != null) node.Format["padding.top"] = mar.TopMargin.Width.Value;
                if (mar.BottomMargin?.Width?.Value != null) node.Format["padding.bottom"] = mar.BottomMargin.Width.Value;
                if (mar.LeftMargin?.Width?.Value != null) node.Format["padding.left"] = mar.LeftMargin.Width.Value;
                if (mar.RightMargin?.Width?.Value != null) node.Format["padding.right"] = mar.RightMargin.Width.Value;
            }
            // Text direction
            if (tcPr.TextDirection?.Val?.Value != null)
                node.Format["textDirection"] = tcPr.TextDirection.Val.InnerText;
            // No wrap
            if (tcPr.NoWrap != null)
                node.Format["nowrap"] = true;
            // BUG-R3-03: cnfStyle (conditional formatting bitfield).
            var cnfRead = tcPr.GetFirstChild<ConditionalFormatStyle>();
            if (cnfRead?.Val?.Value is string cnfVal && !string.IsNullOrEmpty(cnfVal))
                node.Format["cnfStyle"] = cnfVal;
        }
        // BUG-R4-05: when no per-cell tcW is set, synthesize width from the
        // parent table's tblGrid/gridCol so Get always exposes a unit-qualified
        // width (matches the cross-handler width contract). CONSISTENCY(add-set-symmetry):
        // Add intentionally does not stamp per-cell tcW (BUG-R6-06) — width
        // lives in tblGrid as the schema intends — so Get must back-fill.
        if (!node.Format.ContainsKey("width"))
        {
            var parentTbl = cell.Ancestors<Table>().FirstOrDefault();
            var parentRow = cell.Parent as TableRow;
            if (parentTbl != null && parentRow != null)
            {
                var cellIdx = parentRow.Elements<TableCell>().ToList().IndexOf(cell);
                var gridCols = parentTbl.GetFirstChild<TableGrid>()?.Elements<GridColumn>().ToList();
                if (gridCols != null && cellIdx >= 0 && cellIdx < gridCols.Count)
                {
                    // Account for gridSpan — sum spanned cols.
                    var span = (tcPr?.GridSpan?.Val?.Value ?? 1);
                    long total = 0;
                    for (int gi = cellIdx; gi < Math.Min(cellIdx + span, gridCols.Count); gi++)
                    {
                        if (uint.TryParse(gridCols[gi].Width?.Value, out var gv))
                            total += gv;
                    }
                    if (total > 0)
                        node.Format["width"] = total + "dxa";
                }
            }
        }
        // Alignment from first paragraph
        var firstPara = cell.Elements<Paragraph>().FirstOrDefault();
        var just = firstPara?.ParagraphProperties?.Justification?.Val;
        if (just != null)
            node.Format["align"] = just.InnerText;
        // Direction: <w:bidi/> on the first cell paragraph maps to canonical
        // direction=rtl. Mirrors paragraph readback canonical key. R20-bt-2:
        // also surface direction=rtl when the enclosing table carries
        // <w:bidiVisual/> on tblPr — cells inherit table-level visual RTL
        // even without their own pPr.bidi.
        if (firstPara?.ParagraphProperties?.BiDi != null)
            node.Format["direction"] = "rtl";
        else if (cell.Ancestors<Table>().FirstOrDefault()
                     ?.GetFirstChild<TableProperties>()?.GetFirstChild<BiDiVisual>() != null)
            node.Format["direction"] = "rtl";
        // Run-level formatting from first run (mirrors PPTX table cell behavior)
        var firstRun = cell.Descendants<Run>().FirstOrDefault();
        if (firstRun?.RunProperties != null)
        {
            var rPr = firstRun.RunProperties;
            if (rPr.RunFonts?.Ascii?.Value != null) node.Format["font"] = rPr.RunFonts.Ascii.Value;
            if (rPr.FontSize?.Val?.Value != null) node.Format["size"] = $"{int.Parse(rPr.FontSize.Val.Value) / 2.0:0.##}pt";
            if (rPr.Bold != null) node.Format["bold"] = IsToggleOn(rPr.Bold);
            if (rPr.Italic != null) node.Format["italic"] = IsToggleOn(rPr.Italic);
            if (rPr.Color?.Val?.Value != null) node.Format["color"] = ParseHelpers.FormatHexColor(rPr.Color.Val.Value);
            else if (rPr.Color?.ThemeColor?.HasValue == true) node.Format["color"] = rPr.Color.ThemeColor.InnerText;
            if (rPr.Underline?.Val != null) node.Format["underline"] = rPr.Underline.Val.InnerText;
            // CONSISTENCY(underline-color): backfilled from style Get edc8f884.
            if (rPr.Underline?.Color?.Value != null)
                node.Format["underline.color"] = ParseHelpers.FormatHexColor(rPr.Underline.Color.Value);
            if (rPr.Strike != null) node.Format["strike"] = true;
            if (rPr.Highlight?.Val != null) node.Format["highlight"] = rPr.Highlight.Val.InnerText;
        }
    }

    private static void ReadBorder(BorderType? border, string key, DocumentNode node)
    {
        if (border?.Val == null) return;
        // CONSISTENCY(canonical-keys): emit val on the parent key plus .sz/.color/.space sub-keys
        // (matches Excel border.* schema). No compound semicolon-joined string — that was a private
        // encoding that diverged from both OOXML and the rest of the project.
        node.Format[key] = border.Val?.InnerText ?? "none";
        if (border.Size?.Value is uint sz) node.Format[$"{key}.sz"] = sz;
        if (border.Color?.Value is { } c) node.Format[$"{key}.color"] = ParseHelpers.FormatHexColor(c);
        if (border.Space?.Value is uint sp) node.Format[$"{key}.space"] = sp;
    }

    // OOXML localNames that curated style/paragraph/run readers already map
    // to canonical keys. FillUnknownChildProps skips these so the long-tail
    // fallback doesn't re-expose them under their bare OOXML names alongside
    // the canonical key (e.g. avoid emitting both `bold: true` and `b: true`).
    private static readonly System.Collections.Generic.HashSet<string> CuratedStyleLocalNames =
        new(System.StringComparer.Ordinal)
    {
        // rPr-side (covered by curated style/paragraph/run readers)
        "b", "bCs", "i", "iCs", "sz", "szCs", "u", "color", "strike", "rFonts",
        "highlight", "caps", "smallCaps", "dstrike", "vanish",
        "outline", "shadow", "emboss", "imprint", "noProof", "rtl",
        "vertAlign", "spacing", "shd",
        // BUG-DUMP22-08: <w:bdr/> is multi-attribute (val+sz+color+space).
        // Curated reader emits the colon-encoded compound form; suppress
        // the long-tail fallback so the bare `bdr=single` name doesn't
        // co-emit alongside the canonical encoded value.
        "bdr",
        // BUG-DUMP10-01: <w:eastAsianLayout/> is a multi-attribute element
        // surfaced by the curated reader as eastAsianLayout.vert / .combine
        // dotted keys. Skip the long-tail fallback so it doesn't double-emit
        // the bare element name with a `true` value.
        "eastAsianLayout",
        // pPr-side
        "jc", "ind", "outlineLvl", "widowControl",
        "keepNext", "keepLines", "pageBreakBefore", "contextualSpacing",
        "pBdr", "numPr", "tabs", "pStyle",
        // bidi maps to canonical `direction` in style/paragraph readback;
        // skip the long-tail fallback to avoid emitting both `direction: rtl`
        // and `bidi: true` for the same <w:bidi/> child element.
        "bidi",
        // Container elements covered by the curated paragraph-mark / run-property
        // reader (see paraRp block ~line 1004). Without this, an empty <w:rPr/>
        // left behind by Set bold=false (etc.) would surface as `rPr: true` via
        // the long-tail fallback. fuzz-1.
        "rPr",
        // BUG-R7-09 / F-3: <w:lang/> is a multi-slot element (val=latin /
        // eastAsia / bidi). The curated reader emits each slot as
        // lang.latin / lang.ea / lang.cs. Word/WPS occasionally write a bare
        // <w:lang/> with no attributes as a "reset to default language"
        // sentinel — the long-tail fallback would then surface that as
        // `lang: true`, which Set parses as a BCP-47 tag and rejects with
        // "Invalid BCP-47 'true'". Skip lang here so the canonical .latin/
        // .ea/.cs reader stays the single source of truth.
        "lang",
    };

    // Long-tail OOXML fallback: walk a properties container (rPr/pPr/...) and
    // surface every leaf child whose localName isn't already covered by the
    // curated reader. Shape is symmetric with the Add/Set side:
    //
    //   - child with no attrs            → Format[name] = true
    //     (toggle, matches GenericXmlQuery.TryCreateTypedChild bare-toggle).
    //   - child with one `val` attr only → Format[name] = val
    //     (scalar, matches GenericXmlQuery.TryCreateTypedChild val-leaf).
    //   - child with any other attrs     → Format[name.attr] = value per attr
    //     (dotted, matches TypedAttributeFallback.TrySet single-level shape
    //     `elementLocal.attrLocal`). Every typed attr surfaces, including
    //     `val` when accompanied by other attrs (so themed colors / multi-
    //     slot indents / spacing round-trip in full).
    //
    // Nested-children elements are emitted as raw flag toggles only — the
    // dotted reflection covers leaf attrs, and 3+ segment nested reflection
    // is intentionally out of scope (raw-XML escape handles the deep cases).
    private static void FillUnknownChildProps(OpenXmlElement? container, DocumentNode node)
    {
        if (container == null) return;
        foreach (var child in container.ChildElements)
        {
            var name = child.LocalName;
            if (string.IsNullOrEmpty(name)) continue;
            if (CuratedStyleLocalNames.Contains(name)) continue;
            if (child.ChildElements.Count > 0) continue;

            var typedAttrs = new System.Collections.Generic.List<DocumentFormat.OpenXml.OpenXmlAttribute>();
            foreach (var a in child.GetAttributes()) typedAttrs.Add(a);

            if (typedAttrs.Count == 0)
            {
                if (!node.Format.ContainsKey(name))
                    node.Format[name] = true;
                continue;
            }

            if (typedAttrs.Count == 1
                && typedAttrs[0].LocalName.Equals("val", System.StringComparison.OrdinalIgnoreCase))
            {
                if (!node.Format.ContainsKey(name))
                {
                    var raw = typedAttrs[0].Value ?? "";
                    // CT_OnOff toggles (kinsoku, snapToGrid, webHidden, suppressLineNumbers,
                    // …) expose a typed Val of type OnOffValue. Emit a canonical bool so the
                    // long-tail toggles match curated booleans (bold/keepNext → true/false)
                    // instead of the raw "1"/"0" the XML carries. Enum/string children
                    // (em, effect, textAlignment, …) keep their literal val.
                    //
                    // Every CT_OnOff element derives from the shared OnOffType base, so a
                    // direct `is` check is equivalent to inspecting Val's runtime type —
                    // and is trim/AOT-safe where GetType().GetProperty("Val") (IL2075) is not.
                    node.Format[name] =
                        child is DocumentFormat.OpenXml.Wordprocessing.OnOffType
                            ? OfficeCli.Core.ParseHelpers.IsTruthySafe(raw)
                            : raw;
                }
                continue;
            }

            // Multi-attribute element → dotted `<name>.<attr>` keys. Symmetric
            // with TypedAttributeFallback.TrySet on the Add/Set side, so
            // dump→replay round-trips through the same reflection path that
            // already accepts `ind.firstLine=240`, `spacing.line=480`, etc.
            foreach (var a in typedAttrs)
            {
                if (string.IsNullOrEmpty(a.LocalName)) continue;
                var key = $"{name}.{a.LocalName}";
                if (node.Format.ContainsKey(key)) continue;
                node.Format[key] = a.Value ?? "";
            }
        }
    }
}
