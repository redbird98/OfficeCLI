// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler
{
    private string AddChart(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        // CONSISTENCY(host-part-rel): same routing as AddPicture (round23 E) and
        // AddHyperlink (round23 C). When the parent paragraph lives in a Header/Footer
        // part, the chart rel must live on that part — otherwise r:id in headerN.xml
        // points to a rel only present in document.xml.rels and Word reports broken.
        OpenXmlPart chartMainPart = _doc.MainDocumentPart!;
        // parent may itself be a Header/Footer (e.g. /header[1]) when the chart is
        // appended directly, or a descendant paragraph (e.g. /header[1]/p[N]).
        var chartHeaderAnc = parent as Header ?? parent.Ancestors<Header>().FirstOrDefault();
        if (chartHeaderAnc != null)
        {
            var hp = _doc.MainDocumentPart!.HeaderParts
                .FirstOrDefault(p => ReferenceEquals(p.Header, chartHeaderAnc));
            if (hp != null) chartMainPart = hp;
        }
        else
        {
            var chartFooterAnc = parent as Footer ?? parent.Ancestors<Footer>().FirstOrDefault();
            if (chartFooterAnc != null)
            {
                var fp = _doc.MainDocumentPart!.FooterParts
                    .FirstOrDefault(p => ReferenceEquals(p.Footer, chartFooterAnc));
                if (fp != null) chartMainPart = fp;
            }
        }

        // Parse chart data. Use TryGetValue(case-insensitive) so reads
        // are recorded by TrackingPropertyDictionary.
        string chartType = "column";
        if (properties.TryGetValue("charttype", out var ctVal) || properties.TryGetValue("type", out ctVal))
            chartType = ctVal;
        var chartTitle = properties.GetValueOrDefault("title");
        var categories = Core.ChartHelper.ParseCategories(properties);
        var seriesData = Core.ChartHelper.ParseSeriesData(properties);

        if (seriesData.Count == 0)
            throw new ArgumentException("Chart requires data. Use: data=\"Series1:1,2,3;Series2:4,5,6\" " +
                "or series1=\"Revenue:100,200,300\"");

        // Dimensions (default: 15cm x 10cm)
        long chartCx = properties.TryGetValue("width", out var chartWStr) ? ParseEmu(chartWStr) : 5400000;
        long chartCy = properties.TryGetValue("height", out var chStr) ? ParseEmu(chStr) : 3600000;

        var docPropId = NextDocPropId();
        // BUG-R7-02 (T-2): explicit `name` prop was previously ignored —
        // dump emitted name=… on round-trip but Add silently dropped it,
        // so the chart's shape name reverted to its title every replay.
        // Honor caller intent first; fall back to title, then synthesize.
        // CONSISTENCY(empty-string-fallback): mirror AddPicture's
        // !IsNullOrEmpty guard — `??` only short-circuits on null, so a
        // literal name="" would otherwise pin the chart's shape name to
        // empty instead of falling through to title.
        var chartName = (properties.TryGetValue("name", out var chartNameOverride)
                         && !string.IsNullOrEmpty(chartNameOverride))
            ? chartNameOverride
            : (chartTitle ?? $"Chart {docPropId}");

        // Extended chart types (cx:chart) — funnel, treemap, sunburst, boxWhisker, histogram
        if (Core.ChartExBuilder.IsExtendedChartType(chartType))
        {
            var cxChartSpace = Core.ChartExBuilder.BuildExtendedChartSpace(
                chartType, chartTitle, categories, seriesData, properties);
            var extChartPart = chartMainPart.AddNewPart<ExtendedChartPart>();
            extChartPart.ChartSpace = cxChartSpace;
            extChartPart.ChartSpace.Save();

            // CONSISTENCY(chartex-sidecars): see PowerPointHandler.Add.Media.cs
            // for the full rationale. Word's chartEx host has the same hard
            // requirement on rId1 (embedded xlsx) + rId2 (style) + rId3 (colors).
            var embPart = extChartPart.AddNewPart<EmbeddedPackagePart>(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "rId1");
            var xlsxBytes = Core.ChartExResources.BuildMinimalEmbeddedXlsx(categories, seriesData);
            using (var emsr = new MemoryStream(xlsxBytes))
                embPart.FeedData(emsr);

            var stylePart = extChartPart.AddNewPart<ChartStylePart>("rId2");
            using (var styleStream = Core.ChartExResources.OpenChartStyleXml())
                stylePart.FeedData(styleStream);

            var colorPart = extChartPart.AddNewPart<ChartColorStylePart>("rId3");
            using (var colorStream = Core.ChartExResources.OpenChartColorStyleXml())
                colorPart.FeedData(colorStream);

            var cxRelId = chartMainPart.GetIdOfPart(extChartPart);
            var cxChartRef = new DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.RelId { Id = cxRelId };

            var cxInline = new DW.Inline(
                new DW.Extent { Cx = chartCx, Cy = chartCy },
                new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new DW.DocProperties { Id = docPropId, Name = chartName },
                new DW.NonVisualGraphicFrameDrawingProperties(),
                new A.Graphic(
                    new A.GraphicData(cxChartRef)
                    { Uri = "http://schemas.microsoft.com/office/drawing/2014/chartex" }
                )
            )
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            };

            var cxRun = new Run(new Drawing(cxInline));
            Paragraph cxPara;
            if (parent is Paragraph existingCxPara)
            {
                // CONSISTENCY(add-index): honor --index / --after / --before (#76).
                var cxChildren = existingCxPara.ChildElements.ToList();
                if (index.HasValue && index.Value < cxChildren.Count)
                    existingCxPara.InsertBefore(cxRun, cxChildren[index.Value]);
                else
                    existingCxPara.AppendChild(cxRun);
                cxPara = existingCxPara;
            }
            else
            {
                cxPara = new Paragraph(cxRun);
                AssignParaId(cxPara);
                InsertAtIndexOrAppend(parent, cxPara, index);
            }

            // Return document-order position so it matches the resolver
            // (GetAllWordCharts). CountWordCharts is insertion-order and
            // disagrees whenever --before/--after inserts mid-document.
            var cxAllCharts = GetAllWordCharts();
            var cxDocOrderIdx = cxAllCharts.FindIndex(c => ReferenceEquals(c.Inline, cxInline));
            return $"/chart[{(cxDocOrderIdx >= 0 ? cxDocOrderIdx + 1 : cxAllCharts.Count)}]";
        }

        // BUG-R6A(BUG2): Build chart content BEFORE adding part (invalid type or
        // a chart-type cardinality violation throws, which must not leave an
        // empty/malformed ChartPart on disk). Mirrors the Excel/Pptx ordering.
        var chartSpace = Core.ChartHelper.BuildChartSpace(chartType, chartTitle, categories, seriesData, properties);
        var chartPart = chartMainPart.AddNewPart<ChartPart>();
        chartPart.ChartSpace = chartSpace;

        // Apply deferred properties (axisTitle, dataLabels, etc.) via SetChartProperties
        // Must be called BEFORE Save() so the in-memory DOM is still available.
        // CONSISTENCY(tracking-deferred-filter): see PowerPointHandler.Add.Media.cs —
        // .Where() over TrackingPropertyDictionary marks every key consumed and
        // silently swallows real typos. Iterate Keys + TryGetValue per match instead.
        var deferredProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var dk in properties.Keys.ToList())
        {
            if (Core.ChartHelper.IsDeferredKey(dk) && properties.TryGetValue(dk, out var dv))
                deferredProps[dk] = dv;
        }
        if (deferredProps.Count > 0)
            Core.ChartHelper.SetChartProperties(chartPart, deferredProps);
        else
            chartPart.ChartSpace.Save();

        var chartRelId = chartMainPart.GetIdOfPart(chartPart);

        // Build Drawing/Inline with ChartReference
        var inline = new DW.Inline(
            new DW.Extent { Cx = chartCx, Cy = chartCy },
            new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
            new DW.DocProperties { Id = docPropId, Name = chartName },
            new DW.NonVisualGraphicFrameDrawingProperties(),
            new A.Graphic(
                new A.GraphicData(
                    new DocumentFormat.OpenXml.Drawing.Charts.ChartReference { Id = chartRelId }
                )
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }
            )
        )
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U
        };

        var chartRun = new Run(new Drawing(inline));
        Paragraph chartPara;
        if (parent is Paragraph existingChartPara)
        {
            // CONSISTENCY(add-index): honor --index / --after / --before (#76).
            var chartChildren = existingChartPara.ChildElements.ToList();
            if (index.HasValue && index.Value < chartChildren.Count)
                existingChartPara.InsertBefore(chartRun, chartChildren[index.Value]);
            else
                existingChartPara.AppendChild(chartRun);
            chartPara = existingChartPara;
        }
        else
        {
            chartPara = new Paragraph(chartRun);
            AssignParaId(chartPara);
            InsertAtIndexOrAppend(parent, chartPara, index);
        }

        // Return document-order position (matches GetAllWordCharts resolver).
        var allCharts = GetAllWordCharts();
        var docOrderIdx = allCharts.FindIndex(c => ReferenceEquals(c.Inline, inline));
        return $"/chart[{(docOrderIdx >= 0 ? docOrderIdx + 1 : allCharts.Count)}]";
    }

    private string AddPicture(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        if (!properties.TryGetValue("path", out var imgPath) && !properties.TryGetValue("src", out imgPath))
            throw new ArgumentException("'src' property is required for picture type");

        // Buffer the image bytes so we can both feed the image part and sniff
        // the native pixel dimensions for auto aspect-ratio calculations.
        var (rawStream, imgPartType) = OfficeCli.Core.ImageSource.Resolve(imgPath);
        using var rawStreamDispose = rawStream;
        using var imgStream = new MemoryStream();
        rawStream.CopyTo(imgStream);
        imgStream.Position = 0;

        var mainPart = _doc.MainDocumentPart!;
        // BUG-R14A: route through ResolveHostPart so the ImagePart and its
        // r:embed relationship land on whatever part actually holds the
        // <w:drawing> — header/footer AND footnote/endnote/comment. OOXML
        // resolves r:embed against the rels of the part containing the
        // drawing, so a picture added into a footnote/endnote/comment body
        // must register its image rel on word/_rels/footnotes.xml.rels (etc.),
        // not document.xml.rels — otherwise the blip dangles and validation
        // fails ([Semantic] r:embed does not exist). Previously this path did
        // its own Header/Footer-only resolution and defaulted footnote/endnote/
        // comment to MainDocumentPart. Mirrors the comment-hyperlink-rel fix
        // (BUG-R13B(BUG2)) that added the comments branch to ResolveHostPart.
        OpenXmlPart imgHostPart = ResolveHostPart(parent);

        // AddImagePart is a generic extension on parts implementing
        // ISupportedRelationship<…, ImagePart> (MainDocument / Header / Footer /
        // Footnotes / Endnotes / Comments). Dispatch by runtime type so the
        // rel lands on the correct part.
        ImagePart AddImg(PartTypeInfo t) => imgHostPart switch
        {
            MainDocumentPart mdp => mdp.AddImagePart(t),
            HeaderPart hp => hp.AddImagePart(t),
            FooterPart fp => fp.AddImagePart(t),
            FootnotesPart fnp => fnp.AddImagePart(t),
            EndnotesPart enp => enp.AddImagePart(t),
            WordprocessingCommentsPart cp => cp.AddImagePart(t),
            _ => throw new InvalidOperationException(
                $"Host part type {imgHostPart.GetType().Name} does not support image parts"),
        };

        string relId;
        string? svgRelId = null;
        Stream? fallbackDimStream = null;  // source for TryGetDimensions when raster is the fallback
        if (imgPartType == ImagePartType.Svg)
        {
            // OOXML SVG embedding: main blip points to a PNG fallback, and
            // a:blip/a:extLst carries an asvg:svgBlip referencing the SVG
            // part. Modern Office picks up the SVG; older versions render
            // the PNG. See SvgImageHelper for namespace/URI details.
            var svgPart = AddImg(ImagePartType.Svg);
            svgPart.FeedData(imgStream);
            imgStream.Position = 0;
            svgRelId = imgHostPart.GetIdOfPart(svgPart);

            MemoryStream pngStream;
            if (properties.TryGetValue("fallback", out var fallbackPath) && !string.IsNullOrWhiteSpace(fallbackPath))
            {
                var (fbRaw, fbType) = OfficeCli.Core.ImageSource.Resolve(fallbackPath);
                using var fbDispose = fbRaw;
                pngStream = new MemoryStream();
                fbRaw.CopyTo(pngStream);
                pngStream.Position = 0;
                var fbPart = AddImg(fbType);
                fbPart.FeedData(pngStream);
                pngStream.Position = 0;
                relId = imgHostPart.GetIdOfPart(fbPart);
            }
            else
            {
                var pngPart = AddImg(ImagePartType.Png);
                pngPart.FeedData(new MemoryStream(OfficeCli.Core.SvgImageHelper.TransparentPng1x1, writable: false));
                relId = imgHostPart.GetIdOfPart(pngPart);
                pngStream = new MemoryStream(OfficeCli.Core.SvgImageHelper.TransparentPng1x1, writable: false);
            }
            fallbackDimStream = pngStream;
        }
        else
        {
            var imagePart = AddImg(imgPartType);
            imagePart.FeedData(imgStream);
            imgStream.Position = 0;
            relId = imgHostPart.GetIdOfPart(imagePart);
        }

        // Determine dimensions. When only one axis is supplied, compute the
        // other from the image's native pixel aspect ratio. When neither is
        // supplied, width defaults to 6 inches and height follows the aspect
        // ratio (or a 4 inch fallback when the image header cannot be read).
        // CONSISTENCY(picture-size-alias): accept "w"/"h" as short aliases for
        // "width"/"height" — mirrors pptx shape and xlsx picture behavior.
        bool hasWidth = properties.TryGetValue("width", out var widthStr)
            || properties.TryGetValue("w", out widthStr);
        bool hasHeight = properties.TryGetValue("height", out var heightStr)
            || properties.TryGetValue("h", out heightStr);
        long cxEmu = hasWidth ? ParseEmu(widthStr!) : 5486400;  // 6 inches fallback
        long cyEmu = hasHeight ? ParseEmu(heightStr!) : 3657600; // 4 inches fallback

        if (!hasWidth || !hasHeight)
        {
            var dims = OfficeCli.Core.ImageSource.TryGetDimensions(imgStream);
            if (dims is { Width: > 0, Height: > 0 } d)
            {
                double ratio = (double)d.Height / d.Width;
                if (hasWidth && !hasHeight)
                    cyEmu = (long)(cxEmu * ratio);
                else if (!hasWidth && hasHeight)
                    cxEmu = (long)(cyEmu / ratio);
                else
                    cyEmu = (long)(cxEmu * ratio);
            }
        }

        // BUG-R5-02: data URIs (data:image/png;base64,iVBOR...) contain
        // multiple slashes inside the base64 payload, so Path.GetFileName
        // returns a meaningless tail like "png;base64,iVBOR..." which then
        // becomes both the picture name AND the alt text. Detect data: /
        // base64-blob inputs and fall back to a neutral placeholder unless
        // the caller supplied an explicit alt= or name=.
        string DefaultPictureName()
        {
            if (string.IsNullOrEmpty(imgPath)) return "image";
            if (imgPath.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) return "image";
            // Heuristic for raw base64 (no scheme): no path separator and length
            // is implausibly long for a real filename.
            if (imgPath.Length > 256 && imgPath.IndexOf('/') < 0 && imgPath.IndexOf('\\') < 0) return "image";
            try { return Path.GetFileName(imgPath); }
            catch { return "image"; }
        }
        // `name` is the picture's docPr Name (Word's Object Browser label;
        // typically "Picture 1", "Picture 2", …). `alt` is the docPr
        // Description (Word's Alt Text field, used by screen readers).
        // Source documents almost always set these to DIFFERENT values, so
        // collapsing both onto `altText` poisons dump round-trip — d2 would
        // emit the long alt text as the name. Take them separately; fall
        // back through name → alt → DefaultPictureName so callers passing
        // only one still get a sensible result.
        var pictureName = properties.TryGetValue("name", out var nameOverride) && !string.IsNullOrEmpty(nameOverride)
            ? nameOverride
            : (properties.TryGetValue("alt", out var altOverride) && !string.IsNullOrEmpty(altOverride)
                ? altOverride
                : DefaultPictureName());
        // altText: explicit `alt=` wins. When absent, leave the Description
        // attribute blank — auto-stamping `altText = pictureName` meant Get
        // surfaced a phantom `alt=<name>` key on dump round-trip for sources
        // whose docPr had no Description (07_example_online_convert.docx).
        // Empty string keeps the OOXML attr off entirely (DocProperties
        // serialises Description="" as no attr).
        var altText = properties.TryGetValue("alt", out var altOverride2) && !string.IsNullOrEmpty(altOverride2)
            ? altOverride2
            : "";

        var imgDocPropId = NextDocPropId();
        // BUG-DUMP-R29-1: parse the optional effectExtent prop ("l,t,r,b", 4
        // EMU ints — may be negative) captured by the dump emit. Word's inline
        // layout height depends on this margin; without it the rebuild
        // collapses to 0/0/0/0 and downstream content drifts up. Absent →
        // null → CreateImageRun/CreateAnchorImageRun default to 0/0/0/0
        // (interactive Add back-compat). Tolerant of whitespace/bad input.
        (long L, long T, long R, long B)? effectExtent = null;
        if (properties.TryGetValue("effectExtent", out var eeStr) && !string.IsNullOrWhiteSpace(eeStr))
        {
            var eeParts = eeStr.Split(',');
            if (eeParts.Length == 4
                && long.TryParse(eeParts[0].Trim(), out var eeL)
                && long.TryParse(eeParts[1].Trim(), out var eeT)
                && long.TryParse(eeParts[2].Trim(), out var eeR)
                && long.TryParse(eeParts[3].Trim(), out var eeB))
            {
                effectExtent = (eeL, eeT, eeR, eeB);
            }
        }
        Run imgRun;
        // BUG-R4-BT3: a non-"none" `wrap` value implies floating placement —
        // wrap only has meaning on a <wp:anchor>. Previously, callers passing
        // `wrap=square|tight|topBottom|behind|inFront` without an explicit
        // `anchor=true` got an inline picture and the wrap was silently
        // dropped (also affected dump round-trip of floating pictures).
        bool wrapImpliesAnchor = properties.TryGetValue("wrap", out var implicitWrap)
            && !string.IsNullOrEmpty(implicitWrap)
            && !string.Equals(implicitWrap, "none", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(implicitWrap, "inline", StringComparison.OrdinalIgnoreCase);
        // BUG-DUMP11-06: `anchor` is overloaded — historically a bool flag for
        // floating placement, but Get also surfaces the hyperlink anchor name
        // when the picture's run is wrapped in <w:hyperlink w:anchor="...">.
        // Treat bool-recognized values (true/false/yes/no/0/1/on/off) as the
        // floating switch; treat any other non-empty string as a hyperlink
        // bookmark name attached to the picture's drawing.
        bool hasAnchorProp = properties.TryGetValue("anchor", out var anchorVal)
            && !string.IsNullOrEmpty(anchorVal);
        bool anchorIsBool = hasAnchorProp && ParseHelpers.IsValidBooleanString(anchorVal);
        bool anchorIsFloating = hasAnchorProp && anchorIsBool && IsTruthy(anchorVal);
        string? hyperlinkAnchorName = hasAnchorProp && !anchorIsBool ? anchorVal : null;
        if (anchorIsFloating || wrapImpliesAnchor)
        {
            var wrapType = properties.GetValueOrDefault("wrap", "none");
            long hPos = properties.TryGetValue("hposition", out var hPosStr) ? ParseEmu(hPosStr) : 0;
            long vPos = properties.TryGetValue("vposition", out var vPosStr) ? ParseEmu(vPosStr) : 0;
            var hRel = properties.TryGetValue("hrelative", out var hRelStr)
                ? ParseHorizontalRelative(hRelStr)
                : DW.HorizontalRelativePositionValues.Margin;
            var vRel = properties.TryGetValue("vrelative", out var vRelStr)
                ? ParseVerticalRelative(vRelStr)
                : DW.VerticalRelativePositionValues.Margin;
            var behind = properties.TryGetValue("behindtext", out var behindStr) && IsTruthy(behindStr);
            // Relative <wp:align> keyword per axis (left/center/right;
            // top/bottom/center/inside/outside). When present it overrides the
            // posOffset form so an aligned float honours Word's relative
            // placement instead of collapsing to the margin origin.
            var hAlign = properties.TryGetValue("halign", out var hAlignStr) && !string.IsNullOrEmpty(hAlignStr)
                ? hAlignStr : null;
            var vAlign = properties.TryGetValue("valign", out var vAlignStr) && !string.IsNullOrEmpty(vAlignStr)
                ? vAlignStr : null;
            // BUG-DUMP-R26-1: round-trip the anchor z-order. CreateImageNode now
            // surfaces relativeHeight; honour it here so overlapping floats keep
            // their distinct stacking order instead of collapsing to 1U.
            uint relHeight = properties.TryGetValue("relativeHeight", out var relHeightStr)
                && uint.TryParse(relHeightStr, out var rh) ? rh : 1U;
            // Wrap distances (gap between the float and the text wrapping around
            // it): "T,B,L,R" EMU prop captured by the dump. Absent → null →
            // CreateAnchorImageRun keeps the interactive defaults.
            (uint T, uint B, uint L, uint R)? wrapDist = null;
            if (properties.TryGetValue("wrapDist", out var wdStr) && !string.IsNullOrWhiteSpace(wdStr))
            {
                var wd = wdStr.Split(',');
                if (wd.Length == 4
                    && uint.TryParse(wd[0].Trim(), out var wdT)
                    && uint.TryParse(wd[1].Trim(), out var wdB)
                    && uint.TryParse(wd[2].Trim(), out var wdL)
                    && uint.TryParse(wd[3].Trim(), out var wdR))
                {
                    wrapDist = (wdT, wdB, wdL, wdR);
                }
            }
            imgRun = CreateAnchorImageRun(relId, cxEmu, cyEmu, altText, wrapType, hPos, vPos, hRel, vRel, behind, imgDocPropId, pictureName, hAlign, vAlign, relHeight, effectExtent, wrapDist);
        }
        else
        {
            imgRun = CreateImageRun(relId, cxEmu, cyEmu, altText, imgDocPropId, pictureName, effectExtent);
        }

        // Wire the asvg:svgBlip extension after the run is built. Walking
        // the Drawing to find the Blip keeps CreateImageRun /
        // CreateAnchorImageRun signature-stable for non-SVG callers.
        if (svgRelId != null)
        {
            var addedBlip = imgRun.Descendants<A.Blip>().FirstOrDefault();
            if (addedBlip != null)
                OfficeCli.Core.SvgImageHelper.AppendSvgExtension(addedBlip, svgRelId);
        }

        // BUG-DUMP-R45-4: re-inject image-level visual content the dump captured
        // as verbatim XML (blip recolor/alpha children + spPr effectLst/outerShdw
        // drop shadow). These have no flat-key representation; the fixed
        // CreateImageRun/CreateAnchorImageRun rebuild would otherwise drop them.
        // Each prop is present only when the source carried that content, so a
        // plain picture is untouched.
        if (properties.TryGetValue("blipEffects", out var blipEffectsXml)
            && !string.IsNullOrWhiteSpace(blipEffectsXml))
        {
            ApplyBlipEffects(imgRun, blipEffectsXml);
        }
        // Whole-spPr verbatim (xfrm flip flags, a content <a:ext> that differs
        // from the frame's wp:extent, bwMode, explicit noFill/ln, effectLst) —
        // supersedes the narrower spEffects injection when present.
        if (properties.TryGetValue("spPrXml", out var spPrXmlVal)
            && !string.IsNullOrWhiteSpace(spPrXmlVal))
        {
            ApplySpPrVerbatim(imgRun, spPrXmlVal);
        }
        else if (properties.TryGetValue("spEffects", out var spEffectsXml)
            && !string.IsNullOrWhiteSpace(spEffectsXml))
        {
            ApplySpPrEffects(imgRun, spEffectsXml);
        }

        // BUG-DUMP-R51-1: round-trip a click-hyperlink on the image. A clickable
        // image (e.g. a logo linking to a URL) carries <a:hlinkClick r:id="…"> on
        // its <pic:cNvPr> (and, for external targets, a TargetMode="External"
        // relationship). The dump now surfaces the resolved URL as `link`; re-create
        // the external rel on the drawing's host part and attach the hlinkClick so
        // the image stays clickable instead of flattening to a plain image. A bare
        // anchor (internal bookmark, no scheme/relative target) is stored as the
        // hlinkClick @w:anchor attribute with no rel — mirrors the run-level
        // <w:hyperlink w:anchor> path. Reuses AddHyperlinkRelationship, the same
        // rel-creation helper the run-level hyperlink Add/Set paths use.
        if (properties.TryGetValue("link", out var linkVal) && !string.IsNullOrEmpty(linkVal))
        {
            var picCNvPr = imgRun.Descendants<PIC.NonVisualDrawingProperties>().FirstOrDefault();
            if (picCNvPr != null)
            {
                // <a:hlinkClick> always references its target through an r:id
                // relationship (DrawingML has no bare @anchor like w:hyperlink).
                // An absolute URI or relative target round-trips as a
                // TargetMode=External rel; a fragment "#anchor" round-trips as an
                // internal (isExternal=false) rel with Target="#anchor", matching
                // the run-level hyperlink Add path's fragment handling.
                bool isFragment = linkVal.StartsWith('#');
                Uri? linkUri;
                if (isFragment)
                {
                    linkUri = new Uri(linkVal, UriKind.Relative);
                }
                else if (Uri.TryCreate(linkVal, UriKind.Absolute, out linkUri))
                {
                    Core.HyperlinkUriValidator.RequireSafeScheme(linkVal, "link");
                }
                else
                {
                    Uri.TryCreate(linkVal, UriKind.Relative, out linkUri);
                }
                if (linkUri != null)
                {
                    var linkRelId = imgHostPart switch
                    {
                        MainDocumentPart mdp => mdp.AddHyperlinkRelationship(linkUri, !isFragment).Id,
                        HeaderPart hp => hp.AddHyperlinkRelationship(linkUri, !isFragment).Id,
                        FooterPart fp => fp.AddHyperlinkRelationship(linkUri, !isFragment).Id,
                        FootnotesPart fnp => fnp.AddHyperlinkRelationship(linkUri, !isFragment).Id,
                        EndnotesPart enp => enp.AddHyperlinkRelationship(linkUri, !isFragment).Id,
                        WordprocessingCommentsPart cp => cp.AddHyperlinkRelationship(linkUri, !isFragment).Id,
                        _ => mainPart.AddHyperlinkRelationship(linkUri, !isFragment).Id,
                    };
                    picCNvPr.HyperlinkOnClick = new A.HyperlinkOnClick { Id = linkRelId };
                }
            }
        }

        string resultPath;
        Paragraph imgPara;
        if (parent is Paragraph existingPara)
        {
            // Use ChildElements for index lookup to match ResolveAnchorPosition
            // (which counts pPr). If index points at pPr, clamp forward.
            var imgChildren = existingPara.ChildElements.ToList();
            if (index.HasValue && index.Value < imgChildren.Count)
            {
                var refElement = imgChildren[index.Value];
                if (refElement is ParagraphProperties)
                {
                    if (index.Value + 1 < imgChildren.Count)
                        existingPara.InsertBefore(imgRun, imgChildren[index.Value + 1]);
                    else
                        existingPara.AppendChild(imgRun);
                }
                else
                {
                    existingPara.InsertBefore(imgRun, refElement);
                }
            }
            else
            {
                existingPara.AppendChild(imgRun);
            }
            imgPara = existingPara;
            // CONSISTENCY(run-path-index): align the returned r[N] index with
            // navigation's r[N] resolution, which uses Descendants<Run>() and
            // skips comment-reference runs. GetAllRuns encapsulates both rules.
            var imgRunIdx = GetAllRuns(existingPara).IndexOf(imgRun) + 1;
            // CONSISTENCY(para-path-canonical): canonicalize to paraId-form.
            resultPath = $"{ReplaceTrailingParaSegment(parentPath, existingPara)}/r[{imgRunIdx}]";
        }
        else if (parent is TableCell imgCell)
        {
            // Insert image into existing first paragraph if empty, otherwise create new paragraph
            var firstCellPara = imgCell.Elements<Paragraph>().FirstOrDefault();
            if (firstCellPara != null && !firstCellPara.Elements<Run>().Any())
            {
                firstCellPara.AppendChild(imgRun);
                imgPara = firstCellPara;
            }
            else
            {
                imgPara = new Paragraph(imgRun);
                AssignParaId(imgPara);
                // Prevent fixed line spacing (inherited from Normal style) from
                // clipping the image to the text line height.
                imgPara.PrependChild(new ParagraphProperties(
                    new SpacingBetweenLines { Line = "240", LineRule = LineSpacingRuleValues.Auto }));
                imgCell.AppendChild(imgPara);
            }
            var imgPIdx = imgCell.Elements<Paragraph>().ToList().IndexOf(imgPara) + 1;
            resultPath = $"{parentPath}/{BuildParaPathSegment(imgPara, imgPIdx)}";
        }
        else
        {
            imgPara = new Paragraph(imgRun);
            AssignParaId(imgPara);
            // Prevent fixed line spacing (inherited from Normal style) from
            // clipping the image to the text line height.
            imgPara.PrependChild(new ParagraphProperties(
                new SpacingBetweenLines { Line = "240", LineRule = LineSpacingRuleValues.Auto }));

            // Use ChildElements for index lookup so that tables and sectPr
            // siblings do not shift the effective insertion position. This
            // matches ResolveAnchorPosition, which computes anchor indices
            // against ChildElements.
            var allChildren = parent.ChildElements.ToList();
            if (index.HasValue && index.Value < allChildren.Count)
            {
                var refElement = allChildren[index.Value];
                parent.InsertBefore(imgPara, refElement);
                var imgPIdx = parent.Elements<Paragraph>().ToList().IndexOf(imgPara) + 1;
                resultPath = $"{parentPath}/{BuildParaPathSegment(imgPara, imgPIdx)}";
            }
            else
            {
                AppendToParent(parent, imgPara);
                var imgPIdx = parent.Elements<Paragraph>().Count();
                resultPath = $"{parentPath}/{BuildParaPathSegment(imgPara, imgPIdx)}";
            }
        }

        // Apply run-level properties carried on the picture's Format (lang.*,
        // noproof, bold, color, …). Sources often stamp <w:lang> / <w:noProof>
        // on the picture's run — without this pass they were silently dropped
        // on replay, surfacing as phantom keys-missing on the second dump.
        // Iterate AddPicture's full property bag and apply anything the run-
        // formatting helper recognises; AddPicture-specific keys (width,
        // height, alt, name, wrap, anchor, …) are not in the helper's
        // vocabulary so they pass through untouched.
        OpenXmlCompositeElement? imgRunRPr = null;
        foreach (var (key, value) in properties)
        {
            var lk = key.ToLowerInvariant();
            if (lk is "width" or "height" or "alt" or "name" or "src"
                or "wrap" or "anchor" or "hposition" or "vposition"
                or "hrelative" or "vrelative" or "halign" or "valign" or "behindtext"
                or "tooltip" or "tgtframe" or "tgtframe" or "history" or "url"
                or "relid" or "id" or "contenttype" or "filesize"
                or "src.svg"
                // BUG-DUMP-R45-4: verbatim-XML props applied above, not rPr keys.
                or "blipeffects" or "speffects")
                continue;
            // BUG-DUMP-CROP: crop is a blipFill property, not a run-rPr key —
            // ApplyRunFormatting would silently drop it, so the dump→batch
            // `add picture --prop crop=…` op (emitted by CreateImageNode's crop
            // readback) never re-applied the source rectangle. Route it through
            // the shared writer that Set uses.
            if (lk is "crop" or "cropleft" or "cropright" or "croptop" or "cropbottom")
            {
                var blipFillAdd = imgRun.GetFirstChild<Drawing>()
                    ?.Descendants<DocumentFormat.OpenXml.Drawing.Pictures.BlipFill>().FirstOrDefault();
                if (blipFillAdd != null) ApplyCropToBlipFill(blipFillAdd, key, value);
                continue;
            }
            imgRunRPr ??= imgRun.GetFirstChild<RunProperties>() ?? imgRun.PrependChild(new RunProperties());
            ApplyRunFormatting(imgRunRPr, key, value);
        }

        // BUG-DUMP11-06: a hyperlink-wrapped picture's `anchor` attr (the
        // Word-level <w:hyperlink w:anchor="bookmark"> wrapping) round-trips
        // by re-wrapping the inserted Run in a fresh Hyperlink. Navigation's
        // run-parent-is-hyperlink branch already surfaces the anchor on the
        // picture node. Pass-through the optional metadata attrs (tooltip /
        // tgtFrame / history / url) for symmetry with AddHyperlink.
        if (hyperlinkAnchorName != null)
        {
            var hlWrap = new Hyperlink { Anchor = hyperlinkAnchorName };
            if (properties.TryGetValue("tooltip", out var picTip)) hlWrap.Tooltip = picTip;
            if ((properties.TryGetValue("tgtFrame", out var picTgt)
                 || properties.TryGetValue("tgtframe", out picTgt))
                && !string.IsNullOrEmpty(picTgt))
                hlWrap.TargetFrame = picTgt;
            if (properties.TryGetValue("history", out var picHist) && IsTruthy(picHist))
                hlWrap.History = OnOffValue.FromBoolean(true);

            var imgRunParent = imgRun.Parent;
            if (imgRunParent != null)
            {
                // Replace the run in-place with a Hyperlink wrapper so
                // sibling order and the resultPath (which addresses the run
                // via Descendants<Run>()) remain valid.
                imgRun.InsertAfterSelf(hlWrap);
                imgRun.Remove();
                hlWrap.AppendChild(imgRun);
            }
        }

        return resultPath;
    }

    // ==================== OLE Object Insertion ====================
    //
    // Inserts an <w:object> wrapper containing:
    //   1. VML shapetype _x0000_t75 (picture frame, well-known shape ID)
    //   2. VML v:shape bound to an icon preview ImagePart
    //   3. o:OLEObject naming the ProgID and referencing an
    //      EmbeddedObjectPart / EmbeddedPackagePart (the binary payload)
    //
    // Defaults are tuned so callers can just say `--type ole --prop src=...`:
    //   - ProgID auto-detected from src extension (via OleHelper)
    //   - Backing part kind auto-chosen (Package for .docx/.xlsx/.pptx, Object otherwise)
    //   - Icon preview = tiny PNG placeholder
    //   - Dimensions default to 2in × 0.75in (matches Office's show-as-icon frame)
    //
    // Caller can override: progId, width, height, icon (png/jpg/emf file path),
    // display (icon|content). display=content flips DrawAspect to "Content".
    // dump→batch round-trip for an ActiveX form-control run (<w:object> hosting
    // <w:control r:id> + a VML preview <v:imagedata r:id> — no o:OLEObject).
    // props carry the verbatim <w:r> XML plus one part{N}.relId/part{N}.data
    // pair per package part the object references (preview image, activeX
    // persistence XML) and part{N}.child{M}.* for parts nested under those
    // (the activeX binary blob). Parts are recreated with FRESH relationship
    // ids — the source ids would collide with the rebuilt main part's existing
    // rels — and the run XML's r:id refs are rewritten to match. Child parts
    // keep their SOURCE rel ids: they are scoped to the freshly created parent
    // part, so the verbatim part bytes' internal refs resolve untouched.
    private string AddActiveX(OpenXmlElement parent, string parentPath, Dictionary<string, string> properties)
    {
        properties ??= new Dictionary<string, string>();
        if (!properties.TryGetValue("runXml", out var axMarker) || string.IsNullOrEmpty(axMarker)
            || !axMarker.Contains("<w:control", StringComparison.Ordinal))
            throw new ArgumentException("activex requires --prop runXml containing a <w:control> element");
        return AddInlinedPartsRun(parent, parentPath, properties, "activex");
    }

    // SmartArt diagram run — same self-contained carrier as `add activex`:
    // verbatim run XML (the <w:drawing> whose <dgm:relIds> references the
    // data / layout / quickStyle / colors parts via r:dm/r:lo/r:qs/r:cs) plus
    // part{N}.* payloads, including the data part's nested rendered-drawing
    // child (diagramDrawing+xml, what Word actually rasterizes).
    private string AddDiagram(OpenXmlElement parent, string parentPath, Dictionary<string, string> properties)
    {
        properties ??= new Dictionary<string, string>();
        if (!properties.TryGetValue("runXml", out var dgMarker) || string.IsNullOrEmpty(dgMarker)
            || !dgMarker.Contains("relIds", StringComparison.Ordinal))
            throw new ArgumentException("diagram requires --prop runXml containing a <dgm:relIds> element");
        return AddInlinedPartsRun(parent, parentPath, properties, "diagram");
    }

    // Legacy VML shape run (<w:pict>) — textboxes whose content carries
    // hyperlinks (external rels) or v:imagedata image parts. Same carrier.
    private string AddVmlShape(OpenXmlElement parent, string parentPath, Dictionary<string, string> properties)
    {
        properties ??= new Dictionary<string, string>();
        if (!properties.TryGetValue("runXml", out var vmlMarker) || string.IsNullOrEmpty(vmlMarker)
            || !vmlMarker.Contains("<w:pict", StringComparison.Ordinal))
            throw new ArgumentException("vmlshape requires --prop runXml containing a <w:pict> element");
        return AddInlinedPartsRun(parent, parentPath, properties, "vmlshape");
    }

    // Modern DrawingML shape run (<w:drawing> hosting wps:wsp) whose spPr /
    // blipFill references image parts (a cover-page fern graphic). Same
    // carrier; previously the emitter scrubbed the blipFill to a neutral
    // solid fill and the bitmap was lost.
    private string AddDrawingShape(OpenXmlElement parent, string parentPath, Dictionary<string, string> properties)
    {
        properties ??= new Dictionary<string, string>();
        if (!properties.TryGetValue("runXml", out var dsMarker) || string.IsNullOrEmpty(dsMarker)
            || !dsMarker.Contains("<w:drawing", StringComparison.Ordinal))
            throw new ArgumentException("drawingshape requires --prop runXml containing a <w:drawing> element");
        return AddInlinedPartsRun(parent, parentPath, properties, "drawingshape");
    }

    private string AddInlinedPartsRun(OpenXmlElement parent, string parentPath, Dictionary<string, string> properties, string opName)
    {
        var runXml = properties["runXml"];
        var mainPart = _doc.MainDocumentPart!;
        // CONSISTENCY(host-part-rel): same routing as AddOle — parts referenced
        // from a header/footer-hosted run must attach to that part.
        OpenXmlPart hostPart = mainPart;
        {
            var headerAncestor = parent as Header ?? parent.Ancestors<Header>().FirstOrDefault();
            if (headerAncestor != null)
            {
                var hp = mainPart.HeaderParts.FirstOrDefault(p => ReferenceEquals(p.Header, headerAncestor));
                if (hp != null) hostPart = hp;
            }
            else
            {
                var footerAncestor = parent as Footer ?? parent.Ancestors<Footer>().FirstOrDefault();
                if (footerAncestor != null)
                {
                    var fp = mainPart.FooterParts.FirstOrDefault(p => ReferenceEquals(p.Footer, footerAncestor));
                    if (fp != null) hostPart = fp;
                }
            }
        }

        var rewriteRelIds = MaterializeInlinedParts(hostPart, properties, opName);

        var rewritten = rewriteRelIds(runXml);

        var axRun = new Run(rewritten);

        string resultPath;
        if (parent is Paragraph axPara)
        {
            axPara.AppendChild(axRun);
            var axRunIdx = GetAllRuns(axPara).IndexOf(axRun) + 1;
            // CONSISTENCY(para-path-canonical): canonicalize to paraId-form.
            resultPath = $"{ReplaceTrailingParaSegment(parentPath, axPara)}/r[{axRunIdx}]";
        }
        else if (parent is TableCell axCell)
        {
            var firstCellPara = axCell.Elements<Paragraph>().FirstOrDefault();
            Paragraph hostPara;
            if (firstCellPara != null && !firstCellPara.Elements<Run>().Any())
            {
                firstCellPara.AppendChild(axRun);
                hostPara = firstCellPara;
            }
            else
            {
                hostPara = new Paragraph(axRun);
                AssignParaId(hostPara);
                axCell.AppendChild(hostPara);
            }
            var axPIdx = axCell.Elements<Paragraph>().ToList().IndexOf(hostPara) + 1;
            var axCellRunIdx = GetAllRuns(hostPara).IndexOf(axRun) + 1;
            resultPath = $"{parentPath}/{BuildParaPathSegment(hostPara, axPIdx)}/r[{axCellRunIdx}]";
        }
        else
        {
            var hostPara = new Paragraph(axRun);
            AssignParaId(hostPara);
            AppendToParent(parent, hostPara);
            var axPIdx = parent.Elements<Paragraph>().ToList().IndexOf(hostPara) + 1;
            resultPath = $"{parentPath}/{BuildParaPathSegment(hostPara, axPIdx)}/r[1]";
        }
        return resultPath;
    }

    // Shared by AddInlinedPartsRun and the sdtXml carrier in AddSdt: create
    // every part{N} (with part{N}.child{M}) and ext{N} relationship on
    // <paramref name="hostPart"/>, then return the two-phase rel-id rewriter
    // mapping every source id to its freshly assigned one.
    private Func<string, string> MaterializeInlinedParts(OpenXmlPart hostPart, Dictionary<string, string> properties, string opName)
    {
        // Pass 1: create every top-level part empty, so the complete old→new
        // id map exists before any content is written. A collected XML part's
        // bytes can reference a SIBLING part's host-part relationship (the
        // SmartArt data part's <dsp:dataModelExt relId> points at the
        // rendered-drawing part), so the rewrite below must cover all ids.
        var idMap = new List<(string OldId, string NewId)>();
        var pending = new List<(OpenXmlPart Part, byte[] Bytes, string Ct, int Pi)>();
        for (int pi = 1; properties.TryGetValue($"part{pi}.relId", out var oldRelId); pi++)
        {
            var dataUri = properties.GetValueOrDefault($"part{pi}.data");
            if (string.IsNullOrEmpty(oldRelId)
                || !OfficeCli.Core.OleHelper.TryDecodeDataUri(dataUri, out var bytes, out var ct)
                || bytes.Length == 0)
                throw new ArgumentException($"{opName} part{pi} requires relId and a non-empty data: URI");

            var created = CreateInlinedPart(hostPart, ct)
                ?? throw new ArgumentException($"{opName} part{pi}: unsupported content type '{ct}'");
            pending.Add((created, bytes, ct, pi));
            idMap.Add((oldRelId!, hostPart.GetIdOfPart(created)));
        }

        // External relationships (hyperlinks inside a VML textbox, linked
        // content): no part bytes — recreate the rel on the host part with a
        // fresh id and route the source id through the same rewrite map.
        const string HyperlinkRelType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
        for (int ei = 1; properties.TryGetValue($"ext{ei}.relId", out var extOldId); ei++)
        {
            var extType = properties.GetValueOrDefault($"ext{ei}.type");
            var extTarget = properties.GetValueOrDefault($"ext{ei}.target");
            if (string.IsNullOrEmpty(extOldId) || string.IsNullOrEmpty(extType) || string.IsNullOrEmpty(extTarget))
                throw new ArgumentException($"{opName} ext{ei} requires relId, type and target");
            var extUri = new Uri(extTarget, UriKind.RelativeOrAbsolute);
            var newExtId = extType == HyperlinkRelType
                ? hostPart.AddHyperlinkRelationship(extUri, true).Id
                : hostPart.AddExternalRelationship(extType, extUri).Id;
            idMap.Add((extOldId!, newExtId));
        }

        // Two-phase rewrite (shared by the run XML and every inlined XML
        // part's bytes): a freshly assigned id can equal a *different* source
        // id still pending replacement, so route through unique placeholders.
        string RewriteRelIds(string xml)
        {
            for (int i = 0; i < idMap.Count; i++)
                xml = xml.Replace($"\"{idMap[i].OldId}\"", $"\"__OCLI_AXREL_{i}__\"", StringComparison.Ordinal);
            for (int i = 0; i < idMap.Count; i++)
                xml = xml.Replace($"\"__OCLI_AXREL_{i}__\"", $"\"{idMap[i].NewId}\"", StringComparison.Ordinal);
            return xml;
        }

        // Pass 2: feed content (XML parts get their host-part rel refs
        // rewritten; binary parts stay verbatim) and attach child parts.
        foreach (var (created, bytes, ct, pi) in pending)
        {
            var feedBytes = bytes;
            if (ct.EndsWith("+xml", StringComparison.OrdinalIgnoreCase))
            {
                var text = System.Text.Encoding.UTF8.GetString(bytes);
                var rewrittenText = RewriteRelIds(text);
                if (!ReferenceEquals(rewrittenText, text))
                    feedBytes = System.Text.Encoding.UTF8.GetBytes(rewrittenText);
            }
            using (var ms = new MemoryStream(feedBytes))
                created.FeedData(ms);

            for (int ci = 1; properties.TryGetValue($"part{pi}.child{ci}.relId", out var childRelId); ci++)
            {
                var childUri = properties.GetValueOrDefault($"part{pi}.child{ci}.data");
                if (string.IsNullOrEmpty(childRelId)
                    || !OfficeCli.Core.OleHelper.TryDecodeDataUri(childUri, out var cbytes, out var cct)
                    || cbytes.Length == 0)
                    throw new ArgumentException($"{opName} part{pi}.child{ci} requires relId and a non-empty data: URI");
                var childPart = CreateInlinedChildPart(created, cct, childRelId!)
                    ?? throw new ArgumentException($"{opName} part{pi}.child{ci}: unsupported content type '{cct}'");
                using var cms = new MemoryStream(cbytes);
                childPart.FeedData(cms);
            }
        }

        return RewriteRelIds;
    }

    // Part factory for the inlined-parts carriers. Top-level parts hang off the
    // run's host part (main document / header / footer); returns null for an
    // unrecognized content type so the caller surfaces a clear error.
    private static OpenXmlPart? CreateInlinedPart(OpenXmlPart hostPart, string ct) => ct switch
    {
        "application/vnd.ms-office.activeX+xml"
            => hostPart.AddNewPart<EmbeddedControlPersistencePart>(ct, null),
        "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"
            => hostPart.AddNewPart<DiagramDataPart>(ct, null),
        "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"
            => hostPart.AddNewPart<DiagramLayoutDefinitionPart>(ct, null),
        "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"
            => hostPart.AddNewPart<DiagramStylePart>(ct, null),
        "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"
            => hostPart.AddNewPart<DiagramColorsPart>(ct, null),
        // The rendered-drawing part referenced from the data part's
        // dataModelExt extension — its relationship lives on the MAIN part
        // (ms diagramDrawing rel type), so it is a top-level part here.
        "application/vnd.ms-office.drawingml.diagramDrawing+xml"
            => hostPart.AddNewPart<DiagramPersistLayoutPart>(ct, null),
        _ when ct.StartsWith("image/", StringComparison.OrdinalIgnoreCase)
            => hostPart.AddNewPart<ImagePart>(ct, null),
        _ => null,
    };

    // Child parts keep their SOURCE rel id (scoped to the freshly created
    // parent part, so the verbatim parent bytes' internal refs resolve).
    private static OpenXmlPart? CreateInlinedChildPart(OpenXmlPart parent, string ct, string relId) => ct switch
    {
        "application/vnd.ms-office.activeX" or "application/vnd.ms-office.activeX+xml"
            => parent.AddNewPart<EmbeddedControlPersistenceBinaryDataPart>(ct, relId),
        // The rendered-drawing child of a diagram data part — what Word
        // actually rasterizes for modern SmartArt.
        "application/vnd.ms-office.drawingml.diagramDrawing+xml"
            => parent.AddNewPart<DiagramPersistLayoutPart>(ct, relId),
        _ when ct.StartsWith("image/", StringComparison.OrdinalIgnoreCase)
            => parent.AddNewPart<ImagePart>(ct, relId),
        _ => null,
    };


    private string AddOle(OpenXmlElement parent, string parentPath, int? index, Dictionary<string, string> properties)
    {
        properties ??= new Dictionary<string, string>();
        var srcPath = OfficeCli.Core.OleHelper.RequireSource(properties);
        OfficeCli.Core.OleHelper.WarnOnUnknownOleProps(properties);

        var mainPart = _doc.MainDocumentPart!;

        // Determine the host part that owns the parent element.
        // For /header[N] or /footer[N], the parent lives inside a
        // HeaderPart/FooterPart, so the embedded payload AND icon ImagePart
        // relationships must be attached to that part — not to
        // MainDocumentPart — otherwise OpenXmlValidator rejects the
        // cross-part r:id with a NullReferenceException.
        OpenXmlPart hostPart = mainPart;
        {
            var headerAncestor = parent as Header ?? parent.Ancestors<Header>().FirstOrDefault();
            if (headerAncestor != null)
            {
                var hp = mainPart.HeaderParts.FirstOrDefault(p => ReferenceEquals(p.Header, headerAncestor));
                if (hp != null) hostPart = hp;
            }
            else
            {
                var footerAncestor = parent as Footer ?? parent.Ancestors<Footer>().FirstOrDefault();
                if (footerAncestor != null)
                {
                    var fp = mainPart.FooterParts.FirstOrDefault(p => ReferenceEquals(p.Footer, footerAncestor));
                    if (fp != null) hostPart = fp;
                }
            }
        }

        // 1. Create the embedded binary payload part and rel id on the host part.
        // 2. Resolve ProgID.
        string embedRelId;
        string progId;
        if (OfficeCli.Core.OleHelper.TryDecodeDataUri(srcPath, out var embedBytes, out var dataCt))
        {
            // dump→batch round-trip: src is a data: URI carrying the embedded
            // payload exactly as stored (raw package, or CFB-wrapped Ole10Native
            // for generic objects). The dump also forwards oleKind / contentType
            // / embedExt so we rebuild the same part class + content type +
            // target extension and feed the bytes verbatim (no extension sniffing,
            // no CFB re-wrap).
            var oleKind = properties.GetValueOrDefault("oleKind")
                ?? properties.GetValueOrDefault("olekind")
                ?? (dataCt.Contains("oleObject", StringComparison.OrdinalIgnoreCase) ? "object" : "package");
            var contentType = properties.GetValueOrDefault("contentType")
                ?? properties.GetValueOrDefault("contenttype")
                ?? (string.IsNullOrEmpty(dataCt) ? "application/octet-stream" : dataCt);
            var embedExt = properties.GetValueOrDefault("embedExt")
                ?? properties.GetValueOrDefault("embedext");
            (embedRelId, _) = OfficeCli.Core.OleHelper.AddEmbeddedPartFromBytes(
                hostPart, embedBytes, oleKind, contentType, embedExt);
            // No file extension to sniff ProgID from — the dump always forwards
            // it, so require it explicitly rather than guessing.
            progId = properties.GetValueOrDefault("progId")
                ?? properties.GetValueOrDefault("progid")
                ?? throw new ArgumentException(
                    "inline ole payload (data: src) requires an explicit --prop progId");
            OfficeCli.Core.OleHelper.ValidateProgId(progId);
        }
        else
        {
            (embedRelId, _) = OfficeCli.Core.OleHelper.AddEmbeddedPart(hostPart, srcPath, _filePath);
            // ProgID: explicit > auto-detected from extension.
            progId = OfficeCli.Core.OleHelper.ResolveProgId(properties, srcPath);
        }

        // 3. Create the icon preview ImagePart on the host part (same part
        //    that owns the OLE element itself). Attaching to MainDocumentPart
        //    when the OLE lives in a header/footer would produce a dangling
        //    cross-part relationship — see host part resolution above.
        var (_, iconRelId) = OfficeCli.Core.OleHelper.CreateIconPart(hostPart, properties);

        // 4. Dimensions. Word VML shapes take points in their style string.
        //    Defaults match OleHelper's 2in × 0.75in icon frame.
        long cxEmu = properties.TryGetValue("width", out var wStr)
            ? ParseEmu(wStr) : OfficeCli.Core.OleHelper.DefaultOleWidthEmu;
        long cyEmu = properties.TryGetValue("height", out var hStr)
            ? ParseEmu(hStr) : OfficeCli.Core.OleHelper.DefaultOleHeightEmu;
        // EMU → points (914400 EMU/inch, 72 points/inch).
        double cxPt = cxEmu / EmuConverter.EmuPerPointF;
        double cyPt = cyEmu / EmuConverter.EmuPerPointF;
        // Twips for w:dxaOrig/w:dyaOrig (20 twips/point).
        long cxTwips = (long)(cxPt * 20);
        long cyTwips = (long)(cyPt * 20);

        // 5. DrawAspect: "Icon" (default) or "Content" (live preview).
        // Strict validation: unknown values throw rather than silently
        // falling back to Icon — see OleHelper.NormalizeOleDisplay.
        var display = OfficeCli.Core.OleHelper.NormalizeOleDisplay(
            properties.GetValueOrDefault("display", "icon"));
        var drawAspect = display == "content" ? "Content" : "Icon";

        // 6. ObjectID: VML requires a unique "_nnnnnnnnnn" token.
        //    Count existing OLE objects and assign a monotonic id so two
        //    OLEs added within the same wallclock second don't collide
        //    (the old scheme used ToUnixTimeSeconds()).
        var existingOleCount = mainPart.Document?.Body?.Descendants<EmbeddedObject>().Count() ?? 0;
        var oleSeq = existingOleCount + 1;
        var objectId = "_" + (1000000000 + oleSeq);

        // 7. Build the w:object XML. The shapetype + shape + OLEObject
        //    triple is the canonical form Word itself writes for OLE.
        //    ShapeID must also be unique per OLE in the document — base it
        //    on the OLE sequence (not NextDocPropId, which is shared with
        //    Drawing DocProperties and can collide). D4 gives 9999 slots.
        var shapeId = $"_x0000_i1{oleSeq:D4}";

        // Optional friendly name → v:shape alt="..." attribute.
        // CONSISTENCY(ole-name): the VML CT_OleObject complex type has no
        // Name attribute (valid attrs: Type/ProgID/ShapeID/DrawAspect/
        // ObjectID/r:id/UpdateMode/LinkType/LockedField/FieldCodes — see
        // DocumentFormat.OpenXml.Vml.Office.OleObject). Writing Name= on
        // o:OLEObject produces a schema validation error. Use the
        // surrounding v:shape element's "alt" attribute (Alternate Text,
        // closest semantic match in VML) for the friendly name. Get reads
        // it back from the same place, preserving Format["name"] round-trip.
        var shapeAltAttr = "";
        if (properties.TryGetValue("name", out var oleName) && !string.IsNullOrEmpty(oleName))
            shapeAltAttr = $" alt=\"{System.Security.SecurityElement.Escape(oleName)}\"";

        // CONSISTENCY(ole-shapetype-dedup): v:shapetype id="_x0000_t75" must be
        // unique across the whole document.xml — OOXML validation rejects
        // duplicate shapetype ids. If the document already has an
        // _x0000_t75 shapetype (left over from a prior picture/OLE insert),
        // skip re-emitting it and reference the existing one from v:shape.
        var shapetypeAlreadyExists = false;
        foreach (var existingObj in mainPart.Document?.Body?.Descendants<EmbeddedObject>() ?? Enumerable.Empty<EmbeddedObject>())
        {
            foreach (var st in existingObj.Descendants().Where(e => e.LocalName == "shapetype"))
            {
                var idAttr = st.GetAttributes().FirstOrDefault(a => a.LocalName == "id");
                if (idAttr.Value == "_x0000_t75") { shapetypeAlreadyExists = true; break; }
            }
            if (shapetypeAlreadyExists) break;
        }

        var shapetypeXml = shapetypeAlreadyExists ? "" : """
<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
<v:stroke joinstyle="miter"/>
<v:formulas>
<v:f eqn="if lineDrawn pixelLineWidth 0"/>
<v:f eqn="sum @0 1 0"/>
<v:f eqn="sum 0 0 @1"/>
<v:f eqn="prod @2 1 2"/>
<v:f eqn="prod @3 21600 pixelWidth"/>
<v:f eqn="prod @3 21600 pixelHeight"/>
<v:f eqn="sum @0 0 1"/>
<v:f eqn="prod @6 1 2"/>
<v:f eqn="prod @7 21600 pixelWidth"/>
<v:f eqn="sum @8 21600 0"/>
<v:f eqn="prod @7 21600 pixelHeight"/>
<v:f eqn="sum @10 21600 0"/>
</v:formulas>
<v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
<o:lock v:ext="edit" aspectratio="t"/>
</v:shapetype>
""";

        var oleXml = $"""
<w:object xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:dxaOrig="{cxTwips}" w:dyaOrig="{cyTwips}">
{shapetypeXml}<v:shape id="{shapeId}" type="#_x0000_t75" style="width:{cxPt.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}pt;height:{cyPt.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)}pt" o:ole=""{shapeAltAttr}>
<v:imagedata r:id="{iconRelId}" o:title=""/>
</v:shape>
<o:OLEObject Type="Embed" ProgID="{System.Security.SecurityElement.Escape(progId)}" ShapeID="{shapeId}" DrawAspect="{drawAspect}" ObjectID="{objectId}" r:id="{embedRelId}"/>
</w:object>
""";
        var oleObject = new EmbeddedObject(oleXml);

        // 8. Wrap in a Run and insert it, mirroring the AddPicture positional logic.
        var oleRun = new Run(oleObject);

        // If the parent is a block-level SDT, insert into its SdtContentBlock
        // (creating it if missing) instead of appending directly to the SdtBlock.
        // Direct SdtBlock child paragraphs violate the schema and get silently
        // stripped by Word on reload — which previously broke OLE persistence
        // across reopen when added inside an SDT container. See
        // OleTestTeamRound6.Word_OleInsideSdt_QueryFindsOle.
        if (parent is SdtBlock sdtBlockParent)
        {
            var contentBlock = sdtBlockParent.GetFirstChild<SdtContentBlock>();
            if (contentBlock == null)
            {
                contentBlock = new SdtContentBlock();
                sdtBlockParent.AppendChild(contentBlock);
            }
            parent = contentBlock;
        }
        // Inline SDT runs live inside a w:p parent: route the OLE to that
        // surrounding paragraph so insertion follows the normal run path.
        else if (parent is SdtRun sdtRunParent)
        {
            var contentRun = sdtRunParent.GetFirstChild<SdtContentRun>();
            if (contentRun != null)
                contentRun.AppendChild(oleRun);
            else
                sdtRunParent.AppendChild(new SdtContentRun(oleRun));
            var parentParaInline = sdtRunParent.Ancestors<Paragraph>().FirstOrDefault();
            if (parentParaInline != null)
            {
                var runs = GetAllRuns(parentParaInline);
                var runIdxInline = runs.IndexOf(oleRun) + 1;
                // CONSISTENCY(para-path-canonical): canonicalize when the
                // SDT lives directly inside a paragraph (parentPath ends in
                // /p[...]); otherwise (SDT in a cell) parentPath does not
                // end in /p[...] and ReplaceTrailingParaSegment is a no-op.
                return $"{ReplaceTrailingParaSegment(parentPath, parentParaInline)}/r[{runIdxInline}]";
            }
            return parentPath + "/r[1]";
        }

        string resultPath;
        if (parent is Paragraph existingPara)
        {
            // Use ChildElements for index lookup to match ResolveAnchorPosition.
            var oleChildren = existingPara.ChildElements.ToList();
            if (index.HasValue && index.Value < oleChildren.Count)
            {
                var refElement = oleChildren[index.Value];
                if (refElement is ParagraphProperties)
                {
                    if (index.Value + 1 < oleChildren.Count)
                        existingPara.InsertBefore(oleRun, oleChildren[index.Value + 1]);
                    else
                        existingPara.AppendChild(oleRun);
                }
                else
                {
                    existingPara.InsertBefore(oleRun, refElement);
                }
            }
            else
            {
                existingPara.AppendChild(oleRun);
            }
            var olePIdx = 1;
            foreach (var para in parent.Parent?.Elements<Paragraph>() ?? Enumerable.Empty<Paragraph>())
            {
                if (ReferenceEquals(para, existingPara)) break;
                olePIdx++;
            }
            var oleRunIdx = GetAllRuns(existingPara).IndexOf(oleRun) + 1;
            // CONSISTENCY(para-path-canonical): canonicalize to paraId-form.
            resultPath = $"{ReplaceTrailingParaSegment(parentPath, existingPara)}/r[{oleRunIdx}]";
        }
        else if (parent is TableCell oleCell)
        {
            var firstCellPara = oleCell.Elements<Paragraph>().FirstOrDefault();
            Paragraph olePara;
            if (firstCellPara != null && !firstCellPara.Elements<Run>().Any())
            {
                firstCellPara.AppendChild(oleRun);
                olePara = firstCellPara;
            }
            else
            {
                olePara = new Paragraph(oleRun);
                AssignParaId(olePara);
                oleCell.AppendChild(olePara);
            }
            var olePIdx = oleCell.Elements<Paragraph>().ToList().IndexOf(olePara) + 1;
            // CONSISTENCY(ole-run-path): same /r[1] suffix as the else branch
            // below — the OLE run is the addressable target, not the paragraph.
            var oleCellRunIdx = GetAllRuns(olePara).IndexOf(oleRun) + 1;
            resultPath = $"{parentPath}/{BuildParaPathSegment(olePara, olePIdx)}/r[{oleCellRunIdx}]";
        }
        else
        {
            var olePara = new Paragraph(oleRun);
            AssignParaId(olePara);
            var allChildren = parent.ChildElements.ToList();
            if (index.HasValue && index.Value < allChildren.Count)
            {
                var refElement = allChildren[index.Value];
                parent.InsertBefore(olePara, refElement);
            }
            else
            {
                AppendToParent(parent, olePara);
            }
            var olePIdx = parent.Elements<Paragraph>().ToList().IndexOf(olePara) + 1;
            // Return the /r[1] address so callers can Set/Get/Remove the
            // OLE run directly. Picture's Add returns a paragraph-level
            // path because the paragraph Set is meaningful (font, style);
            // for OLE, the only interesting target is the run itself.
            resultPath = $"{parentPath}/{BuildParaPathSegment(olePara, olePIdx)}/r[1]";
        }
        return resultPath;
    }
}
