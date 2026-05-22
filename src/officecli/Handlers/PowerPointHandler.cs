// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class PowerPointHandler : IDocumentHandler
{
    private readonly PresentationDocument _doc;
    private readonly string _filePath;
    private HashSet<uint> _usedShapeIds = new();
    private uint _nextShapeId = 10000;
    public int LastFindMatchCount { get; internal set; }

    /// <summary>
    /// Set true by Add/Set/Remove/RawSet, consumed by Save/Dispose to decide
    /// whether to stamp <c>docProps/custom.xml</c> with an OfficeCLI audit
    /// trail. Pure Get/Query sessions leave this false.
    /// </summary>
    internal bool Modified { get; set; }

    // Backing FileStream when we open via stream (shared-read mode). null
    // when the package owns its own file handle via PresentationDocument.Open(path).
    private FileStream? _backingStream;

    public PowerPointHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        // Open via a shared FileStream so external readers (e.g. test harness
        // ZipFile.OpenRead while the handler is alive) don't hit the macOS
        // flock exclusive lock that PresentationDocument.Open(path, editable)
        // would acquire. The package writes through to the stream; we call
        // _doc.Save() in Dispose() to flush before closing the stream.
        var share = editable ? FileShare.Read : FileShare.ReadWrite;
        var access = editable ? FileAccess.ReadWrite : FileAccess.Read;
        _backingStream = new FileStream(filePath, FileMode.Open, access, share);
        _doc = PresentationDocument.Open(_backingStream, editable);
        if (editable)
            InitShapeIdCounter();
    }

    /// <summary>
    /// Get the slide dimensions from the presentation. Falls back to 16:9 (33.867cm × 19.05cm).
    /// </summary>
    private (long width, long height) GetSlideSize()
    {
        var sldSz = _doc.PresentationPart?.Presentation?.GetFirstChild<SlideSize>();
        return (sldSz?.Cx?.Value ?? SlideSizeDefaults.Widescreen16x9Cx, sldSz?.Cy?.Value ?? SlideSizeDefaults.Widescreen16x9Cy);
    }

    // ==================== Raw Layer ====================

    // CONSISTENCY(zip-uri-lookup): see ExcelHandler.cs / RawXmlHelper —
    // any partPath ending in `.xml` is resolved as a literal zip URI via
    // the package's part tree, no per-handler alias table needed.

    public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null)
    {
        if (partPath == null) throw new ArgumentNullException(nameof(partPath));
        var presentationPart = _doc.PresentationPart;
        if (presentationPart == null) return "(empty)";

        if (RawXmlHelper.IsZipUriPath(partPath))
        {
            var xml = RawXmlHelper.TryReadByZipUri(_doc, _filePath, partPath)
                ?? throw new ArgumentException(
                    $"Unknown part: {partPath}. The path was treated as a zip-internal URI " +
                    $"but no matching part exists in the package. " +
                    $"Use semantic paths (/presentation, /slide[N], /slideMaster[N]) for stable identification.");
            return xml;
        }

        if (partPath == "/" || partPath == "/presentation")
            return presentationPart.Presentation?.OuterXml ?? "(empty)";

        if (partPath == "/theme")
            return presentationPart.ThemePart?.Theme?.OuterXml ?? "(no theme)";

        var slideMatch = Regex.Match(partPath, @"^/slide\[(\d+)\]$");
        if (slideMatch.Success)
        {
            var idx = int.Parse(slideMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx >= 1 && idx <= slideParts.Count)
                return GetSlide(slideParts[idx - 1]).OuterXml;
            throw new ArgumentException($"slide[{idx}] not found (total: {slideParts.Count})");
        }

        // CONSISTENCY(raw-rawset-symmetry): RawSet supports master/layout/noteSlide;
        // Raw must too, otherwise users can't read back what they just wrote.
        var masterMatch = Regex.Match(partPath, @"^/slideMaster\[(\d+)\]$");
        if (masterMatch.Success)
        {
            var idx = int.Parse(masterMatch.Groups[1].Value);
            var masters = presentationPart.SlideMasterParts.ToList();
            if (idx < 1 || idx > masters.Count)
                throw new ArgumentException($"slideMaster[{idx}] not found (total: {masters.Count})");
            return masters[idx - 1].SlideMaster?.OuterXml
                ?? throw new InvalidOperationException("Corrupt file: slide master data missing");
        }

        var layoutMatch = Regex.Match(partPath, @"^/slideLayout\[(\d+)\]$");
        if (layoutMatch.Success)
        {
            var idx = int.Parse(layoutMatch.Groups[1].Value);
            var layouts = presentationPart.SlideMasterParts
                .SelectMany(m => m.SlideLayoutParts).ToList();
            if (idx < 1 || idx > layouts.Count)
                throw new ArgumentException($"slideLayout[{idx}] not found (total: {layouts.Count})");
            return layouts[idx - 1].SlideLayout?.OuterXml
                ?? throw new InvalidOperationException("Corrupt file: slide layout data missing");
        }

        var noteMatch = Regex.Match(partPath, @"^/noteSlide\[(\d+)\]$");
        if (noteMatch.Success)
        {
            var idx = int.Parse(noteMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx < 1 || idx > slideParts.Count)
                throw new ArgumentException($"slide[{idx}] not found (total: {slideParts.Count})");
            var notesPart = slideParts[idx - 1].NotesSlidePart
                ?? throw new ArgumentException($"Slide {idx} has no notes");
            return notesPart.NotesSlide?.OuterXml
                ?? throw new InvalidOperationException("Corrupt file: notes slide data missing");
        }

        // CONSISTENCY(raw-rawset-symmetry): /notesMaster surfaces the
        // presentation-level NotesMasterPart's XML so PptxBatchEmitter can
        // raw-set it on replay (mirrors theme/master/layout treatment).
        if (partPath == "/notesMaster")
        {
            return presentationPart.NotesMasterPart?.NotesMaster?.OuterXml
                ?? throw new ArgumentException("No notes master part");
        }

        throw new ArgumentException($"Unknown part: {partPath}. Available: /presentation, /theme, /slide[N], /slideMaster[N], /slideLayout[N], /noteSlide[N], /notesMaster");
    }

    public void RawSet(string partPath, string xpath, string action, string? xml)
    {
        Modified = true;
        if (partPath == null) throw new ArgumentNullException(nameof(partPath));
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("No presentation part");

        if (RawXmlHelper.IsZipUriPath(partPath))
        {
            var part = RawXmlHelper.FindPartByZipUri(_doc, partPath)
                ?? throw new ArgumentException(
                    $"Unknown part: {partPath}. The path was treated as a zip-internal URI " +
                    $"but no matching part exists in the package. " +
                    $"Use semantic paths (/presentation, /slide[N], /slideMaster[N]) for stable identification.");
            RawXmlHelper.Execute(part, xpath, action, xml);
            return;
        }

        OpenXmlPartRootElement rootElement;

        if (partPath is "/" or "/presentation")
        {
            rootElement = presentationPart.Presentation
                ?? throw new InvalidOperationException("No presentation");
        }
        else if (partPath == "/theme")
        {
            rootElement = presentationPart.ThemePart?.Theme
                ?? throw new ArgumentException("No theme part");
        }
        else if (Regex.Match(partPath, @"^/slide\[(\d+)\]$") is { Success: true } slideMatch)
        {
            var idx = int.Parse(slideMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx < 1 || idx > slideParts.Count)
                throw new ArgumentException($"Slide {idx} not found (total: {slideParts.Count})");
            rootElement = GetSlide(slideParts[idx - 1]);
        }
        else if (Regex.Match(partPath, @"^/slideMaster\[(\d+)\]$") is { Success: true } masterMatch)
        {
            var idx = int.Parse(masterMatch.Groups[1].Value);
            var masters = presentationPart.SlideMasterParts.ToList();
            if (idx < 1)
                throw new ArgumentException($"SlideMaster {idx} not found (total: {masters.Count})");
            if (idx > masters.Count)
            {
                // CONSISTENCY(grow-on-rawset): mirrors the slideLayout branch.
                // Source decks with multiple slideMasters (template kits, decks
                // assembled from several themes) emit raw-set on /slideMaster[2..N];
                // blank target only stamped /slideMaster[1], so the replay used to
                // fail every additional master AND every layout owned by those
                // missing masters. Auto-grow to idx so the raw-set replace has a
                // root element to swap out; the replace then carries in the real
                // sldLayoutIdLst, which GrowSlideLayoutParts consults when the
                // subsequent /slideLayout[K] raw-sets land.
                GrowSlideMasterParts(idx);
                masters = presentationPart.SlideMasterParts.ToList();
                if (idx > masters.Count)
                    throw new ArgumentException($"SlideMaster {idx} not found (total: {masters.Count})");
            }
            rootElement = masters[idx - 1].SlideMaster
                ?? throw new InvalidOperationException("Corrupt file: slide master data missing");
        }
        else if (Regex.Match(partPath, @"^/slideLayout\[(\d+)\]$") is { Success: true } layoutMatch)
        {
            var idx = int.Parse(layoutMatch.Groups[1].Value);
            var layouts = presentationPart.SlideMasterParts
                .SelectMany(m => m.SlideLayoutParts).ToList();
            if (idx < 1)
                throw new ArgumentException($"SlideLayout {idx} not found (total: {layouts.Count})");
            if (idx > layouts.Count)
            {
                // BUG-J: Replay scenario — source deck has N layouts (e.g. 11) but
                // the blank target only stamped K (5). EmitMasterRaw already
                // replaced the master's sldLayoutIdLst to reference all N rIds,
                // but the SlideLayoutPart objects for K+1..N don't exist yet, so
                // raw-set fails with "SlideLayout {idx} not found". Auto-grow the
                // missing layout parts under the appropriate master based on the
                // post-master-replace sldLayoutIdLst, then re-resolve.
                GrowSlideLayoutParts(idx);
                layouts = presentationPart.SlideMasterParts
                    .SelectMany(m => m.SlideLayoutParts).ToList();
                if (idx > layouts.Count)
                    throw new ArgumentException($"SlideLayout {idx} not found (total: {layouts.Count})");
            }
            rootElement = layouts[idx - 1].SlideLayout
                ?? throw new InvalidOperationException("Corrupt file: slide layout data missing");
        }
        else if (Regex.Match(partPath, @"^/noteSlide\[(\d+)\]$") is { Success: true } noteMatch)
        {
            var idx = int.Parse(noteMatch.Groups[1].Value);
            var slideParts = GetSlideParts().ToList();
            if (idx < 1 || idx > slideParts.Count)
                throw new ArgumentException($"Slide {idx} not found (total: {slideParts.Count})");
            var notesPart = slideParts[idx - 1].NotesSlidePart
                ?? throw new ArgumentException($"Slide {idx} has no notes");
            rootElement = notesPart.NotesSlide
                ?? throw new InvalidOperationException("Corrupt file: notes slide data missing");
        }
        else if (partPath == "/notesMaster")
        {
            // CONSISTENCY(grow-on-rawset): blank pptx files have no
            // NotesMasterPart, but PptxBatchEmitter emits a raw-set /notesMaster
            // on any deck that has one. Create the part on demand so dump-replay
            // can stamp the source notes master back in (mirrors GrowSlideLayoutParts
            // in the slideLayout branch above).
            var nmPart = presentationPart.NotesMasterPart;
            if (nmPart == null)
            {
                nmPart = presentationPart.AddNewPart<NotesMasterPart>();
                // Seed a minimal placeholder so the raw-set "replace" action has
                // a NotesMaster root element to swap out; raw-replace builds the
                // real content from the supplied XML on the next line.
                nmPart.NotesMaster = new NotesMaster(
                    new CommonSlideData(new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties { Id = 1, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new DocumentFormat.OpenXml.Drawing.TransformGroup()))),
                    new ColorMap
                    {
                        Background1 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Light1,
                        Text1 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Dark1,
                        Background2 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Light2,
                        Text2 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Dark2,
                        Accent1 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent1,
                        Accent2 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent2,
                        Accent3 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent3,
                        Accent4 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent4,
                        Accent5 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent5,
                        Accent6 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent6,
                        Hyperlink = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Hyperlink,
                        FollowedHyperlink = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.FollowedHyperlink,
                    });
            }
            rootElement = nmPart.NotesMaster!;
        }
        else
        {
            throw new ArgumentException($"Unknown part: {partPath}. Available: /presentation, /theme, /slide[N], /slideMaster[N], /slideLayout[N], /noteSlide[N], /notesMaster");
        }

        var affected = RawXmlHelper.Execute(rootElement, xpath, action, xml);
        rootElement.Save();
        // After a /slideMaster[N] raw-set the master's <p:sldLayoutIdLst> is
        // the source's authoritative layout count. Blank decks ship with a
        // pre-stamped 5-layout master, so a 1-layout source replays to a
        // 5-layout deck — dump→batch→dump grows from 8 ops to 12 because
        // the 4 extra blank layouts survive. Prune SlideLayoutParts whose
        // rId is no longer in the post-replace sldLayoutIdLst so the
        // replayed deck mirrors the source's layout set exactly. The grow
        // path (line ~203) handles the opposite case (source has MORE).
        if (Regex.IsMatch(partPath, @"^/slideMaster\[\d+\]$") && rootElement is SlideMaster sm)
        {
            var mp = sm.SlideMasterPart;
            if (mp != null)
            {
                var declaredRids = new HashSet<string>(
                    sm.SlideLayoutIdList?.Elements<SlideLayoutId>()
                        .Select(e => e.RelationshipId?.Value ?? "")
                        .Where(s => !string.IsNullOrEmpty(s))
                    ?? Enumerable.Empty<string>(),
                    StringComparer.Ordinal);
                foreach (var pair in mp.Parts.ToList())
                {
                    if (pair.OpenXmlPart is SlideLayoutPart lp
                        && !declaredRids.Contains(pair.RelationshipId))
                    {
                        // The orphan layout's rels (theme/image links etc.) drop
                        // with the part; DeletePart cascades.
                        mp.DeletePart(lp);
                    }
                }
            }
        }
        // BUG-R43: raw-set may have inserted/removed shape XML directly (incl.
        // cNvPr ids). The cached _usedShapeIds set is now stale, so the next
        // Add() can hand out an id that already exists in the tree, producing
        // duplicate cNvPr ids that PowerPoint silently rejects. Rebuild the
        // shape-id index from the live tree after every raw-set.
        InitShapeIdCounter();
        // BUG-R5-01: silent — CLI wrappers print their own structured message.
        _ = affected;
    }

    // BUG-J: Auto-grow SlideLayoutPart objects so raw-set replay can target
    // /slideLayout[targetGlobalIdx] even when the blank deck only stamped a
    // subset. The master's sldLayoutIdLst (already replaced by EmitMasterRaw)
    // is the source of truth: each entry holds the rId that the new
    // SlideLayoutPart must register with so the master-layout relationship
    // matches. We walk masters in declaration order, compute each master's
    // declared layout count from sldLayoutIdLst, and create missing parts
    // under whichever master's range contains targetGlobalIdx. Newly created
    // parts get a stub root (the imminent replace overwrites it).
    private void GrowSlideLayoutParts(int targetGlobalIdx)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("No presentation part");
        var masters = presentationPart.SlideMasterParts.ToList();
        int seen = 0;
        foreach (var mp in masters)
        {
            var declared = mp.SlideMaster?.SlideLayoutIdList?.Elements<SlideLayoutId>().ToList()
                ?? new List<SlideLayoutId>();
            int declaredCount = declared.Count;
            int existingCount = mp.SlideLayoutParts.Count();
            // This master "owns" global indices (seen+1)..(seen+declaredCount).
            int rangeStart = seen + 1;
            int rangeEnd = seen + declaredCount;
            if (targetGlobalIdx >= rangeStart && targetGlobalIdx <= rangeEnd)
            {
                // Create missing parts for slots existingCount+1 .. declaredCount.
                for (int slot = existingCount; slot < declaredCount; slot++)
                {
                    var declaredId = declared[slot];
                    var rId = declaredId.RelationshipId?.Value;
                    SlideLayoutPart newPart;
                    if (!string.IsNullOrEmpty(rId) && !mp.Parts.Any(p => p.RelationshipId == rId))
                    {
                        newPart = mp.AddNewPart<SlideLayoutPart>(rId);
                    }
                    else
                    {
                        // Either rId missing in sldLayoutIdLst or already taken
                        // (corruption guard) — let OpenXml allocate a new one
                        // and patch the sldLayoutIdLst entry to match.
                        newPart = mp.AddNewPart<SlideLayoutPart>();
                        var newRid = mp.GetIdOfPart(newPart);
                        declaredId.RelationshipId = newRid;
                    }
                    // Stub root — the raw-set replace immediately rewrites it.
                    newPart.SlideLayout = new SlideLayout(
                        new CommonSlideData(
                            new ShapeTree(
                                new NonVisualGroupShapeProperties(
                                    new NonVisualDrawingProperties { Id = 1, Name = "" },
                                    new NonVisualGroupShapeDrawingProperties(),
                                    new ApplicationNonVisualDrawingProperties()),
                                new GroupShapeProperties()))
                    ) { Type = SlideLayoutValues.Blank };
                    newPart.SlideLayout.Save();
                    // Layouts must point back to their master.
                    newPart.AddPart(mp);
                }
                if (mp.SlideMaster != null) mp.SlideMaster.Save();
                return;
            }
            seen += declaredCount;
        }
        // targetGlobalIdx is beyond every master's declared range; caller will
        // raise the canonical "not found" error after we return.
    }

    // CONSISTENCY(grow-on-rawset): mirror of GrowSlideLayoutParts for the
    // SlideMasterPart side. Multi-master source decks (template kits, decks
    // assembled from multiple themes) emit raw-set on /slideMaster[2..N], but
    // BlankDocCreator only stamps one master. Create enough placeholder
    // SlideMasterParts (each with a minimal SlideMaster root plus its own
    // SlideLayoutIdList stub) and register them in the presentation's
    // sldMasterIdLst so the raw-set replace has a root element to swap, and
    // so subsequent /slideLayout[K] raw-sets can find their owning master via
    // GrowSlideLayoutParts.
    private void GrowSlideMasterParts(int targetIdx)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("No presentation part");
        var presentation = presentationPart.Presentation
            ?? throw new InvalidOperationException("No presentation");
        var sldMasterIdLst = presentation.SlideMasterIdList
            ?? throw new InvalidOperationException("Presentation has no SlideMasterIdList");

        var existing = presentationPart.SlideMasterParts.Count();
        if (targetIdx <= existing) return;

        // Pick a SlideMasterId base that won't collide with the existing IDs.
        var existingIds = sldMasterIdLst.Elements<SlideMasterId>()
            .Select(e => e.Id?.Value ?? 0u)
            .ToHashSet();
        uint nextId = 2147483648u;
        while (existingIds.Contains(nextId)) nextId++;

        for (int i = existing; i < targetIdx; i++)
        {
            var newPart = presentationPart.AddNewPart<SlideMasterPart>();
            var rId = presentationPart.GetIdOfPart(newPart);
            // Minimal SlideMaster root with an empty SlideLayoutIdList so
            // GrowSlideLayoutParts sees the right declared count once the
            // imminent raw-set replace overwrites this stub with the real
            // master XML (which carries the source's actual sldLayoutIdLst).
            newPart.SlideMaster = new SlideMaster(
                new CommonSlideData(new ShapeTree(
                    new NonVisualGroupShapeProperties(
                        new NonVisualDrawingProperties { Id = 1, Name = "" },
                        new NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new DocumentFormat.OpenXml.Drawing.TransformGroup()))),
                new ColorMap
                {
                    Background1 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Light1,
                    Text1 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Dark1,
                    Background2 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Light2,
                    Text2 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Dark2,
                    Accent1 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent1,
                    Accent2 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent2,
                    Accent3 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent3,
                    Accent4 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent4,
                    Accent5 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent5,
                    Accent6 = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Accent6,
                    Hyperlink = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.Hyperlink,
                    FollowedHyperlink = DocumentFormat.OpenXml.Drawing.ColorSchemeIndexValues.FollowedHyperlink,
                },
                new SlideLayoutIdList()
            );
            newPart.SlideMaster.Save();

            // Every SlideMasterPart must reference a ThemePart. The replay's
            // raw-set replaces only the master XML, not the package
            // relationships — without a theme link here the package fails
            // validation. Share the presentation's primary theme; a richer
            // multi-theme deck can later raw-set its own theme parts.
            if (presentationPart.ThemePart != null)
                newPart.AddPart(presentationPart.ThemePart);

            sldMasterIdLst.AppendChild(new SlideMasterId { Id = nextId++, RelationshipId = rId });
        }
        presentation.Save();
    }

    public (string RelId, string PartPath) AddPart(string parentPartPath, string partType, Dictionary<string, string>? properties = null)
    {
        var presentationPart = _doc.PresentationPart
            ?? throw new InvalidOperationException("No presentation part");

        switch (partType.ToLowerInvariant())
        {
            case "chart":
                // Charts go under a SlidePart
                var slideMatch = System.Text.RegularExpressions.Regex.Match(
                    parentPartPath, @"^/slide\[(\d+)\]$");
                if (!slideMatch.Success)
                    throw new ArgumentException(
                        "Chart must be added under a slide: add-part <file> '/slide[N]' --type chart");

                var slideIdx = int.Parse(slideMatch.Groups[1].Value);
                var slideParts = GetSlideParts().ToList();
                if (slideIdx < 1 || slideIdx > slideParts.Count)
                    throw new ArgumentException($"Slide index {slideIdx} out of range");

                var slidePart = slideParts[slideIdx - 1];
                var chartPart = slidePart.AddNewPart<DocumentFormat.OpenXml.Packaging.ChartPart>();
                var relId = slidePart.GetIdOfPart(chartPart);

                chartPart.ChartSpace = new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace(
                    new DocumentFormat.OpenXml.Drawing.Charts.Chart(
                        new DocumentFormat.OpenXml.Drawing.Charts.PlotArea(
                            new DocumentFormat.OpenXml.Drawing.Charts.Layout()
                        )
                    )
                );
                chartPart.ChartSpace.Save();

                var chartIdx = slidePart.ChartParts.ToList().IndexOf(chartPart);
                return (relId, $"/slide[{slideIdx}]/chart[{chartIdx + 1}]");

            case "smartart":
                // SmartArt graphicFrame references four separate OOXML
                // sub-parts under the owning SlidePart: data (dgm:dataModel),
                // layout (dgm:layoutDef), colors (dgm:colorsDef), style
                // (dgm:styleDef). The graphicFrame's <dgm:relIds> carries the
                // four rIds, so dump→batch→replay byte-equality requires that
                // the rIds match the source's. Callers MAY pass explicit rIds
                // via properties {data, layout, colors, quickStyle}; when
                // omitted the SDK allocates fresh ones. Each part is seeded
                // with a minimal typed root so subsequent raw-set replace
                // ops can target /dgm:dataModel etc.
                var saSlideMatch = System.Text.RegularExpressions.Regex.Match(
                    parentPartPath, @"^/slide\[(\d+)\]$");
                if (!saSlideMatch.Success)
                    throw new ArgumentException(
                        "SmartArt must be added under a slide: add-part <file> '/slide[N]' --type smartart");
                var saSlideIdx = int.Parse(saSlideMatch.Groups[1].Value);
                var saSlideParts = GetSlideParts().ToList();
                if (saSlideIdx < 1 || saSlideIdx > saSlideParts.Count)
                    throw new ArgumentException($"Slide index {saSlideIdx} out of range");
                var saSlidePart = saSlideParts[saSlideIdx - 1];

                string? dataRid     = properties != null && properties.TryGetValue("data", out var dv) ? dv : null;
                string? layoutRid   = properties != null && properties.TryGetValue("layout", out var lv) ? lv : null;
                string? colorsRid   = properties != null && properties.TryGetValue("colors", out var cv) ? cv : null;
                string? qsRid       = properties != null && properties.TryGetValue("quickStyle", out var qv) ? qv : null;

                DiagramDataPart   dataPart   = !string.IsNullOrEmpty(dataRid)
                    ? saSlidePart.AddNewPart<DiagramDataPart>(dataRid)
                    : saSlidePart.AddNewPart<DiagramDataPart>();
                DiagramLayoutDefinitionPart layoutPart = !string.IsNullOrEmpty(layoutRid)
                    ? saSlidePart.AddNewPart<DiagramLayoutDefinitionPart>(layoutRid)
                    : saSlidePart.AddNewPart<DiagramLayoutDefinitionPart>();
                DiagramColorsPart colorsPart = !string.IsNullOrEmpty(colorsRid)
                    ? saSlidePart.AddNewPart<DiagramColorsPart>(colorsRid)
                    : saSlidePart.AddNewPart<DiagramColorsPart>();
                DiagramStylePart  stylePart  = !string.IsNullOrEmpty(qsRid)
                    ? saSlidePart.AddNewPart<DiagramStylePart>(qsRid)
                    : saSlidePart.AddNewPart<DiagramStylePart>();

                // Minimal typed roots — raw-set replace immediately overwrites.
                dataPart.DataModelRoot = new DocumentFormat.OpenXml.Drawing.Diagrams.DataModelRoot(
                    new DocumentFormat.OpenXml.Drawing.Diagrams.PointList(),
                    new DocumentFormat.OpenXml.Drawing.Diagrams.ConnectionList());
                dataPart.DataModelRoot.Save(dataPart);
                layoutPart.LayoutDefinition = new DocumentFormat.OpenXml.Drawing.Diagrams.LayoutDefinition();
                layoutPart.LayoutDefinition.Save(layoutPart);
                colorsPart.ColorsDefinition = new DocumentFormat.OpenXml.Drawing.Diagrams.ColorsDefinition();
                colorsPart.ColorsDefinition.Save(colorsPart);
                stylePart.StyleDefinition = new DocumentFormat.OpenXml.Drawing.Diagrams.StyleDefinition();
                stylePart.StyleDefinition.Save(stylePart);

                // Encode all four rIds in the RelId field — callers (batch
                // emit / replay) need to know each part's id to write the
                // matching dgm:relIds on the graphicFrame. Format:
                // "data=rIdX;layout=rIdY;colors=rIdZ;quickStyle=rIdW".
                var dataActualRid   = saSlidePart.GetIdOfPart(dataPart);
                var layoutActualRid = saSlidePart.GetIdOfPart(layoutPart);
                var colorsActualRid = saSlidePart.GetIdOfPart(colorsPart);
                var styleActualRid  = saSlidePart.GetIdOfPart(stylePart);
                var encoded = $"data={dataActualRid};layout={layoutActualRid};colors={colorsActualRid};quickStyle={styleActualRid}";
                return (encoded, parentPartPath);

            case "video":
            case "audio":
                // Phase 3c-media. Mirror Phase 3b SmartArt: create the
                // underlying parts (MediaDataPart + ImagePart thumbnail)
                // and pin all three rIds via properties so the post-replay
                // <p:pic> appended by raw-set finds the same rIds it carried
                // in the source. The graphicFrame analogue here is the
                // <p:pic> referencing <a:videoFile r:link=…/> +
                // <p14:media r:embed=…/> in nvPr, plus <a:blip r:embed=…/>
                // in blipFill for the thumbnail.
                //
                // Props (all optional except data + thumbnail-data):
                //   data                   = base64 binary (mp4/m4a/…)
                //   content-type           = "video/mp4" / "audio/mpeg" / …
                //   extension              = ".mp4" / ".m4a" (for the
                //                            MediaDataPart URI extension;
                //                            best-effort from content-type
                //                            when omitted)
                //   thumbnail-data         = base64 image binary
                //   thumbnail-content-type = "image/png" / "image/jpeg"
                //   video-rid / audio-rid  = pinned VideoReference / AudioReference rId
                //   media-rid              = pinned p14:media MediaReference rId
                //   thumbnail-rid          = pinned ImagePart rId
                //
                // Audio uses AddAudioReferenceRelationship; video uses
                // AddVideoReferenceRelationship. Both ALSO add a
                // MediaReferenceRelationship (the p14:media r:embed is
                // distinct from the legacy r:link).
                var mediaSlideMatch = System.Text.RegularExpressions.Regex.Match(
                    parentPartPath, @"^/slide\[(\d+)\]$");
                if (!mediaSlideMatch.Success)
                    throw new ArgumentException(
                        $"{partType} must be added under a slide: add-part <file> '/slide[N]' --type {partType}");
                var mediaSlideIdx = int.Parse(mediaSlideMatch.Groups[1].Value);
                var mediaSlidePartsList = GetSlideParts().ToList();
                if (mediaSlideIdx < 1 || mediaSlideIdx > mediaSlidePartsList.Count)
                    throw new ArgumentException($"Slide index {mediaSlideIdx} out of range");
                var mediaSlidePart = mediaSlidePartsList[mediaSlideIdx - 1];

                if (properties == null || !properties.TryGetValue("data", out var mediaB64) || string.IsNullOrEmpty(mediaB64))
                    throw new ArgumentException(
                        $"add-part {partType} requires property 'data' (base64 binary)");
                byte[] mediaBytes;
                try { mediaBytes = Convert.FromBase64String(mediaB64); }
                catch (FormatException) { throw new ArgumentException($"add-part {partType}: 'data' is not valid base64"); }

                var mediaContentType = properties.TryGetValue("content-type", out var mct) && !string.IsNullOrEmpty(mct)
                    ? mct
                    : (partType == "video" ? "video/mp4" : "audio/mpeg");
                var mediaExt = properties.TryGetValue("extension", out var mxt) && !string.IsNullOrEmpty(mxt)
                    ? mxt
                    : mediaContentType switch {
                        "video/mp4" => ".mp4", "video/x-msvideo" => ".avi",
                        "video/x-ms-wmv" => ".wmv", "video/mpeg" => ".mpg",
                        "video/quicktime" => ".mov",
                        "audio/mpeg" => ".mp3", "audio/wav" => ".wav",
                        "audio/x-ms-wma" => ".wma", "audio/mp4" => ".m4a",
                        _ => ".bin" };

                var mediaDataPart = _doc.CreateMediaDataPart(mediaContentType, mediaExt);
                using (var inStream = new MemoryStream(mediaBytes))
                    mediaDataPart.FeedData(inStream);

                string? pinnedVideoRid = properties.TryGetValue("video-rid", out var vr) ? vr : null;
                string? pinnedAudioRid = properties.TryGetValue("audio-rid", out var ar) ? ar : null;
                string? pinnedMediaRid = properties.TryGetValue("media-rid", out var mr) ? mr : null;
                string? pinnedThumbRid = properties.TryGetValue("thumbnail-rid", out var tr) ? tr : null;

                string linkRelId;
                if (partType == "video")
                {
                    linkRelId = !string.IsNullOrEmpty(pinnedVideoRid)
                        ? mediaSlidePart.AddVideoReferenceRelationship(mediaDataPart, pinnedVideoRid).Id
                        : mediaSlidePart.AddVideoReferenceRelationship(mediaDataPart).Id;
                }
                else
                {
                    linkRelId = !string.IsNullOrEmpty(pinnedAudioRid)
                        ? mediaSlidePart.AddAudioReferenceRelationship(mediaDataPart, pinnedAudioRid).Id
                        : mediaSlidePart.AddAudioReferenceRelationship(mediaDataPart).Id;
                }
                var mediaEmbedRid = !string.IsNullOrEmpty(pinnedMediaRid)
                    ? mediaSlidePart.AddMediaReferenceRelationship(mediaDataPart, pinnedMediaRid).Id
                    : mediaSlidePart.AddMediaReferenceRelationship(mediaDataPart).Id;

                // Thumbnail (poster) — required so the <a:blip r:embed> in
                // the <p:pic>'s blipFill resolves on replay. Caller MAY
                // omit thumbnail-data; we then seed a 1x1 transparent PNG
                // (mirrors the AddMedia helper's placeholder path).
                byte[] thumbBytes;
                PartTypeInfo thumbType;
                if (properties.TryGetValue("thumbnail-data", out var tdB64) && !string.IsNullOrEmpty(tdB64))
                {
                    try { thumbBytes = Convert.FromBase64String(tdB64); }
                    catch (FormatException) { throw new ArgumentException($"add-part {partType}: 'thumbnail-data' is not valid base64"); }
                    var thumbCT = properties.TryGetValue("thumbnail-content-type", out var tct) && !string.IsNullOrEmpty(tct)
                        ? tct
                        : "image/png";
                    thumbType = thumbCT switch {
                        "image/png" => ImagePartType.Png, "image/jpeg" => ImagePartType.Jpeg,
                        "image/gif" => ImagePartType.Gif, "image/bmp" => ImagePartType.Bmp,
                        "image/tiff" => ImagePartType.Tiff, _ => ImagePartType.Png };
                }
                else
                {
                    thumbBytes = new byte[]
                    {
                        0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                        0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                        0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,0x89,
                        0x00,0x00,0x00,0x0D,0x49,0x44,0x41,0x54,
                        0x08,0xD7,0x63,0x60,0x60,0x60,0x60,0x00,0x00,0x00,0x05,0x00,0x01,0x87,0xA1,0x4E,0xD4,
                        0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,0x42,0x60,0x82
                    };
                    thumbType = ImagePartType.Png;
                }
                var thumbImagePart = !string.IsNullOrEmpty(pinnedThumbRid)
                    ? mediaSlidePart.AddImagePart(thumbType, pinnedThumbRid)
                    : mediaSlidePart.AddImagePart(thumbType);
                using (var thumbStream = new MemoryStream(thumbBytes))
                    thumbImagePart.FeedData(thumbStream);
                var thumbActualRid = mediaSlidePart.GetIdOfPart(thumbImagePart);

                // Encode three rIds — emitter / replay caller may use any
                // of them when writing the <p:pic> XML via raw-set. The
                // (RelId, PartPath) tuple's RelId is consumed by callers
                // who do their own bookkeeping; format mirrors smartart.
                var mediaKey = partType == "video" ? "video" : "audio";
                var encodedMedia = $"{mediaKey}={linkRelId};media={mediaEmbedRid};thumbnail={thumbActualRid}";
                return (encodedMedia, parentPartPath);

            case "model3d":
            case "3dmodel":
                // Phase 3c-3d. Mirrors Phase 3c-media (video/audio). The
                // PPT 3D model lives inside <mc:AlternateContent>:
                //   <mc:Choice Requires="am3d">
                //     <p:graphicFrame>... <am3d:model3d r:embed=…>
                //         <am3d:raster><am3d:blip r:embed=…/></am3d:raster>
                //   <mc:Fallback><p:pic>... <a:blip r:embed=…/></p:pic>
                // Two rels back the slice:
                //  - relType .../office/2017/06/relationships/model3d
                //    -> the .glb binary (created via AddExtendedPart;
                //       SDK lacks a typed Model3DPart)
                //  - relType .../officeDocument/2006/relationships/image
                //    -> a static thumbnail PNG (shared by am3d:raster's
                //       blip AND the Fallback p:pic's blipFill)
                //
                // Props (all optional except data + thumbnail-data):
                //   data                   = base64 .glb bytes
                //   content-type           = "model/gltf.binary" (default)
                //   extension              = ".glb" (default)
                //   model3d-rid            = pinned model3d ExtendedPart rId
                //   thumbnail-data         = base64 thumbnail image bytes
                //   thumbnail-content-type = "image/png" (default) / "image/jpeg"
                //   thumbnail-rid          = pinned thumbnail ImagePart rId
                //
                // Returned encoded relId: "model3d=rIdA;thumbnail=rIdB".
                // No shape XML is inserted under the slide; the companion
                // raw-set append carries the full <mc:AlternateContent>
                // verbatim with the matching pinned rIds.
                var m3dSlideMatch = System.Text.RegularExpressions.Regex.Match(
                    parentPartPath, @"^/slide\[(\d+)\]$");
                if (!m3dSlideMatch.Success)
                    throw new ArgumentException(
                        $"{partType} must be added under a slide: add-part <file> '/slide[N]' --type {partType}");
                var m3dSlideIdx = int.Parse(m3dSlideMatch.Groups[1].Value);
                var m3dSlideParts = GetSlideParts().ToList();
                if (m3dSlideIdx < 1 || m3dSlideIdx > m3dSlideParts.Count)
                    throw new ArgumentException($"Slide index {m3dSlideIdx} out of range");
                var m3dSlidePart = m3dSlideParts[m3dSlideIdx - 1];

                // 'data' may be absent ONLY when callers (rare) are pre-
                // declaring rels with empty parts; the dump emitter always
                // passes it.
                if (properties == null || !properties.TryGetValue("data", out var m3dB64) || string.IsNullOrEmpty(m3dB64))
                    throw new ArgumentException(
                        $"add-part {partType} requires property 'data' (base64 .glb bytes)");
                byte[] m3dBytes;
                try { m3dBytes = Convert.FromBase64String(m3dB64); }
                catch (FormatException) { throw new ArgumentException($"add-part {partType}: 'data' is not valid base64"); }

                var m3dContentType = properties.TryGetValue("content-type", out var m3dct) && !string.IsNullOrEmpty(m3dct)
                    ? m3dct
                    : "model/gltf.binary";
                var m3dExt = properties.TryGetValue("extension", out var m3dxt) && !string.IsNullOrEmpty(m3dxt)
                    ? (m3dxt.StartsWith('.') ? m3dxt : "." + m3dxt)
                    : ".glb";

                string? pinnedM3dRid   = properties.TryGetValue("model3d-rid", out var m3dr) ? m3dr : null;
                string? pinnedM3dThumb = properties.TryGetValue("thumbnail-rid", out var m3dtr) ? m3dtr : null;

                const string m3dRelType = "http://schemas.microsoft.com/office/2017/06/relationships/model3d";
                var m3dPart = !string.IsNullOrEmpty(pinnedM3dRid)
                    ? m3dSlidePart.AddExtendedPart(m3dRelType, m3dContentType, m3dExt, pinnedM3dRid)
                    : m3dSlidePart.AddExtendedPart(m3dRelType, m3dContentType, m3dExt);
                using (var s = new MemoryStream(m3dBytes)) m3dPart.FeedData(s);
                var m3dActualRid = m3dSlidePart.GetIdOfPart(m3dPart);

                // Thumbnail (am3d:raster blip + Fallback blipFill share one
                // ImagePart). Caller may omit thumbnail-data; we then seed
                // the same 1x1 transparent PNG used by the video/audio
                // placeholder path.
                byte[] m3dThumbBytes;
                PartTypeInfo m3dThumbType;
                if (properties.TryGetValue("thumbnail-data", out var m3dtdB64) && !string.IsNullOrEmpty(m3dtdB64))
                {
                    try { m3dThumbBytes = Convert.FromBase64String(m3dtdB64); }
                    catch (FormatException) { throw new ArgumentException($"add-part {partType}: 'thumbnail-data' is not valid base64"); }
                    var m3dThumbCT = properties.TryGetValue("thumbnail-content-type", out var m3dtct) && !string.IsNullOrEmpty(m3dtct)
                        ? m3dtct
                        : "image/png";
                    m3dThumbType = m3dThumbCT switch {
                        "image/png" => ImagePartType.Png, "image/jpeg" => ImagePartType.Jpeg,
                        "image/gif" => ImagePartType.Gif, "image/bmp" => ImagePartType.Bmp,
                        "image/tiff" => ImagePartType.Tiff, _ => ImagePartType.Png };
                }
                else
                {
                    m3dThumbBytes = new byte[]
                    {
                        0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                        0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                        0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,0x89,
                        0x00,0x00,0x00,0x0D,0x49,0x44,0x41,0x54,
                        0x08,0xD7,0x63,0x60,0x60,0x60,0x60,0x00,0x00,0x00,0x05,0x00,0x01,0x87,0xA1,0x4E,0xD4,
                        0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,0x42,0x60,0x82
                    };
                    m3dThumbType = ImagePartType.Png;
                }
                var m3dThumbPart = !string.IsNullOrEmpty(pinnedM3dThumb)
                    ? m3dSlidePart.AddImagePart(m3dThumbType, pinnedM3dThumb)
                    : m3dSlidePart.AddImagePart(m3dThumbType);
                using (var s = new MemoryStream(m3dThumbBytes)) m3dThumbPart.FeedData(s);
                var m3dThumbActualRid = m3dSlidePart.GetIdOfPart(m3dThumbPart);

                var encodedM3d = $"model3d={m3dActualRid};thumbnail={m3dThumbActualRid}";
                return (encodedM3d, parentPartPath);

            default:
                throw new ArgumentException(
                    $"Unknown part type: {partType}. Supported: chart, smartart, video, audio, model3d");
        }
    }

    public List<ValidationError> Validate() => RawXmlHelper.ValidateDocument(_doc);

    public void Save()
    {
        // _doc writes through to _backingStream; force the FileStream buffer
        // out to disk so external readers see the latest bytes immediately.
        if (Modified)
        {
            try { OfficeCli.Core.OfficeCliMetadata.StampOnSave(_doc); }
            catch { /* best-effort audit trail */ }
        }
        _doc.Save();
        _backingStream?.Flush();
    }

    public void Dispose()
    {
        // Save through the package (flush in-memory edits to the underlying
        // stream) before disposing. When we own the backing FileStream, the
        // package would otherwise leave the on-disk file in whatever state
        // the last auto-flush left it — for the stream-Open path this can
        // truncate to zero bytes and look like a corrupted zip on reopen.
        if (Modified)
        {
            try { OfficeCli.Core.OfficeCliMetadata.StampOnSave(_doc); }
            catch { /* best-effort audit trail */ }
        }
        try { _doc.Save(); } catch { /* read-only or already disposed */ }
        _doc.Dispose();
        _backingStream?.Dispose();
        _backingStream = null;
    }

    // Internal accessors used by PptxBatchEmitter (resource enumeration).
    // Keep the PresentationPart itself private; expose only the counts and
    // a binary getter that the emitter needs.
    internal int SlideMasterCount =>
        _doc.PresentationPart?.SlideMasterParts.Count() ?? 0;
    internal int SlideLayoutCount =>
        _doc.PresentationPart?.SlideMasterParts.SelectMany(m => m.SlideLayoutParts).Count() ?? 0;
    internal bool HasNotesMaster =>
        _doc.PresentationPart?.NotesMasterPart != null;
    // Exposed for PptxBatchEmitter so it can iterate slides without going
    // through Get("/") — Get("/") fans out into per-slide deep walks that
    // can throw at SDK validation time on vendor templates with foreign
    // attributes (gov_bja, 1.pptx, ...). The emitter now uses this count
    // plus per-slide try/catch to keep the dump going on partial corruption.
    internal int SlideCount => GetSlideParts().Count();

    internal bool SlideHasNotes(int slideIdx)
    {
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return false;
        return parts[slideIdx - 1].NotesSlidePart != null;
    }

    /// <summary>
    /// Per-slide SmartArt info for PptxBatchEmitter passthrough. Returns
    /// one entry per <p:graphicFrame> on the slide that carries a
    /// <dgm:relIds> child (= SmartArt host frame). Each entry includes the
    /// source's four rIds and the four diagram parts' XML so the emitter
    /// can issue an `add-part smartart` + four `raw-set` rows that
    /// round-trip byte-equal.
    /// </summary>
    internal readonly record struct SmartArtInfo(
        string GraphicFrameXml,
        string DataRelId,
        string LayoutRelId,
        string ColorsRelId,
        string QuickStyleRelId,
        string DataXml,
        string LayoutXml,
        string ColorsXml,
        string QuickStyleXml);

    internal IReadOnlyList<SmartArtInfo> GetSmartArtsOnSlide(int slideIdx)
    {
        var result = new List<SmartArtInfo>();
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return result;
        var slidePart = parts[slideIdx - 1];
        var slide = GetSlide(slidePart);
        var spTree = slide.CommonSlideData?.ShapeTree;
        if (spTree == null) return result;

        var ns = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
        foreach (var gf in spTree.Descendants<DocumentFormat.OpenXml.Presentation.GraphicFrame>())
        {
            var relIds = gf.Descendants().FirstOrDefault(e =>
                e.LocalName == "relIds" && e.NamespaceUri == ns);
            if (relIds == null) continue;

            string? dRid = null, lRid = null, cRid = null, qRid = null;
            foreach (var a in relIds.GetAttributes())
            {
                var ln = a.LocalName;
                var v = a.Value;
                if (ln == "dm") dRid = v;
                else if (ln == "lo") lRid = v;
                else if (ln == "cs") cRid = v;
                else if (ln == "qs") qRid = v;
            }
            if (dRid == null || lRid == null || cRid == null || qRid == null) continue;

            string? xmlFor(string rid)
            {
                try
                {
                    var part = slidePart.GetPartById(rid);
                    if (part is DiagramDataPart d) return d.DataModelRoot?.OuterXml;
                    if (part is DiagramLayoutDefinitionPart l) return l.LayoutDefinition?.OuterXml;
                    if (part is DiagramColorsPart c) return c.ColorsDefinition?.OuterXml;
                    if (part is DiagramStylePart s) return s.StyleDefinition?.OuterXml;
                }
                catch { }
                return null;
            }

            var dXml = xmlFor(dRid);
            var lXml = xmlFor(lRid);
            var cXml = xmlFor(cRid);
            var qXml = xmlFor(qRid);
            if (dXml == null || lXml == null || cXml == null || qXml == null) continue;

            result.Add(new SmartArtInfo(
                GraphicFrameXml: gf.OuterXml,
                DataRelId: dRid, LayoutRelId: lRid, ColorsRelId: cRid, QuickStyleRelId: qRid,
                DataXml: dXml, LayoutXml: lXml, ColorsXml: cXml, QuickStyleXml: qXml));
        }
        return result;
    }

    /// <summary>
    /// Resolve a SmartArt sub-part's zip-URI for raw-set targeting. Given a
    /// slide index and a rId (data/layout/colors/quickStyle), returns
    /// e.g. "/ppt/diagrams/data1.xml". Returns null if the rId does not
    /// resolve to a known diagram part type.
    /// </summary>
    internal string? GetSmartArtPartUri(int slideIdx, string relId)
    {
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return null;
        var slidePart = parts[slideIdx - 1];
        try
        {
            var part = slidePart.GetPartById(relId);
            if (part is DiagramDataPart or DiagramLayoutDefinitionPart
                or DiagramColorsPart or DiagramStylePart)
            {
                return part.Uri.OriginalString;
            }
        }
        catch { }
        return null;
    }

    /// <summary>
    /// Per-slide video/audio info for PptxBatchEmitter Phase 3c-media
    /// passthrough. Returns one entry per &lt;p:pic&gt; on the slide whose
    /// nvPr carries &lt;a:videoFile&gt; or &lt;a:audioFile&gt;. Each entry
    /// includes the &lt;p:pic&gt; XML verbatim plus the source's three rIds
    /// (link/media/thumbnail) and the underlying binary streams, so the
    /// emitter can issue an `add-part video` (or `audio`) + a `raw-set`
    /// append on /p:sld/p:cSld/p:spTree that round-trips byte-equal.
    /// </summary>
    internal readonly record struct MediaInfo(
        string PicXml,
        bool IsVideo,
        string LinkRelId,
        string MediaEmbedRelId,
        string ThumbnailRelId,
        byte[] MediaBytes,
        string MediaContentType,
        string MediaExtension,
        byte[] ThumbnailBytes,
        string ThumbnailContentType);

    internal IReadOnlyList<MediaInfo> GetMediaOnSlide(int slideIdx)
    {
        var result = new List<MediaInfo>();
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return result;
        var slidePart = parts[slideIdx - 1];
        var slide = GetSlide(slidePart);
        var spTree = slide.CommonSlideData?.ShapeTree;
        if (spTree == null) return result;

        foreach (var pic in spTree.Descendants<Picture>())
        {
            var nvPr = pic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties;
            if (nvPr == null) continue;

            var videoFile = nvPr.GetFirstChild<DocumentFormat.OpenXml.Drawing.VideoFromFile>();
            var audioFile = nvPr.GetFirstChild<DocumentFormat.OpenXml.Drawing.AudioFromFile>();
            bool isVideo = videoFile != null;
            bool isAudio = audioFile != null;
            if (!isVideo && !isAudio) continue;

            string? linkRid = isVideo ? videoFile?.Link?.Value : audioFile?.Link?.Value;
            if (string.IsNullOrEmpty(linkRid)) continue;

            // Locate the p14:media extension carrying the MediaReference rId.
            string? mediaEmbedRid = null;
            var p14Media = nvPr.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().FirstOrDefault();
            if (p14Media?.Embed?.Value != null) mediaEmbedRid = p14Media.Embed.Value;
            if (string.IsNullOrEmpty(mediaEmbedRid)) continue;

            // Thumbnail rId from blipFill.
            var blip = pic.BlipFill?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Blip>();
            var thumbRid = blip?.Embed?.Value;
            if (string.IsNullOrEmpty(thumbRid)) continue;

            // Resolve media binary via either rId. Both VideoReference and
            // MediaReference point at the same MediaDataPart.
            byte[]? mediaBytes = null;
            string? mediaCT = null;
            string? mediaExt = null;
            try
            {
                MediaDataPart? mdp = null;
                foreach (var rel in slidePart.DataPartReferenceRelationships)
                {
                    if (rel.Id == linkRid || rel.Id == mediaEmbedRid)
                    {
                        if (rel.DataPart is MediaDataPart mdp2) { mdp = mdp2; break; }
                    }
                }
                if (mdp != null)
                {
                    using var s = mdp.GetStream();
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    mediaBytes = ms.ToArray();
                    mediaCT = mdp.ContentType;
                    // Extract extension from the part Uri.
                    var u = mdp.Uri.OriginalString;
                    var dot = u.LastIndexOf('.');
                    mediaExt = dot > 0 ? u[dot..] : (isVideo ? ".mp4" : ".mp3");
                }
            }
            catch { }
            if (mediaBytes == null || mediaCT == null || mediaExt == null) continue;

            // Resolve thumbnail binary.
            byte[]? thumbBytes = null;
            string? thumbCT = null;
            try
            {
                var p = slidePart.GetPartById(thumbRid);
                if (p is ImagePart ip)
                {
                    using var s = ip.GetStream();
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    thumbBytes = ms.ToArray();
                    thumbCT = ip.ContentType;
                }
            }
            catch { }
            if (thumbBytes == null || thumbCT == null) continue;

            result.Add(new MediaInfo(
                PicXml: pic.OuterXml,
                IsVideo: isVideo,
                LinkRelId: linkRid,
                MediaEmbedRelId: mediaEmbedRid,
                ThumbnailRelId: thumbRid,
                MediaBytes: mediaBytes,
                MediaContentType: mediaCT,
                MediaExtension: mediaExt,
                ThumbnailBytes: thumbBytes,
                ThumbnailContentType: thumbCT));
        }
        return result;
    }

    /// <summary>
    /// Per-slide am3d 3D-model info for PptxBatchEmitter Phase 3c-3d
    /// passthrough. Returns one entry per &lt;mc:AlternateContent&gt; block
    /// whose &lt;mc:Choice Requires="am3d"&gt; carries an &lt;am3d:model3d&gt;
    /// element. Each entry includes the AlternateContent XML verbatim plus
    /// the source's two rIds (the model3d ExtendedPart and the shared
    /// thumbnail ImagePart) and the underlying binary streams, so the
    /// emitter can issue an `add-part model3d` + a `raw-set` append on
    /// /p:sld/p:cSld/p:spTree that round-trips byte-equal.
    /// </summary>
    internal readonly record struct Model3dInfo(
        string AlternateContentXml,
        string Model3dRelId,
        string ThumbnailRelId,
        byte[] Model3dBytes,
        string Model3dContentType,
        string Model3dExtension,
        byte[] ThumbnailBytes,
        string ThumbnailContentType);

    private static string InjectAmbientXmlnsOnRoot(string sliceXml,
        (string Prefix, string Uri)[] decls)
    {
        if (string.IsNullOrEmpty(sliceXml) || sliceXml[0] != '<') return sliceXml;
        int gt = sliceXml.IndexOf('>');
        if (gt <= 0) return sliceXml;
        var head = sliceXml[..gt];
        var tail = sliceXml[gt..];
        // Skip injecting decls that the root already carries (avoids
        // duplicate xmlns attributes which is illegal).
        var sb = new System.Text.StringBuilder(head);
        foreach (var (prefix, uri) in decls)
        {
            if (head.Contains($"xmlns:{prefix}=\"", StringComparison.Ordinal)) continue;
            // Only inject if the slice actually references this prefix.
            if (!sliceXml.Contains($"<{prefix}:", StringComparison.Ordinal)
                && !sliceXml.Contains($" {prefix}:", StringComparison.Ordinal))
                continue;
            sb.Append(" xmlns:").Append(prefix).Append("=\"").Append(uri).Append('"');
        }
        sb.Append(tail);
        return sb.ToString();
    }

    internal IReadOnlyList<Model3dInfo> GetModel3dOnSlide(int slideIdx)
    {
        var result = new List<Model3dInfo>();
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return result;
        var slidePart = parts[slideIdx - 1];

        // Read raw slide XML — am3d lives under <mc:AlternateContent> which
        // the typed SDK tree exposes only as OpenXmlUnknownElement. Walking
        // the raw stream avoids the awkward typed traversal and gives us a
        // straight slice for raw-set passthrough.
        string slideXml;
        using (var s = slidePart.GetStream())
        using (var sr = new StreamReader(s))
            slideXml = sr.ReadToEnd();

        // Parse the slide XML and walk for <mc:AlternateContent> elements
        // whose Choice has Requires="am3d". We extract slices by element
        // identity rather than textual regex because the SDK may re-prefix
        // the relationships namespace (e.g. p10:embed instead of r:embed)
        // when re-serialising an unknown subtree on round-trip — text-only
        // matching misses these.
        System.Xml.Linq.XDocument slideDoc;
        try { slideDoc = System.Xml.Linq.XDocument.Parse(slideXml); }
        catch { return result; }
        System.Xml.Linq.XNamespace mcNs = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        System.Xml.Linq.XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        System.Xml.Linq.XNamespace am3dNs = "http://schemas.microsoft.com/office/drawing/2017/model3d";
        System.Xml.Linq.XNamespace aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

        foreach (var ac in slideDoc.Descendants(mcNs + "AlternateContent").ToList())
        {
            var choice = ac.Element(mcNs + "Choice");
            var requires = choice?.Attribute("Requires")?.Value;
            if (choice == null || requires == null || !requires.Contains("am3d", StringComparison.Ordinal)) continue;

            var model3d = choice.Descendants(am3dNs + "model3d").FirstOrDefault();
            var m3dRidAttr = model3d?.Attribute(rNs + "embed");
            if (m3dRidAttr == null) continue;
            var m3dRid = m3dRidAttr.Value;

            // Thumbnail rId: prefer <am3d:blip r:embed=…> in <am3d:raster>;
            // fall back to <a:blip r:embed=…> in the Fallback <p:pic>.
            string? thumbRid = model3d!.Descendants(am3dNs + "blip")
                .Select(b => b.Attribute(rNs + "embed")?.Value)
                .FirstOrDefault(v => !string.IsNullOrEmpty(v));
            if (string.IsNullOrEmpty(thumbRid))
            {
                thumbRid = ac.Descendants(aNs + "blip")
                    .Select(b => b.Attribute(rNs + "embed")?.Value)
                    .FirstOrDefault(v => !string.IsNullOrEmpty(v));
            }
            if (string.IsNullOrEmpty(thumbRid)) continue;

            // Re-serialise the AlternateContent subtree as the slice. This
            // gives a canonical XML form that NormalizeSlideRawSlice can
            // further harmonise — both round-1 (read from source) and
            // round-2 (read from B.pptx) paths funnel through XLinq here,
            // so namespace prefix drift (p10:embed vs r:embed) reconciles.
            var slice = ac.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);

            // Resolve the model3d binary via the slide's ExtendedParts
            // relationship dictionary. AddExtendedPart routes to
            // ExtendedParts, NOT to a typed part of slidePart.Parts; we
            // must reach in by rel id.
            byte[]? m3dBytes = null;
            string? m3dCT = null;
            string m3dExt = ".glb";
            try
            {
                // Extended parts (created via AddExtendedPart) live in
                // slidePart.Parts under their pinned rel id; GetPartById
                // resolves any internal child part by relationship id
                // regardless of typing.
                var part = slidePart.GetPartById(m3dRid);
                if (part != null)
                {
                    using var s = part.GetStream();
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    m3dBytes = ms.ToArray();
                    m3dCT = part.ContentType;
                    var u = part.Uri.OriginalString;
                    var dot = u.LastIndexOf('.');
                    if (dot > 0) m3dExt = u[dot..];
                }
            }
            catch { }
            if (m3dBytes == null || string.IsNullOrEmpty(m3dCT)) continue;

            // Resolve thumbnail bytes from the shared ImagePart.
            byte[]? thumbBytes = null;
            string? thumbCT = null;
            try
            {
                var tp = slidePart.GetPartById(thumbRid);
                if (tp is ImagePart ip)
                {
                    using var s = ip.GetStream();
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    thumbBytes = ms.ToArray();
                    thumbCT = ip.ContentType;
                }
            }
            catch { }
            if (thumbBytes == null || string.IsNullOrEmpty(thumbCT)) continue;

            result.Add(new Model3dInfo(
                AlternateContentXml: slice,
                Model3dRelId: m3dRid,
                ThumbnailRelId: thumbRid,
                Model3dBytes: m3dBytes,
                Model3dContentType: m3dCT!,
                Model3dExtension: m3dExt,
                ThumbnailBytes: thumbBytes,
                ThumbnailContentType: thumbCT!));
        }
        return result;
    }

    // Resolve a /slide[N]/picture[M] path's image bytes for base64-inline emit.
    // Mirrors WordHandler.GetImageBinary's contract: returns null if the path
    // does not resolve to a Picture with an embedded ImagePart.
    public (byte[] Bytes, string ContentType)? GetImageBinary(string picturePath)
    {
        // Accept both `picture[N]` positional and `picture[@id=N]` cNvPr-id
        // segment forms (BuildElementPathSegment emits @id= when the shape
        // carries a cNvPr id, which Pictures always do).
        var m = Regex.Match(picturePath,
            @"^/slide\[(\d+)\]/(?:.+/)?picture\[(?:@id=)?(\d+)\]$");
        if (!m.Success) return null;
        var slideIdx = int.Parse(m.Groups[1].Value);
        var idOrIdx = int.Parse(m.Groups[2].Value);
        var byId = picturePath.Contains("@id=", StringComparison.Ordinal);
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return null;
        var slidePart = parts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (shapeTree == null) return null;
        var pictures = shapeTree.Descendants<Picture>().ToList();
        Picture? pic = null;
        if (byId)
        {
            pic = pictures.FirstOrDefault(p =>
            {
                var pid = p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value;
                return pid.HasValue && pid.Value == (uint)idOrIdx;
            });
        }
        else
        {
            if (idOrIdx >= 1 && idOrIdx <= pictures.Count) pic = pictures[idOrIdx - 1];
        }
        if (pic == null) return null;
        var blip = pic.BlipFill?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Blip>();
        var embedId = blip?.Embed?.Value;
        if (string.IsNullOrEmpty(embedId)) return null;
        try
        {
            var part = slidePart.GetPartById(embedId);
            using var src = part.GetStream();
            using var ms = new MemoryStream();
            src.CopyTo(ms);
            return (ms.ToArray(), part.ContentType);
        }
        catch { return null; }
    }

    // ==================== Private Helpers ====================

    private static Slide GetSlide(SlidePart part) =>
        part.Slide ?? throw new InvalidOperationException("Corrupt file: slide data missing");

    private IEnumerable<SlidePart> GetSlideParts()
    {
        var presentation = _doc.PresentationPart?.Presentation;
        var slideIdList = presentation?.GetFirstChild<SlideIdList>();
        if (slideIdList == null) yield break;

        foreach (var slideId in slideIdList.Elements<SlideId>())
        {
            var relId = slideId.RelationshipId?.Value;
            if (relId == null) continue;
            yield return (SlidePart)_doc.PresentationPart!.GetPartById(relId);
        }
    }

}
