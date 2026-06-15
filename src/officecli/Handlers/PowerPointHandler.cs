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
    // Number of elements a no-slash selector Set matched and mutated (Sheet1!row[...]).
    // Read by the CLI/resident to echo the multi-element change count.
    public int LastSelectorSetCount { get; internal set; }

    /// <summary>
    /// Set true by Add/Set/Remove/RawSet, consumed by Save/Dispose to decide
    /// whether to stamp <c>docProps/custom.xml</c> with an OfficeCLI audit
    /// trail. Pure Get/Query sessions leave this false.
    /// </summary>
    internal bool Modified { get; set; }

    /// <summary>
    /// Enumerate every <see cref="OpenXmlPart"/> in the package (transitive
    /// walk via the SDK's own <c>GetAllParts</c> extension) yielding each
    /// part's zip-URI (<c>OpenXmlPart.Uri.OriginalString</c>). Used by the
    /// batch emitter's auxiliary-parts scan to surface warnings for parts the
    /// dump surface does not round-trip (tableStyles, viewProps,
    /// handoutMasters, printerSettings, customXml, embedded fonts, tags,
    /// user docProps). Mirrors <see cref="WordHandler.EnumeratePartUris"/>.
    /// </summary>
    internal IEnumerable<string> EnumeratePartUris()
    {
        // Two complementary sources (same rationale as WordHandler):
        //   1. SDK part graph (GetAllParts) — only reachable via the
        //      relationship graph; misses orphan parts but matches every
        //      real pptx file produced by PowerPoint / python-pptx / WPS.
        //   2. Raw zip entries — catches orphan parts dropped by failed
        //      saves and anything the SDK refuses to surface.
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var part in _doc.GetAllParts())
        {
            var u = part.Uri.OriginalString;
            if (seen.Add(u)) yield return u;
        }
        if (File.Exists(_filePath))
        {
            List<string> zipEntries;
            try
            {
                using var fs = File.Open(_filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var zip = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Read);
                zipEntries = zip.Entries.Select(e => e.FullName).ToList();
            }
            catch { yield break; }
            foreach (var name in zipEntries)
            {
                if (string.IsNullOrEmpty(name)) continue;
                var u = name.StartsWith("/") ? name : "/" + name;
                if (seen.Add(u)) yield return u;
            }
        }
    }

    /// <summary>
    /// Surface the <c>docProps/custom.xml</c> property names (if any).
    /// Used by the batch emitter to detect user-defined custom document
    /// properties whose values are silently dropped by dump (only
    /// <c>OfficeCLI.*</c> values are auto-restamped on save; user-authored
    /// props vanish). Mirrors <see cref="WordHandler.EnumerateCustomDocPropertyNames"/>.
    /// </summary>
    internal IReadOnlyList<string> EnumerateCustomDocPropertyNames()
    {
        var part = _doc.CustomFilePropertiesPart;
        if (part?.Properties == null) return Array.Empty<string>();
        var names = new List<string>();
        foreach (var p in part.Properties.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>())
        {
            var n = p.Name?.Value;
            if (!string.IsNullOrEmpty(n)) names.Add(n!);
        }
        return names;
    }

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
        if (xpath == null) throw new ArgumentNullException(nameof(xpath));
        if (action == null) throw new ArgumentNullException(nameof(action));
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
            bool nmCreated = nmPart == null;
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
            // Register <p:notesMasterIdLst> in presentation.xml. AddNewPart only
            // wires the relationship; without the IdLst element the part is an
            // orphan reference and PowerPoint refuses the file with a schema
            // error ("unexpected child element in /p:presentation"). Schema order
            // (CT_Presentation): sldMasterIdLst -> notesMasterIdLst ->
            // handoutMasterIdLst -> sldIdLst -> sldSz -> notesSz -> ...
            if (nmCreated)
            {
                var pres = presentationPart.Presentation
                    ?? throw new InvalidOperationException("No presentation");
                if (pres.NotesMasterIdList == null)
                {
                    var nmRelId = presentationPart.GetIdOfPart(nmPart);
                    var nmIdLst = new NotesMasterIdList(
                        new NotesMasterId { Id = nmRelId });
                    // Insert after sldMasterIdLst if present, else prepend.
                    var sldMasterIdLst = pres.SlideMasterIdList;
                    if (sldMasterIdLst != null)
                        pres.InsertAfter(nmIdLst, sldMasterIdLst);
                    else
                        pres.PrependChild(nmIdLst);
                }
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

            // Every SlideMasterPart must reference a ThemePart. Give each grown
            // master its OWN distinct (placeholder) theme part rather than sharing
            // the presentation's primary theme: `add-part theme` overwrites it
            // with the source's per-master theme content, and a later
            // delete-of-its-own-theme must not orphan a theme another master
            // shares. Seed minimal valid theme content (replaced on replay when
            // the deck carries per-master themes).
            var grownTheme = newPart.AddNewPart<ThemePart>();
            if (presentationPart.ThemePart?.Theme != null)
                grownTheme.Theme = (DocumentFormat.OpenXml.Drawing.Theme)presentationPart.ThemePart.Theme.CloneNode(true);
            else
                grownTheme.Theme = new DocumentFormat.OpenXml.Drawing.Theme();
            grownTheme.Theme.Save();

            sldMasterIdLst.AppendChild(new SlideMasterId { Id = nextId++, RelationshipId = rId });
        }
        presentation.Save();
    }

    // PowerPoint requires every <p:sldMasterId>/@id AND every
    // <p:sldLayoutId>/@id (across all masters) to be unique within one shared
    // id space — PowerPoint's own decks number them as a single monotonic run
    // (master=N, its layouts=N+1.., next master continues after the last
    // layout). The SDK and our GrowSlideMasterParts pick master ids starting at
    // 2147483648 while only avoiding existing *master* ids, so a grown master
    // can land on a value a source master's sldLayoutIdLst already uses
    // (e.g. master2 id 2147483649 == master1 layout1 id 2147483649). The SDK
    // schema-validates this duplicate fine, but PowerPoint rejects the package
    // with 0x80070570 ("file or directory corrupted"). Reconcile at save: keep
    // layout ids fixed (slides never reference them), reassign only colliding
    // master ids to fresh unused values. Runs once per save cycle, after every
    // raw-set has landed the masters' real sldLayoutIdLst.
    private void ReconcileSlideMasterIds()
    {
        var presentation = _doc.PresentationPart?.Presentation;
        var sldMasterIdLst = presentation?.SlideMasterIdList;
        if (presentation == null || sldMasterIdLst == null) return;

        // Collect every layout id (these stay put) plus seed with the values
        // we will keep for masters as we walk them.
        var used = new HashSet<uint>();
        foreach (var mp in _doc.PresentationPart!.SlideMasterParts)
        {
            var layoutIds = mp.SlideMaster?.SlideLayoutIdList?.Elements<SlideLayoutId>();
            if (layoutIds == null) continue;
            foreach (var lid in layoutIds)
                if (lid.Id?.Value is uint v) used.Add(v);
        }

        bool changed = false;
        foreach (var smId in sldMasterIdLst.Elements<SlideMasterId>())
        {
            uint id = smId.Id?.Value ?? 2147483648u;
            if (used.Contains(id))
            {
                uint repl = 2147483648u;
                while (used.Contains(repl)) repl++;
                smId.Id = repl;
                id = repl;
                changed = true;
            }
            used.Add(id);
        }

        if (changed) presentation.Save();
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

                // Inline diagram part content. The dump emitter carries each
                // sub-part's verbatim XML here so add-part writes it directly
                // into the freshly-created part. This supersedes the legacy
                // "create seed + separate raw-set replace" flow: the SDK
                // allocates the diagram parts under /ppt/graphics/dataN.xml
                // (its own naming base, N incrementing package-globally),
                // NOT the source's /ppt/diagrams/data1.xml, so a raw-set
                // pre-targeted at the source URI never resolved (FindPartByZipUri
                // miss) and the parts persisted EMPTY → blank/broken SmartArt.
                // Writing content at creation time is URI-agnostic and robust.
                string? dataXml     = properties != null && properties.TryGetValue("dataXml", out var dxv) ? dxv : null;
                string? layoutXml   = properties != null && properties.TryGetValue("layoutXml", out var lxv) ? lxv : null;
                string? colorsXml   = properties != null && properties.TryGetValue("colorsXml", out var cxv) ? cxv : null;
                string? qsXml       = properties != null && properties.TryGetValue("quickStyleXml", out var qxv) ? qxv : null;
                string? drawingXml  = properties != null && properties.TryGetValue("drawingXml", out var drxv) ? drxv : null;
                string? drawingRelId = properties != null && properties.TryGetValue("drawingRelId", out var drrv) ? drrv : null;

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

                // Write the real content when supplied; else seed a minimal
                // typed root (keeps direct CLI `add-part smartart` usable).
                WriteDiagramPartXml(dataPart, dataXml, () =>
                    new DocumentFormat.OpenXml.Drawing.Diagrams.DataModelRoot(
                        new DocumentFormat.OpenXml.Drawing.Diagrams.PointList(),
                        new DocumentFormat.OpenXml.Drawing.Diagrams.ConnectionList()));
                WriteDiagramPartXml(layoutPart, layoutXml,
                    () => new DocumentFormat.OpenXml.Drawing.Diagrams.LayoutDefinition());
                WriteDiagramPartXml(colorsPart, colorsXml,
                    () => new DocumentFormat.OpenXml.Drawing.Diagrams.ColorsDefinition());
                WriteDiagramPartXml(stylePart, qsXml,
                    () => new DocumentFormat.OpenXml.Drawing.Diagrams.StyleDefinition());

                // The DSP cached-drawing part is referenced from the data XML
                // via <dsp:dataModelExt relId="...">. That relId resolves
                // against the SLIDE part's relationships (the drawing part is
                // a slide-level part of type .../2007/relationships/diagramDrawing,
                // sibling to the data/layout/colors/qs rels — NOT a child of
                // the data part). Create it on saSlidePart with the pinned
                // relId so the reference resolves; otherwise PowerPoint
                // refuses the file (0x80070570).
                if (!string.IsNullOrEmpty(drawingXml) && !string.IsNullOrEmpty(drawingRelId))
                {
                    var drawingPart = saSlidePart.AddNewPart<DiagramPersistLayoutPart>(
                        "application/vnd.ms-office.drawingml.diagramDrawing+xml", drawingRelId);
                    // drawingXml is always present on this branch; the seed
                    // fallback is unreachable here but supplied for the typed
                    // signature.
                    WriteDiagramPartXml(drawingPart, drawingXml,
                        () => new DocumentFormat.OpenXml.Drawing.Diagrams.DataModelRoot());
                }

                // Encode all four rIds in the RelId field — callers (batch
                // emit / replay) need to know each part's id to write the
                // matching dgm:relIds on the graphicFrame. Format:
                // "data=rIdX;layout=rIdY;colors=rIdZ;quickStyle=rIdW".
                var dataActualRid   = saSlidePart.GetIdOfPart(dataPart);
                var layoutActualRid = saSlidePart.GetIdOfPart(layoutPart);
                var colorsActualRid = saSlidePart.GetIdOfPart(colorsPart);
                var styleActualRid  = saSlidePart.GetIdOfPart(stylePart);

                // Inject a minimal <p:graphicFrame> into the slide's spTree
                // so GetSmartArtsOnSlide (which finds SmartArt only via the
                // graphicFrame + dgm:relIds anchor) can see this SmartArt on
                // the next dump. Without the host frame, a SmartArt created
                // by direct `add-part smartart` is silently dropped on dump.
                // The dump-emitter path passes `skip-frame=true` because it
                // raw-set appends the source's full graphicFrame (with real
                // position/size/name) immediately after — otherwise we'd
                // emit a stub-plus-real pair on every replay.
                var skipFrame = properties != null
                    && properties.TryGetValue("skip-frame", out var sf)
                    && (sf == "true" || sf == "1");
                if (!skipFrame)
                {
                    var saSlide = GetSlide(saSlidePart);
                    var saSpTree = saSlide.CommonSlideData?.ShapeTree;
                    if (saSpTree != null)
                    {
                        // Allocate a non-colliding cNvPr id within the slide.
                        // R42-T1: spTree carries pptx-namespaced <p:nvSpPr>/<p:nvGrpSpPr>/<p:nvGraphicFramePr>
                        // wrappers whose cNvPr maps to DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties,
                        // NOT the Drawing-namespace type. The wrong SDK type matched zero descendants,
                        // pinning nextId at 1 and colliding with nvGrpSpPr cNvPr id=1.
                        uint nextId = 1;
                        foreach (var nv in saSpTree.Descendants<DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties>())
                        {
                            if (nv.Id?.Value >= nextId) nextId = nv.Id!.Value + 1;
                        }
                        // CT_GraphicalObjectFrame: nvGraphicFramePr / xfrm /
                        // graphic. xfrm carries a default 6"x4.5" host area
                        // anchored at (1in, 1in) — PowerPoint will rescale
                        // on first render anyway; the values just keep the
                        // shape selectable.
                        const long EmuIn = 914400;
                        var gf = new DocumentFormat.OpenXml.Presentation.GraphicFrame(
                            new DocumentFormat.OpenXml.Presentation.NonVisualGraphicFrameProperties(
                                new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties { Id = nextId, Name = $"Diagram {nextId}" },
                                new DocumentFormat.OpenXml.Presentation.NonVisualGraphicFrameDrawingProperties(
                                    new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks { NoChangeAspect = true }),
                                new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()),
                            new DocumentFormat.OpenXml.Presentation.Transform(
                                new DocumentFormat.OpenXml.Drawing.Offset { X = 1 * EmuIn, Y = 1 * EmuIn },
                                new DocumentFormat.OpenXml.Drawing.Extents { Cx = 6 * EmuIn, Cy = (long)(4.5 * EmuIn) }),
                            new DocumentFormat.OpenXml.Drawing.Graphic(
                                new DocumentFormat.OpenXml.Drawing.GraphicData(
                                    new DocumentFormat.OpenXml.Drawing.Diagrams.RelationshipIds
                                    {
                                        DataPart = dataActualRid,
                                        LayoutPart = layoutActualRid,
                                        ColorPart = colorsActualRid,
                                        StylePart = styleActualRid,
                                    })
                                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/diagram" }));
                        saSpTree.AppendChild(gf);
                        saSlide.Save();
                    }
                }

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

            case "ole":
            case "oleobject":
                // Phase 3c-ole. Mirrors Phase 3c-media (video/audio) and
                // Phase 3c-3d. PPT OLE embed lives inside a <p:graphicFrame>:
                //   <a:graphicData uri=".../presentationml/2006/ole">
                //     <p:oleObj progId="…" showAsIcon="1" r:id="rIdA" …>
                //       <p:embed/>
                //       <p:pic>... <a:blip r:embed="rIdB"/> ...</p:pic>
                //     </p:oleObj>
                // Two rels back the slice:
                //  - relType .../officeDocument/2006/relationships/oleObject
                //    -> the embedded payload. EmbeddedPackagePart for OOXML
                //       containers (.xlsx/.docx/.pptx + macro/template
                //       siblings, content type vnd.openxmlformats-…), else
                //       EmbeddedObjectPart for generic binaries (.bin OLE10,
                //       legacy .doc/.xls, PDF, etc.).
                //  - relType .../officeDocument/2006/relationships/image
                //    -> the icon/thumbnail ImagePart for the inner <p:pic>.
                //
                // Props (all optional except data + thumbnail-data):
                //   data                   = base64 OLE payload bytes
                //   content-type           = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                //                            (for .xlsx), or any package /
                //                            generic OLE content-type. Drives
                //                            the Package vs Object auto-select.
                //   extension              = ".xlsx" / ".docx" / ".bin" (drives
                //                            the part URI extension when we
                //                            fall through to EmbeddedObjectPart;
                //                            EmbeddedPackagePart picks its own
                //                            extension from the PartTypeInfo).
                //   ole-rid                = pinned OLE part rId
                //   thumbnail-data         = base64 icon image bytes
                //   thumbnail-content-type = "image/png" (default) / "image/jpeg"
                //   thumbnail-rid          = pinned thumbnail ImagePart rId
                //
                // Returned encoded relId: "ole=rIdA;thumbnail=rIdB".
                // No shape XML is inserted under the slide; the companion
                // raw-set append carries the full <p:graphicFrame> verbatim
                // with the matching pinned rIds.
                var oleSlideMatchAP = System.Text.RegularExpressions.Regex.Match(
                    parentPartPath, @"^/slide\[(\d+)\]$");
                if (!oleSlideMatchAP.Success)
                    throw new ArgumentException(
                        $"{partType} must be added under a slide: add-part <file> '/slide[N]' --type {partType}");
                var oleSlideIdxAP = int.Parse(oleSlideMatchAP.Groups[1].Value);
                var oleSlidePartsAP = GetSlideParts().ToList();
                if (oleSlideIdxAP < 1 || oleSlideIdxAP > oleSlidePartsAP.Count)
                    throw new ArgumentException($"Slide index {oleSlideIdxAP} out of range");
                var oleSlidePartAP = oleSlidePartsAP[oleSlideIdxAP - 1];

                if (properties == null || !properties.TryGetValue("data", out var oleB64) || string.IsNullOrEmpty(oleB64))
                    throw new ArgumentException(
                        $"add-part {partType} requires property 'data' (base64 OLE payload bytes)");
                byte[] oleBytes;
                try { oleBytes = Convert.FromBase64String(oleB64); }
                catch (FormatException) { throw new ArgumentException($"add-part {partType}: 'data' is not valid base64"); }

                var oleContentTypeAP = properties.TryGetValue("content-type", out var olct) && !string.IsNullOrEmpty(olct)
                    ? olct
                    : "application/vnd.openxmlformats-officedocument.oleObject";
                var oleExtAP = properties.TryGetValue("extension", out var olxt) && !string.IsNullOrEmpty(olxt)
                    ? (olxt.StartsWith('.') ? olxt : "." + olxt)
                    : ".bin";

                string? pinnedOleRid   = properties.TryGetValue("ole-rid", out var olr) ? olr : null;
                string? pinnedOleThumb = properties.TryGetValue("thumbnail-rid", out var oltr) ? oltr : null;

                // Auto-select EmbeddedPackagePart vs EmbeddedObjectPart by
                // content-type. The OOXML package family carries content
                // types starting with "application/vnd.openxmlformats-".
                // Everything else (oleObject generic, application/pdf,
                // application/octet-stream, vnd.ms-excel for legacy .xls,
                // etc.) goes through EmbeddedObjectPart with its raw
                // content-type preserved on the part. EmbeddedPackagePart
                // requires a typed PartTypeInfo (EmbeddedPackagePartType.*);
                // map extension → typed value via OleHelper, fall back to
                // EmbeddedObjectPart if the extension is unrecognized.
                OpenXmlPart olePart;
                PartTypeInfo? packagePti = null;
                bool oleIsPackage = oleContentTypeAP.StartsWith(
                    "application/vnd.openxmlformats-officedocument.",
                    StringComparison.OrdinalIgnoreCase)
                    && !oleContentTypeAP.Equals(
                        "application/vnd.openxmlformats-officedocument.oleObject",
                        StringComparison.OrdinalIgnoreCase);
                if (oleIsPackage)
                {
                    // GetPackagePartTypeInfo takes a path; we feed a synthetic
                    // path with the extension. Returns null for unknown exts
                    // — in that case route to EmbeddedObjectPart so the bytes
                    // still survive (with a generic part type).
                    packagePti = OfficeCli.Core.OleHelper.GetPackagePartTypeInfo("x" + oleExtAP);
                    if (packagePti == null) oleIsPackage = false;
                }
                if (oleIsPackage)
                {
                    olePart = !string.IsNullOrEmpty(pinnedOleRid)
                        ? oleSlidePartAP.AddEmbeddedPackagePart(packagePti!.Value, pinnedOleRid)
                        : oleSlidePartAP.AddEmbeddedPackagePart(packagePti!.Value);
                }
                else
                {
                    olePart = !string.IsNullOrEmpty(pinnedOleRid)
                        ? oleSlidePartAP.AddEmbeddedObjectPart(oleContentTypeAP, pinnedOleRid)
                        : oleSlidePartAP.AddEmbeddedObjectPart(oleContentTypeAP);
                }
                using (var s = new MemoryStream(oleBytes)) olePart.FeedData(s);
                var oleActualRid = oleSlidePartAP.GetIdOfPart(olePart);

                // Thumbnail icon image. Caller may omit thumbnail-data; we
                // then seed the same 1x1 transparent PNG used by video/audio
                // and the model3d placeholder paths.
                byte[] oleThumbBytes;
                PartTypeInfo oleThumbType;
                if (properties.TryGetValue("thumbnail-data", out var oltdB64) && !string.IsNullOrEmpty(oltdB64))
                {
                    try { oleThumbBytes = Convert.FromBase64String(oltdB64); }
                    catch (FormatException) { throw new ArgumentException($"add-part {partType}: 'thumbnail-data' is not valid base64"); }
                    var oleThumbCT = properties.TryGetValue("thumbnail-content-type", out var oltct) && !string.IsNullOrEmpty(oltct)
                        ? oltct
                        : "image/png";
                    oleThumbType = oleThumbCT switch {
                        "image/png" => ImagePartType.Png, "image/jpeg" => ImagePartType.Jpeg,
                        "image/gif" => ImagePartType.Gif, "image/bmp" => ImagePartType.Bmp,
                        "image/tiff" => ImagePartType.Tiff,
                        "image/x-emf" => ImagePartType.Emf, "image/x-wmf" => ImagePartType.Wmf,
                        _ => ImagePartType.Png };
                }
                else
                {
                    oleThumbBytes = new byte[]
                    {
                        0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                        0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                        0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,0x89,
                        0x00,0x00,0x00,0x0D,0x49,0x44,0x41,0x54,
                        0x08,0xD7,0x63,0x60,0x60,0x60,0x60,0x00,0x00,0x00,0x05,0x00,0x01,0x87,0xA1,0x4E,0xD4,
                        0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,0x42,0x60,0x82
                    };
                    oleThumbType = ImagePartType.Png;
                }
                var oleThumbPart = !string.IsNullOrEmpty(pinnedOleThumb)
                    ? oleSlidePartAP.AddImagePart(oleThumbType, pinnedOleThumb)
                    : oleSlidePartAP.AddImagePart(oleThumbType);
                using (var s = new MemoryStream(oleThumbBytes)) oleThumbPart.FeedData(s);
                var oleThumbActualRid = oleSlidePartAP.GetIdOfPart(oleThumbPart);

                var encodedOle = $"ole={oleActualRid};thumbnail={oleThumbActualRid}";
                return (encodedOle, parentPartPath);

            case "image":
                // Generic ImagePart attached to a slide / slideMaster / slideLayout
                // / notesMaster. Used by dump→replay to round-trip image
                // references that live in raw-set'd master/layout XML — the
                // master raw XML carries `r:embed="rIdN"` on <p:pic> blipFills,
                // but the underlying ImagePart is enumerated separately by the
                // SDK and was never re-emitted, so post-replay validate flagged
                // "rIdN does not exist" on slideMaster2 etc.
                //
                // Required properties:
                //   data         = base64 image bytes
                // Optional properties:
                //   content-type = "image/png" / "image/jpeg" / "image/gif" / …
                //                  (default "image/png")
                //   rid          = pinned relationship id; when omitted the SDK
                //                  allocates one
                OpenXmlPart imageHost = parentPartPath switch
                {
                    "/" => presentationPart,
                    "/notesMaster" => (OpenXmlPart?)presentationPart.NotesMasterPart
                        ?? throw new ArgumentException("add-part image /notesMaster: no notes master in deck"),
                    // A theme's <a:fmtScheme>/<a:fillStyleLst>/<a:blipFill> can
                    // reference a texture image via r:embed; the raw-set'd theme
                    // XML carries that reference but the ImagePart lives in the
                    // theme's own .rels, enumerated separately. Without carrying
                    // it the rId dangles and PowerPoint refuses to open the deck.
                    "/theme" => (OpenXmlPart?)presentationPart.ThemePart
                        ?? throw new ArgumentException("add-part image /theme: presentation has no theme part"),
                    _ => null!,
                };
                if (imageHost == null)
                {
                    var smMatch = System.Text.RegularExpressions.Regex.Match(parentPartPath, @"^/slideMaster\[(\d+)\]$");
                    var slMatch = smMatch.Success ? null : System.Text.RegularExpressions.Regex.Match(parentPartPath, @"^/slideLayout\[(\d+)\]$");
                    var sldMatch = (smMatch.Success || (slMatch?.Success ?? false)) ? null : System.Text.RegularExpressions.Regex.Match(parentPartPath, @"^/slide\[(\d+)\]$");
                    // CONSISTENCY(notes-image-host): mirror /slide[N]/master/layout
                    // entries — dump emits add-part image /noteSlide[N] BEFORE the
                    // notes raw-set replace so r:embed references in the notesSlide
                    // XML resolve to a real ImagePart on replay. Without this the
                    // post-replay notesSlide carries a dangling rId and PowerPoint
                    // renders the speaker-notes picture as a broken placeholder.
                    var nsMatch = (smMatch.Success || (slMatch?.Success ?? false) || (sldMatch?.Success ?? false))
                        ? null
                        : System.Text.RegularExpressions.Regex.Match(parentPartPath, @"^/noteSlide\[(\d+)\]$");
                    if (smMatch.Success)
                    {
                        var smIdx = int.Parse(smMatch.Groups[1].Value);
                        var smParts = presentationPart.SlideMasterParts.ToList();
                        if (smIdx < 1) throw new ArgumentException($"slideMaster index {smIdx} out of range");
                        if (smIdx > smParts.Count)
                        {
                            // CONSISTENCY(grow-on-rawset): mirror RawSet's auto-grow.
                            // Dump emits add-part image /slideMaster[N] BEFORE the
                            // raw-set replace, so on a blank target the master slot
                            // doesn't exist yet — grow to idx so the ImagePart can
                            // attach to the right host.
                            GrowSlideMasterParts(smIdx);
                            smParts = presentationPart.SlideMasterParts.ToList();
                            if (smIdx > smParts.Count)
                                throw new ArgumentException($"slideMaster index {smIdx} out of range (total {smParts.Count})");
                        }
                        imageHost = smParts[smIdx - 1];
                    }
                    else if (slMatch != null && slMatch.Success)
                    {
                        var slIdx = int.Parse(slMatch.Groups[1].Value);
                        var slParts = presentationPart.SlideMasterParts.SelectMany(m => m.SlideLayoutParts).ToList();
                        if (slIdx < 1) throw new ArgumentException($"slideLayout index {slIdx} out of range");
                        if (slIdx > slParts.Count)
                        {
                            GrowSlideLayoutParts(slIdx);
                            slParts = presentationPart.SlideMasterParts.SelectMany(m => m.SlideLayoutParts).ToList();
                            if (slIdx > slParts.Count)
                                throw new ArgumentException($"slideLayout index {slIdx} out of range (total {slParts.Count})");
                        }
                        imageHost = slParts[slIdx - 1];
                    }
                    else if (sldMatch != null && sldMatch.Success)
                    {
                        var sldIdx = int.Parse(sldMatch.Groups[1].Value);
                        var sldParts = GetSlideParts().ToList();
                        if (sldIdx < 1 || sldIdx > sldParts.Count)
                            throw new ArgumentException($"slide index {sldIdx} out of range");
                        imageHost = sldParts[sldIdx - 1];
                    }
                    else if (nsMatch != null && nsMatch.Success)
                    {
                        var nsIdx = int.Parse(nsMatch.Groups[1].Value);
                        var sldParts = GetSlideParts().ToList();
                        if (nsIdx < 1 || nsIdx > sldParts.Count)
                            throw new ArgumentException($"noteSlide index {nsIdx} out of range");
                        // NotesSlidePart may not exist yet on the replay target.
                        // EmitNotes always lands its typed `add notes` row BEFORE
                        // this add-part, but on a blank target with no prior
                        // notes row the part is still absent; create it on
                        // demand so the ImagePart has a host to attach to.
                        // CONSISTENCY(grow-on-rawset): mirrors slideMaster/layout
                        // auto-grow above and /notesMaster on-demand creation.
                        var hostSlide = sldParts[nsIdx - 1];
                        imageHost = hostSlide.NotesSlidePart
                            ?? hostSlide.AddNewPart<NotesSlidePart>();
                    }
                    else
                        throw new ArgumentException(
                            "add-part image: parent must be /slide[N], /slideMaster[N], /slideLayout[N], /noteSlide[N], or /notesMaster");
                }

                if (properties == null || !properties.TryGetValue("data", out var imgB64) || string.IsNullOrEmpty(imgB64))
                    throw new ArgumentException("add-part image requires property 'data' (base64 binary)");
                byte[] imgBytes;
                try { imgBytes = Convert.FromBase64String(imgB64); }
                catch (FormatException) { throw new ArgumentException("add-part image: 'data' is not valid base64"); }

                var imgCT = properties.TryGetValue("content-type", out var ict) && !string.IsNullOrEmpty(ict)
                    ? ict
                    : "image/png";
                var imgPartType = imgCT switch
                {
                    "image/png" => ImagePartType.Png,
                    "image/jpeg" => ImagePartType.Jpeg,
                    "image/jpg" => ImagePartType.Jpeg,
                    "image/gif" => ImagePartType.Gif,
                    "image/bmp" => ImagePartType.Bmp,
                    "image/tiff" => ImagePartType.Tiff,
                    "image/x-emf" => ImagePartType.Emf,
                    "image/x-wmf" => ImagePartType.Wmf,
                    _ => ImagePartType.Png,
                };
                string? pinnedImgRid = properties.TryGetValue("rid", out var prid) && !string.IsNullOrEmpty(prid) ? prid : null;
                // rId-collision guard. The blank scaffold ships a fixed set of
                // slideLayout relationships on its master (rId1..rId5). A source
                // master whose own bg-image relationship numerically lands inside
                // that range (e.g. rId5 = master bg image, but the scaffold already
                // wired rId5 = slideLayout5) cannot pin its source rId: AddImagePart
                // with an occupied id throws, and even if it didn't, the master XML
                // raw-set's <p:bg> r:embed="rId5" would resolve to a slideLayout
                // (broken-link background → blank slide). Free the pinned id first
                // by re-homing the colliding scaffold relationship onto a fresh id.
                // For a scaffold-leftover SlideLayoutPart we additionally patch the
                // master's sldLayoutIdLst entry so the re-homed layout stays
                // referenced (mirrors GrowSlideLayoutParts' own collision fallback).
                if (!string.IsNullOrEmpty(pinnedImgRid) && imageHost is OpenXmlPartContainer collHost)
                {
                    var occupant = collHost.Parts.FirstOrDefault(p => p.RelationshipId == pinnedImgRid);
                    if (occupant.OpenXmlPart != null)
                    {
                        // Re-home the colliding scaffold relationship onto a fresh,
                        // collision-free id so the pinned id is free for the image.
                        var newRid = "Rimg" + Guid.NewGuid().ToString("N").Substring(0, 12);
                        collHost.ChangeIdOfPart(occupant.OpenXmlPart, newRid);
                        // If a master's sldLayoutIdLst still points at the old id,
                        // repoint it so the re-homed layout remains declared.
                        if (collHost is SlideMasterPart smHost && smHost.SlideMaster?.SlideLayoutIdList != null)
                        {
                            foreach (var lid in smHost.SlideMaster.SlideLayoutIdList.Elements<SlideLayoutId>())
                            {
                                if (lid.RelationshipId?.Value == pinnedImgRid)
                                {
                                    lid.RelationshipId = newRid;
                                    smHost.SlideMaster.Save();
                                    break;
                                }
                            }
                        }
                    }
                }
                ImagePart imgPart = imageHost switch
                {
                    SlidePart sp           => !string.IsNullOrEmpty(pinnedImgRid) ? sp.AddImagePart(imgPartType, pinnedImgRid) : sp.AddImagePart(imgPartType),
                    SlideMasterPart smp    => !string.IsNullOrEmpty(pinnedImgRid) ? smp.AddImagePart(imgPartType, pinnedImgRid) : smp.AddImagePart(imgPartType),
                    SlideLayoutPart sllp   => !string.IsNullOrEmpty(pinnedImgRid) ? sllp.AddImagePart(imgPartType, pinnedImgRid) : sllp.AddImagePart(imgPartType),
                    NotesMasterPart nmp    => !string.IsNullOrEmpty(pinnedImgRid) ? nmp.AddImagePart(imgPartType, pinnedImgRid) : nmp.AddImagePart(imgPartType),
                    NotesSlidePart nsp     => !string.IsNullOrEmpty(pinnedImgRid) ? nsp.AddImagePart(imgPartType, pinnedImgRid) : nsp.AddImagePart(imgPartType),
                    ThemePart tp           => !string.IsNullOrEmpty(pinnedImgRid) ? tp.AddImagePart(imgPartType, pinnedImgRid) : tp.AddImagePart(imgPartType),
                    _ => throw new ArgumentException($"add-part image: unsupported host part type {imageHost.GetType().Name}"),
                };
                using (var imgStream = new MemoryStream(imgBytes))
                    imgPart.FeedData(imgStream);
                var imgActualRid = imageHost.GetIdOfPart(imgPart);
                return (imgActualRid, parentPartPath);

            case "hyperlink":
            {
                // Re-create an EXTERNAL hyperlink relationship on a host part with
                // a pinned relationship id. Layouts/masters are emitted via raw-set
                // (wholesale XML carrying <a:hlinkClick r:id="rIdN">), but the
                // referenced external relationship is not an embedded part, so the
                // ImagePart carrier path never re-created it — the renumbered
                // rebuilt layout's .rels lost rIdN and PowerPoint rejected the file
                // ("rIdN referenced by hlinkClick does not exist"). Pinning the id
                // here makes the raw-set'd r:id="rIdN" resolve again. The host path
                // is the SOURCE-index /slideLayout[N] (or master/slide); on replay
                // GrowSlideLayoutParts maps it to the renumbered part, so the rel
                // lands on the same part the raw-set replaces.
                if (properties == null
                    || !properties.TryGetValue("target", out var hlTarget) || string.IsNullOrEmpty(hlTarget))
                    throw new ArgumentException("add-part hyperlink requires property 'target' (the external URI)");
                if (!properties.TryGetValue("rid", out var hlRid) || string.IsNullOrEmpty(hlRid))
                    throw new ArgumentException("add-part hyperlink requires property 'rid' (the relationship id to pin)");

                OpenXmlPartContainer hlHost;
                var hlSmMatch = Regex.Match(parentPartPath, @"^/slideMaster\[(\d+)\]$");
                var hlSlMatch = hlSmMatch.Success ? null : Regex.Match(parentPartPath, @"^/slideLayout\[(\d+)\]$");
                var hlSldMatch = (hlSmMatch.Success || (hlSlMatch?.Success ?? false))
                    ? null : Regex.Match(parentPartPath, @"^/slide\[(\d+)\]$");
                if (hlSmMatch.Success)
                {
                    var i = int.Parse(hlSmMatch.Groups[1].Value);
                    var parts = presentationPart.SlideMasterParts.ToList();
                    if (i > parts.Count) { GrowSlideMasterParts(i); parts = presentationPart.SlideMasterParts.ToList(); }
                    if (i < 1 || i > parts.Count) throw new ArgumentException($"slideMaster index {i} out of range");
                    hlHost = parts[i - 1];
                }
                else if (hlSlMatch != null && hlSlMatch.Success)
                {
                    var i = int.Parse(hlSlMatch.Groups[1].Value);
                    var parts = presentationPart.SlideMasterParts.SelectMany(m => m.SlideLayoutParts).ToList();
                    if (i > parts.Count) { GrowSlideLayoutParts(i); parts = presentationPart.SlideMasterParts.SelectMany(m => m.SlideLayoutParts).ToList(); }
                    if (i < 1 || i > parts.Count) throw new ArgumentException($"slideLayout index {i} out of range");
                    hlHost = parts[i - 1];
                }
                else if (hlSldMatch != null && hlSldMatch.Success)
                {
                    var i = int.Parse(hlSldMatch.Groups[1].Value);
                    var parts = GetSlideParts().ToList();
                    if (i < 1 || i > parts.Count) throw new ArgumentException($"slide index {i} out of range");
                    hlHost = parts[i - 1];
                }
                else
                    throw new ArgumentException(
                        "add-part hyperlink: parent must be /slideLayout[N], /slideMaster[N], or /slide[N]");

                // Idempotent: if a relationship with this id already exists, don't
                // re-add (AddHyperlinkRelationship throws on a duplicate id).
                if (hlHost.HyperlinkRelationships.Any(r => r.Id == hlRid))
                    return (hlRid, parentPartPath);
                hlHost.AddHyperlinkRelationship(new Uri(hlTarget, UriKind.RelativeOrAbsolute), isExternal: true, hlRid);
                return (hlRid, parentPartPath);
            }

            case "tags":
            {
                // Re-create a UserDefinedTags part (programmability metadata,
                // <p:tagLst>) on a host part with a pinned relationship id and the
                // source tag XML. Layouts/masters are emitted via raw-set (wholesale
                // XML carrying <p:custDataLst><p:tags r:id="rIdN"/>), but the tags
                // part lives in the host's own .rels enumerated separately, so the
                // ImagePart/hyperlink carriers never re-created it — the renumbered
                // rebuilt layout's r:id="rIdN" dangled and PowerPoint rejected the
                // whole deck (0x80070570 OPC corrupt). Pinning the id here makes the
                // raw-set'd reference resolve. The host path is the SOURCE-index
                // /slideLayout[N] (or master); on replay GrowSlideLayoutParts maps
                // it to the renumbered part, so the tags rel lands on the same part
                // the raw-set replaces. (mirrors the add-part hyperlink pattern)
                if (properties == null
                    || !properties.TryGetValue("data", out var tagXml) || string.IsNullOrEmpty(tagXml))
                    throw new ArgumentException("add-part tags requires property 'data' (the <p:tagLst> XML)");
                if (!properties.TryGetValue("rid", out var tagRid) || string.IsNullOrEmpty(tagRid))
                    throw new ArgumentException("add-part tags requires property 'rid' (the relationship id to pin)");

                OpenXmlPartContainer tagHost;
                var tgSmMatch = Regex.Match(parentPartPath, @"^/slideMaster\[(\d+)\]$");
                var tgSlMatch = tgSmMatch.Success ? null : Regex.Match(parentPartPath, @"^/slideLayout\[(\d+)\]$");
                var tgSldMatch = (tgSmMatch.Success || (tgSlMatch?.Success ?? false))
                    ? null : Regex.Match(parentPartPath, @"^/slide\[(\d+)\]$");
                if (tgSmMatch.Success)
                {
                    var i = int.Parse(tgSmMatch.Groups[1].Value);
                    var parts = presentationPart.SlideMasterParts.ToList();
                    if (i > parts.Count) { GrowSlideMasterParts(i); parts = presentationPart.SlideMasterParts.ToList(); }
                    if (i < 1 || i > parts.Count) throw new ArgumentException($"slideMaster index {i} out of range");
                    tagHost = parts[i - 1];
                }
                else if (tgSlMatch != null && tgSlMatch.Success)
                {
                    var i = int.Parse(tgSlMatch.Groups[1].Value);
                    var parts = presentationPart.SlideMasterParts.SelectMany(m => m.SlideLayoutParts).ToList();
                    if (i > parts.Count) { GrowSlideLayoutParts(i); parts = presentationPart.SlideMasterParts.SelectMany(m => m.SlideLayoutParts).ToList(); }
                    if (i < 1 || i > parts.Count) throw new ArgumentException($"slideLayout index {i} out of range");
                    tagHost = parts[i - 1];
                }
                else if (tgSldMatch != null && tgSldMatch.Success)
                {
                    var i = int.Parse(tgSldMatch.Groups[1].Value);
                    var parts = GetSlideParts().ToList();
                    if (i < 1 || i > parts.Count) throw new ArgumentException($"slide index {i} out of range");
                    tagHost = parts[i - 1];
                }
                else
                    throw new ArgumentException(
                        "add-part tags: parent must be /slideLayout[N], /slideMaster[N], or /slide[N]");

                // Idempotent: if a relationship with this id already exists, don't
                // re-add (AddNewPart with a duplicate id throws).
                if (tagHost.Parts.Any(p => p.RelationshipId == tagRid))
                    return (tagRid, parentPartPath);
                var newTagPart = tagHost switch
                {
                    SlidePart sp        => sp.AddNewPart<UserDefinedTagsPart>(tagRid),
                    SlideLayoutPart slp => slp.AddNewPart<UserDefinedTagsPart>(tagRid),
                    SlideMasterPart smp => smp.AddNewPart<UserDefinedTagsPart>(tagRid),
                    _ => throw new ArgumentException($"add-part tags: unsupported host part type {tagHost.GetType().Name}"),
                };
                using (var sw = new StreamWriter(newTagPart.GetStream(FileMode.Create, FileAccess.Write), new System.Text.UTF8Encoding(false)))
                    sw.Write(tagXml);
                return (tagHost.GetIdOfPart(newTagPart), parentPartPath);
            }

            case "theme":
            {
                // Attach a DISTINCT theme part to a slideMaster / notesMaster with
                // a pinned relationship id and the source theme XML. Multi-master
                // decks give each master its own theme; the blank scaffold +
                // GrowSlideMasterParts share the presentation's primary theme,
                // which loses theme2/theme3 content and (worse) makes masters
                // reference the wrong / a shared theme — PowerPoint refuses such a
                // deck. This re-creates each master's own theme so the package
                // matches the source's 1:1 master:theme topology.
                if (properties == null
                    || !properties.TryGetValue("data", out var themeXml) || string.IsNullOrEmpty(themeXml))
                    throw new ArgumentException("add-part theme requires property 'data' (the theme XML)");
                // The theme is a part relationship of the master, NOT referenced
                // by r:id anywhere in the master XML body, so the relationship id
                // need not be pinned — and pinning it risks colliding with the
                // master's other relationships (images/layouts). Let the SDK
                // assign a fresh id; only honour an explicit rid when it's free.
                string? themeRid = properties.TryGetValue("rid", out var trid) && !string.IsNullOrEmpty(trid) ? trid : null;

                OpenXmlPartContainer themeHost;
                var tmMatch = Regex.Match(parentPartPath, @"^/slideMaster\[(\d+)\]$");
                if (tmMatch.Success)
                {
                    var i = int.Parse(tmMatch.Groups[1].Value);
                    var parts = presentationPart.SlideMasterParts.ToList();
                    if (i > parts.Count) { GrowSlideMasterParts(i); parts = presentationPart.SlideMasterParts.ToList(); }
                    if (i < 1 || i > parts.Count) throw new ArgumentException($"slideMaster index {i} out of range");
                    themeHost = parts[i - 1];
                }
                else if (parentPartPath == "/notesMaster")
                {
                    themeHost = presentationPart.NotesMasterPart
                        ?? throw new ArgumentException("add-part theme: notesMaster part does not exist yet");
                }
                else
                    throw new ArgumentException(
                        "add-part theme: parent must be /slideMaster[N] or /notesMaster");

                // Remove any existing (shared/placeholder) theme part on this host
                // so it gets its OWN distinct part rather than pointing at the
                // primary theme. Then add a fresh ThemePart with the pinned rId.
                var existingTheme = themeHost switch
                {
                    SlideMasterPart smp => (ThemePart?)smp.ThemePart,
                    NotesMasterPart nmp => nmp.ThemePart,
                    _ => null,
                };
                if (existingTheme != null)
                    themeHost.DeletePart(existingTheme);

                // Only pin the rId if it's not already taken by another rel on the
                // host (after deleting the old theme). Otherwise auto-assign.
                bool ridFree = !string.IsNullOrEmpty(themeRid)
                    && !themeHost.Parts.Any(p => p.RelationshipId == themeRid)
                    && !themeHost.ExternalRelationships.Any(r => r.Id == themeRid)
                    && !themeHost.HyperlinkRelationships.Any(r => r.Id == themeRid);
                ThemePart newTheme = themeHost switch
                {
                    SlideMasterPart smp => ridFree ? smp.AddNewPart<ThemePart>(themeRid!) : smp.AddNewPart<ThemePart>(),
                    NotesMasterPart nmp => ridFree ? nmp.AddNewPart<ThemePart>(themeRid!) : nmp.AddNewPart<ThemePart>(),
                    _ => throw new ArgumentException($"add-part theme: unsupported host {themeHost.GetType().Name}"),
                };
                using (var ts = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(themeXml)))
                    newTheme.FeedData(ts);
                var themeActualRid = themeHost.GetIdOfPart(newTheme);
                return (themeActualRid, parentPartPath);
            }

            default:
                throw new ArgumentException(
                    $"Unknown part type: {partType}. Supported: chart, smartart, video, audio, model3d, ole, image, hyperlink, theme");
        }
    }

    // Write verbatim XML into a freshly-created SmartArt diagram sub-part, or
    // seed a minimal typed root when no content was supplied. Writing the raw
    // bytes directly (not via the typed root) preserves the source's exact
    // namespace declarations / extension prefixes that the dump emitter
    // canonicalised, and — crucially — lands the content regardless of the
    // SDK's part-naming base (the parts land under /ppt/graphics/, not the
    // source's /ppt/diagrams/, so a URI-targeted raw-set could not reach them).
    private static void WriteDiagramPartXml(
        OpenXmlPart part, string? xml, Func<OpenXmlElement> seedFactory)
    {
        if (!string.IsNullOrEmpty(xml))
        {
            const string prolog = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
            using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = new StreamWriter(stream, new System.Text.UTF8Encoding(false));
            writer.Write(prolog);
            writer.Write(xml);
            return;
        }
        // Fallback: minimal typed root so direct CLI add-part stays usable.
        var seed = seedFactory();
        using var seedStream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var xw = System.Xml.XmlWriter.Create(seedStream,
            new System.Xml.XmlWriterSettings { OmitXmlDeclaration = false, Encoding = new System.Text.UTF8Encoding(false) });
        seed.WriteTo(xw);
    }

    public List<ValidationError> Validate() => RawXmlHelper.ValidateDocument(_doc, _filePath);

    public void Save()
    {
        // _doc writes through to _backingStream; force the FileStream buffer
        // out to disk so external readers see the latest bytes immediately.
        if (Modified)
        {
            try { ReconcileSlideMasterIds(); }
            catch { /* best-effort id reconcile */ }
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
            try { ReconcileSlideMasterIds(); }
            catch { /* best-effort id reconcile */ }
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

    /// <summary>
    /// Enumerate ImageParts attached to a slideMaster — one entry per
    /// (rId, content-type, base64 bytes). Used by PptxBatchEmitter to emit
    /// `add-part image` rows so a raw-set'd master XML that references
    /// <c>r:embed="rIdN"</c> on <c>&lt;p:pic&gt;</c> blipFills replays
    /// against an ImagePart with the same rId. Returns an empty list when
    /// the master has no embedded images or the index is out of range.
    /// </summary>
    internal readonly record struct MasterImageInfo(string RelId, string ContentType, string Base64Data);

    internal IReadOnlyList<MasterImageInfo> GetMasterImageParts(int masterIdx)
    {
        var result = new List<MasterImageInfo>();
        var pp = _doc.PresentationPart;
        if (pp == null) return result;
        var masters = pp.SlideMasterParts.ToList();
        if (masterIdx < 1 || masterIdx > masters.Count) return result;
        var master = masters[masterIdx - 1];
        foreach (var img in master.ImageParts)
        {
            var rid = master.GetIdOfPart(img);
            using var s = img.GetStream();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            result.Add(new MasterImageInfo(rid, img.ContentType, Convert.ToBase64String(ms.ToArray())));
        }
        return result;
    }

    /// <summary>
    /// A slide master's own ThemePart: its relationship id (as the master XML's
    /// package wires it) and the theme XML. Multi-master decks attach a DISTINCT
    /// theme to each master; the rebuild must re-create each one rather than
    /// sharing the presentation's primary theme (PowerPoint refuses a deck whose
    /// masters share or mis-reference themes). Returns null when the master has no
    /// theme part.
    /// </summary>
    internal (string RelId, string ThemeXml)? GetMasterTheme(int masterIdx)
    {
        var pp = _doc.PresentationPart;
        if (pp == null) return null;
        var masters = pp.SlideMasterParts.ToList();
        if (masterIdx < 1 || masterIdx > masters.Count) return null;
        var master = masters[masterIdx - 1];
        var themePart = master.ThemePart;
        if (themePart?.Theme == null) return null;
        return (master.GetIdOfPart(themePart), themePart.Theme.OuterXml);
    }

    /// <summary>The notes master's own ThemePart (rel id + XML), or null.</summary>
    internal (string RelId, string ThemeXml)? GetNotesMasterTheme()
    {
        var pp = _doc.PresentationPart;
        var nmp = pp?.NotesMasterPart;
        var themePart = nmp?.ThemePart;
        if (nmp == null || themePart?.Theme == null) return null;
        return (nmp.GetIdOfPart(themePart), themePart.Theme.OuterXml);
    }

    /// <summary>
    /// 1-based ordinal of a slide's layout within the same
    /// SlideMasterParts→SlideLayoutParts enumeration that
    /// <see cref="ResolveSlideLayout"/> indexes (its numeric-index match path)
    /// and that the raw-set <c>/slideLayout[N]</c> emission walks. Returns null
    /// when the slide or its layout can't be resolved.
    ///
    /// The batch dump emits this as the slide's `layout=` so replay re-binds to
    /// the EXACT source layout. Emitting the layout NAME is ambiguous: decks
    /// routinely carry several layouts sharing a name (e.g. two "标题幻灯片"
    /// under different masters), and ResolveSlideLayout's name match returns the
    /// first — which can chain the slide to the wrong master and silently drop a
    /// master-level background. The ordinal is unambiguous and stable because
    /// replay reconstructs masters/layouts in this same enumeration order.
    /// </summary>
    internal int? GetSlideLayoutOrdinal(int slideNum)
    {
        var pp = _doc.PresentationPart;
        if (pp == null) return null;
        var slideParts = GetSlideParts().ToList();
        if (slideNum < 1 || slideNum > slideParts.Count) return null;
        var layoutPart = slideParts[slideNum - 1].SlideLayoutPart;
        if (layoutPart == null) return null;
        var allLayouts = pp.SlideMasterParts.SelectMany(m => m.SlideLayoutParts).ToList();
        var idx = allLayouts.IndexOf(layoutPart);
        return idx >= 0 ? idx + 1 : null;
    }

    /// <summary>
    /// Images attached to the presentation's main ThemePart — referenced by an
    /// <c>&lt;a:fmtScheme&gt;&lt;a:fillStyleLst&gt;&lt;a:blipFill&gt;</c> texture
    /// fill via r:embed. The theme XML is raw-set verbatim but its ImageParts are
    /// enumerated separately; without re-emitting them the embed rId dangles and
    /// PowerPoint refuses to open the deck. Same shape as
    /// <see cref="GetMasterImageParts"/>.
    /// </summary>
    internal IReadOnlyList<MasterImageInfo> GetThemeImageParts()
    {
        var result = new List<MasterImageInfo>();
        var theme = _doc.PresentationPart?.ThemePart;
        if (theme == null) return result;
        foreach (var img in theme.ImageParts)
        {
            var rid = theme.GetIdOfPart(img);
            using var s = img.GetStream();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            result.Add(new MasterImageInfo(rid, img.ContentType, Convert.ToBase64String(ms.ToArray())));
        }
        return result;
    }

    /// <summary>Same as <see cref="GetMasterImageParts"/> for slideLayouts.</summary>
    internal IReadOnlyList<MasterImageInfo> GetLayoutImageParts(int layoutIdx)
    {
        var result = new List<MasterImageInfo>();
        var pp = _doc.PresentationPart;
        if (pp == null) return result;
        var layouts = pp.SlideMasterParts.SelectMany(m => m.SlideLayoutParts).ToList();
        if (layoutIdx < 1 || layoutIdx > layouts.Count) return result;
        var layout = layouts[layoutIdx - 1];
        foreach (var img in layout.ImageParts)
        {
            var rid = layout.GetIdOfPart(img);
            using var s = img.GetStream();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            result.Add(new MasterImageInfo(rid, img.ContentType, Convert.ToBase64String(ms.ToArray())));
        }
        return result;
    }

    /// <summary>
    /// External (TargetMode="External") hyperlink relationships on a slideLayout.
    /// The layout XML is replayed via raw-set carrying <c>&lt;a:hlinkClick r:id="rIdN"/&gt;</c>,
    /// but the referenced relationship is external (a URL, not an embedded part),
    /// so the ImagePart carrier never re-creates it — the renumbered rebuilt
    /// layout's .rels lost rIdN and PowerPoint refused the file. Surfaced as
    /// (rId, target) pairs so PptxBatchEmitter can emit an `add-part hyperlink`
    /// row that pins each id before the layout raw-set replace.
    /// </summary>
    internal IReadOnlyList<(string RelId, string Target)> GetLayoutExternalHyperlinks(int layoutIdx)
    {
        var result = new List<(string, string)>();
        var pp = _doc.PresentationPart;
        if (pp == null) return result;
        var layouts = pp.SlideMasterParts.SelectMany(m => m.SlideLayoutParts).ToList();
        if (layoutIdx < 1 || layoutIdx > layouts.Count) return result;
        var layout = layouts[layoutIdx - 1];
        foreach (var rel in layout.HyperlinkRelationships)
        {
            if (rel.IsExternal)
                result.Add((rel.Id, rel.Uri.OriginalString));
        }
        return result;
    }

    /// <summary>
    /// UserDefinedTags parts (programmability metadata, <c>&lt;p:tagLst&gt;</c>)
    /// attached to a slideLayout, surfaced as (rId, verbatim tag XML) pairs.
    /// The layout XML is replayed via raw-set carrying
    /// <c>&lt;p:custDataLst&gt;&lt;p:tags r:id="rIdN"/&gt;</c>, but the tags part
    /// lives in the layout's own .rels (enumerated separately) and was never
    /// re-emitted — the rebuilt layout's <c>r:id="rIdN"</c> then dangled and
    /// PowerPoint refused the whole deck (0x80070570 OPC corrupt). Emitting an
    /// <c>add-part tags</c> row that pins each source rId before the layout
    /// raw-set replace makes the reference resolve. Same shape as
    /// <see cref="GetLayoutImageParts"/>.
    /// </summary>
    internal IReadOnlyList<(string RelId, string TagXml)> GetLayoutTagParts(int layoutIdx)
    {
        var pp = _doc.PresentationPart;
        if (pp == null) return Array.Empty<(string, string)>();
        var layouts = pp.SlideMasterParts.SelectMany(m => m.SlideLayoutParts).ToList();
        if (layoutIdx < 1 || layoutIdx > layouts.Count) return Array.Empty<(string, string)>();
        return ReadTagParts(layouts[layoutIdx - 1]);
    }

    /// <summary>Same as <see cref="GetLayoutTagParts"/> for a slideMaster.</summary>
    internal IReadOnlyList<(string RelId, string TagXml)> GetMasterTagParts(int masterIdx)
    {
        var pp = _doc.PresentationPart;
        if (pp == null) return Array.Empty<(string, string)>();
        var masters = pp.SlideMasterParts.ToList();
        if (masterIdx < 1 || masterIdx > masters.Count) return Array.Empty<(string, string)>();
        return ReadTagParts(masters[masterIdx - 1]);
    }

    private static IReadOnlyList<(string RelId, string TagXml)> ReadTagParts(OpenXmlPartContainer host)
    {
        var result = new List<(string, string)>();
        foreach (var idp in host.Parts)
        {
            if (idp.OpenXmlPart is not UserDefinedTagsPart tagPart) continue;
            using var s = tagPart.GetStream(FileMode.Open, FileAccess.Read);
            using var sr = new StreamReader(s);
            result.Add((idp.RelationshipId, sr.ReadToEnd()));
        }
        return result;
    }

    /// <summary>
    /// Same as <see cref="GetMasterImageParts"/> for a slide's NotesSlidePart.
    /// Used by PptxBatchEmitter.EmitNotes so a notesSlide raw-set replace that
    /// references <c>r:embed="rIdN"</c> on a <c>&lt;p:pic&gt;</c> blipFill
    /// (image pasted into speaker notes) replays against an ImagePart with
    /// the same rId. Without this the post-replay notesSlide carries a
    /// dangling rId and PowerPoint shows a broken picture placeholder.
    /// Returns an empty list when the slide has no notesSlide or no embedded
    /// images.
    /// </summary>
    internal IReadOnlyList<MasterImageInfo> GetNoteSlideImageParts(int slideIdx)
    {
        var result = new List<MasterImageInfo>();
        var pp = _doc.PresentationPart;
        if (pp == null) return result;
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count) return result;
        var notesPart = slideParts[slideIdx - 1].NotesSlidePart;
        if (notesPart == null) return result;
        foreach (var img in notesPart.ImageParts)
        {
            var rid = notesPart.GetIdOfPart(img);
            using var s = img.GetStream();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            result.Add(new MasterImageInfo(rid, img.ContentType, Convert.ToBase64String(ms.ToArray())));
        }
        return result;
    }

    /// <summary>
    /// ImageParts on a slide whose relationship id is one of <paramref name="relIds"/>.
    /// Used by PptxBatchEmitter.EmitRawSlideBgSlice so a slide-level
    /// <c>&lt;p:bg&gt;&lt;p:bgPr&gt;&lt;a:blipFill&gt;&lt;a:blip r:embed="rIdN"&gt;</c>
    /// (background image) raw-set replays against an ImagePart carrying the
    /// SAME source rId. The bg slice is emitted verbatim via raw-set, so its
    /// <c>r:embed="rIdN"</c> only resolves if a matching ImagePart with that
    /// pinned rId exists on the rebuilt slide. We scope to the bg-referenced
    /// rIds only — slide pictures already round-trip through the typed
    /// <c>add picture</c> (fresh rId) path, so re-creating every ImagePart
    /// here would double-create them. Returns an empty list when the slide
    /// is out of range or carries none of the requested rIds.
    /// </summary>
    internal IReadOnlyList<MasterImageInfo> GetSlideImagePartsByRelId(
        int slideIdx, IReadOnlyCollection<string> relIds)
    {
        var result = new List<MasterImageInfo>();
        if (relIds.Count == 0) return result;
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count) return result;
        var slide = slideParts[slideIdx - 1];
        var wanted = new HashSet<string>(relIds, StringComparer.Ordinal);
        foreach (var img in slide.ImageParts)
        {
            var rid = slide.GetIdOfPart(img);
            if (!wanted.Contains(rid)) continue;
            using var s = img.GetStream();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            result.Add(new MasterImageInfo(rid, img.ContentType, Convert.ToBase64String(ms.ToArray())));
        }
        return result;
    }

    /// <summary>
    /// Enumerate <c>r:embed</c> / <c>r:link</c> attribute values referenced by
    /// the source notesSlide XML. Used by PptxBatchEmitter.EmitNotes to detect
    /// rIds that the typed Add/Set surface cannot reproduce (anything not an
    /// ImagePart enumerated by <see cref="GetNoteSlideImageParts"/>). When such
    /// orphan rIds exist, the emitter surfaces a
    /// <c>notes_unresolved_rid</c> warning so callers know the post-replay
    /// notesSlide may have dangling references that PowerPoint will render
    /// as broken placeholders.
    /// </summary>
    internal IReadOnlyList<string> GetNoteSlideExternalRelIds(int slideIdx)
    {
        var result = new List<string>();
        var pp = _doc.PresentationPart;
        if (pp == null) return result;
        var slideParts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > slideParts.Count) return result;
        var notesPart = slideParts[slideIdx - 1].NotesSlidePart;
        if (notesPart == null) return result;
        var xml = notesPart.NotesSlide?.OuterXml;
        if (string.IsNullOrEmpty(xml)) return result;
        var rx = new Regex(@"r:(?:embed|link|id)=""([^""]+)""");
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (Match m in rx.Matches(xml))
        {
            var rid = m.Groups[1].Value;
            if (seen.Add(rid)) result.Add(rid);
        }
        return result;
    }

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
        string QuickStyleXml,
        string? DrawingXml,
        string? DrawingRelId);

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

            // The data part references a 5th part — the DSP cached-drawing
            // part — via <dsp:dataModelExt relId="..."> in the data XML. That
            // relId resolves against the SLIDE part's relationships (the
            // drawing part is a slide-level part of type
            // .../2007/relationships/diagramDrawing, sibling to the
            // data/layout/colors/qs rels — NOT a child of the data part).
            // PowerPoint refuses the file if that relId dangles, so carry the
            // drawing XML + the relId. Leave both null when absent (older /
            // simpler SmartArt without a cached drawing) → keep behavior.
            string? drawingXml = null, drawingRelId = null;
            try
            {
                if (slidePart.GetPartById(dRid) is DiagramDataPart ddp)
                {
                    const string dspNs = "http://schemas.microsoft.com/office/drawing/2008/diagram";
                    var ext = ddp.DataModelRoot?.Descendants().FirstOrDefault(e =>
                        e.LocalName == "dataModelExt" && e.NamespaceUri == dspNs);
                    if (ext != null)
                    {
                        foreach (var a in ext.GetAttributes())
                            if (a.LocalName == "relId") { drawingRelId = a.Value; break; }
                    }
                    if (!string.IsNullOrEmpty(drawingRelId))
                    {
                        try
                        {
                            if (slidePart.GetPartById(drawingRelId) is DiagramPersistLayoutPart drawingPart)
                            {
                                using var s = drawingPart.GetStream(FileMode.Open, FileAccess.Read);
                                using var r = new StreamReader(s);
                                drawingXml = r.ReadToEnd();
                                // Strip XML prolog so emit/replay re-adds it
                                // uniformly (WriteDiagramPartXml re-prepends).
                                int lt = drawingXml.IndexOf('<', StringComparison.Ordinal);
                                int decl = drawingXml.IndexOf("<?xml", StringComparison.Ordinal);
                                if (decl == 0)
                                {
                                    int end = drawingXml.IndexOf("?>", StringComparison.Ordinal);
                                    if (end >= 0) drawingXml = drawingXml.Substring(end + 2).TrimStart();
                                }
                                else if (lt > 0) drawingXml = drawingXml.Substring(lt);
                            }
                        }
                        catch { drawingXml = null; }
                    }
                }
            }
            catch { drawingXml = null; drawingRelId = null; }
            if (drawingXml == null || drawingRelId == null) { drawingXml = null; drawingRelId = null; }

            result.Add(new SmartArtInfo(
                GraphicFrameXml: gf.OuterXml,
                DataRelId: dRid, LayoutRelId: lRid, ColorsRelId: cRid, QuickStyleRelId: qRid,
                DataXml: dXml, LayoutXml: lXml, ColorsXml: cXml, QuickStyleXml: qXml,
                DrawingXml: drawingXml, DrawingRelId: drawingRelId));
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

    /// <summary>
    /// Per-slide OLE-embed info for PptxBatchEmitter Phase 3c-ole passthrough.
    /// Returns one entry per &lt;p:graphicFrame&gt; whose
    /// &lt;a:graphicData uri="…/presentationml/2006/ole"&gt; carries a
    /// &lt;p:oleObj&gt; element. Each entry includes the graphicFrame XML
    /// verbatim plus the source's two rIds (the OLE part and the icon
    /// thumbnail ImagePart) and the underlying binary streams, so the
    /// emitter can issue an `add-part ole` + a `raw-set` append on
    /// /p:sld/p:cSld/p:spTree that round-trips byte-equal.
    /// </summary>
    internal readonly record struct OleInfo(
        string GraphicFrameXml,
        string OleRelId,
        string ThumbnailRelId,
        byte[] OleBytes,
        string OleContentType,
        string OleExtension,
        byte[] ThumbnailBytes,
        string ThumbnailContentType);

    internal IReadOnlyList<OleInfo> GetOlesOnSlide(int slideIdx)
    {
        var result = new List<OleInfo>();
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return result;
        var slidePart = parts[slideIdx - 1];

        // Read raw slide XML — the OleObject child of GraphicData is
        // typed in the SDK (DocumentFormat.OpenXml.Presentation.OleObject)
        // but the whole graphicFrame slice is what we want to raw-set, so
        // a textual XLinq walk is cleaner. Matches the model3d Phase 3c-3d
        // approach: read raw, parse, slice by element identity (no regex)
        // to dodge SDK namespace-prefix drift on round 2.
        string slideXml;
        using (var s = slidePart.GetStream())
        using (var sr = new StreamReader(s))
            slideXml = sr.ReadToEnd();

        System.Xml.Linq.XDocument slideDoc;
        try { slideDoc = System.Xml.Linq.XDocument.Parse(slideXml); }
        catch { return result; }
        System.Xml.Linq.XNamespace pNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
        System.Xml.Linq.XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        System.Xml.Linq.XNamespace aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const string oleUri = "http://schemas.openxmlformats.org/presentationml/2006/ole";

        foreach (var gf in slideDoc.Descendants(pNs + "graphicFrame").ToList())
        {
            var gd = gf.Descendants(aNs + "graphicData").FirstOrDefault();
            var uri = gd?.Attribute("uri")?.Value;
            if (uri == null || !uri.Equals(oleUri, StringComparison.Ordinal)) continue;

            var oleObj = gd!.Element(pNs + "oleObj");
            if (oleObj == null) continue;
            var oleRidAttr = oleObj.Attribute(rNs + "id");
            if (oleRidAttr == null) continue;
            var oleRid = oleRidAttr.Value;

            // Thumbnail icon: <p:pic>/<p:blipFill>/<a:blip r:embed="…"/>
            // inside the <p:oleObj>. PowerPoint always emits one (even when
            // showAsIcon="0", the fallback render still needs an icon).
            var thumbRid = oleObj.Descendants(aNs + "blip")
                .Select(b => b.Attribute(rNs + "embed")?.Value)
                .FirstOrDefault(v => !string.IsNullOrEmpty(v));
            if (string.IsNullOrEmpty(thumbRid)) continue;

            // Re-serialise the graphicFrame subtree as the slice. Funnels
            // both rounds through XLinq for prefix-drift reconciliation.
            var slice = gf.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);

            // Resolve the OLE payload. Could be EmbeddedPackagePart (modern
            // OOXML container) or EmbeddedObjectPart (generic binary / .bin
            // OLE10 stream / legacy .doc/.xls). GetPartById resolves either.
            byte[]? oleBytes = null;
            string? oleCT = null;
            string oleExt = ".bin";
            try
            {
                var part = slidePart.GetPartById(oleRid);
                if (part != null)
                {
                    using var s = part.GetStream();
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    oleBytes = ms.ToArray();
                    oleCT = part.ContentType;
                    var u = part.Uri.OriginalString;
                    var dot = u.LastIndexOf('.');
                    if (dot > 0) oleExt = u[dot..];
                }
            }
            catch { }
            if (oleBytes == null || string.IsNullOrEmpty(oleCT)) continue;

            // Resolve the thumbnail icon ImagePart bytes.
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

            result.Add(new OleInfo(
                GraphicFrameXml: slice,
                OleRelId: oleRid,
                ThumbnailRelId: thumbRid,
                OleBytes: oleBytes,
                OleContentType: oleCT!,
                OleExtension: oleExt,
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

    // Return the verbatim <a:clrChange> outer XML on a picture's <a:blip>,
    // or null when absent. Used by PptxBatchEmitter.EmitPicture to round-
    // trip color-change adjustments (recolor) — there is no typed Set
    // vocabulary for clrChange today, so the emitter copies the element
    // through a raw-set passthrough on the freshly-added picture's blip.
    // Mirrors the GetImageBinary path-resolution preamble verbatim.
    public string? GetPictureBlipClrChangeXml(string picturePath)
    {
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
        Picture? pic = byId
            ? pictures.FirstOrDefault(p =>
                p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value
                    == (uint)idOrIdx)
            : (idOrIdx >= 1 && idOrIdx <= pictures.Count ? pictures[idOrIdx - 1] : null);
        if (pic == null) return null;
        var blip = pic.BlipFill?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Blip>();
        var clrChange = blip?.GetFirstChild<DocumentFormat.OpenXml.Drawing.ColorChange>();
        return clrChange?.OuterXml;
    }

    // R56 bt-6: return outer XML for every <a:blip> child the typed Add/Set
    // surface does NOT cover, so dump→batch can re-inject them via raw-set
    // passthrough. Standard typed-handled children (filtered out): a:alphaModFix
    // (opacity), a:biLevel, a:duotone, a:lum + legacy lumOff/lumMod
    // (brightness/contrast), a:clrChange (already round-tripped via the
    // dedicated GetPictureBlipClrChangeXml + EmitPicture raw-set).
    // Everything else — alphaBiLevel, alphaCeiling, alphaFloor, alphaInv,
    // alphaMod, alphaRepl, blur, clrRepl, fillOverlay, grayscl, hsl, tint,
    // extLst, plus non-schema extension elements like <a:colorMod> seen in
    // the wild — was silently dropped on dump. Mirrors the path-resolution
    // preamble in GetPictureBlipClrChangeXml verbatim.
    public IReadOnlyList<string> GetPictureBlipPassthroughChildrenXml(string picturePath)
    {
        var empty = (IReadOnlyList<string>)Array.Empty<string>();
        var m = Regex.Match(picturePath,
            @"^/slide\[(\d+)\]/(?:.+/)?picture\[(?:@id=)?(\d+)\]$");
        if (!m.Success) return empty;
        var slideIdx = int.Parse(m.Groups[1].Value);
        var idOrIdx = int.Parse(m.Groups[2].Value);
        var byId = picturePath.Contains("@id=", StringComparison.Ordinal);
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return empty;
        var slidePart = parts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (shapeTree == null) return empty;
        var pictures = shapeTree.Descendants<Picture>().ToList();
        Picture? pic = byId
            ? pictures.FirstOrDefault(p =>
                p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value
                    == (uint)idOrIdx)
            : (idOrIdx >= 1 && idOrIdx <= pictures.Count ? pictures[idOrIdx - 1] : null);
        if (pic == null) return empty;
        var blip = pic.BlipFill?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Blip>();
        if (blip == null) return empty;
        var result = new List<string>();
        foreach (var kid in blip.ChildElements)
        {
            // Skip the typed-handled blip children — Add/Set already
            // round-trips these via Format keys.
            if (kid is DocumentFormat.OpenXml.Drawing.AlphaModulationFixed) continue;
            if (kid is DocumentFormat.OpenXml.Drawing.BiLevel) continue;
            if (kid is DocumentFormat.OpenXml.Drawing.Duotone) continue;
            if (kid is DocumentFormat.OpenXml.Drawing.LuminanceEffect) continue;
            if (kid is DocumentFormat.OpenXml.Drawing.ColorChange) continue;
            // Legacy invalid markup written by older builds — NodeBuilder
            // already maps these to brightness/contrast Format keys so the
            // Set-side path re-writes them as <a:lum>. Don't double-emit.
            if (kid.LocalName is "lumOff" or "lumMod"
                && kid.NamespaceUri == "http://schemas.openxmlformats.org/drawingml/2006/main")
                continue;
            result.Add(kid.OuterXml);
        }
        return result;
    }

    // Probe whether a shape's NonVisualDrawingProperties carries a
    // hlinkClick child. Used by PptxBatchEmitter.EmitShape to disambiguate
    // a Format["link"] surfaced by NodeBuilder's single-run shortcut
    // (run-level hlinkClick promoted onto the shape Format bag for Get
    // convenience) from a true shape-level hlinkClick. The dump path
    // should emit shape-level link= on Add only when the link is truly on
    // the cNvPr — otherwise the run-level emit duplicates and AddShape
    // fabricates a shape-level hyperlink that the source never had.
    internal bool ShapeHasCNvPrHyperlink(string shapePath)
        => GetShapeCNvPrHyperlinkInfo(shapePath).HasShapeLink;

    // CONSISTENCY(shape-link-source-readback): when a shape carries BOTH a
    // cNvPr.hlinkClick AND a first-run rPr.hlinkClick, NodeBuilder promotes
    // the RUN url onto Format["link"] (first-run wins; line ~960). The dump
    // emitter needs the actual shape-level url so it can emit `add shape
    // link=<shape-url>` instead of inheriting the run url. Return the
    // shape-level url + tooltip (or null/empty) alongside the boolean.
    internal (bool HasShapeLink, string? Url, string? Tooltip) GetShapeCNvPrHyperlinkInfo(string shapePath)
    {
        // Accept positional /shape[N] and @id= forms. NodeBuilder emits the
        // @id= form for shapes with a known cNvPr.Id (the typical case);
        // the typed dump walk passes shapeNode.Path verbatim.
        var m = Regex.Match(shapePath,
            @"^/slide\[(\d+)\]((?:/group\[\d+\])*)/(?:shape|textbox|title|equation|placeholder)\[(@id=)?(\d+)\]$");
        if (!m.Success) return (false, null, null);
        var slideIdx = int.Parse(m.Groups[1].Value);
        var grpChain = m.Groups[2].Value;
        var byId = m.Groups[3].Value.Length > 0;
        var shapeIdx = int.Parse(m.Groups[4].Value);
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return (false, null, null);
        var slidePart = parts[slideIdx - 1];
        OpenXmlCompositeElement? scope = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (scope == null) return (false, null, null);
        foreach (Match gm in Regex.Matches(grpChain, @"/group\[(\d+)\]"))
        {
            var gIdx = int.Parse(gm.Groups[1].Value);
            var groupsHere = scope.Elements<GroupShape>().ToList();
            if (gIdx < 1 || gIdx > groupsHere.Count) return (false, null, null);
            scope = groupsHere[gIdx - 1];
        }
        Shape? shape;
        if (byId)
        {
            shape = scope.Elements<Shape>().FirstOrDefault(
                s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value == (uint)shapeIdx);
            if (shape == null) return (false, null, null);
        }
        else
        {
            var shapes = scope.Elements<Shape>().ToList();
            if (shapeIdx < 1 || shapeIdx > shapes.Count) return (false, null, null);
            shape = shapes[shapeIdx - 1];
        }
        var nvDp = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
        var hlClick = nvDp?.GetFirstChild<DocumentFormat.OpenXml.Drawing.HyperlinkOnClick>();
        if (hlClick == null) return (false, null, null);
        var url = ReadHyperlinkOnClickUrl(hlClick, slidePart);
        var tip = hlClick.Tooltip?.Value;
        return (true, url, tip);
    }

    // Return the verbatim <p:style> outer XML on a shape (the lnRef / fillRef
    // / effectRef / fontRef theme-reference block) plus the shape's ordinal
    // among <p:sp> siblings in its parent shapeTree/group, or null when
    // either the shape has no <p:style> or the path doesn't resolve. Used by
    // PptxBatchEmitter.EmitShape to round-trip the style reference block
    // through a raw-set passthrough — there is no typed Add/Set vocabulary
    // for these theme-style refs today and the source block was silently
    // dropped on dump->replay before this hook.
    // CONSISTENCY: mirrors GetShapeCNvPrHyperlinkInfo's path-resolution
    // preamble (group-chain + @id= / positional shape index).
    internal (string Xml, int SpOrdinal)? GetShapeStyleXmlWithOrdinal(string shapePath)
    {
        var m = Regex.Match(shapePath,
            @"^/slide\[(\d+)\]((?:/group\[\d+\])*)/(?:shape|textbox|title|equation|placeholder)\[(@id=)?(\d+)\]$");
        if (!m.Success) return null;
        var slideIdx = int.Parse(m.Groups[1].Value);
        var grpChain = m.Groups[2].Value;
        var byId = m.Groups[3].Value.Length > 0;
        var shapeIdx = int.Parse(m.Groups[4].Value);
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return null;
        var slidePart = parts[slideIdx - 1];
        OpenXmlCompositeElement? scope = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (scope == null) return null;
        foreach (Match gm in Regex.Matches(grpChain, @"/group\[(\d+)\]"))
        {
            var gIdx = int.Parse(gm.Groups[1].Value);
            var groupsHere = scope.Elements<GroupShape>().ToList();
            if (gIdx < 1 || gIdx > groupsHere.Count) return null;
            scope = groupsHere[gIdx - 1];
        }
        var shapes = scope.Elements<Shape>().ToList();
        Shape? shape;
        int ordinal;
        if (byId)
        {
            shape = shapes.FirstOrDefault(
                s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value == (uint)shapeIdx);
            if (shape == null) return null;
            ordinal = shapes.IndexOf(shape) + 1;
        }
        else
        {
            if (shapeIdx < 1 || shapeIdx > shapes.Count) return null;
            shape = shapes[shapeIdx - 1];
            ordinal = shapeIdx;
        }
        var styleEl = shape.GetFirstChild<ShapeStyle>();
        if (styleEl == null) return null;
        return (styleEl.OuterXml, ordinal);
    }

    // Resolve a shape path's blipFill image bytes (image fill on a non-Picture
    // <p:sp>). Mirrors GetImageBinary but walks <p:sp> (Shape) instead of
    // <p:pic> (Picture). Accepts /slide[N]/shape[K] and /slide[N]/shape[@id=ID]
    // path forms (group-nested shape paths are NOT supported in this initial
    // pass — covers the common slide-level blipFill case used by examples).
    public (byte[] Bytes, string ContentType)? GetShapeImageFillBinary(string shapePath)
    {
        var m = Regex.Match(shapePath,
            @"^/slide\[(\d+)\]/shape\[(@id=)?(\d+)\]$");
        if (!m.Success) return null;
        var slideIdx = int.Parse(m.Groups[1].Value);
        var byId = m.Groups[2].Success;
        var idOrIdx = int.Parse(m.Groups[3].Value);
        var parts = GetSlideParts().ToList();
        if (slideIdx < 1 || slideIdx > parts.Count) return null;
        var slidePart = parts[slideIdx - 1];
        var shapeTree = GetSlide(slidePart).CommonSlideData?.ShapeTree;
        if (shapeTree == null) return null;
        var shapes = shapeTree.Elements<Shape>().ToList();
        Shape? shape = null;
        if (byId)
        {
            shape = shapes.FirstOrDefault(s =>
            {
                var sid = s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
                return sid.HasValue && sid.Value == (uint)idOrIdx;
            });
        }
        else
        {
            if (idOrIdx >= 1 && idOrIdx <= shapes.Count) shape = shapes[idOrIdx - 1];
        }
        if (shape == null) return null;
        var blipFill = shape.ShapeProperties?.GetFirstChild<DocumentFormat.OpenXml.Drawing.BlipFill>();
        var embedId = blipFill?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Blip>()?.Embed?.Value;
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
