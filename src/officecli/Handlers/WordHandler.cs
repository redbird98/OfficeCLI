// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class WordHandler : IDocumentHandler
{
    private readonly WordprocessingDocument _doc;
    private readonly string _filePath;
    private HashSet<string> _usedParaIds = new(StringComparer.OrdinalIgnoreCase);
    private int _nextParaId = 0x100000;
    public int LastFindMatchCount { get; internal set; }
    // Number of elements a no-slash selector Set matched and mutated (Sheet1!row[...]).
    // Read by the CLI/resident to echo the multi-element change count.
    public int LastSelectorSetCount { get; internal set; }

    // Backing FileStream — mirrors the PPT pattern. Opening via a shared
    // FileStream (FileShare.Read in editable mode) lets external readers
    // observe the file while the handler is alive, which is required for
    // mid-session `save` snapshots to be useful to third-party consumers
    // (issue #114). The package writes through the stream; the on-disk
    // bytes lag _doc until _doc.Save() runs.
    private FileStream? _backingStream;

    /// <summary>
    /// Props that the most recent Add() call could not consume. Surfaced to
    /// the CLI layer so silent-drops on the curated surface (e.g.
    /// `add /styles --prop font.eastAsia=...`) become visible warnings
    /// instead of "Added" lies. Reset at the start of each Add.
    /// </summary>
    public List<string> LastAddUnsupportedProps { get; internal set; } = new();

    /// <summary>
    /// Advisory warnings from the most recent Add() call (e.g. unknown
    /// style id referenced but stored as-is). Surfaced to the CLI layer
    /// as stderr WARNING lines, non-fatal. Reset at the start of each Add.
    /// </summary>
    public List<string> LastAddWarnings { get; internal set; } = new();

    /// <summary>
    /// Advisory warnings from the most recent Set() call (e.g. unknown
    /// style id referenced as-is). Surfaced to the CLI layer as stderr
    /// WARNING lines, non-fatal — kept out of the unsupported-prop list so
    /// the write still counts as applied. Reset at the start of each Set.
    /// </summary>
    public List<string> LastSetWarnings { get; internal set; } = new();

    /// <summary>
    /// Set true by Add/Set/Remove/RawSet, consumed by Save/Dispose to decide
    /// whether to stamp <c>docProps/custom.xml</c> with an OfficeCLI audit
    /// trail. Pure Get/Query sessions leave this false and never touch the
    /// file's metadata.
    /// </summary>
    internal bool Modified { get; set; }

    /// <summary>
    /// When true, per-mutation <c>Document.Save()</c> calls are skipped — the
    /// in-memory DOM stays authoritative and is serialized once at Dispose
    /// (AutoSave) / explicit flush. Set by the batch driver around a replay so
    /// N mutations cost O(N) instead of O(N²) (each Save re-serializes the whole
    /// growing part). Single-command paths leave this false and save eagerly.
    /// </summary>
    public bool DeferSave { get; set; }

    /// <summary>
    /// Serialize the main document part to the backing store, unless
    /// <see cref="DeferSave"/> is set (batch replay defers to one save at the
    /// end). Every mutation path (Add/Set/Remove/Move/Swap) routes its
    /// per-operation <c>Document.Save()</c> through here so the eager-vs-deferred
    /// decision lives in ONE place — N batch mutations cost one serialize, not N.
    /// </summary>
    private void SaveDoc()
    {
        if (!DeferSave) _doc.MainDocumentPart?.Document?.Save();
    }

    /// <summary>
    /// Enumerate every <see cref="OpenXmlPart"/> in the package (transitive
    /// walk via the SDK's own <c>GetAllParts</c> extension) yielding each
    /// part's zip-URI (<c>OpenXmlPart.Uri.OriginalString</c>). Used by the
    /// batch emitter's auxiliary-parts scan to surface warnings for parts the
    /// dump surface does not round-trip (customXml, glossary, webSettings,
    /// fontTable, embedded fonts, modern-comment metadata, user docProps).
    /// </summary>
    internal IEnumerable<string> EnumeratePartUris()
    {
        // Two complementary sources:
        //   1. SDK part graph (GetAllParts) — only reachable via the
        //      relationship graph; misses orphan parts but matches every
        //      real Word file produced by Office / python-docx / WPS.
        //   2. Raw zip entries — catches orphan parts (fuzzer-injected
        //      content, partial-rel files dropped by failed save), parts
        //      whose `_rels/.rels` reference disappeared, and anything the
        //      SDK refuses to surface as a typed part.
        // Union them so the aux-parts scan can warn on both cases. Skips
        // duplicates inside the consumer.
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var part in _doc.GetAllParts())
        {
            var u = part.Uri.OriginalString;
            if (seen.Add(u)) yield return u;
        }
        // Raw zip walk — opens the file independently. Use FileShare.ReadWrite
        // because the SDK's _backingStream already holds the file handle in
        // editable mode (FileShare.Read on our side). Falling back to the SDK
        // package via reflection is also possible (RawXmlHelper does this for
        // TryReadByZipUri), but the raw zip is simpler and sees the full
        // on-disk truth.
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
    /// props vanish). Returns an empty list if no custom part exists.
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

    public WordHandler(string filePath, bool editable)
    {
        _filePath = filePath;
        var share = editable ? FileShare.Read : FileShare.ReadWrite;
        var access = editable ? FileAccess.ReadWrite : FileAccess.Read;
        _backingStream = new FileStream(filePath, FileMode.Open, access, share);
        try
        {
            _doc = WordprocessingDocument.Open(_backingStream, editable);
            WordStrictAttributeSanitizer.Sanitize(_doc);
            if (editable)
            {
                EnsureAllParaIds();
                EnsureDocPropIds();
                EnsureSdtIds();
            }
        }
        catch
        {
            // A failed open must not leak the backing FileStream — the
            // factory's repair-and-retry paths (FixXmlEncoding /
            // StripDanglingPackageRels) reopen the file for in-place fixes
            // and would hit "file is being used by another process".
            _doc?.Dispose();
            _backingStream.Dispose();
            throw;
        }
    }

    /// <summary>
    /// Resolve a picture-run path to the embedded image's bytes and content
    /// type. Returns null if the path doesn't point at a Drawing-bearing
    /// run, or the run carries no resolvable rId/embed target.
    ///
    /// <para>
    /// Used by <c>WordBatchEmitter</c> to round-trip pictures through batch
    /// dumps — the bytes are encoded as a data URI in the emitted
    /// `src=` prop and re-imported via <c>ImageSource.Resolve</c> on replay.
    /// </para>
    /// </summary>
    /// <summary>
    /// Returns true if the run at <paramref name="runPath"/> wraps a chart
    /// (c:chart inside a Drawing's graphicData). WordBatchEmitter uses this to
    /// distinguish chart-bearing runs from picture/OLE/background runs that
    /// also surface as type="picture" in Get — without this, an unsupported
    /// drawing's failed image extraction would consume the next chart spec
    /// and render at the wrong paragraph.
    /// </summary>
    public bool IsChartRun(string runPath)
    {
        var segments = ParsePath(runPath);
        var element = NavigateToElement(segments);
        if (element is not Run run) return false;
        var drawing = run.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Drawing>();
        if (drawing == null) return false;
        return drawing
            .Descendants<DocumentFormat.OpenXml.Drawing.Charts.ChartReference>()
            .Any();
    }

    /// <summary>
    /// Outer XML of the element at <paramref name="path"/>. WordBatchEmitter
    /// uses this as a raw-XML fallback for content that has no typed Add
    /// path — wps:wsp background shapes being the motivating case. Returns
    /// null if the path doesn't resolve.
    /// </summary>
    public string? GetElementXml(string path)
    {
        try
        {
            var segments = ParsePath(path);
            var element = NavigateToElement(segments);
            return element?.OuterXml;
        }
        catch
        {
            return null;
        }
    }

    public (byte[] Bytes, string ContentType)? GetImageBinary(string runPath)
    {
        // Parse + navigate via the same machinery Get/Set use so paraId
        // anchors and positional indices behave consistently.
        var segments = ParsePath(runPath);
        var element = NavigateToElement(segments);
        if (element is not Run run) return null;

        var drawing = run.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Drawing>();
        if (drawing == null) return null;

        var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
        var embedId = blip?.Embed?.Value;
        if (string.IsNullOrEmpty(embedId)) return null;

        // CONSISTENCY(host-part-rel): mirror the AddPicture host-part lookup
        // — image part may be attached to a header/footer part rather than
        // the main document part, depending on where the run lives.
        var hostPart = ResolveImageHostPart(run);
        try
        {
            var part = hostPart.GetPartById(embedId);
            using var src = part.GetStream();
            using var ms = new MemoryStream();
            src.CopyTo(ms);
            return (ms.ToArray(), part.ContentType);
        }
        catch
        {
            return null;
        }
    }

    // dump→batch round-trip carrier for a Word OLE object run. The dump emits
    // these so AddOle can rebuild the embedded part + icon + VML frame without
    // an external src file. EmbeddedBytes is the payload exactly as stored
    // (raw package, or CFB-wrapped Ole10Native) — fed back verbatim.
    internal sealed record OleEmitData(
        byte[] EmbeddedBytes, string OleKind, string EmbeddedContentType, string EmbeddedExt,
        byte[]? IconBytes, string? IconContentType,
        string? ProgId, string? Display, string? Width, string? Height, string? Name);

    private const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>
    /// dump→batch: extract everything needed to faithfully re-add an OLE object
    /// run — the embedded payload bytes (raw) plus part kind / content type /
    /// target extension, the icon image bytes, and the VML display metadata
    /// (progId, drawAspect→display, width/height, friendly name). Returns null
    /// when the run is not an OLE object or its parts can't be resolved, so the
    /// emitter falls back to the warn-and-drop path.
    /// </summary>
    internal OleEmitData? GetOleEmitData(string runPath)
    {
        OpenXmlElement? element;
        try { element = NavigateToElement(ParsePath(runPath)); }
        catch { return null; }
        if (element is not Run run) return null;
        var oleObj = run.GetFirstChild<EmbeddedObject>();
        if (oleObj == null) return null;

        var oleElement = oleObj.Descendants().FirstOrDefault(e => e.LocalName == "OLEObject");
        if (oleElement == null) return null;
        string? embedRelId = null, progId = null, drawAspect = null;
        foreach (var a in oleElement.GetAttributes())
        {
            if (a.LocalName == "ProgID") progId = a.Value;
            else if (a.LocalName == "DrawAspect") drawAspect = a.Value;
            else if (a.LocalName == "id" && a.NamespaceUri == RelNs) embedRelId = a.Value;
        }
        if (string.IsNullOrEmpty(embedRelId)) return null;

        var hostPart = ResolveImageHostPart(run);
        OpenXmlPart embedPart;
        try { embedPart = hostPart.GetPartById(embedRelId); }
        catch { return null; }
        byte[] embeddedBytes;
        try
        {
            using var s = embedPart.GetStream();
            using var ms = new MemoryStream();
            s.CopyTo(ms);
            embeddedBytes = ms.ToArray();
        }
        catch { return null; }
        var oleKind = embedPart is EmbeddedPackagePart ? "package" : "object";
        var embedExt = System.IO.Path.GetExtension(embedPart.Uri.ToString()).TrimStart('.');

        byte[]? iconBytes = null;
        string? iconCt = null;
        var imageData = oleObj.Descendants().FirstOrDefault(e => e.LocalName == "imagedata");
        if (imageData != null)
        {
            var iconRelId = imageData.GetAttributes()
                .FirstOrDefault(a => a.LocalName == "id" && a.NamespaceUri == RelNs).Value;
            if (!string.IsNullOrEmpty(iconRelId))
            {
                try
                {
                    var ip = hostPart.GetPartById(iconRelId);
                    using var s = ip.GetStream();
                    using var ms = new MemoryStream();
                    s.CopyTo(ms);
                    iconBytes = ms.ToArray();
                    iconCt = ip.ContentType;
                }
                catch { /* icon is best-effort; AddOle falls back to placeholder */ }
            }
        }

        string? display = string.IsNullOrEmpty(drawAspect)
            ? null
            : (drawAspect.Equals("Content", StringComparison.OrdinalIgnoreCase) ? "content" : "icon");

        string? width = null, height = null, name = null;
        var shape = oleObj.Descendants().FirstOrDefault(e => e.LocalName == "shape");
        if (shape != null)
        {
            var alt = shape.GetAttributes().FirstOrDefault(a => a.LocalName == "alt").Value;
            if (!string.IsNullOrEmpty(alt)) name = alt;
            var style = shape.GetAttributes().FirstOrDefault(a => a.LocalName == "style").Value;
            if (!string.IsNullOrEmpty(style))
            {
                foreach (var seg in style.Split(';', StringSplitOptions.RemoveEmptyEntries))
                {
                    var kv = seg.Split(':', 2);
                    if (kv.Length != 2) continue;
                    var k = kv[0].Trim().ToLowerInvariant();
                    // Keep the raw VML literal (e.g. "77pt") — AddOle's ParseEmu
                    // accepts pt/cm/in, so the frame dimensions round-trip exactly.
                    if (k == "width") width = kv[1].Trim();
                    else if (k == "height") height = kv[1].Trim();
                }
            }
        }

        return new OleEmitData(embeddedBytes, oleKind, embedPart.ContentType, embedExt,
            iconBytes, iconCt, progId, display, width, height, name);
    }

    // dump→batch round-trip carrier for an ActiveX form-control run — a
    // <w:object> hosting <w:control r:id> plus a VML preview <v:imagedata r:id>
    // instead of an o:OLEObject. The dump ships the verbatim run XML and every
    // referenced package part's bytes (preview image, activeX persistence XML,
    // and that part's nested binary blob); AddActiveX rebuilds the parts and
    // rewrites the run's r:id refs, so the control needs no external files.
    internal sealed record ActiveXPartData(
        string RelId, byte[] Bytes, string ContentType, List<ActiveXPartData> Children);
    // Externals: relationships with TargetMode=External (hyperlinks inside a
    // VML textbox, linked content) — no part bytes, just type + target URI.
    internal sealed record ActiveXExternalData(string RelId, string Type, string Target);
    internal sealed record ActiveXEmitData(
        string RunXml, List<ActiveXPartData> Parts, List<ActiveXExternalData> Externals);

    /// <summary>
    /// dump→batch: extract the verbatim run XML and all referenced parts for an
    /// ActiveX control run. Returns null when the run hosts no &lt;w:control&gt;
    /// or any referenced part can't be resolved, so the emitter falls back to
    /// the warn-and-drop path.
    /// </summary>
    internal ActiveXEmitData? GetActiveXEmitData(string runPath)
    {
        OpenXmlElement? element;
        try { element = NavigateToElement(ParsePath(runPath)); }
        catch { return null; }
        if (element is not Run run) return null;
        var obj = run.GetFirstChild<EmbeddedObject>();
        if (obj == null) return null;
        if (!obj.Descendants().Any(e => e.LocalName == "control")) return null;
        return CollectInlinedPartsEmitData(run, obj);
    }

    /// <summary>
    /// dump→batch: extract the verbatim run XML and all referenced parts for a
    /// SmartArt diagram run. The &lt;dgm:relIds&gt; element references the
    /// data / layout / quickStyle / colors parts via r:dm / r:lo / r:qs / r:cs,
    /// and the data part nests the rendered-drawing part. Same carrier shape as
    /// the ActiveX path; returns null when the run hosts no diagram or any
    /// referenced part can't be resolved.
    /// </summary>
    internal ActiveXEmitData? GetDiagramEmitData(string runPath)
    {
        OpenXmlElement? element;
        try { element = NavigateToElement(ParsePath(runPath)); }
        catch { return null; }
        if (element is not Run run) return null;
        var drawing = run.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Drawing>();
        if (drawing == null) return null;
        if (!drawing.Descendants().Any(e => e.LocalName == "relIds")) return null;
        return CollectInlinedPartsEmitData(run, drawing);
    }

    /// <summary>
    /// dump→batch: extract the verbatim run XML and all referenced parts /
    /// external relationships for a legacy VML shape run (&lt;w:pict&gt;).
    /// Covers textbox content carrying hyperlinks (external rels) and
    /// v:imagedata image parts. Returns null when the run hosts no pict or a
    /// reference can't be resolved.
    /// </summary>
    internal ActiveXEmitData? GetVmlShapeEmitData(string runPath)
    {
        OpenXmlElement? element;
        try { element = NavigateToElement(ParsePath(runPath)); }
        catch { return null; }
        if (element is not Run run) return null;
        // LibreOffice/Word export VML textboxes wrapped in mc:AlternateContent
        // (Choice = modern wps drawing, Fallback = w:pict) — probe the whole
        // run subtree, and collect rel refs from the whole run so the Choice
        // branch's references resolve on replay too.
        if (!run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Picture>().Any())
            return null;
        return CollectInlinedPartsEmitData(run, run);
    }

    /// <summary>
    /// dump→batch: extract the verbatim &lt;w:sdt&gt; XML and every referenced
    /// part / external relationship for a rich BLOCK content control (a cover
    /// page with anchored textboxes and a logo image). Same carrier shape as
    /// the ActiveX / diagram / vmlshape paths; returns null when the SDT
    /// references nothing (the plain raw-set passthrough already handles that)
    /// or a reference can't be resolved.
    /// </summary>
    /// <summary>
    /// dump→batch: extract the verbatim run XML and referenced parts for a
    /// modern DrawingML shape run (&lt;w:drawing&gt; hosting wps:wsp with an
    /// image blipFill). Same carrier shape as the vmlshape path.
    /// </summary>
    internal ActiveXEmitData? GetDrawingShapeEmitData(string runPath)
    {
        OpenXmlElement? element;
        try { element = NavigateToElement(ParsePath(runPath)); }
        catch { return null; }
        if (element is not Run run) return null;
        if (!run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any())
            return null;
        return CollectInlinedPartsEmitData(run, run);
    }

    internal ActiveXEmitData? GetSdtEmitData(string sdtPath)
    {
        OpenXmlElement? element;
        try { element = NavigateToElement(ParsePath(sdtPath)); }
        catch { return null; }
        if (element is not SdtBlock sdt) return null;
        return CollectInlinedPartsEmitData(sdt, sdt);
    }

    // Shared collector for the `add activex` / `add diagram` carriers: every
    // relationship-namespace attribute in the subtree (r:id, r:dm, r:lo, …)
    // names a part on the run's host part; ship each part's bytes plus its
    // direct children (activeX binary blob, diagram rendered drawing).
    private ActiveXEmitData? CollectInlinedPartsEmitData(OpenXmlElement run, OpenXmlElement subtree)
    {
        var hostPart = ResolveImageHostPart(run);
        var relIds = new List<string>();
        foreach (var el in subtree.Descendants())
        {
            foreach (var a in el.GetAttributes())
            {
                if (a.NamespaceUri == RelNs
                    && !string.IsNullOrEmpty(a.Value) && !relIds.Contains(a.Value!))
                    relIds.Add(a.Value!);
            }
        }
        if (relIds.Count == 0) return null;

        var parts = new List<ActiveXPartData>();
        var externals = new List<ActiveXExternalData>();
        for (int idx = 0; idx < relIds.Count; idx++)
        {
            var relId = relIds[idx];
            OpenXmlPart part;
            byte[] bytes;
            try
            {
                part = hostPart.GetPartById(relId);
                bytes = ReadPartBytes(part);
            }
            catch
            {
                // Not a part — maybe an EXTERNAL relationship (a hyperlink
                // inside a VML textbox, linked content). Those carry no bytes;
                // ship type + target so the apply side recreates the rel.
                var hyper = hostPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                if (hyper != null)
                {
                    externals.Add(new ActiveXExternalData(
                        relId, hyper.RelationshipType, hyper.Uri.OriginalString));
                    continue;
                }
                var ext = hostPart.ExternalRelationships.FirstOrDefault(x => x.Id == relId);
                if (ext != null)
                {
                    externals.Add(new ActiveXExternalData(
                        relId, ext.RelationshipType, ext.Uri.OriginalString));
                    continue;
                }
                return null;
            }
            var children = new List<ActiveXPartData>();
            foreach (var child in part.Parts)
            {
                byte[] cb;
                try { cb = ReadPartBytes(child.OpenXmlPart); }
                catch { return null; }
                children.Add(new ActiveXPartData(
                    child.RelationshipId, cb, child.OpenXmlPart.ContentType, new List<ActiveXPartData>()));
            }
            parts.Add(new ActiveXPartData(relId, bytes, part.ContentType, children));

            // A collected XML part's CONTENT can reference further host-part
            // relationships via bare relId="…" attributes — SmartArt's data
            // part points at its rendered-drawing part through
            // <dsp:dataModelExt relId="rId12"> resolved against the MAIN part,
            // not the data part. Chase those so the rendered drawing ships too.
            if (part.ContentType.EndsWith("+xml", StringComparison.OrdinalIgnoreCase))
            {
                var xmlText = System.Text.Encoding.UTF8.GetString(bytes);
                foreach (System.Text.RegularExpressions.Match m in
                         System.Text.RegularExpressions.Regex.Matches(xmlText, "relId=\"([^\"]+)\""))
                {
                    var extra = m.Groups[1].Value;
                    if (!string.IsNullOrEmpty(extra) && !relIds.Contains(extra))
                    {
                        try { hostPart.GetPartById(extra); }
                        catch { continue; }
                        relIds.Add(extra);
                    }
                }
            }
        }
        // A run that references ONLY external rels (a VML textbox whose sole
        // refs are hyperlinks) is still a valid carrier.
        if (parts.Count == 0 && externals.Count == 0) return null;
        return new ActiveXEmitData(run.OuterXml, parts, externals);
    }

    private static byte[] ReadPartBytes(OpenXmlPart part)
    {
        using var s = part.GetStream();
        using var ms = new MemoryStream();
        s.CopyTo(ms);
        return ms.ToArray();
    }

    private OpenXmlPart ResolveImageHostPart(OpenXmlElement run)
    {
        var headerAncestor = run.Ancestors<Header>().FirstOrDefault();
        if (headerAncestor != null)
        {
            var hp = _doc.MainDocumentPart!.HeaderParts
                .FirstOrDefault(p => ReferenceEquals(p.Header, headerAncestor));
            if (hp != null) return hp;
        }
        var footerAncestor = run.Ancestors<Footer>().FirstOrDefault();
        if (footerAncestor != null)
        {
            var fp = _doc.MainDocumentPart!.FooterParts
                .FirstOrDefault(p => ReferenceEquals(p.Footer, footerAncestor));
            if (fp != null) return fp;
        }
        return _doc.MainDocumentPart!;
    }

    // BUG-DUMP-R45-1: enumerate the embedded font binaries (word/fonts/*.odttf)
    // attached to the FontTablePart, keyed by the relationship id the
    // fontTable.xml <w:embed*> elements reference. The dump emits each as a
    // companion `raw-set /fontTable --action embed-binary` carrying the bytes as
    // a data: URI; the apply side rebuilds the FontPart with the SAME r:id so the
    // round-tripped embed refs (kept verbatim in the fontTable XML) resolve.
    // The .odttf bytes are obfuscated with the w:fontKey already in the XML —
    // they round-trip raw, no de/re-obfuscation needed.
    // customXml data stores (item.xml + itemProps.xml pairs). SDT content
    // controls bind to a store through the itemProps datastore-item GUID, not
    // a relationship id, so recreating the parts with their bytes verbatim is
    // sufficient for the bindings to resolve. The dump ships each pair as
    // `raw-set /customXml --action embed-binary` (item) plus
    // `raw-set /customXml/<itemRelId> --action embed-binary` (its props child).
    internal List<(string RelId, byte[] Bytes, string ContentType,
                   (string RelId, byte[] Bytes, string ContentType)? Props)> GetCustomXmlEmitData()
    {
        var result = new List<(string, byte[], string, (string, byte[], string)?)>();
        var mainPart = _doc.MainDocumentPart;
        if (mainPart == null) return result;
        foreach (var part in mainPart.CustomXmlParts)
        {
            byte[] bytes;
            try { bytes = ReadPartBytes(part); }
            catch { continue; }
            (string, byte[], string)? props = null;
            var pp = part.CustomXmlPropertiesPart;
            if (pp != null)
            {
                try { props = (part.GetIdOfPart(pp), ReadPartBytes(pp), pp.ContentType); }
                catch { /* props are optional; item still round-trips */ }
            }
            result.Add((mainPart.GetIdOfPart(part), bytes, part.ContentType, props));
        }
        return result;
    }

    internal List<(string RelId, byte[] Bytes, string ContentType)> GetEmbeddedFontParts()
    {
        var result = new List<(string, byte[], string)>();
        var ftp = _doc.MainDocumentPart?.FontTablePart;
        if (ftp == null) return result;
        foreach (var fp in ftp.FontParts)
        {
            string relId;
            try { relId = ftp.GetIdOfPart(fp); }
            catch { continue; }
            if (string.IsNullOrEmpty(relId)) continue;
            try
            {
                using var s = fp.GetStream();
                using var ms = new MemoryStream();
                s.CopyTo(ms);
                result.Add((relId, ms.ToArray(), fp.ContentType));
            }
            catch { /* unreadable font part — skip; emitter falls back to strip */ }
        }
        return result;
    }

    // BUG-DUMP-R45-2: enumerate the image binaries (word/media/*) attached to
    // the NumberingDefinitionsPart, keyed by the relationship id the
    // numbering.xml <w:numPicBullet> VML <v:imagedata r:id> references. Mirrors
    // GetEmbeddedFontParts — the dump emits each as a companion
    // `raw-set /numbering --action embed-binary` so the apply side rebuilds the
    // ImagePart with the SAME r:id and the kept numPicBullet VML resolves.
    internal List<(string RelId, byte[] Bytes, string ContentType)> GetNumberingImageParts()
    {
        var result = new List<(string, byte[], string)>();
        var ndp = _doc.MainDocumentPart?.NumberingDefinitionsPart;
        if (ndp == null) return result;
        foreach (var ip in ndp.ImageParts)
        {
            string relId;
            try { relId = ndp.GetIdOfPart(ip); }
            catch { continue; }
            if (string.IsNullOrEmpty(relId)) continue;
            try
            {
                using var s = ip.GetStream();
                using var ms = new MemoryStream();
                s.CopyTo(ms);
                result.Add((relId, ms.ToArray(), ip.ContentType));
            }
            catch { /* unreadable image part — skip */ }
        }
        return result;
    }

    // ==================== Raw Layer ====================

    public string Raw(string partPath, int? startRow = null, int? endRow = null, HashSet<string>? cols = null)
    {
        if (partPath == null) throw new ArgumentNullException(nameof(partPath));
        var mainPart = _doc.MainDocumentPart;
        if (mainPart == null) return "(no main part)";

        // CONSISTENCY(zip-uri-lookup): see RawXmlHelper. Any path ending in
        // .xml or .rels is resolved against the package directly.
        if (RawXmlHelper.IsZipUriPath(partPath))
        {
            var xml = RawXmlHelper.TryReadByZipUri(_doc, _filePath, partPath)
                ?? throw new ArgumentException(
                    $"Unknown part: {partPath}. The path was treated as a zip-internal URI " +
                    $"but no matching part exists in the package. " +
                    $"Use semantic paths (/document, /styles, /header[N]) for stable identification.");
            return xml;
        }

        return partPath.ToLowerInvariant() switch
        {
            "/document" => mainPart.Document?.OuterXml ?? "",
            "/styles" => mainPart.StyleDefinitionsPart?.Styles?.OuterXml ?? "(no styles)",
            "/settings" => mainPart.DocumentSettingsPart?.Settings?.OuterXml ?? "(no settings)",
            "/numbering" => mainPart.NumberingDefinitionsPart?.Numbering?.OuterXml ?? "(no numbering)",
            "/comments" => mainPart.WordprocessingCommentsPart?.Comments?.OuterXml ?? "(no comments)",
            "/theme" => mainPart.ThemePart?.Theme?.OuterXml ?? "(no theme)",
            // BUG-DUMP-R42-3: word/fontTable.xml declares font faces +
            // altName substitutions (e.g. 方正小标宋简体 altName=方正舒体).
            // Without it Word substitutes different faces — a real visual
            // change. FontTablePart has a typed Fonts root we read verbatim.
            "/fonttable" => mainPart.FontTablePart?.Fonts?.OuterXml ?? "(no fontTable)",
            "/websettings" => mainPart.WebSettingsPart?.WebSettings?.OuterXml ?? "(no webSettings)",
            _ when partPath.StartsWith("/header") => GetHeaderRawXml(partPath),
            _ when partPath.StartsWith("/footer") => GetFooterRawXml(partPath),
            _ when partPath.StartsWith("/chart") => GetChartRawXml(partPath),
            _ => throw new ArgumentException($"Unknown part: {partPath}. Available: /document, /styles, /settings, /numbering, /comments, /theme, /header[n], /footer[n], /chart[n]")
        };
    }

    public void RawSet(string partPath, string xpath, string action, string? xml)
    {
        Modified = true;
        ClearBodyChildIndex(); // raw-set may rewrite the body / its paragraph set
        if (partPath == null) throw new ArgumentNullException(nameof(partPath));
        var mainPart = _doc.MainDocumentPart
            ?? throw new InvalidOperationException("No main document part");

        // BUG-DUMP-R45-1 / BUG-DUMP-R45-2: companion binary-attach action for the
        // /fontTable and /numbering raw-set replace. The dump emits the part XML
        // verbatim (keeping <w:embed*> / <w:numPicBullet> refs) PLUS one
        // `embed-binary` op per referenced binary, carrying the bytes as a
        // `data:<ct>;base64,…` URI in `xml` and the SOURCE relationship id in
        // `xpath`. Rebuild the FontPart / ImagePart with that exact r:id so the
        // already-replaced XML's refs resolve — no XML rewrite, no dangling rel.
        // Mirrors the OLE base64-inline round-trip (OleHelper.AddEmbeddedPartFromBytes).
        if (string.Equals(action, "embed-binary", StringComparison.OrdinalIgnoreCase))
        {
            RawEmbedBinary(mainPart, partPath, xpath, xml);
            return;
        }

        if (RawXmlHelper.IsZipUriPath(partPath))
        {
            var part = RawXmlHelper.FindPartByZipUri(_doc, partPath)
                ?? throw new ArgumentException(
                    $"Unknown part: {partPath}. The path was treated as a zip-internal URI " +
                    $"but no matching part exists in the package. " +
                    $"Use semantic paths (/document, /styles, /header[N]) for stable identification.");
            RawXmlHelper.Execute(part, xpath, action, xml);
            return;
        }

        OpenXmlPartRootElement rootElement;
        var lowerPath = partPath.ToLowerInvariant();

        if (lowerPath is "/document" or "/")
            rootElement = mainPart.Document ?? throw new InvalidOperationException("No document");
        else if (lowerPath is "/styles")
            rootElement = mainPart.StyleDefinitionsPart?.Styles ?? throw new InvalidOperationException("No styles part");
        else if (lowerPath is "/settings")
            rootElement = mainPart.DocumentSettingsPart?.Settings ?? throw new InvalidOperationException("No settings part");
        else if (lowerPath is "/numbering")
        {
            // CONSISTENCY(raw-set-create-missing-part): see /theme branch.
            var numPart = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
            if (numPart.Numbering == null)
            {
                numPart.Numbering = new Numbering();
                numPart.Numbering.Save();
            }
            rootElement = numPart.Numbering;
        }
        else if (lowerPath is "/websettings")
        {
            // CONSISTENCY(raw-set-create-missing-part): blank docs have no
            // webSettings part; dump→batch from a Word-authored source emits
            // raw-set /webSettings replace, so create the part lazily.
            var wsPart = mainPart.WebSettingsPart ?? mainPart.AddNewPart<WebSettingsPart>();
            if (wsPart.WebSettings == null)
            {
                wsPart.WebSettings = new WebSettings();
                wsPart.WebSettings.Save();
            }
            rootElement = wsPart.WebSettings;
        }
        else if (lowerPath is "/comments")
            rootElement = mainPart.WordprocessingCommentsPart?.Comments ?? throw new InvalidOperationException("No comments part");
        else if (lowerPath is "/theme")
        {
            // BUG-DUMP-R37-5: a theme-LESS source must round-trip with NO theme
            // part — the blank rebuild template auto-stamps theme1.xml, whose
            // minorFont resolves CJK glyphs to a different fallback than Word's
            // app-default theme, tightening CJK line height and drifting body
            // text. EmitThemeRaw emits `action=remove` on /theme when the source
            // had none; delete the whole ThemePart (DeletePart also drops the
            // document.xml.rels theme relationship, so no dangling ref remains).
            // Mirrors EmitDocDefaultsRaw's remove path for a source lacking
            // docDefaults. Idempotent: a re-dump of the rebuilt (theme-less) doc
            // sees no theme and emits the same remove.
            if (string.Equals(action, "remove", StringComparison.OrdinalIgnoreCase)
                && string.Equals(partPath, "/theme", StringComparison.OrdinalIgnoreCase)
                && string.Equals(xpath, "/a:theme", StringComparison.OrdinalIgnoreCase))
            {
                if (mainPart.ThemePart != null)
                    mainPart.DeletePart(mainPart.ThemePart);
                return;
            }
            // CONSISTENCY(raw-set-create-missing-part): blank docs created via
            // BlankDocCreator have no ThemePart; dump→batch round-trip from a
            // real Word/python-docx file emits raw-set /theme replace which
            // would otherwise abort the whole batch. Lazily add the theme part
            // and an empty <a:theme> root so RawXmlHelper.Execute can match
            // /a:theme and replace it with the dumped XML.
            var themePart = mainPart.ThemePart ?? mainPart.AddNewPart<ThemePart>();
            if (themePart.Theme == null)
            {
                themePart.Theme = new DocumentFormat.OpenXml.Drawing.Theme(
                    new DocumentFormat.OpenXml.Drawing.ThemeElements());
                themePart.Theme.Save();
            }
            rootElement = themePart.Theme;
        }
        else if (lowerPath is "/fonttable")
        {
            // BUG-DUMP-R42-3: round-trip word/fontTable.xml (font-face
            // declarations + altName substitutions). The blank rebuild template
            // has no FontTablePart; lazily create one with an empty <w:fonts>
            // root so RawXmlHelper.Execute can match /w:fonts and replace it
            // with the dumped XML. The emitter strips any embedded-font
            // references (w:embedRegular/Bold/Italic/BoldItalic r:id) first, so
            // no dangling rel survives (the .odttf binaries are not round-tripped).
            // CONSISTENCY(raw-set-create-missing-part): mirrors the /theme +
            // /numbering branches.
            var fontTablePart = mainPart.FontTablePart ?? mainPart.AddNewPart<FontTablePart>();
            if (fontTablePart.Fonts == null)
            {
                fontTablePart.Fonts = new DocumentFormat.OpenXml.Wordprocessing.Fonts();
                fontTablePart.Fonts.Save();
            }
            rootElement = fontTablePart.Fonts;
        }
        else if (lowerPath.StartsWith("/header"))
        {
            var idx = 0;
            var bracketIdx = partPath.IndexOf('[');
            if (bracketIdx >= 0)
                int.TryParse(partPath[(bracketIdx + 1)..].TrimEnd(']'), out idx);
            var headerPart = mainPart.HeaderParts.ElementAtOrDefault(idx - 1)
                ?? throw new ArgumentException($"header[{idx}] not found");
            rootElement = headerPart.Header ?? throw new InvalidOperationException($"Corrupt file: header[{idx}] data missing");
        }
        else if (lowerPath.StartsWith("/footer"))
        {
            var idx = 0;
            var bracketIdx = partPath.IndexOf('[');
            if (bracketIdx >= 0)
                int.TryParse(partPath[(bracketIdx + 1)..].TrimEnd(']'), out idx);
            var footerPart = mainPart.FooterParts.ElementAtOrDefault(idx - 1)
                ?? throw new ArgumentException($"footer[{idx}] not found");
            rootElement = footerPart.Footer ?? throw new InvalidOperationException($"Corrupt file: footer[{idx}] data missing");
        }
        else if (lowerPath.StartsWith("/chart"))
        {
            var idx = 0;
            var bracketIdx = partPath.IndexOf('[');
            if (bracketIdx >= 0)
                int.TryParse(partPath[(bracketIdx + 1)..].TrimEnd(']'), out idx);
            var chartPart = mainPart.ChartParts.ElementAtOrDefault(idx - 1)
                ?? throw new ArgumentException($"chart[{idx}] not found");
            rootElement = chartPart.ChartSpace ?? throw new InvalidOperationException($"Corrupt file: chart[{idx}] data missing");
        }
        else if (lowerPath is "/commentsextended")
        {
            // BUG-DUMP-R26-4: round-trip word/commentsExtended.xml (modern
            // comment-reply threading; w15:commentEx paraIdParent links a reply
            // to its parent comment, keyed by the comment paragraphs' w14:paraId).
            // The part has no clean typed-root API across SDK versions, so feed
            // the whole part stream directly. `replace` (the only action the dump
            // emits) overwrites the part body verbatim; the part is lazily created
            // on a blank rebuilt doc. The linkage stays valid because EmitComments
            // now preserves the comment paragraphs' source w14:paraId.
            var exPart = mainPart.WordprocessingCommentsExPart
                ?? mainPart.AddNewPart<WordprocessingCommentsExPart>();
            if (!string.Equals(action, "replace", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException(
                    $"/commentsExtended supports only the 'replace' action (got '{action}').");
            if (string.IsNullOrEmpty(xml))
                throw new ArgumentException("/commentsExtended replace requires XML.");
            using (var ms = new System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml)))
                exPart.FeedData(ms);
            EnsureAllParaIds();
            return;
        }
        else
            throw new ArgumentException($"Unknown part: {partPath}. Available: /document, /styles, /settings, /numbering, /header[n], /footer[n], /chart[n]");

        var affected = RawXmlHelper.Execute(rootElement, xpath, action, xml);
        rootElement.Save();
        // CONSISTENCY(paraid-global-uniqueness): RawSet may inject paragraphs
        // carrying paraIds the handler hasn't seen — without re-scanning,
        // _usedParaIds and _nextParaId stay stale and the next AddBreak /
        // AddParagraph could allocate a colliding paraId. Especially
        // dangerous in resident mode where one process serves many commands
        // across the same _usedParaIds set. Re-run EnsureAllParaIds after
        // every successful raw mutation so the global pool stays accurate.
        EnsureAllParaIds();
        // CONSISTENCY(docpr-global-uniqueness): same hazard for <wp:docPr> ids.
        // RawSet may inject a drawing whose docPr id collides with one a typed
        // add already allocated via NextDocPropId (e.g. dump preserving a
        // non-textbox shape replays its source id verbatim while the textbox
        // alongside it was renumbered onto the same value). EnsureDocPropIds
        // otherwise only runs at open, so an in-session raw mutation would
        // leave the duplicate on disk; re-run it here to dedupe to the lowest
        // free id — the same correction paraId gets above.
        EnsureDocPropIds();
        // CONSISTENCY(sdt-global-uniqueness): same hazard for sdt <w:id>.
        // NextSdtId allocates max+1 for typed adds, but a raw-set can inject an
        // sdt carrying a colliding id. EnsureSdtIds otherwise runs only at open.
        EnsureSdtIds();
        // BUG-R5-01: do not emit chatter from inside the handler — the CLI
        // wrappers (CommandBuilder.Raw raw-set + batch run raw-set) print
        // their own structured message. Writing here pollutes batch --json
        // output (extra stdout lines escaped into result.message strings).
        _ = affected;
    }

    // BUG-DUMP-R45-1 / BUG-DUMP-R45-2: apply an `embed-binary` companion op.
    // partPath selects the host part (/fontTable → FontPart, /numbering →
    // ImagePart); xpath carries the source relationship id to reuse; xml is a
    // `data:<contentType>;base64,…` URI. The FontTablePart / NumberingDefinitions
    // Part is created lazily by the preceding `raw-set replace`, so it already
    // exists here. AddNewPart<T>(contentType, id) attaches the part with the
    // exact r:id, so the verbatim <w:embed*> / <v:imagedata> ref resolves.
    private void RawEmbedBinary(MainDocumentPart mainPart, string partPath, string? xpath, string? xml)
    {
        var relId = xpath;
        if (string.IsNullOrEmpty(relId))
            throw new ArgumentException("embed-binary requires the relationship id in 'xpath'.");
        if (!OleHelper.TryDecodeDataUri(xml, out var bytes, out var contentType) || bytes.Length == 0)
            throw new ArgumentException("embed-binary requires a non-empty data: URI in 'xml'.");

        var lowerPath = partPath.ToLowerInvariant();
        if (lowerPath is "/fonttable")
        {
            var ftp = mainPart.FontTablePart
                ?? throw new InvalidOperationException("No fontTable part for embed-binary");
            var ct = string.IsNullOrEmpty(contentType)
                ? "application/vnd.openxmlformats-officedocument.obfuscatedFont"
                : contentType;
            var fp = ftp.AddNewPart<FontPart>(ct, relId);
            using var ms = new MemoryStream(bytes);
            fp.FeedData(ms);
        }
        else if (lowerPath is "/numbering")
        {
            var ndp = mainPart.NumberingDefinitionsPart
                ?? throw new InvalidOperationException("No numbering part for embed-binary");
            var ct = string.IsNullOrEmpty(contentType) ? "image/png" : contentType;
            var ip = ndp.AddNewPart<ImagePart>(ct, relId);
            using var ms = new MemoryStream(bytes);
            ip.FeedData(ms);
        }
        else if (lowerPath is "/customxml")
        {
            // customXml data store item — recreate with the SOURCE rel id so
            // the companion props op can address it as /customXml/<relId>.
            // Nothing in any XML references this id, so reuse is collision-
            // safe (the rebuilt main part's auto ids use a different scheme).
            var ct = string.IsNullOrEmpty(contentType) ? "application/xml" : contentType;
            var cxp = mainPart.AddNewPart<CustomXmlPart>(ct, relId);
            using var ms = new MemoryStream(bytes);
            cxp.FeedData(ms);
        }
        else if (lowerPath.StartsWith("/customxml/", StringComparison.Ordinal))
        {
            // itemProps child of a previously attached customXml item part.
            var itemRelId = partPath.Substring("/customXml/".Length);
            OpenXmlPart item;
            try { item = mainPart.GetPartById(itemRelId); }
            catch
            {
                throw new InvalidOperationException(
                    $"embed-binary: no customXml item part with rel id '{itemRelId}'");
            }
            if (item is not CustomXmlPart customItem)
                throw new InvalidOperationException(
                    $"embed-binary: rel id '{itemRelId}' is not a customXml part");
            var ct = string.IsNullOrEmpty(contentType)
                ? "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
                : contentType;
            var propsPart = customItem.AddNewPart<CustomXmlPropertiesPart>(ct, relId);
            using var ms = new MemoryStream(bytes);
            propsPart.FeedData(ms);
        }
        else
        {
            throw new ArgumentException(
                $"embed-binary not supported for part '{partPath}' (only /fontTable, /numbering, /customXml).");
        }
    }

    /// <summary>
    /// BUG-DUMP-R26-4: read word/commentsExtended.xml verbatim for the dump
    /// emitter, or null when the part is absent. The part carries modern
    /// comment-reply threading (w15:commentEx paraIdParent) keyed by comment
    /// paragraph w14:paraId; EmitComments raw-sets this XML back on replay and
    /// preserves the comment paraIds so the linkage survives round-trip.
    /// </summary>
    internal string? GetCommentsExtendedXml()
    {
        var exPart = _doc.MainDocumentPart?.WordprocessingCommentsExPart;
        if (exPart == null) return null;
        using var reader = new System.IO.StreamReader(exPart.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read));
        var xml = reader.ReadToEnd();
        return string.IsNullOrWhiteSpace(xml) ? null : xml;
    }

    public List<ValidationError> Validate() => RawXmlHelper.ValidateDocument(_doc, _filePath);

    /// <summary>
    /// Run the global id-uniqueness passes once, just before a deferred-batch
    /// flush. During <see cref="DeferSave"/> replay the per-mutation id
    /// management that normally runs after each raw-set is skipped, so an
    /// `add p` carrying a paragraph-mark revision id preserved from the source
    /// can collide with a content-run ins/del id generated by a later item.
    /// EnsureAllParaIds (which also dedupes revision w:ids), EnsureDocPropIds and
    /// EnsureSdtIds reconcile the whole document in one O(N) sweep at the end,
    /// matching the after-raw-set normalization but without the per-op O(N²) cost
    /// the deferred path exists to avoid. Best-effort; never block the save.
    /// </summary>
    private void FinalizeDeferredIds()
    {
        try
        {
            EnsureAllParaIds();
            EnsureDocPropIds();
            EnsureSdtIds();
        }
        catch { /* id normalization is best-effort; never block the flush */ }
    }

    /// <summary>
    /// BUG-R7B(BUG2): document-wide id reconciliation, callable mid-session.
    /// Under <see cref="DeferSave"/> the per-raw-set id passes
    /// (EnsureDocPropIds / EnsureAllParaIds / EnsureSdtIds) are skipped, so a
    /// batch that raw-sets header/footer parts can leave the in-memory document
    /// with colliding wp:docPr ids until the next save/close runs
    /// <see cref="FinalizeDeferredIds"/>. A `validate` (or any flush) issued
    /// between the batch and the save would then see the duplicates. The
    /// resident batch path calls this right after the batch loop so the
    /// in-memory tree matches the on-disk state save/close would produce —
    /// the same document-scoped passes, run once.
    /// </summary>
    public void ReconcileGlobalIds() => FinalizeDeferredIds();

    public void Save()
    {
        if (DeferSave) FinalizeDeferredIds();
        // Mid-session flush. The Dispose-time NormalizeSelfClosingInDocx step
        // is intentionally skipped here — it requires opening the file as a
        // Zip with read-write access, which can't be done while the backing
        // stream still holds the file. The on-disk snapshot will have
        // `<w:br />` form instead of the canonical `<w:br/>` form; both are
        // schema-valid OOXML.
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
        // Mirror the PPT pattern: when we own the backing FileStream the
        // package would otherwise leave the on-disk file in whatever state
        // the last auto-flush left it (potentially truncated for the
        // stream-Open path). Save first, then dispose.
        if (DeferSave) FinalizeDeferredIds();
        if (Modified)
        {
            try { OfficeCli.Core.OfficeCliMetadata.StampOnSave(_doc); }
            catch { /* best-effort audit trail */ }
        }
        try { _doc.Save(); } catch { /* read-only or already disposed */ }
        _doc.Dispose();
        _backingStream?.Dispose();
        _backingStream = null;
        // CONSISTENCY(word-self-close): the OpenXml SDK serializes empty
        // elements with a space before the self-close (`<w:br />`). Several
        // downstream consumers (and test regexes) look for the canonical
        // `<w:br/>` / `<w:tab/>` form. Normalize the persisted document.xml
        // in place so the saved package matches the canonical short form.
        // Only applied to word/document.xml; styles/settings/numbering are
        // left untouched since the space form is schema-equivalent.
        try { NormalizeSelfClosingInDocx(_filePath); } catch { /* best-effort */ }
    }

    private static void NormalizeSelfClosingInDocx(string path)
    {
        if (!System.IO.File.Exists(path)) return;
        using var fs = new System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite);
        using var za = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: false);
        var entry = za.GetEntry("word/document.xml");
        if (entry == null) return;
        string xml;
        using (var rs = entry.Open())
        using (var sr = new System.IO.StreamReader(rs))
            xml = sr.ReadToEnd();
        // Collapse "<w:br />" → "<w:br/>" and "<w:tab />" → "<w:tab/>"
        // (no-attribute empty elements only).
        var normalized = System.Text.RegularExpressions.Regex.Replace(
            xml, @"<w:(br|tab) />", "<w:$1/>");
        if (normalized == xml) return;
        entry.Delete();
        var newEntry = za.CreateEntry("word/document.xml");
        using var ws = newEntry.Open();
        using var sw = new System.IO.StreamWriter(ws, new System.Text.UTF8Encoding(false));
        sw.Write(normalized);
    }

    // (private helpers, navigation, selector, style/list, image helpers moved to Word/ partial files)
}
