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
        _doc = WordprocessingDocument.Open(_backingStream, editable);
        WordStrictAttributeSanitizer.Sanitize(_doc);
        if (editable)
        {
            EnsureAllParaIds();
            EnsureDocPropIds();
            EnsureSdtIds();
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

    private OpenXmlPart ResolveImageHostPart(Run run)
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
        else if (lowerPath is "/comments")
            rootElement = mainPart.WordprocessingCommentsPart?.Comments ?? throw new InvalidOperationException("No comments part");
        else if (lowerPath is "/theme")
        {
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

    public List<ValidationError> Validate() => RawXmlHelper.ValidateDocument(_doc);

    public void Save()
    {
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
