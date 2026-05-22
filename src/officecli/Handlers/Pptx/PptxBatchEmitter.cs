// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace OfficeCli.Handlers;

// CONSISTENCY(emit-X-mirror): scaffold mirrors WordBatchEmitter.cs — same
// public entry shape (full-doc + subtree overloads), same Get-driven
// transcription, same partial-class split (entry / Filters / Shape / Notes).
//
// PR1 scope (text-only): slide / shape / textbox / title / connector /
// group / placeholder + paragraph + run. Tables, pictures, charts, notes
// bodies, layout/master/theme raw — PR2.
public static partial class PptxBatchEmitter
{
    /// <summary>
    /// Carry-state for one emit run. Mirrors WordBatchEmitter.BodyEmitContext
    /// but trimmed for PR1 (no footnote/endnote/chart cursors yet —
    /// PowerPoint has no notes-with-numbering concept; chart/table content
    /// lands in PR2).
    /// </summary>
    internal sealed record SlideEmitContext(
        List<UnsupportedWarning> Unsupported)
    {
        // Forward slide-jump links (e.g. shape[1] on slide[1] linking to
        // slide[3]) must replay AFTER every slide is added — otherwise the
        // `link=slide[N]` prop on shape Add resolves against a deck where
        // the target slide does not yet exist and ResolveHyperlinkTarget
        // throws "Slide jump target out of range". Defer those props into
        // a second set-pass appended at the end of EmitPptx.
        public List<BatchItem> DeferredLinks { get; } = new();
    }

    /// <summary>
    /// Captured at emit time when a slide carries content we cannot round-trip
    /// through the existing handler vocabulary (animations, SmartArt, OLE,
    /// video/audio, exotic transitions). The slide itself is emitted; the
    /// unsupported element is dropped silently from `items` but recorded
    /// here so the CLI can surface a warning bundle to the caller.
    /// </summary>
    public sealed record UnsupportedWarning(string Element, string SlidePath, string Reason);

    /// <summary>
    /// Emit a full PowerPoint document as a sequence of BatchItem rows.
    /// Returns the items plus any unsupported-element warnings.
    /// </summary>
    public static (List<BatchItem> Items, List<UnsupportedWarning> Warnings) EmitPptx(PowerPointHandler ppt)
    {
        var items = new List<BatchItem>();
        var ctx = new SlideEmitContext(new List<UnsupportedWarning>());

        // Clear the target deck's slides FIRST so replay onto a non-empty
        // target lands on a clean slate. Without this, `add slide` items
        // append after existing slides while every `add shape parent=/slide[N]`
        // path still resolves to the original slide[N] — the target ends up
        // with 2× the slide count (existing + freshly added empties) on each
        // round-trip. `remove /slide[*]` is a no-op on a deck with 0 slides,
        // so this is safe for the clean-target case too.
        items.Add(new BatchItem { Command = "remove", Path = "/slide[*]" });

        // Resource parts FIRST — theme, notesMaster, masters, layouts.
        // Order matters: replay's raw-set must overwrite the blank deck's
        // seeded baseline before slide content is added so per-slide
        // layout refs (sld@layout="rId4") resolve against the source's
        // layout set, not blank's. Mirrors docx's
        // settings → theme → numbering → styles → body ordering.
        EmitThemeRaw(ppt, items);
        EmitNotesMasterRaw(ppt, items);
        EmitMasterRaw(ppt, items);
        EmitLayoutRaw(ppt, items);
        // R8-5: emit presentation-level slide dimensions so custom sldSz
        // round-trips through dump → batch. Previously EmitPptx skipped the
        // root node entirely; replay always landed on the blank-deck default
        // (33.87cm × 19.05cm widescreen), silently resizing decks built for
        // 4:3, A4, custom banners, etc.
        EmitPresentationProps(ppt, items);

        // CONSISTENCY(slide-order): always iterate via the handler's
        // GetSlideParts() (sldIdLst-driven). Walking SlideParts off the
        // package returns parts in zip URI order — `slide12.xml` sorts
        // before `slide3.xml`, scrambling user-visible order.
        // CONSISTENCY(emit-skip-on-validate): a non-standard attribute or
        // element on a single slide must not abort the whole dump. The
        // OpenXml SDK throws a flat InvalidOperationException ("The element
        // does not allow the specified attribute.") when its strict-mode
        // validator catches a foreign/extension attribute (common in vendor
        // templates: gov_bja_template, 1.pptx, ...). Iterate slides one by
        // one and surface OOXML validation failures as unsupported_element
        // warnings instead of crashing the whole dump.
        var slideCount = ppt.SlideCount;
        for (int slideNum = 1; slideNum <= slideCount; slideNum++)
        {
            var slidePath = $"/slide[{slideNum}]";
            // CONSISTENCY(slide-ordinal-stub): every iteration MUST contribute
            // exactly one `add slide` so subsequent set paths /slide[N+1]/…
            // resolve to the same N+1 slot on replay. Pre-R5 we just
            // `continue`d on validation failure, emitting zero items for the
            // skipped slide — every later set drifted by one slot and
            // dump → batch on a deck with one bad slide could orphan
            // hundreds of items.
            DocumentNode slideNode;
            int preCount = items.Count;
            try { slideNode = ppt.Get(slidePath); }
            catch (Exception ex) when (ex.Message.Contains("does not allow", StringComparison.Ordinal)
                                    || ex.Message.Contains("not allowed", StringComparison.Ordinal))
            {
                ctx.Unsupported.Add(new UnsupportedWarning(
                    Element: "slide.ooxml_validation",
                    SlidePath: slidePath,
                    Reason: ex.Message));
                items.Add(new BatchItem { Command = "add", Parent = "/", Type = "slide" });
                continue;
            }
            try
            {
                EmitSlide(ppt, slideNode, slideNum, items, ctx);
            }
            catch (Exception ex) when (ex.Message.Contains("does not allow", StringComparison.Ordinal)
                                    || ex.Message.Contains("not allowed", StringComparison.Ordinal))
            {
                ctx.Unsupported.Add(new UnsupportedWarning(
                    Element: "slide.ooxml_validation",
                    SlidePath: slidePath,
                    Reason: ex.Message));
                // Roll back partial emits from the failing slide and replace
                // with a single blank-slide stub to keep ordinals aligned.
                if (items.Count > preCount)
                    items.RemoveRange(preCount, items.Count - preCount);
                items.Add(new BatchItem { Command = "add", Parent = "/", Type = "slide" });
            }
        }

        // Flush deferred slide-jump link sets — every target slide now exists,
        // so `ResolveHyperlinkTarget` can map slide[N] to the relationship.
        if (ctx.DeferredLinks.Count > 0)
            items.AddRange(ctx.DeferredLinks);

        return (items, ctx.Unsupported);
    }

    // R8-5: emit a single `set /` carrying slideWidth/slideHeight when the
    // source deck deviates from the blank-baseline 33.87cm × 19.05cm
    // widescreen. The blank-doc default is hard-coded inside BlankDocCreator,
    // not surfaced by Get, so we string-compare the canonical FormatEmu
    // output. EmitPresentationProps is a no-op for the default case to keep
    // unchanged decks from gaining a spurious item on round-trip.
    private const string DefaultSlideWidth = "33.87cm";
    private const string DefaultSlideHeight = "19.05cm";

    // Presentation-level Format keys that TrySetPresentationSetting accepts
    // on `set /`. The Get side surfaces these via PopulatePresentationSettings
    // (Set.Presentation.cs); without this allowlist, only slideWidth/Height
    // round-tripped — firstSlideNum, show.loop, print.*, compatMode, etc.
    // were silently dropped on dump.
    //
    // Get emits `direction = rtl` for RTL presentations but the setter case
    // key is `rtl`. We rewrite the key on emit so replay's TrySetPresentationSetting
    // accepts it. Mirrors the `direction → rtl` alias that already lives in
    // Set.cs path-pattern dispatch.
    private static readonly HashSet<string> PresentationEmitKeys =
        new(StringComparer.OrdinalIgnoreCase)
        {
            "firstSlideNum", "compatMode", "removePersonalInfo",
            "print.what", "print.colorMode", "print.hiddenSlides",
            "print.scaleToFitPaper", "print.frameSlides",
            "show.loop", "show.narration", "show.animation", "show.useTimings",
        };

    private static void EmitPresentationProps(PowerPointHandler ppt, List<BatchItem> items)
    {
        DocumentNode root;
        try { root = ppt.Get("/"); }
        catch { return; }
        var props = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (root.Format.TryGetValue("slideWidth", out var wObj) && wObj is string w
            && !string.Equals(w, DefaultSlideWidth, StringComparison.OrdinalIgnoreCase))
            props["slideWidth"] = w;
        if (root.Format.TryGetValue("slideHeight", out var hObj) && hObj is string h
            && !string.Equals(h, DefaultSlideHeight, StringComparison.OrdinalIgnoreCase))
            props["slideHeight"] = h;

        // Presentation attributes / print / show settings — only emit non-default
        // values (Get omits keys that match the OOXML defaults).
        foreach (var key in PresentationEmitKeys)
        {
            if (!root.Format.TryGetValue(key, out var v) || v == null) continue;
            var s = v switch { bool b => b ? "true" : "false", _ => v.ToString() ?? "" };
            if (s.Length == 0) continue;
            props[key] = s;
        }

        // direction → rtl: Get emits `direction = rtl`, setter accepts `rtl`.
        if (root.Format.TryGetValue("direction", out var dObj) && dObj is string ds
            && ds.Equals("rtl", StringComparison.OrdinalIgnoreCase))
            props["rtl"] = "true";

        if (props.Count == 0) return;
        items.Add(new BatchItem
        {
            Command = "set",
            Path = "/",
            Props = props,
        });
    }

    /// <summary>
    /// Emit a subtree of a PowerPoint document. Supported subtree paths:
    /// `/slide[N]`, `/theme`, `/notesMaster`, `/slideMaster[N]`, `/slideLayout[N]`,
    /// `/noteSlide[N]`, `/presentation`. Resource subtrees emit a single raw-set
    /// replace; replay onto a foreign deck does NOT carry cross-part dependency
    /// closure (e.g. a `/slideLayout[K]` dump only stamps the layout's XML — the
    /// referenced master, theme, and per-slide layout rId rewiring are NOT
    /// included). Mirrors WordBatchEmitter's raw-emit subtree surface
    /// (/theme, /settings, /numbering, /styles).
    /// </summary>
    public static (List<BatchItem> Items, List<UnsupportedWarning> Warnings) EmitPptx(
        PowerPointHandler ppt, string path)
    {
        const string SupportedHint = "Supported: /, /presentation, /slide[N], /theme, /notesMaster, /slideMaster[N], /slideLayout[N], /noteSlide[N]";

        if (string.IsNullOrEmpty(path))
            throw new CliException($"dump path cannot be empty. Use '/' for the full document or a subtree path like /slide[N]. {SupportedHint}")
                { Code = "invalid_path" };
        if (path == "/") return EmitPptx(ppt);

        var items = new List<BatchItem>();
        var ctx = new SlideEmitContext(new List<UnsupportedWarning>());

        // CONSISTENCY(case-insensitive-subtree): paths with [N] go through
        // regex `IgnoreCase`, so case-folding the literal-prefix branches too
        // aligns the dispatcher to a single rule and matches the docx subtree
        // dispatcher (WordBatchEmitter uses `path.ToLowerInvariant()`).
        var lp = path.ToLowerInvariant();

        if (lp == "/presentation")
        {
            EmitPresentationProps(ppt, items);
            return (items, ctx.Unsupported);
        }
        if (lp == "/theme")
        {
            EmitThemeRaw(ppt, items);
            return (items, ctx.Unsupported);
        }
        if (lp == "/notesmaster")
        {
            EmitNotesMasterRaw(ppt, items);
            return (items, ctx.Unsupported);
        }

        // Index parsing: regex restricts to ASCII [0-9]+ (not Unicode \d, which
        // matches Arabic-Indic numerals etc. that int.Parse rejects under
        // InvariantCulture). int.TryParse guards against Int32 overflow.
        static int ParseIndexOrThrow(string raw, string fullPath)
        {
            if (!int.TryParse(raw, System.Globalization.NumberStyles.Integer,
                              System.Globalization.CultureInfo.InvariantCulture, out var n))
                throw new CliException($"dump path not found: {fullPath} (index '{raw}' out of range or not an integer)")
                    { Code = "path_not_found" };
            return n;
        }

        var slideMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/slide\[([0-9]+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (slideMatch.Success)
        {
            var idx = ParseIndexOrThrow(slideMatch.Groups[1].Value, path);
            DocumentNode slideNode;
            try { slideNode = ppt.Get(path); }
            catch (Exception ex)
            {
                throw new CliException($"dump path not found: {path} ({ex.Message})") { Code = "path_not_found" };
            }
            EmitSlide(ppt, slideNode, idx, items, ctx);
            // CONSISTENCY(deferred-link-flush): subtree slide dump must flush
            // ctx.DeferredLinks before returning, otherwise any `link=slide[N]`
            // prop on a shape inside the dumped slide is silently dropped from
            // the output (DeferSlideJumpLink moves it out of the shape's
            // prop bag into ctx.DeferredLinks expecting the full-doc EmitPptx
            // tail flush, which the subtree path never reaches).
            //
            // Cross-slide targets (e.g. dump /slide[1] when the shape links
            // to /slide[3]) still emit the set row — replay against a deck
            // missing the target slide fails with ResolveHyperlinkTarget's
            // "Slide jump target out of range", which is a clearer error
            // than silent prop loss.
            if (ctx.DeferredLinks.Count > 0)
                items.AddRange(ctx.DeferredLinks);
            return (items, ctx.Unsupported);
        }

        var masterMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/slideMaster\[([0-9]+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (masterMatch.Success)
        {
            var idx = ParseIndexOrThrow(masterMatch.Groups[1].Value, path);
            if (idx < 1 || idx > ppt.SlideMasterCount)
                throw new CliException($"dump path not found: {path} (total slideMasters: {ppt.SlideMasterCount})")
                    { Code = "path_not_found" };
            if (!EmitMasterRawOne(ppt, idx, items))
                throw new CliException($"dump path not found: {path} (slideMaster {idx} raw read failed)")
                    { Code = "path_not_found" };
            return (items, ctx.Unsupported);
        }

        var layoutMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/slideLayout\[([0-9]+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (layoutMatch.Success)
        {
            var idx = ParseIndexOrThrow(layoutMatch.Groups[1].Value, path);
            if (idx < 1 || idx > ppt.SlideLayoutCount)
                throw new CliException($"dump path not found: {path} (total slideLayouts: {ppt.SlideLayoutCount})")
                    { Code = "path_not_found" };
            if (!EmitLayoutRawOne(ppt, idx, items))
                throw new CliException($"dump path not found: {path} (slideLayout {idx} raw read failed)")
                    { Code = "path_not_found" };
            return (items, ctx.Unsupported);
        }

        var noteMatch = System.Text.RegularExpressions.Regex.Match(path, @"^/noteSlide\[([0-9]+)\]$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        if (noteMatch.Success)
        {
            var idx = ParseIndexOrThrow(noteMatch.Groups[1].Value, path);
            if (idx < 1 || idx > ppt.SlideCount)
                throw new CliException($"dump path not found: {path} (total slides: {ppt.SlideCount})")
                    { Code = "path_not_found" };
            if (!EmitNoteSlideRawOne(ppt, idx, items))
                throw new CliException($"dump path not found: {path} (slide {idx} has no notes)")
                    { Code = "path_not_found" };
            return (items, ctx.Unsupported);
        }

        throw new CliException(
            $"dump path not supported: {path}. {SupportedHint}")
            { Code = "unsupported_path" };
    }

    private static void EmitSlide(PowerPointHandler ppt, DocumentNode slideNode, int slideNum,
                                  List<BatchItem> items, SlideEmitContext ctx)
    {
        var slidePath = slideNode.Path;
        ProbeUnsupportedOnSlide(ppt, slidePath, ctx);

        // Detect exotic transition / timing content that the semantic emit
        // can't faithfully reproduce (morph + p14/p15 transitions, motion
        // paths, sequence groupings). When present, we suppress the semantic
        // emit for that category and emit a raw-set passthrough at the end
        // of the slide — single source of truth per slide-per-category.
        var exotic = ScanSlideExoticContent(ppt, slidePath);

        // Pull the full slide node so layout / hidden / background etc. surface
        // even when the entry passed us a depth-truncated tree from "/".
        var fullSlide = ppt.Get(slidePath);
        var slideProps = FilterEmittableProps(fullSlide.Format);

        if (exotic.HasExoticTransition)
        {
            // Strip transition-related props so the add slide doesn't write a
            // semantic <p:transition> that would then collide with the
            // raw-set append below (schema permits only one <p:transition>).
            foreach (var k in new[] { "transition", "transitionSpeed", "transitionDuration",
                                       "advanceTime", "advanceClick" })
                slideProps.Remove(k);
        }

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = "/",
            Type = "slide",
            Props = slideProps.Count > 0 ? slideProps : null,
        });

        // ShapeToNode tags placeholder shapes as plain "textbox"/"title". To
        // emit them as `add placeholder` we cross-reference each shape's cNvPr
        // id with the slide's Query("placeholder") result.
        // Only index placeholders defined on the slide itself. Query also
        // returns layout-inherited placeholders (Format["inheritedFrom"]
        // = "layout") whose ph index/id can collide with auto-assigned
        // textbox cNvPr ids on the slide (python-pptx starts at 2, layout
        // ftr/dt/sldNum live at id 2..4) — without this filter, the second
        // textbox would be misclassified as `ftr` and crash placeholder
        // type parsing, or silently disappear in dump.
        var placeholderById = new Dictionary<string, DocumentNode>(StringComparer.Ordinal);
        foreach (var ph in ppt.Query("placeholder"))
        {
            if (!ph.Path.StartsWith(slidePath + "/", StringComparison.Ordinal)) continue;
            if (ph.Format.TryGetValue("inheritedFrom", out var inh) && inh as string == "layout") continue;
            if (ph.Format.TryGetValue("id", out var phId) && phId != null)
                placeholderById[phId.ToString()!] = ph;
        }

        // Children: walk shape-tree level. Get already routed group/connector/
        // textbox/title/equation into typed nodes, so just iterate and dispatch.
        if (fullSlide.Children == null) return;
        // CONSISTENCY(positional-emit): dump references its own added elements
        // by positional `/slide[N]/shape[K]` (mirrors docx /body/p[K]) rather
        // than cNvPr `@id=N`. Add accepts caller-supplied id but emit chooses
        // not to use it — id collisions with layout-inherited placeholders
        // would otherwise break replay (animations/video deck cascade).
        //
        // CONSISTENCY(unified-shape-counter): placeholders are <p:sp> siblings
        // of plain shapes in the OOXML shape tree, so ResolveShape counts them
        // together. AddPlaceholder also appends a <p:sp> and returns
        // `/slide[N]/shape[<count>]` (Add.Misc.cs). The emitter must therefore
        // share a SINGLE positional counter across textbox/title/shape/equation
        // /placeholder and emit replay paths as `/slide[N]/shape[K]` for ALL
        // of them — otherwise a placeholder dispatched first leaves the
        // shape counter at 1, and the next textbox emits `set
        // /slide[N]/shape[1]/...` which on replay clobbers the placeholder.
        // Previously the emitter kept separate `shape`/`placeholder` counters
        // and emitted `/slide[N]/placeholder[K]` for placeholders, but the
        // replay paths for paragraph/run inside that placeholder still used
        // the same `/slide[N]/shape[K]` form — see EmitTextBody — so every
        // shape after a placeholder collided.
        // Pre-build the per-slide animation index keyed by source shape @id
        // (or positional fallback). EmitAnimationsForShape pulls per-shape
        // entries from this map as we emit each <p:sp>.
        //
        // When the slide has exotic timing content (motion paths, sequence
        // groupings, custom triggers the Query doesn't enumerate), we skip
        // semantic per-shape animation emits entirely and rely on a raw-set
        // passthrough of the whole <p:timing> tree appended at slide end.
        // Mixing the two would silently corrupt the replay (semantic add
        // would inject duplicate effect nodes alongside the raw-set tree).
        var animIndex = exotic.HasExoticTiming
            ? new Dictionary<string, List<DocumentNode>>(StringComparer.Ordinal)
            : BuildSlideAnimationIndex(ppt, slideNum);

        var ord = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var child in fullSlide.Children)
        {
            // Placeholder dispatch first — overrides textbox/title type.
            if ((child.Type == "textbox" || child.Type == "title" || child.Type == "shape")
                && child.Format.TryGetValue("id", out var cid) && cid != null
                && placeholderById.TryGetValue(cid.ToString()!, out var phNode))
            {
                ord["shape"] = ord.GetValueOrDefault("shape", 0) + 1;
                var phReplay = $"{slidePath}/shape[{ord["shape"]}]";
                EmitPlaceholder(ppt, phNode, slidePath, phReplay, items, ctx);
                EmitAnimationsForShape(GetAnimationsForChild(animIndex, child, ord["shape"]), phReplay, items);
                continue;
            }
            switch (child.Type)
            {
                case "textbox":
                case "title":
                case "shape":
                case "equation":
                    ord["shape"] = ord.GetValueOrDefault("shape", 0) + 1;
                    {
                        var shReplay = $"{slidePath}/shape[{ord["shape"]}]";
                        EmitShape(ppt, child, slidePath, shReplay, items, ctx);
                        EmitAnimationsForShape(GetAnimationsForChild(animIndex, child, ord["shape"]), shReplay, items);
                    }
                    break;
                case "placeholder":
                    ord["shape"] = ord.GetValueOrDefault("shape", 0) + 1;
                    {
                        var phReplay2 = $"{slidePath}/shape[{ord["shape"]}]";
                        EmitPlaceholder(ppt, child, slidePath, phReplay2, items, ctx);
                        EmitAnimationsForShape(GetAnimationsForChild(animIndex, child, ord["shape"]), phReplay2, items);
                    }
                    break;
                case "connector":
                    ord["connector"] = ord.GetValueOrDefault("connector", 0) + 1;
                    EmitConnector(ppt, child, slidePath, items, ctx);
                    break;
                case "group":
                    ord["group"] = ord.GetValueOrDefault("group", 0) + 1;
                    EmitGroup(ppt, child, slidePath, $"{slidePath}/group[{ord["group"]}]", items, ctx);
                    break;
                case "table":
                    ord["table"] = ord.GetValueOrDefault("table", 0) + 1;
                    EmitTable(ppt, child, slidePath, $"{slidePath}/table[{ord["table"]}]", items, ctx);
                    break;
                case "picture":
                    ord["picture"] = ord.GetValueOrDefault("picture", 0) + 1;
                    EmitPicture(ppt, child, slidePath, $"{slidePath}/picture[{ord["picture"]}]", items, ctx);
                    break;
                case "chart":
                    ord["chart"] = ord.GetValueOrDefault("chart", 0) + 1;
                    EmitChart(ppt, child, slidePath, items, ctx, ord["chart"]);
                    break;
                case "video":
                case "audio":
                    // Phase 3c-media: routed through EmitMediaForSlide (a
                    // per-slide pass at slide-end) so the entire <p:pic>
                    // including <a:videoFile>/<a:audioFile>/p14:media rel
                    // references survive via add-part + raw-set passthrough.
                    // The typed walk skips them here to avoid double-emit
                    // (EmitPicture would only re-emit the picture shape
                    // without the media wiring).
                    break;
                case "3dmodel":
                case "model3d":
                    // Phase 3c-3d: routed through EmitModel3dForSlide (a
                    // per-slide pass at slide-end) so the <mc:AlternateContent>
                    // wrapper (Choice am3d:model3d + Fallback p:pic) and
                    // the underlying ExtendedPart .glb + thumbnail ImagePart
                    // survive via add-part + raw-set passthrough. The typed
                    // walk skips them here to avoid double-emit.
                    break;
                case "ole":
                case "zoom":
                    // PR3+ scope. ProbeUnsupportedOnSlide already records the
                    // OLE markers via raw-XML sniff; this branch catches
                    // the children that surfaced via the typed Get tree
                    // (when NodeBuilder learns to tag them).
                    ctx.Unsupported.Add(new UnsupportedWarning(
                        Element: child.Type ?? "unknown",
                        SlidePath: slidePath,
                        Reason: "deferred to later PR"));
                    break;
                default:
                    ctx.Unsupported.Add(new UnsupportedWarning(
                        Element: child.Type ?? "unknown",
                        SlidePath: slidePath,
                        Reason: "unrecognized child type"));
                    break;
            }
        }

        // Raw-XML passthrough for exotic transition / timing content. Emitted
        // AFTER all shape/animation rows so they replace anything the semantic
        // emit produced (defensive — slideProps already stripped, animIndex
        // already nulled, but raw-set is the authoritative payload).
        // Append into /p:sld preserves OOXML schema order because we removed
        // the corresponding props upstream: the slide carries neither
        // <p:transition> nor <p:timing> at this point in replay.
        if (exotic.HasExoticTransition && exotic.TransitionXml != null)
            EmitRawSlideSlice(slidePath, "p:transition", exotic.TransitionXml, items, ctx);
        if (exotic.HasExoticTiming && exotic.TimingXml != null)
            EmitRawSlideSlice(slidePath, "p:timing", exotic.TimingXml, items, ctx);

        // SmartArt graphicFrames live in /p:sld/p:cSld/p:spTree but are
        // skipped by NodeBuilder (table/chart-only routing). Phase 3b emits
        // them as add-part smartart (creates the four diagram sub-parts with
        // caller-pinned rIds) followed by raw-set rows that fill each part's
        // XML, and a final raw-set append on /p:sld/p:cSld/p:spTree with the
        // graphicFrame XML. Caller-pinned rIds make the graphicFrame's
        // <dgm:relIds> round-trip byte-equal.
        EmitSmartArtsForSlide(ppt, slideNum, slidePath, items, ctx);

        // Phase 3c-media: video/audio <p:pic> hosts with their underlying
        // MediaDataPart + thumbnail ImagePart, mirroring the SmartArt
        // passthrough. The typed walk skipped video/audio children above.
        EmitMediaForSlide(ppt, slideNum, slidePath, items, ctx);

        // Phase 3c-3d: am3d 3D-model AlternateContent blocks with their
        // underlying ExtendedPart .glb + thumbnail ImagePart, mirroring
        // the video/audio passthrough. The typed walk skipped
        // 3dmodel/model3d children above.
        EmitModel3dForSlide(ppt, slideNum, slidePath, items, ctx);

        // Notes body content — stub for PR1. Notes part presence does not
        // surface in the slide subtree's children today (notes live under
        // /slide[N]/notes); PR2 will reach in and emit them.
        EmitNotes(ppt, slidePath, items, ctx);

        // Legacy slide comments — also off the shape tree (SlideCommentsPart).
        // Emit AFTER notes so the per-slide row order is stable: shapes →
        // notes → comments, mirroring how a reader would traverse the slide.
        EmitComments(ppt, slidePath, items, ctx);

        // Modern p188 threaded comments — distinct from legacy p:cm; live in
        // PowerPointCommentPart. Emit after legacy comments to keep a stable
        // per-slide row ordering.
        EmitModernComments(ppt, slidePath, items, ctx);
    }

    // Touch the raw slide XML to find content that has no handler vocabulary
    // yet. Each match adds an UnsupportedWarning entry; we never throw.
    private static void ProbeUnsupportedOnSlide(PowerPointHandler ppt, string slidePath,
                                                SlideEmitContext ctx)
    {
        string xml;
        try { xml = ppt.Raw(slidePath); }
        catch { return; }

        // <p:timing> = slide animation. EmitAnimationsForShape now emits the
        // entrance/exit/emphasis effects per shape via the `animation` Query
        // surface, so the timing tree no longer aborts to an unsupported
        // warning. Exotic timing constructs (motion paths, sequence groupings)
        // still go through the Query — animations the Query doesn't enumerate
        // are silently dropped.

        // SmartArt sits inside a graphicFrame as a dgm:relIds element.
        // Phase 3b: handled by EmitSmartArtsForSlide via add-part smartart +
        // raw-set passthrough; no warning is raised when we can extract the
        // four diagram parts. If extraction fails the SmartArt emit silently
        // falls back to a missing slice — caller sees a degraded slide but
        // no crash.

        // OLE / video / audio / 3D — element names are distinctive enough.
        if (xml.Contains("<p:oleObj", StringComparison.Ordinal))
            ctx.Unsupported.Add(new UnsupportedWarning("oleObj", slidePath,
                "embedded OLE object present"));
        // Phase 3c-media: video/audio <p:pic> hosts round-trip via
        // EmitMediaForSlide (add-part + raw-set). No probe warning here —
        // even if the slide carries a <p:video>/<p:audio> timing node, the
        // <p:pic> shape itself surfaces in the typed Get tree as
        // child.Type = "video"/"audio" and is now handled.
        // Phase 3c-3d: am3d:model3d AlternateContent blocks round-trip via
        // EmitModel3dForSlide (add-part model3d + raw-set). No probe warning
        // is raised — the slice owns the entire emit. The legacy <p:model3d>
        // bare form never appears in PowerPoint-authored decks, so we no
        // longer match it either.

        // Exotic transitions (morph, p15:prstTrans gallery, p14:* like flip/
        // gallery/conveyor) and exotic animation timing (motion paths,
        // sequence groupings) now round-trip via a raw-set passthrough on the
        // <p:transition> / <p:timing> elements — see ScanSlideExoticContent
        // and EmitRawSlideSlice. No UnsupportedWarning is raised for them
        // because the slice is emitted verbatim. The warning is reserved for
        // cases where the slice itself cannot be canonicalised (handled in
        // EmitRawSlideSlice).
    }

    /// <summary>
    /// Result of scanning a slide's raw XML for content that the semantic
    /// emit path cannot reproduce. Both transition and timing fields are
    /// null when the slide carries only vanilla content (or none).
    /// </summary>
    private readonly record struct SlideExoticContent(
        bool HasExoticTransition, string? TransitionXml,
        bool HasExoticTiming, string? TimingXml);

    private static SlideExoticContent ScanSlideExoticContent(PowerPointHandler ppt, string slidePath)
    {
        string xml;
        try { xml = ppt.Raw(slidePath); }
        catch { return default; }

        string? transXml = null;
        bool transExotic = false;
        var tIdx = xml.IndexOf("<p:transition", StringComparison.Ordinal);
        if (tIdx >= 0)
        {
            // Capture the full <p:transition>...</p:transition> slice, OR the
            // self-closing form. Also include any enclosing
            // <mc:AlternateContent> wrapper because morph/p14/p15 transitions
            // live INSIDE that wrapper (the typed <p:transition> sits as the
            // mc:Fallback child); the wrapper is the natural replace target.
            var mcWrapStart = xml.LastIndexOf("<mc:AlternateContent", tIdx, StringComparison.Ordinal);
            // mcWrapStart is valid only if its closing </mc:AlternateContent>
            // tag lies after tIdx (i.e. <p:transition> is nested inside it).
            int sliceStart, sliceEnd;
            if (mcWrapStart >= 0)
            {
                var mcWrapEnd = xml.IndexOf("</mc:AlternateContent>", mcWrapStart, StringComparison.Ordinal);
                if (mcWrapEnd > tIdx)
                {
                    sliceStart = mcWrapStart;
                    sliceEnd = mcWrapEnd + "</mc:AlternateContent>".Length;
                }
                else
                {
                    sliceStart = tIdx;
                    sliceEnd = SliceEnd(xml, tIdx, "p:transition");
                }
            }
            else
            {
                sliceStart = tIdx;
                sliceEnd = SliceEnd(xml, tIdx, "p:transition");
            }
            if (sliceEnd > sliceStart)
            {
                var slice = xml.Substring(sliceStart, sliceEnd - sliceStart);
                // Exotic markers: any markup outside the plain <p:transition>
                // grammar — namespaces other than p:/a: under the transition
                // tree, mc:AlternateContent wrapping, p14/p15/p159 extension
                // elements. Vanilla fade/push/wipe/cut/cover/cut/etc. that the
                // semantic `transition=` prop already round-trips through
                // ReadSlideTransition NEVER carry these markers, so the
                // semantic path remains authoritative for them.
                if (slice.Contains("mc:AlternateContent", StringComparison.Ordinal)
                    || slice.Contains("p159:", StringComparison.Ordinal)
                    || slice.Contains("p15:", StringComparison.Ordinal)
                    || slice.Contains("p14:", StringComparison.Ordinal))
                {
                    transExotic = true;
                    transXml = slice;
                }
            }
        }

        string? timingXml = null;
        bool timingExotic = false;
        var pIdx = xml.IndexOf("<p:timing", StringComparison.Ordinal);
        if (pIdx >= 0)
        {
            var sliceEnd = SliceEnd(xml, pIdx, "p:timing");
            if (sliceEnd > pIdx)
            {
                var slice = xml.Substring(pIdx, sliceEnd - pIdx);
                // Motion paths surface as presetClass="path"; sequence
                // groupings beyond the per-shape entrance/exit/emphasis tree
                // that Query enumerates show up as <p:tnLst> nested under
                // <p:par> with no presetID anchor that we currently parse,
                // OR as <p:set>/<p:anim>/<p:animMotion>/<p:animRot>/etc.
                // direct timing-effect nodes which BuildSlideAnimationIndex
                // doesn't materialise. The cheapest precise signal is
                // presetClass="path" (motion path) OR any <p:animMotion>
                // element OR a presetClass we don't enumerate.
                // Precise signal: `presetClass="path"` flags motion-path
                // animations (the Query selector "animation" excludes
                // presetClass=motion/path entirely, so they vanish under the
                // semantic emit). `<p:animMotion>` is the lower-level
                // OOXML element a motion-path expands to but rarely appears
                // without the presetClass marker on the enclosing effect.
                // Other p:anim* variants (animScale/animRot/animClr) are
                // how the SDK implements ordinary zoom/spin/colorChange
                // EMPHASIS effects that the Query DOES enumerate via
                // PopulateAnimationNode — flagging those would force every
                // emphasis slide through raw-set and break the basic
                // animation round-trip.
                if (slice.Contains("presetClass=\"path\"", StringComparison.Ordinal)
                    || slice.Contains("<p:animMotion", StringComparison.Ordinal))
                {
                    timingExotic = true;
                    timingXml = slice;
                }
            }
        }

        return new SlideExoticContent(transExotic, transXml, timingExotic, timingXml);
    }

    // Normalize a slide raw slice into a stable textual form so the first-pass
    // (clean source XML) and second-pass (post-SDK-round-trip) produce
    // byte-identical raw-set rows. The SDK round-trip aggressively rewrites
    // ambient prefixes: it may render <mc:AlternateContent> as <AlternateContent
    // xmlns="…/markup-compatibility/2006"> (default-namespaced) on output even
    // when the inserted source used the prefixed form. This normalizer parses
    // the slice, forces every element in the four ambient pptx namespaces
    // (p, a, r, mc) onto its canonical prefix, then strips redundant xmlns
    // decls. The result compares equal across rounds regardless of which
    // serialization the SDK picked.
    private static string NormalizeSlideRawSlice(string sliceXml)
    {
        if (string.IsNullOrEmpty(sliceXml) || !sliceXml.StartsWith("<")) return sliceXml;
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(sliceXml);
            if (doc.Root == null) return sliceXml;
            var ambient = new (string Prefix, System.Xml.Linq.XNamespace Ns)[]
            {
                ("p",  "http://schemas.openxmlformats.org/presentationml/2006/main"),
                ("a",  "http://schemas.openxmlformats.org/drawingml/2006/main"),
                ("r",  "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
                ("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
            };
            // Force the root to carry the canonical prefix decls for any
            // ambient namespace it uses on itself or its descendants. We do
            // not strip non-ambient decls (e.g. p159, p14, p15) since those
            // are extension namespaces specific to this slice and must
            // travel with it.
            foreach (var (prefix, ns) in ambient)
            {
                bool used = doc.Root.DescendantsAndSelf().Any(e => e.Name.Namespace == ns)
                            || doc.Root.DescendantsAndSelf().SelectMany(e => e.Attributes())
                                .Any(a => !a.IsNamespaceDeclaration && a.Name.Namespace == ns);
                if (!used) continue;
                // Remove any default-namespace decls pointing at this ambient
                // namespace, anywhere in the tree — they will be supplanted
                // by the prefixed form on the root.
                foreach (var el in doc.Root.DescendantsAndSelf().ToList())
                {
                    var toRemove = el.Attributes()
                        .Where(a => a.IsNamespaceDeclaration && a.Value == ns.NamespaceName)
                        .ToList();
                    foreach (var a in toRemove) a.Remove();
                }
                // Stamp the canonical prefix decl onto the root.
                doc.Root.SetAttributeValue(System.Xml.Linq.XNamespace.Xmlns + prefix, ns.NamespaceName);
            }
            // Drop redundant prefix decls on descendants that match the root's
            // (mirrors CanonicalizeRawXml but on the post-rewrite tree).
            var rootDecls = doc.Root.Attributes()
                .Where(a => a.IsNamespaceDeclaration)
                .ToDictionary(a => a.Name, a => a.Value);
            foreach (var desc in doc.Root.Descendants())
            {
                var dups = desc.Attributes()
                    .Where(a => a.IsNamespaceDeclaration
                                && rootDecls.TryGetValue(a.Name, out var v) && v == a.Value)
                    .ToList();
                foreach (var a in dups) a.Remove();
            }
            // Final pass: drop ambient namespace decls from the slice root.
            // They are guaranteed to be in scope at the /p:sld replay site,
            // so keeping them only causes textual drift between rounds (the
            // SDK re-stamps them on read-back, our source-side extraction
            // may not have them at all).
            //
            // We CANNOT use XLinq's RemoveAttributes here naively — XLinq
            // refuses to remove a namespace declaration that is currently
            // in use by the element's own name or attribute names; doing so
            // would silently break the serialization. So we serialize first,
            // THEN textually drop the ambient decls from the root tag.
            var serialized = doc.Root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
            return StripAmbientXmlnsFromRootTag(serialized);
        }
        catch { return sliceXml; }
    }

    private static string StripAmbientXmlnsFromRootTag(string xml)
    {
        if (string.IsNullOrEmpty(xml) || xml[0] != '<') return xml;
        var gtIdx = xml.IndexOf('>');
        if (gtIdx <= 0) return xml;
        var head = xml.Substring(0, gtIdx);
        var tail = xml.Substring(gtIdx);
        // Remove ` xmlns:p="…/presentationml/2006/main"` etc. only when the
        // URI matches the well-known ambient. Other xmlns decls (xmlns:p159,
        // xmlns:p14, …) stay — they are extension-scoped and must travel.
        var ambientUris = new (string Prefix, string Uri)[]
        {
            ("p",  "http://schemas.openxmlformats.org/presentationml/2006/main"),
            ("a",  "http://schemas.openxmlformats.org/drawingml/2006/main"),
            ("r",  "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
            ("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
        };
        foreach (var (prefix, uri) in ambientUris)
        {
            var pat = $" xmlns:{prefix}=\"{uri}\"";
            head = head.Replace(pat, "");
        }
        return head + tail;
    }

    // Strip well-known pptx slide ambient namespace declarations from the
    // ROOT of a slice destined for raw-set into /p:sld. The slide root
    // always declares xmlns:p, xmlns:a, xmlns:r, xmlns:mc — the SDK's
    // round-trip serialization stamps them onto every direct child of the
    // slide root, so a slice extracted from the source's raw XML (no
    // root-level decls) and a slice extracted from the post-replay raw XML
    // (root-level decls present, courtesy of the SDK) would not compare
    // equal under CanonicalizeRawXml alone — that helper only strips
    // descendant-vs-root duplicates, never the root's own ambient decls.
    private static string StripSlideAmbientXmlns(string xml)
    {
        if (string.IsNullOrEmpty(xml) || !xml.StartsWith("<")) return xml;
        try
        {
            var doc = System.Xml.Linq.XDocument.Parse(xml);
            if (doc.Root == null) return xml;
            // Match the slide root's known ambient namespaces. Any
            // declaration on the slice root pointing at one of these is
            // redundant once the slice is appended under /p:sld.
            var ambient = new Dictionary<string, string>
            {
                ["p"]  = "http://schemas.openxmlformats.org/presentationml/2006/main",
                ["a"]  = "http://schemas.openxmlformats.org/drawingml/2006/main",
                ["r"]  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                ["mc"] = "http://schemas.openxmlformats.org/markup-compatibility/2006",
            };
            var toRemove = doc.Root.Attributes()
                .Where(a => a.IsNamespaceDeclaration
                            && ambient.TryGetValue(a.Name.LocalName, out var u)
                            && u == a.Value)
                .ToList();
            foreach (var a in toRemove) a.Remove();
            return doc.Root.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        catch { return xml; }
    }

    private static int SliceEnd(string xml, int start, string localName)
    {
        // Find the end of the element starting at `start`. Handles both
        // self-closing form (`<p:transition .../>`) and paired form
        // (`<p:transition ...> ... </p:transition>`).
        var gtIdx = xml.IndexOf('>', start);
        if (gtIdx < 0) return -1;
        // Self-closing: char before '>' is '/'.
        if (gtIdx > start && xml[gtIdx - 1] == '/')
            return gtIdx + 1;
        var closeTag = $"</{localName}>";
        var closeIdx = xml.IndexOf(closeTag, gtIdx, StringComparison.Ordinal);
        if (closeIdx < 0) return -1;
        return closeIdx + closeTag.Length;
    }

    // SmartArt passthrough: per slide, scan for <p:graphicFrame> hosts that
    // carry <dgm:relIds>; emit an `add-part smartart` row that creates the
    // four diagram sub-parts (data/layout/colors/quickStyle) under the
    // slide with the SOURCE's rIds pinned via --prop. Then emit four
    // raw-set replace rows (one per diagram part) and one raw-set append
    // on /p:sld/p:cSld/p:spTree carrying the graphicFrame XML verbatim.
    //
    // rId stability: pinning the rIds on add-part makes the graphicFrame's
    // <dgm:relIds dm=... lo=... cs=... qs=...> attributes resolve to the
    // same diagram parts after replay. Without pinning, AddNewPart<T>()
    // would allocate rId{slide+K} sequentially which would NOT match the
    // source's rIds and the SDK's serialized graphicFrame would drift on
    // re-emit.
    //
    // Diagram part XML canonicalization is the same shape-stripping/canon
    // pass as the slide-slice path; both passes need to be idempotent so
    // first emit (raw XML from source) and second emit (raw XML from
    // post-replay SDK-roundtripped doc) compare byte-equal.
    private static void EmitSmartArtsForSlide(PowerPointHandler ppt, int slideNum,
                                              string slidePath, List<BatchItem> items,
                                              SlideEmitContext ctx)
    {
        IReadOnlyList<PowerPointHandler.SmartArtInfo> smartArts;
        try { smartArts = ppt.GetSmartArtsOnSlide(slideNum); }
        catch { return; }
        if (smartArts.Count == 0) return;

        foreach (var sa in smartArts)
        {
            // add-part smartart with pinned rIds. Props carry the source's
            // rIds; replay's AddPart calls AddNewPart<T>(rId) for each.
            items.Add(new BatchItem
            {
                Command = "add-part",
                Parent = slidePath,
                Type = "smartart",
                Props = new Dictionary<string, string>(StringComparer.Ordinal)
                {
                    ["data"] = sa.DataRelId,
                    ["layout"] = sa.LayoutRelId,
                    ["colors"] = sa.ColorsRelId,
                    ["quickStyle"] = sa.QuickStyleRelId,
                },
            });

            // Resolve each rId to its part URI for raw-set targeting.
            // The post-replay file will have the same URIs because the
            // SlidePart's part-name allocator is deterministic for a
            // freshly created sub-part (e.g. /ppt/diagrams/data1.xml).
            string? dUri = ppt.GetSmartArtPartUri(slideNum, sa.DataRelId);
            string? lUri = ppt.GetSmartArtPartUri(slideNum, sa.LayoutRelId);
            string? cUri = ppt.GetSmartArtPartUri(slideNum, sa.ColorsRelId);
            string? qUri = ppt.GetSmartArtPartUri(slideNum, sa.QuickStyleRelId);
            if (dUri == null || lUri == null || cUri == null || qUri == null)
            {
                ctx.Unsupported.Add(new UnsupportedWarning(
                    Element: "smartArt", SlidePath: slidePath,
                    Reason: "SmartArt diagram part URIs could not be resolved; graphicFrame appended without populated parts"));
            }
            else
            {
                EmitDiagramPart(dUri, "dgm:dataModel", sa.DataXml, items);
                EmitDiagramPart(lUri, "dgm:layoutDef", sa.LayoutXml, items);
                EmitDiagramPart(cUri, "dgm:colorsDef", sa.ColorsXml, items);
                EmitDiagramPart(qUri, "dgm:styleDef", sa.QuickStyleXml, items);
            }

            // Append the graphicFrame into /p:sld/p:cSld/p:spTree. The
            // slice carries the <dgm:relIds> with the source's rIds, which
            // resolve to the just-created diagram parts via the pinned rIds.
            string gfCanon;
            try { gfCanon = NormalizeSlideRawSlice(sa.GraphicFrameXml); }
            catch { gfCanon = sa.GraphicFrameXml; }
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = slidePath,
                Xpath = "/p:sld/p:cSld/p:spTree",
                Action = "append",
                Xml = gfCanon,
            });
        }
    }

    private static void EmitDiagramPart(string partUri, string rootName,
                                        string sliceXml, List<BatchItem> items)
    {
        // Canonicalize for round-trip stability: same canonicalizer as the
        // slide slice path. The diagram parts only carry dgm: / a: ambient
        // ns most of the time; NormalizeSlideRawSlice's ambient set covers
        // a:/r:/mc: which is a superset. Extension prefixes specific to the
        // part travel verbatim.
        string canon;
        try { canon = NormalizeSlideRawSlice(sliceXml); }
        catch { canon = sliceXml; }
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = partUri,
            Xpath = "/" + rootName,
            Action = "replace",
            Xml = canon,
        });
    }

    private static void EmitRawSlideSlice(string slidePath, string localName,
                                          string sliceXml, List<BatchItem> items,
                                          SlideEmitContext ctx)
    {
        // The replay target's freshly-added /slide[N] has no <p:transition>
        // and no <p:timing> (we stripped the semantic props upstream), so
        // raw-set "replace" against `/p:sld/p:transition` would fail with
        // "XPath matched no elements". Use append on /p:sld instead — the
        // OOXML schema order (cSld → clrMapOvr → transition → timing) is
        // preserved because we always emit transition before timing, and
        // neither was present before this append.
        string canon;
        try { canon = NormalizeSlideRawSlice(sliceXml); }
        catch
        {
            ctx.Unsupported.Add(new UnsupportedWarning(
                Element: localName,
                SlidePath: slidePath,
                Reason: "raw slice could not be canonicalised; element dropped"));
            return;
        }
        if (string.IsNullOrEmpty(canon))
        {
            ctx.Unsupported.Add(new UnsupportedWarning(
                Element: localName,
                SlidePath: slidePath,
                Reason: "raw slice canonicalised to empty; element dropped"));
            return;
        }
        items.Add(new BatchItem
        {
            Command = "raw-set",
            Part = slidePath,
            Xpath = "/p:sld",
            Action = "append",
            Xml = canon,
        });
    }

    // Emit one `add animation` BatchItem per effect attached to this shape.
    // Replay parent is the shape's positional path in the emitted document
    // (caller-supplied — must match the just-emitted `add shape/placeholder`).
    //
    // Previously animations were caught by ProbeUnsupportedOnSlide and surfaced
    // only as a warning, so dump→batch→replay lost every entrance/exit/emphasis
    // effect plus its trigger/delay/duration. The animation Query surface
    // already produces fine-grained nodes (effect/class/trigger/duration/delay/
    // direction/easein/easeout via PopulateAnimationNode); this helper just
    // forwards each animation's emittable props as an `add animation` row.
    //
    // Direction was added to PopulateAnimationNode in this same change — without
    // it, fly-down would round-trip as fly-up (AddAnimation default).
    //
    // Motion-path animations are excluded by the Query (presetClass="motion"
    // never surfaces under selector "animation"). Other exotic timing
    // constructs (sequence groupings, conditional triggers) are silently
    // dropped — the visible effects round-trip.
    // Per-shape animation emit. Accepts a pre-filtered list of animation
    // nodes whose shape segment matches this shape (resolved by the caller
    // via the @id → positional map built from fullSlide.Children).
    private static void EmitAnimationsForShape(List<DocumentNode> animsForShape,
                                               string replayShapePath, List<BatchItem> items)
    {
        foreach (var anim in animsForShape)
        {
            var animProps = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            // Map Format keys → AddAnimation accepted keys. presetId is
            // derived from effect+class on Add, so emitting it would either
            // be ignored or trigger an unsupported_property warning.
            foreach (var (k, v) in anim.Format)
            {
                if (v == null) continue;
                if (k.Equals("presetId", StringComparison.OrdinalIgnoreCase)) continue;
                var s = v.ToString() ?? "";
                if (s.Length == 0) continue;
                animProps[k] = s;
            }
            if (animProps.Count == 0) continue;

            items.Add(new BatchItem
            {
                Command = "add",
                Parent = replayShapePath,
                Type = "animation",
                Props = animProps,
            });
        }
    }

    // Build a map from source @id (or source positional) to the list of
    // animation nodes on that shape. Query("animation") paths use either
    // /slide[N]/shape[@id=X]/animation[A] or /slide[N]/shape[K]/animation[A]
    // depending on whether cNvPr.Id is present.
    private static Dictionary<string, List<DocumentNode>> BuildSlideAnimationIndex(
        PowerPointHandler ppt, int slideNum)
    {
        var map = new Dictionary<string, List<DocumentNode>>(StringComparer.Ordinal);
        List<DocumentNode> all;
        try { all = ppt.Query("animation"); }
        catch { return map; }

        var slidePrefix = $"/slide[{slideNum}]/";
        var rx = new System.Text.RegularExpressions.Regex(
            @"^/slide\[\d+\]/shape\[([^\]]+)\]/animation\[\d+\]$");
        foreach (var anim in all)
        {
            if (!anim.Path.StartsWith(slidePrefix, StringComparison.Ordinal)) continue;
            var m = rx.Match(anim.Path);
            if (!m.Success) continue;
            var key = m.Groups[1].Value; // either "5" (positional) or "@id=10" form
            if (!map.TryGetValue(key, out var list))
            {
                list = new List<DocumentNode>();
                map[key] = list;
            }
            list.Add(anim);
        }
        return map;
    }

    // Resolve the animation list for the shape currently being emitted.
    // child.Format["id"] (when present) maps to @id=X; otherwise positional.
    private static List<DocumentNode> GetAnimationsForChild(
        Dictionary<string, List<DocumentNode>> map, DocumentNode child, int sourcePositional)
    {
        // Try @id= form first when child carries id.
        if (child.Format.TryGetValue("id", out var cidObj) && cidObj != null)
        {
            var idKey = $"@id={cidObj}";
            if (map.TryGetValue(idKey, out var byId)) return byId;
        }
        // Fall back to positional.
        if (map.TryGetValue(sourcePositional.ToString(System.Globalization.CultureInfo.InvariantCulture),
                            out var byPos))
            return byPos;
        return new List<DocumentNode>();
    }
}
