// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using OfficeCli.Core;

namespace OfficeCli.Handlers;

public static partial class PptxBatchEmitter
{
    // CONSISTENCY(picture-inline-base64): mirrors
    // WordBatchEmitter.Paragraph.TryEmitPictureRun — no size threshold, no
    // sidecar file, always emit `src="data:<contentType>;base64,<bytes>"`.
    // A 50MB picture produces a 70MB batch JSON; accepted by design.
    private static void EmitPicture(PowerPointHandler ppt, DocumentNode picNode,
                                    string parentSlidePath, string replayPath,
                                    List<BatchItem> items,
                                    SlideEmitContext ctx)
    {
        var fullPic = ppt.Get(picNode.Path);
        var props = FilterEmittableProps(fullPic.Format);
        DeferSlideJumpLink(props, replayPath, ctx);

        var binary = ppt.GetImageBinary(picNode.Path);
        if (binary.HasValue)
        {
            var (bytes, contentType) = binary.Value;
            props["src"] = $"data:{contentType};base64,{Convert.ToBase64String(bytes)}";
        }
        else
        {
            // No embedded part — picture is unresolvable on round-trip.
            // Drop to an unsupported warning rather than emit a half-row
            // that AddPicture would reject for missing src.
            ctx.Unsupported.Add(new UnsupportedWarning(
                Element: "picture",
                SlidePath: parentSlidePath,
                Reason: "picture has no resolvable embedded image part"));
            return;
        }

        // Drop Get-only diagnostic keys that AddPicture neither expects nor
        // accepts (mirrors docx WordBatchEmitter picture emit).
        props.Remove("id");
        props.Remove("contentType");
        props.Remove("fileSize");
        props.Remove("alt");
        // Re-add alt only if it was the explicit user-set value (not the
        // "(missing)" placeholder PictureToNode stamps in).
        var altRaw = fullPic.Format.TryGetValue("alt", out var av) ? av?.ToString() : null;
        if (!string.IsNullOrEmpty(altRaw) && altRaw != "(missing)")
            props["alt"] = altRaw;

        // Schema declares brightness/contrast/shadow/glow as add:false, set:true
        // on pptx/picture. AddPicture rejects them with UNSUPPORTED on replay
        // and the values are silently lost. Lift them out of the add bag and
        // defer to a follow-up `set` on the same replay path. Mirrors the
        // DeferSlideJumpLink pattern (deferred-set after every add).
        // Hard-coded picture-level drop list — same precedent as `image=true`,
        // `background=image`, `fill=gradient` drops elsewhere in this emitter.
        var deferredEffects = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var key in PictureSetOnlyEffectKeys)
        {
            if (props.TryGetValue(key, out var val))
            {
                deferredEffects[key] = val;
                props.Remove(key);
            }
        }

        items.Add(new BatchItem
        {
            Command = "add",
            Parent = parentSlidePath,
            Type = "picture",
            Props = props.Count > 0 ? props : null,
        });

        if (deferredEffects.Count > 0)
        {
            ctx.DeferredLinks.Add(new BatchItem
            {
                Command = "set",
                Path = replayPath,
                Props = deferredEffects,
            });
        }
    }

    // Picture effect props with schema `add: false, set: true`. Must NOT ride
    // along inside the add picture op props bag — AddPicture rejects them.
    private static readonly HashSet<string> PictureSetOnlyEffectKeys =
        new(StringComparer.OrdinalIgnoreCase)
    {
        "brightness", "contrast", "shadow", "glow",
    };

    // Phase 3c-media. Mirrors EmitSmartArtsForSlide (Phase 3b). Per slide,
    // scan for <p:pic> hosts that carry <a:videoFile> or <a:audioFile>;
    // emit an `add-part video|audio` row that creates the underlying
    // MediaDataPart + Video/AudioReferenceRelationship + MediaReferenceRel
    // + thumbnail ImagePart with SOURCE rIds pinned via --prop. Then emit
    // one raw-set append on /p:sld/p:cSld/p:spTree carrying the <p:pic>
    // XML verbatim — the pinned rIds make the videoFile/audioFile/p14:media/
    // blip references all resolve to the just-created parts.
    //
    // Skipped by the typed walk: the dispatch in EmitSlide's switch routes
    // child.Type == "video"|"audio" away from EmitPicture (which would
    // re-emit a plain picture without the media rels) into a no-op,
    // letting THIS pass own the entire <p:pic> emit.
    //
    // Audit caveat: the SDK's CreateMediaDataPart allocates a URI like
    // /ppt/media/media1.mp4, NOT the source's /media/mediadata.mp4
    // (legacy zip-root layout). The binary content survives byte-equal;
    // the audit's content-loss check is by content hash (see
    // tools/pptx-roundtrip-audit.py).
    internal static void EmitMediaForSlide(PowerPointHandler ppt, int slideNum,
                                           string slidePath, List<BatchItem> items,
                                           SlideEmitContext ctx)
    {
        IReadOnlyList<PowerPointHandler.MediaInfo> medias;
        try { medias = ppt.GetMediaOnSlide(slideNum); }
        catch { return; }
        if (medias.Count == 0) return;

        foreach (var m in medias)
        {
            var partType = m.IsVideo ? "video" : "audio";
            var ridKey   = m.IsVideo ? "video-rid" : "audio-rid";
            var props = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["data"] = Convert.ToBase64String(m.MediaBytes),
                ["content-type"] = m.MediaContentType,
                ["extension"] = m.MediaExtension,
                ["thumbnail-data"] = Convert.ToBase64String(m.ThumbnailBytes),
                ["thumbnail-content-type"] = m.ThumbnailContentType,
                [ridKey] = m.LinkRelId,
                ["media-rid"] = m.MediaEmbedRelId,
                ["thumbnail-rid"] = m.ThumbnailRelId,
            };
            items.Add(new BatchItem
            {
                Command = "add-part",
                Parent = slidePath,
                Type = partType,
                Props = props,
            });

            // Append the <p:pic> verbatim into the spTree. Canonicalise
            // via the slide-slice canonicaliser so post-replay re-emit
            // hits byte-equal (same trick SmartArt uses).
            string picCanon;
            try { picCanon = NormalizeSlideRawSlice(m.PicXml); }
            catch { picCanon = m.PicXml; }
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = slidePath,
                Xpath = "/p:sld/p:cSld/p:spTree",
                Action = "append",
                Xml = picCanon,
            });
        }
    }

    // Phase 3c-3d. Mirrors EmitMediaForSlide (Phase 3c-media). Per slide,
    // scan for <mc:AlternateContent> blocks whose <mc:Choice Requires="am3d">
    // carries <am3d:model3d>; emit an `add-part model3d` row that creates
    // the underlying ExtendedPart (.glb via AddExtendedPart, since the SDK
    // has no typed Model3DPart) + thumbnail ImagePart with SOURCE rIds
    // pinned via --prop. Then emit a raw-set append on
    // /p:sld/p:cSld/p:spTree carrying the AlternateContent XML verbatim —
    // the pinned rIds make the am3d:model3d r:embed, am3d:raster's
    // am3d:blip r:embed, AND the Fallback p:pic's a:blip r:embed all
    // resolve to the just-created parts.
    //
    // Slice handling: the AlternateContent block is canonicalised via the
    // same NormalizeSlideRawSlice pass as smartart / media. The am3d
    // namespace is NOT in the ambient list (p / a / r / mc), so its
    // xmlns:am3d declaration travels with the slice on the <mc:Choice>
    // — exactly the source form, so first- and post-replay rounds match.
    internal static void EmitModel3dForSlide(PowerPointHandler ppt, int slideNum,
                                              string slidePath, List<BatchItem> items,
                                              SlideEmitContext ctx)
    {
        IReadOnlyList<PowerPointHandler.Model3dInfo> models;
        try { models = ppt.GetModel3dOnSlide(slideNum); }
        catch { return; }
        if (models.Count == 0) return;

        foreach (var m in models)
        {
            var props = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["data"] = Convert.ToBase64String(m.Model3dBytes),
                ["content-type"] = m.Model3dContentType,
                ["extension"] = m.Model3dExtension,
                ["model3d-rid"] = m.Model3dRelId,
                ["thumbnail-data"] = Convert.ToBase64String(m.ThumbnailBytes),
                ["thumbnail-content-type"] = m.ThumbnailContentType,
                ["thumbnail-rid"] = m.ThumbnailRelId,
            };
            items.Add(new BatchItem
            {
                Command = "add-part",
                Parent = slidePath,
                Type = "model3d",
                Props = props,
            });

            string acCanon;
            try { acCanon = NormalizeSlideRawSlice(m.AlternateContentXml); }
            catch { acCanon = m.AlternateContentXml; }
            items.Add(new BatchItem
            {
                Command = "raw-set",
                Part = slidePath,
                Xpath = "/p:sld/p:cSld/p:spTree",
                Action = "append",
                Xml = acCanon,
            });
        }
    }
}
