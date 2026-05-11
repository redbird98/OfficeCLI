// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

namespace OfficeCli.Core;

/// <summary>
/// Central catalogue of <c>view issues --type</c> accepted values. Single
/// source of truth so the CLI front-end (CommandBuilder.View) and the
/// resident server (ResidentServer.ExecuteView) reject typos identically
/// and the cross-handler protocol documentation cannot drift from what the
/// validator actually accepts.
/// </summary>
public static class IssueSubtypes
{
    public const string FormulaNotEvaluated = "formula_not_evaluated";
    public const string FormulaCacheStale = "formula_cache_stale";
    public const string FormulaRefMissingSheet = "formula_ref_missing_sheet";
    public const string FieldNotEvaluated = "field_not_evaluated";
    public const string FieldCacheStale = "field_cache_stale";
    public const string SlideFieldNotEvaluated = "slide_field_not_evaluated";
    public const string ChartSeriesRefMissingSheet = "chart_series_ref_missing_sheet";
    public const string ChartCacheStale = "chart_cache_stale";
    public const string DefinedNameBroken = "definedname_broken";
    public const string DefinedNameTargetMissing = "definedname_target_missing";

    /// <summary>Broad IssueType bucket names — the canonical surface shown
    /// in error messages and help. Single-letter aliases (<see cref="BucketAliases"/>)
    /// are accepted by Validate but kept out of the user-facing list so the
    /// canonical-vs-alias distinction is visible.</summary>
    public static readonly string[] BucketNames =
        new[] { "format", "content", "structure" };

    /// <summary>Single-letter aliases accepted in addition to the canonical
    /// bucket names. Kept separate from <see cref="BucketNames"/> so error
    /// listings don't expose them as first-class values.</summary>
    public static readonly string[] BucketAliases =
        new[] { "f", "c", "s" };

    /// <summary>Combined accepted bucket inputs (canonical + aliases).</summary>
    public static readonly string[] ValidBuckets =
        BucketNames.Concat(BucketAliases).ToArray();

    /// <summary>Every subtype the <c>view issues</c> filter accepts by name.</summary>
    public static readonly string[] ValidSubtypes = new[]
    {
        FormulaNotEvaluated, FormulaCacheStale, FormulaRefMissingSheet,
        FieldNotEvaluated, FieldCacheStale,
        SlideFieldNotEvaluated,
        ChartSeriesRefMissingSheet, ChartCacheStale,
        DefinedNameBroken, DefinedNameTargetMissing,
    };

    /// <summary>
    /// Validate a user-supplied <c>--type</c> argument and return the
    /// canonicalised form. Null, empty, and whitespace-only inputs are
    /// normalised to null (treated as "no filter"). Surrounding whitespace
    /// is trimmed so values copied from shells with extra spaces still
    /// match. Recognised buckets and subtypes (case-insensitive) pass
    /// through unchanged. Anything else raises <see cref="CliException"/>
    /// with the full valid list — turning silent typos into a clear
    /// failure on both the CLI front-end and the resident-server fan-out.
    /// </summary>
    public static string? Validate(string? issueType)
    {
        if (string.IsNullOrWhiteSpace(issueType)) return null;
        var trimmed = issueType.Trim();
        var canonical = trimmed.ToLowerInvariant();
        foreach (var v in ValidBuckets) if (v == canonical) return trimmed;
        foreach (var v in ValidSubtypes) if (v == canonical) return trimmed;
        var all = ValidBuckets.Concat(ValidSubtypes).ToArray();
        throw new CliException(
            $"Invalid --type value: '{issueType}'. Valid buckets: {string.Join(", ", BucketNames)} (alias {string.Join(", ", BucketAliases)}). Valid subtypes: {string.Join(", ", ValidSubtypes)}.")
        { Code = "invalid_issue_type", ValidValues = all };
    }
}
