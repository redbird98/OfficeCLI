// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;

namespace OfficeCli.Core.Plugins;

/// <summary>
/// Shared logic for invoking exporter plugins (docs/plugin-protocol.md §5.2):
/// resolution, source snapshotting, subprocess invocation, exit-code mapping.
/// Used by `view <file> pdf` and any future caller that needs to convert a
/// native document to a foreign target via an installed exporter plugin.
/// </summary>
public static class ExporterInvoker
{
    public sealed record ExportResult(string OutputPath, ResolvedPlugin Plugin, bool ResidentClosed);

    /// <summary>
    /// Resolve an exporter for (sourceExt, targetExt) and run it. On success,
    /// the target file exists at <paramref name="outPath"/> and the result
    /// reports which plugin handled it. On failure, throws CliException with
    /// an appropriate code (exporter_not_found, plugin_failed, ...).
    ///
    /// If a resident is holding the source file, it's closed first to release
    /// the exclusive lock; <see cref="ExportResult.ResidentClosed"/> indicates
    /// this happened so the caller can surface it to the user.
    /// </summary>
    public static ExportResult Run(string sourceFullPath, string targetExt, string outPath)
    {
        var sourceExt = Path.GetExtension(sourceFullPath).ToLowerInvariant();

        var plugin = Resolve(sourceExt, targetExt)
            ?? throw new CliException($"No exporter plugin found for {sourceExt} → {targetExt}.")
            {
                Code = "exporter_not_found",
                Suggestion = "Install an exporter plugin: `officecli plugins list` to see what's available, or see docs/plugin-protocol.md.",
            };

        bool residentClosed = false;
        if (ResidentClient.TryConnect(sourceFullPath, out _))
        {
            if (ResidentClient.SendCloseWithResponse(sourceFullPath, out _))
                residentClosed = true;
        }

        var pluginSource = SnapshotForExport(sourceFullPath);
        try
        {
            var (exitCode, stderr) = RunPlugin(plugin.ExecutablePath, pluginSource, outPath);
            if (exitCode != 0)
                throw new CliException(
                    $"Exporter plugin '{plugin.Manifest.Name}' failed (exit {exitCode}): {Truncate(stderr, 500)}")
                {
                    Code = exitCode switch
                    {
                        2 => "corrupt_input",
                        3 => "unsupported_feature",
                        4 => "license_expired",
                        5 => "protocol_mismatch",
                        _ => "plugin_failed",
                    },
                };

            if (!File.Exists(outPath))
                throw new CliException(
                    $"Exporter plugin '{plugin.Manifest.Name}' reported success but no output file was written at {outPath}.")
                { Code = "plugin_contract_violation" };

            return new ExportResult(outPath, plugin, residentClosed);
        }
        finally
        {
            try { File.Delete(pluginSource); } catch { }
        }
    }

    /// <summary>
    /// Find an exporter for (source, target). Indexed by target extension (the
    /// plugin's declared extensions field); filtered by source via the manifest's
    /// supports list. A plugin missing supports is assumed to accept all native
    /// sources — conservative default for older manifests.
    /// </summary>
    public static ResolvedPlugin? Resolve(string sourceExt, string targetExt)
    {
        var p = PluginRegistry.FindFor(PluginKind.Exporter, targetExt);
        if (p is null) return null;

        if (p.Manifest.Supports is null || p.Manifest.Supports.Count == 0)
            return p;

        var sourceBare = sourceExt.TrimStart('.');
        if (p.Manifest.Supports.Any(s =>
                string.Equals(s, $"from:{sourceBare}", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(s, sourceExt, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(s, sourceBare, StringComparison.OrdinalIgnoreCase)))
            return p;

        return null;
    }

    /// <summary>
    /// Copy the source to a fresh tmp file so the exporter plugin can open it
    /// exclusively even when another process (resident, antivirus, ...) holds
    /// the original. Caller is responsible for deleting the result.
    /// </summary>
    private static string SnapshotForExport(string sourcePath)
    {
        var ext = Path.GetExtension(sourcePath);
        var tmp = Path.Combine(Path.GetTempPath(),
            $"officecli-export-{Guid.NewGuid():N}{ext}");
        using var src = new FileStream(sourcePath, FileMode.Open, FileAccess.Read,
            FileShare.ReadWrite | FileShare.Delete);
        using var dst = new FileStream(tmp, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        src.CopyTo(dst);
        return tmp;
    }

    private static (int exitCode, string stderr) RunPlugin(string exe, string source, string target)
    {
        using var p = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = exe,
                ArgumentList = { "export", source, "--out", target },
                UseShellExecute = false,
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                CreateNoWindow = true,
            }
        };
        p.Start();
        _ = p.StandardOutput.ReadToEndAsync();
        var stderrTask = p.StandardError.ReadToEndAsync();
        p.WaitForExit();
        return (p.ExitCode, stderrTask.Result);
    }

    private static string Truncate(string s, int max) =>
        s.Length <= max ? s : s.Substring(0, max) + "...";
}
