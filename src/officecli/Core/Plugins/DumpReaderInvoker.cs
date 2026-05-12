// Copyright 2025 OfficeCLI (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using System.Diagnostics;
using System.Text;
using System.Text.Json;
using OfficeCli.Handlers;

namespace OfficeCli.Core.Plugins;

/// <summary>
/// Runs a dump-reader plugin per docs/plugin-protocol.md §5.1. The plugin reads
/// a foreign source file (e.g. .doc) and emits a BatchItem[] JSON array on
/// stdout describing the document as a sequence of officecli commands. Main
/// creates a blank .docx scratch file, replays the batch against it, and
/// returns the populated path. Callers open that path as a normal .docx.
///
/// The conversion is one-shot: edits to the returned .docx are not propagated
/// back to the source file.
/// </summary>
public static class DumpReaderInvoker
{
    public sealed record DumpResult(string ConvertedDocxPath, ResolvedPlugin Plugin);

    /// <summary>
    /// Resolve a dump-reader plugin for <paramref name="sourceExt"/>, invoke it
    /// against <paramref name="sourceFullPath"/>, and replay the resulting
    /// command stream into a fresh .docx. Throws CliException on resolution
    /// or invocation failure; otherwise the result references a temp file the
    /// caller must dispose (or leave for OS tmp cleanup).
    /// </summary>
    public static DumpResult Run(string sourceFullPath, string sourceExt)
    {
        var plugin = PluginRegistry.FindFor(PluginKind.DumpReader, sourceExt)
            ?? throw new CliException($"No dump-reader plugin found for {sourceExt}.")
            {
                Code = "dump_reader_not_found",
                Suggestion = "Install a dump-reader plugin (`officecli plugins list` to see installed; docs/plugin-protocol.md for paths).",
            };

        var (exitCode, stdout, stderr) = RunPlugin(plugin.ExecutablePath, sourceFullPath);
        if (exitCode != 0)
            throw new CliException(
                $"Dump-reader plugin '{plugin.Manifest.Name}' failed (exit {exitCode}): {Truncate(stderr, 500)}")
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

        List<BatchItem> items;
        try
        {
            items = JsonSerializer.Deserialize<List<BatchItem>>(stdout, BatchJsonContext.Default.ListBatchItem)
                ?? new List<BatchItem>();
        }
        catch (JsonException ex)
        {
            throw new CliException(
                $"Dump-reader plugin '{plugin.Manifest.Name}' emitted invalid JSON: {ex.Message}")
            { Code = "plugin_contract_violation" };
        }

        var tmpDocx = Path.Combine(Path.GetTempPath(),
            $"officecli-dumpread-{Guid.NewGuid():N}.docx");
        // minimal: true gives a bare-skeleton docx (no Normal style, no theme,
        // no docDefaults). The plugin's batch is expected to define everything
        // it references — round-trip dumps from `officecli dump` do exactly that.
        BlankDocCreator.Create(tmpDocx, locale: null, minimal: true);

        using (var handler = DocumentHandlerFactory.Open(tmpDocx, editable: true))
        {
            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                if (item is null)
                    throw new CliException(
                        $"Dump-reader plugin '{plugin.Manifest.Name}' emitted null at index {i}.")
                    { Code = "plugin_contract_violation" };
                try
                {
                    CommandBuilder.ExecuteBatchItem(handler, item, json: false);
                }
                catch (Exception ex)
                {
                    throw new CliException(
                        $"Dump-reader plugin '{plugin.Manifest.Name}' command #{i} ({item.Command}) failed while replaying: {ex.Message}", ex)
                    { Code = "plugin_command_failed" };
                }
            }
        }

        return new DumpResult(tmpDocx, plugin);
    }

    private static (int exitCode, string stdout, string stderr) RunPlugin(string exe, string source)
    {
        var psi = new ProcessStartInfo
        {
            FileName = exe,
            ArgumentList = { "dump", source },
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
        };

        // Hand the plugin a stable pointer back to the current officecli binary
        // so it can shell out to officecli (e.g. `officecli dump <converted.docx>`)
        // without relying on PATH lookup. Plugins that don't need it can ignore
        // the variable.
        var selfPath = Environment.ProcessPath;
        if (!string.IsNullOrEmpty(selfPath))
            psi.Environment["OFFICECLI_BIN"] = selfPath;

        using var p = new Process { StartInfo = psi };
        p.Start();
        var stdoutTask = p.StandardOutput.ReadToEndAsync();
        var stderrTask = p.StandardError.ReadToEndAsync();
        p.WaitForExit();
        return (p.ExitCode, stdoutTask.Result, stderrTask.Result);
    }

    private static string Truncate(string s, int max) =>
        s.Length <= max ? s : s.Substring(0, max) + "...";
}
