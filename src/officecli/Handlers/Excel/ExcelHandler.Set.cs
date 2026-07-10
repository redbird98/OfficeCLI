// Copyright 2026 OfficeCLI (https://OfficeCLI.AI)
// SPDX-License-Identifier: Apache-2.0

using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeCli.Core;
using Drawing = DocumentFormat.OpenXml.Drawing;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using XDR = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using ThreadedCmt = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;


namespace OfficeCli.Handlers;

public partial class ExcelHandler
{
    public List<string> Set(string path, Dictionary<string, string> properties)
    {
        Modified = true;
        // Batch Set: route to the shared filter engine when the path is a bare
        // selector (no `/`) OR a `/`-scoped path that carries a content filter
        // (e.g. `/Sheet1/cell[value>5000 or value<300]`). The latter would
        // otherwise fall to the positional-index navigator and reject the
        // predicate — query already resolves it, so set must too (parity).
        // Structural `/`-paths (`/Sheet1/A1`, `/Sheet1/row[2]`,
        // `/Sheet1/chart[1]/axis[@role=…]`) stay false in IsContentFilterPath
        // and take the direct dispatch below.
        if (!string.IsNullOrEmpty(path)
            && (!path.StartsWith("/") || AttributeFilter.IsContentFilterPath(path)))
        {
            var unsupported = new List<string>();
            // FilterSelector narrows the mutation set with the same engine `query`
            // uses (CommandBuilder.GetQuery.cs): a pure-AND selector takes the
            // legacy path (handler pre-filter + flat re-apply — so `cell[value>100]`
            // no longer over-matches every cell), and an `or` selector is queried
            // with its filter brackets stripped (Query returns broadly) then
            // narrowed by the boolean expression tree. ResolveCellAttributeAlias
            // maps cell short keys (bold -> font.bold, ...); a no-op on other keys.
            var (targets, _) = AttributeFilter.FilterSelector(path, Query, ResolveCellAttributeAlias);
            if (targets.Count == 0)
                // A selector that resolves to zero rows is an ordinary empty
                // WHERE result, not a tool failure — surface not_found so the
                // JSON envelope reads clean instead of internal_error.
                throw new Core.CliException($"No elements matched selector: {path}") { Code = "not_found" };
            LastSelectorSetCount = targets.Count;
            foreach (var target in targets)
            {
                var targetUnsupported = Set(target.Path, properties);
                foreach (var u in targetUnsupported)
                    if (!unsupported.Contains(u)) unsupported.Add(u);
            }
            return unsupported;
        }

        // Normalize to case-insensitive lookup so camelCase keys match lowercase lookups
        if (properties != null && properties.Comparer != StringComparer.OrdinalIgnoreCase)
            properties = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
        properties ??= new Dictionary<string, string>();

        path = NormalizeExcelPath(path);
        path = ResolveSheetIndexInPath(path);

        // Excel only supports find+replace — reject find without replace early (before path dispatch)
        if (properties.ContainsKey("find") && !properties.ContainsKey("replace"))
            throw new ArgumentException("Excel only supports 'find' with 'replace'. Use 'find' + 'replace' for text replacement. find+format (without replace) is not supported in Excel.");
        // CONSISTENCY(find-regex): Excel find/replace now honours the `r"..."`
        // raw-string regex prefix (see ApplyFindReplace), matching docx/pptx.
        // The legacy `regex=true` flag is normalized to the r"..." form so both
        // spellings work, mirroring the WordHandler.Set.cs find path.
        if (properties.TryGetValue("regex", out var xlRegexFlag)
            && ParseHelpers.IsTruthySafe(xlRegexFlag)
            && properties.TryGetValue("find", out var xlFindRaw)
            && !xlFindRaw.StartsWith("r\"") && !xlFindRaw.StartsWith("r'"))
            properties["find"] = $"r\"{xlFindRaw}\"";

        // Handle root path "/" — document properties
        if (path == "/")
        {
            // Find & Replace: special handling before document properties
            if (properties.TryGetValue("find", out var findText) && properties.TryGetValue("replace", out var replaceText))
            {
                var count = FindAndReplace(findText, replaceText, null);
                LastFindMatchCount = count;
                var remaining = new Dictionary<string, string>(properties, StringComparer.OrdinalIgnoreCase);
                remaining.Remove("find");
                remaining.Remove("replace");
                if (remaining.Count > 0)
                    return Set(path, remaining);
                return [];
            }

            var unsupported = new List<string>();
            var pkg = _doc.PackageProperties;
            foreach (var (key, value) in properties)
            {
                switch (key.ToLowerInvariant())
                {
                    case "title": pkg.Title = value; break;
                    case "author" or "creator": pkg.Creator = value; break;
                    case "subject": pkg.Subject = value; break;
                    case "description": pkg.Description = value; break;
                    case "category": pkg.Category = value; break;
                    case "keywords": pkg.Keywords = value; break;
                    case "lastmodifiedby": pkg.LastModifiedBy = value; break;
                    case "revisionnumber": pkg.Revision = value; break;
                    default:
                        var lowerKey = key.ToLowerInvariant();
                        if (!TrySetWorkbookSetting(lowerKey, value)
                            && !Core.ThemeHandler.TrySetTheme(_doc.WorkbookPart?.ThemePart, lowerKey, value)
                            && !Core.ExtendedPropertiesHandler.TrySetExtendedProperty(
                                Core.ExtendedPropertiesHandler.GetOrCreateExtendedPart(_doc), lowerKey, value))
                            unsupported.Add(key);
                        break;
                }
            }
            return unsupported;
        }

        // Handle /SheetName/sparkline[N]
        var sparklineSetMatch = Regex.Match(path.TrimStart('/'), @"^([^/]+)/sparkline\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (sparklineSetMatch.Success) return SetSparklineByPath(sparklineSetMatch, properties);

        // Handle /namedrange[N] or /namedrange[Name]
        var namedRangeMatch = Regex.Match(path.TrimStart('/'), @"^namedrange\[(.+?)\]$", RegexOptions.IgnoreCase);
        if (namedRangeMatch.Success) return SetNamedRangeByPath(namedRangeMatch, properties);

        // Parse path: /SheetName, /SheetName/A1, /SheetName/A1:D1, /SheetName/col[A], /SheetName/row[1], /SheetName/autofilter
        var segments = path.TrimStart('/').Split('/', 2);
        var sheetName = segments[0];

        var worksheet = FindWorksheet(sheetName);
        if (worksheet == null)
            throw SheetNotFoundException(sheetName);

        // Sheet-level Set (path is just /SheetName)
        if (segments.Length < 2)
        {
            return SetSheetLevel(worksheet, sheetName, properties);
        }

        // BUG-R41-F2: reject cell reference segments that contain control characters
        // (e.g. \n, \r, \t). In .NET, Regex `$` matches before a trailing \n, so
        // without this check "A1\n" would pass ParseCellReference and create a ghost
        // cell with CellReference="A1\n" — an address that never resolves to A1.
        // Reject up-front so the caller gets a clear error instead of silent corruption.
        var cellRef = segments[1];
        if (cellRef.Any(c => c < ' ' && c != '\t' || c == '\x7f'))
            throw new ArgumentException(
                $"Cell reference '{cellRef.Replace("\n", "\\n").Replace("\r", "\\r")}' contains invalid control characters. " +
                $"Expected a clean cell address like 'A1' or 'B2'.");

        // Handle /SheetName/dataValidation[N] (canonical) and
        // /SheetName/validation[N] (legacy alias, R7-bt-6 CONSISTENCY)
        var validationSetMatch = Regex.Match(cellRef, @"^(?:dataValidation|validation)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (validationSetMatch.Success) return SetValidationByPath(validationSetMatch, worksheet, properties);

        // Handle /SheetName/ole[N]
        var oleSetMatch = Regex.Match(cellRef, @"^(?:ole|object|embed)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (oleSetMatch.Success) return SetOleByPath(oleSetMatch, worksheet, properties);

        // Handle /SheetName/picture[N]
        var picSetMatch = Regex.Match(cellRef, @"^picture\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (picSetMatch.Success) return SetPictureByPath(picSetMatch, worksheet, properties);

        // Handle /SheetName/shape[N]
        var shapeSetMatch = Regex.Match(cellRef, @"^shape\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (shapeSetMatch.Success) return SetShapeByPath(shapeSetMatch, worksheet, properties);

        // Handle /SheetName/slicer[N] — caption/style/columnCount/rowHeight/name
        var slicerSetMatch = Regex.Match(cellRef, @"^slicer\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (slicerSetMatch.Success) return SetSlicerByPath(slicerSetMatch, worksheet, properties);

        // Handle /SheetName/table[N]/columns[M] or /SheetName/table[N]/column[M]
        // CONSISTENCY(table-column-path): mirror the col[M].prop= dotted form already
        // accepted on /Sheet/table[N] by exposing the column as a sub-path so users can
        // address it as a node and call Set with a flat property bag.
        var tableColPathMatch = Regex.Match(cellRef,
            @"^table\[(\d+)\]/(?:columns|column)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (tableColPathMatch.Success) return SetTableColumnByPath(tableColPathMatch, worksheet, properties);

        // Handle /SheetName/table[N]
        var tableSetMatch = Regex.Match(cellRef, @"^table\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (tableSetMatch.Success) return SetTableByPath(tableSetMatch, worksheet, properties);

        // Handle /SheetName/comment[N]
        var commentSetMatch = Regex.Match(cellRef, @"^comment\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (commentSetMatch.Success) return SetCommentByPath(commentSetMatch, worksheet, sheetName, properties);

        // Handle /SheetName/autofilter
        if (cellRef.Equals("autofilter", StringComparison.OrdinalIgnoreCase))
        {
            return SetAutoFilter(worksheet, properties);
        }

        // Handle /SheetName/cf[N] or /SheetName/conditionalformatting[N]
        var cfSetMatch = Regex.Match(cellRef, @"^(?:cf|conditionalformatting)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (cfSetMatch.Success) return SetCfByPath(cfSetMatch, worksheet, properties);

        // Handle /SheetName/rowbreak[N] or /SheetName/colbreak[N] — reposition
        // (row/col) and toggle manual. Mirrors the Get readback keys so a queried
        // page break round-trips into set.
        var brkSetMatch = Regex.Match(cellRef, @"^(rowbreak|colbreak)\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (brkSetMatch.Success)
            return SetPageBreak(worksheet,
                brkSetMatch.Groups[1].Value.Equals("rowbreak", StringComparison.OrdinalIgnoreCase),
                int.Parse(brkSetMatch.Groups[2].Value), properties);

        // CONSISTENCY(axis-ref-compat): accept Excel-style whole-column/row
        // references (B:B, B:D, 1:1, 2:5) as input aliases — expand to the
        // canonical col[X]/row[N] segments and apply per column/row, the way
        // range set (A1:D1) applies per cell.
        if (TryExpandAxisRef(cellRef) is { } axisSegments)
        {
            var axisUnsupported = new List<string>();
            foreach (var seg in axisSegments)
            {
                var segResult = Set($"/{sheetName}/{seg}", properties);
                // Intersect: a prop is unsupported only if no column/row took it.
                axisUnsupported = axisSegments[0] == seg
                    ? segResult
                    : axisUnsupported.Intersect(segResult, StringComparer.OrdinalIgnoreCase).ToList();
            }
            return axisUnsupported;
        }

        // Handle /SheetName/col[X] where X is a column letter (A) or numeric index (1)
        var colMatch = Regex.Match(cellRef, @"^col\[([A-Za-z0-9]+)\]$", RegexOptions.IgnoreCase);
        if (colMatch.Success)
        {
            var colValue = colMatch.Groups[1].Value;
            var colName = int.TryParse(colValue, out var colNumIdx) ? IndexToColumnName(colNumIdx) : colValue.ToUpperInvariant();
            return SetColumn(worksheet, colName, properties);
        }

        // Handle /SheetName/row[N]
        var rowMatch = Regex.Match(cellRef, @"^row\[(\d+)\]$");
        if (rowMatch.Success)
        {
            var rowIdx = uint.Parse(rowMatch.Groups[1].Value);
            return SetRow(worksheet, rowIdx, properties);
        }

        // Handle /SheetName/chart[N]/axis[@role=ROLE]
        var chartAxisSetMatch = Regex.Match(cellRef,
            @"^chart\[(\d+)\]/axis\[@role=([a-zA-Z0-9_]+)\]$");
        if (chartAxisSetMatch.Success) return SetChartAxisByPath(chartAxisSetMatch, worksheet, properties);

        // Handle /SheetName/chart[N] or /SheetName/chart[N]/series[K]
        var chartMatch = Regex.Match(cellRef, @"^chart\[(\d+)\](?:/series\[(\d+)\])?$");
        if (chartMatch.Success) return SetChartByPath(chartMatch, worksheet, properties);

        // Handle /SheetName/pivottable[N]
        var pivotSetMatch = Regex.Match(cellRef, @"^pivottable\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (pivotSetMatch.Success) return SetPivotTableByPath(pivotSetMatch, worksheet, properties);

        // Handle /SheetName/A1/run[N] (rich text run)
        var runSetMatch = Regex.Match(cellRef, @"^([A-Z]+\d+)/run\[(\d+)\]$", RegexOptions.IgnoreCase);
        if (runSetMatch.Success) return SetCellRunByPath(runSetMatch, worksheet, properties);

        // Handle /SheetName/A1:D1 (range — merge/unmerge)
        if (cellRef.Contains(':'))
        {
            var firstPartRange = cellRef.Split(':')[0];
            bool isRangeRef = Regex.IsMatch(firstPartRange, @"^[A-Z]+\d+$", RegexOptions.IgnoreCase);
            if (isRangeRef)
            {
                return SetRange(worksheet, cellRef.ToUpperInvariant(), properties);
            }
        }

        // Check if path is a cell reference or generic XML path
        var firstPart = cellRef.Split('/')[0].Split('[')[0];
        bool isCellRef = Regex.IsMatch(firstPart, @"^[A-Z]+\d+", RegexOptions.IgnoreCase);
        if (!isCellRef)
        {
            // Generic XML fallback: navigate to element and set attributes
            var xmlSegments = GenericXmlQuery.ParsePathSegments(cellRef);
            var target = GenericXmlQuery.NavigateByPath(GetSheet(worksheet), xmlSegments);
            if (target == null)
                throw new ArgumentException($"Element not found: {cellRef}");
            var unsup = new List<string>();
            foreach (var (key, value) in properties)
            {
                if (!GenericXmlQuery.SetGenericAttribute(target, key, value))
                    unsup.Add(key);
            }
            SaveWorksheet(worksheet);
            return unsup;
        }

        var sheetData = GetSheet(worksheet).GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            sheetData = new SheetData();
            GetSheet(worksheet).Append(sheetData);
        }

        // Did the cell exist before this Set? If not, FindOrCreateCell
        // materializes it, and a mid-Set validation throw must remove it
        // entirely — restoring an empty clone (below) would leave a ghost
        // <c> stub behind on a failed create (exit 1 must mean no change).
        var cellPreExisted = sheetData.Elements<Row>()
            .SelectMany(r => r.Elements<Cell>())
            .Any(c => string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));

        var cell = FindOrCreateCell(sheetData, cellRef);

        // Clone cell for rollback on failure (atomic: no partial modifications)
        var cellBackup = cell.CloneNode(true);

        try
        {
        return SetCellProperties(cell, cellRef, worksheet, properties);
        }
        catch
        {
            if (cellPreExisted)
                // Rollback: restore cell to pre-modification state.
                cell.Parent?.ReplaceChild(cellBackup, cell);
            else
                // Newly created by this Set — remove it so a failed create
                // leaves no ghost cell.
                cell.Remove();
            throw;
        }
    }

    private List<string> SetCellProperties(Cell cell, string cellRef, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        var unsupported = ApplyCellProperties(cell, cellRef, worksheet, properties);
        // Remove completely empty cells (no value, no formula, no custom style) so that
        // rows with no remaining cells are pruned from XML. This keeps maxRow correct
        // and produces "remove" watch patches instead of "replace" for cleared rows.
        PruneEmptyCell(cell);
        // CONSISTENCY(xlsx/table-autoexpand): eager post-write auto-grow —
        // only fires when the cell still carries a value/formula after prune.
        if (cell.Parent != null && (cell.CellValue != null || cell.CellFormula != null || cell.InlineString != null))
            MaybeExpandTablesForCell(worksheet, cellRef);
        // DATA-CORRUPTION(xlsx/table-header-name): keep <tableColumn name> in
        // sync with a header cell's text after Set — Excel rejects mismatches.
        if (cell.Parent != null)
            MaybeSyncTableHeaderName(worksheet, cellRef);
        // Any mutation to a cell (value, formula, clear) can invalidate the calc chain
        DeleteCalcChainIfPresent();
        SaveWorksheet(worksheet);
        return unsupported;
    }

    private void PruneEmptyCell(Cell cell)
    {
        var hasValue = cell.CellValue != null && !string.IsNullOrEmpty(cell.CellValue.Text);
        var hasFormula = cell.CellFormula != null;
        var hasStyle = cell.StyleIndex != null && cell.StyleIndex.Value != 0;
        if (!hasValue && !hasFormula && !hasStyle)
        {
            var row = cell.Parent as Row;
            cell.Remove();
            if (row != null && !row.Elements<Cell>().Any())
            {
                // Capture sheetData and rowIdx before detaching — row.Parent is null after Remove()
                var sheetData = row.Parent as SheetData;
                var rowIdx = row.RowIndex?.Value;
                row.Remove();
                // Keep row index cache in sync: detached row must not be returned by FindOrCreateRow
                if (sheetData != null && rowIdx.HasValue)
                    _rowIndex?.GetValueOrDefault(sheetData)?.Remove(rowIdx.Value);
            }
        }
    }

    /// <summary>
    /// Remove every &lt;hyperlink&gt; on <paramref name="cellRef"/> AND delete the
    /// external relationship each referenced — unless another surviving
    /// hyperlink still uses the same rId. Without the rel cleanup, repointing
    /// or clearing a cell link leaves an orphan External relationship in
    /// .rels on every edit (a per-file leak invisible to Get).
    /// </summary>
    private static void RemoveCellHyperlinksAndRels(WorksheetPart worksheet, Hyperlinks hyperlinksEl, string cellRef)
    {
        var toRemove = hyperlinksEl.Elements<Hyperlink>()
            .Where(h => h.Reference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true)
            .ToList();
        var removedIds = toRemove.Select(h => h.Id?.Value).Where(id => !string.IsNullOrEmpty(id)).ToList();
        foreach (var h in toRemove) h.Remove();
        foreach (var id in removedIds.Distinct())
        {
            // Skip if a surviving hyperlink still references this rId.
            if (hyperlinksEl.Elements<Hyperlink>().Any(h => h.Id?.Value == id)) continue;
            try { worksheet.DeleteReferenceRelationship(id!); } catch { /* already gone */ }
        }
    }

    /// <summary>Apply cell properties without saving — caller is responsible for SaveWorksheet.</summary>
    private List<string> ApplyCellProperties(Cell cell, WorksheetPart worksheet, Dictionary<string, string> properties)
        => ApplyCellProperties(cell, cell.CellReference?.Value ?? "", worksheet, properties);

    private List<string> ApplyCellProperties(Cell cell, string cellRef, WorksheetPart worksheet, Dictionary<string, string> properties)
    {
        // Separate content props from style props
        var styleProps = new Dictionary<string, string>();
        var unsupported = new List<string>();

        foreach (var (key, value) in properties)
        {
            if (value is null) continue;
            if (ExcelStyleManager.IsStyleKey(key))
            {
                styleProps[key] = value;
                continue;
            }

            switch (key.ToLowerInvariant())
            {
                case "value" or "text":
                    // bt-3: if the cell already carries a text number format
                    // ("@", numFmtId 49) from a prior `set numberformat=@`,
                    // honor it on subsequent value updates by forcing the cell
                    // to String storage. Skip when the user is overriding the
                    // numberformat in this same call (styleProps captures that
                    // path via IsTextNumberFormat already).
                    bool existingIsTextFmt = false;
                    if (!properties.ContainsKey("numberformat")
                        && !properties.ContainsKey("numfmt")
                        && !properties.ContainsKey("format")
                        && !properties.ContainsKey("type"))
                    {
                        var (existingNumFmtId, existingFmtCode) = ExcelDataFormatter.GetCellFormat(cell, _doc.WorkbookPart);
                        if (existingNumFmtId == 49
                            || (existingFmtCode != null && existingFmtCode.Trim() == "@"))
                            existingIsTextFmt = true;
                    }
                    // R28-B4 — leading apostrophe is Excel's "force text" idiom.
                    // Strip the apostrophe from the stored value and stamp
                    // quotePrefix=1 on the cell xf so Excel renders the value
                    // literally as text without the apostrophe glyph. Cell type
                    // is forced to String below via the local quotePrefixForce flag
                    // (we can't safely add to `properties` mid-foreach).
                    bool quotePrefixForce = false;
                    string effectiveValue = value;
                    if (effectiveValue.StartsWith('\'') && effectiveValue.Length > 1)
                    {
                        effectiveValue = effectiveValue.Substring(1);
                        styleProps["quoteprefix"] = "true";
                        quotePrefixForce = true;
                    }
                    // R13-1: enforce Excel's 32767-char per-cell limit.
                    EnsureCellValueLength(effectiveValue, cell.CellReference?.Value);
                    // R13-3: warn if both value= and formula= supplied — formula
                    // takes precedence below (explicit-formula case runs last and
                    // clears CellValue), so the literal value is silently discarded.
                    if (properties.Any(p => p.Key.Equals("formula", StringComparison.OrdinalIgnoreCase)))
                    {
                        Console.Error.WriteLine(
                            "Warning: Both value= and formula= supplied — using formula, value ignored.");
                    }
                    // Auto-detect formula: value starting with '=' is treated as
                    // formula — UNLESS the caller forced text via the leading
                    // apostrophe (quotePrefixForce) or an explicit type=string.
                    // Without the gate, '=TEXT and type=string both got coerced
                    // into a #NAME? formula, and dump→batch replay of any string
                    // cell starting with '=' (which the emitter pins via the
                    // apostrophe idiom) reproduced the corruption.
                    var setForcedString = quotePrefixForce
                        || (properties.TryGetValue("type", out var setTypeVal)
                            && setTypeVal.Equals("string", StringComparison.OrdinalIgnoreCase));
                    if (!setForcedString && effectiveValue.StartsWith('=') && effectiveValue.Length > 1)
                        goto case "formula";
                    // CONSISTENCY(text-escape-boundary): \n / \t resolution is
                    // applied at the CLI --prop parse boundary
                    // (CommandBuilder.ParsePropsArray); the value arrives here
                    // with real newlines/tabs already decoded.
                    var cellValue = effectiveValue;
                    // Warn when overwriting an existing formula with a literal value.
                    // Without this, `set --prop value=N` on a formula cell silently
                    // drops the formula — the same conflict-class as supplying both
                    // value= and formula= in one call (handled above), but split
                    // across two calls. Skipped when formula= is also in this call
                    // (already warned above) and when the literal coerces into a
                    // formula (value starting with '=', which gotos the formula case).
                    if (cell.CellFormula != null
                        && !properties.Any(p => p.Key.Equals("formula", StringComparison.OrdinalIgnoreCase)))
                    {
                        var oldFormula = cell.CellFormula.Text;
                        Console.Error.WriteLine(
                            $"Warning: Cell {cell.CellReference?.Value ?? cellRef} has formula \"={oldFormula}\"; replacing with literal value. Use --prop formula=… to update the formula instead.");
                    }
                    cell.CellFormula = null; // Clear formula when explicit value is set
                    // CONSISTENCY(value-child-uniqueness): drop any stale <is>
                    // inline-string child before writing <v>. Table header cells
                    // carry an <is><t>ColumnN</t></is> placeholder; leaving both
                    // <v> and <is> on one <c> is invalid OOXML and real Excel
                    // refuses to open the file (0x800A03EC).
                    cell.RemoveAllChildren<InlineString>();
                    // If cell is already boolean type, convert true/false to 1/0
                    if (cell.DataType?.Value == CellValues.Boolean)
                    {
                        var bv = cellValue.Trim().ToLowerInvariant();
                        if (bv is "true" or "yes" or "1") cell.CellValue = new CellValue("1");
                        else if (bv is "false" or "no" or "0") cell.CellValue = new CellValue("0");
                        else
                            // A t="b" cell whose value isn't 0/1 makes Excel
                            // refuse the whole file (0x800A03EC). Reject rather
                            // than write garbage into the existing boolean cell.
                            throw new ArgumentException(
                                $"Cannot store '{cellValue}' in a boolean cell; value must be true/false, yes/no, or 1/0. " +
                                "Set type=string first to store literal text.");
                    }
                    else
                    {
                        // Check if user explicitly set type
                        var hasExplicitType = properties.Any(p => p.Key.Equals("type", StringComparison.OrdinalIgnoreCase));
                        var explicitTypeIsString = quotePrefixForce || existingIsTextFmt || (hasExplicitType && properties
                            .Where(p => p.Key.Equals("type", StringComparison.OrdinalIgnoreCase))
                            .Select(p => p.Value?.ToLowerInvariant())
                            .Any(v => v is "string" or "str"));
                        var explicitTypeIsNumber = hasExplicitType && properties
                            .Where(p => p.Key.Equals("type", StringComparison.OrdinalIgnoreCase))
                            .Select(p => p.Value?.ToLowerInvariant())
                            .Any(v => v is "number" or "num");
                        var explicitTypeIsDate = hasExplicitType && properties
                            .Where(p => p.Key.Equals("type", StringComparison.OrdinalIgnoreCase))
                            .Select(p => p.Value?.ToLowerInvariant())
                            .Any(v => v is "date");

                        // BUG-FIX(B10): when caller explicitly says type=date, the
                        // value MUST parse as a real date. Falling through to the
                        // generic else-branch would store an invalid date-shaped
                        // string in a numeric-styled cell. Reject up-front (mirrors
                        // explicitTypeIsNumber's guard against non-numeric input).
                        if (explicitTypeIsDate && !TryParseIsoDateFlexible(cellValue, out _))
                            throw new ArgumentException(
                                $"Cannot store '{cellValue}' as date; value must be ISO 8601 (yyyy-MM-dd) " +
                                $"and represent a real calendar day. Use type=string to keep the literal text.");

                        // Auto-detect ISO date (only if user did NOT explicitly set type=string)
                        // R13-2: accept date-with-time variants (T and space separators).
                        if (!explicitTypeIsString && TryParseIsoDateFlexible(cellValue, out var dt))
                        {
                            // Excel's date serial epoch is 1899-12-30 (preserving the
                            // 1900 leap bug). Dates earlier than that map to negative
                            // serials, which Excel renders as ####### or silently
                            // clamps to the epoch — neither is what the user asked
                            // for. Reject up front so the round-trip is honest
                            // instead of writing serial 0 and reading back "1899-12-30".
                            if (dt < new System.DateTime(1900, 1, 1))
                                throw new ArgumentException(
                                    $"Cannot store '{cellValue}' as date; Excel does not support dates before 1900-01-01 " +
                                    $"(serial epoch is 1899-12-30). Use type=string to keep the literal text.");
                            cell.CellValue = new CellValue(ExcelDataFormatter.ToExcelSerial(dt).ToString(System.Globalization.CultureInfo.InvariantCulture));
                            cell.DataType = null;
                            if (!properties.ContainsKey("numberformat") && !properties.ContainsKey("numfmt") && !properties.ContainsKey("format"))
                                styleProps["numberformat"] = "yyyy-mm-dd";
                        }
                        // Auto-detect strings that look like numbers but should be text
                        else if (!explicitTypeIsNumber
                            && ((cellValue.Length > 1 && cellValue.StartsWith('0') && !cellValue.StartsWith("0.") && !cellValue.StartsWith("0,") && cellValue.All(c => char.IsDigit(c)))
                                || (cellValue.All(char.IsDigit) && cellValue.Length > 15)))
                        {
                            cell.CellValue = new CellValue(cellValue);
                            cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        }
                        else if (explicitTypeIsString)
                        {
                            // R15-2: honor explicit type=string even for
                            // numeric-looking literals. Without this, Excel
                            // renders 123 as a number despite user intent.
                            cell.CellValue = new CellValue(cellValue);
                            cell.DataType = new EnumValue<CellValues>(CellValues.String);
                        }
                        else if (explicitTypeIsNumber)
                        {
                            // R15-2: honor explicit type=number — refuse
                            // non-numeric values rather than silently storing
                            // as string. R32-1: also refuse NaN/Infinity even
                            // though TryParse may accept them — they are not
                            // valid xs:double cell content.
                            if (!double.TryParse(cellValue, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var numDbl)
                                || !double.IsFinite(numDbl))
                                throw new ArgumentException(
                                    $"Cannot store '{cellValue}' as number; use type=string or remove type=");
                            // R-fuzz2-1: TryParse accepts "+5" / padded spellings
                            // Excel's <v> parser rejects — store canonical form.
                            cell.CellValue = new CellValue(NormalizeNumericCellText(cellValue, numDbl));
                            cell.DataType = null;
                        }
                        else
                        {
                            // R32-1: double.TryParse("NaN") returns true; without
                            // an IsFinite gate, the cell would be written with
                            // no t= attribute (numeric default) and content
                            // "NaN", which Excel rejects as invalid xs:double.
                            // Force string storage for non-finite doubles,
                            // matching how "Infinity" already behaves.
                            // R-fuzz2-1: numeric text is stored in canonical form
                            // ("+5" / "1,234" parse fine but corrupt <v> verbatim).
                            if (double.TryParse(cellValue, out var dbl) && double.IsFinite(dbl))
                            {
                                cell.CellValue = new CellValue(NormalizeNumericCellText(cellValue, dbl));
                                cell.DataType = null;
                            }
                            else
                            {
                                cell.CellValue = new CellValue(cellValue);
                                cell.DataType = new EnumValue<CellValues>(CellValues.String);
                            }
                        }
                    }
                    break;
                case "formula":
                    // BUG-R36-03 fix: reject empty/whitespace formula strings.
                    // Storing an empty CellFormula (<x:f/>) is invalid OOXML and causes
                    // Get() to return "=" as the cell text. Treat as clear-formula intent.
                    if (string.IsNullOrWhiteSpace(value))
                        throw new ArgumentException(
                            "Formula cannot be empty or whitespace. " +
                            "To clear a formula use --prop value= (set to a plain value) or --prop clear=true.");
                    RejectCrossWorkbookFormula(value);
                    ValidateFormulaCellRefs(value);
                    // BUG R4-A: scrub XML-illegal control chars (U+000B, U+000C, etc.)
                    // from formula text before assignment. CellValue text gets sanitized
                    // elsewhere; without symmetric handling here, save throws
                    // ArgumentException("invalid character") from XmlUtf8RawTextWriter.
                    var setFormulaText = value.TrimStart('=');
                    var setCellFormula = new CellFormula(Core.PivotTableHelper.SanitizeXmlText(Core.ModernFunctionQualifier.Qualify(Core.ModernFunctionQualifier.AutoQuoteSheetRefs(setFormulaText))));
                    // Dynamic-array spill metadata — see ExcelHandler.DynamicArray.cs
                    // for rationale. t="array" + the XLDAPR cm metadata make Excel 365
                    // spill SORT/FILTER/UNIQUE/SEQUENCE/etc.; t="array" alone is a
                    // legacy CSE array locked to the anchor cell.
                    if (Core.ModernFunctionQualifier.IsDynamicArrayFormula(setFormulaText) && cell.CellReference?.Value != null)
                    {
                        setCellFormula.FormulaType = CellFormulaValues.Array;
                        setCellFormula.Reference = cell.CellReference.Value;
                        EnsureDynamicArrayMetadata(cell);
                    }
                    // CONSISTENCY(value-child-uniqueness): drop any stale <is>
                    // placeholder so the cell holds a single value child.
                    cell.RemoveAllChildren<InlineString>();
                    cell.CellFormula = setCellFormula;
                    // Try to evaluate and cache the result immediately. A formula
                    // that references a sheet which does not exist yet (the sheet
                    // is added later in the session) would evaluate to #REF! and
                    // cache that wrong value; leave the cache empty instead so the
                    // persist-time RefreshStaleFormulaCaches sweep fills it once the
                    // referenced sheet's data exists. Mirrors the import path.
                    if (!FormulaReferencesMissingSheet(value.TrimStart('=')))
                    {
                        var evalSheetData = GetSheet(worksheet).GetFirstChild<SheetData>();
                        var evaluator = new Core.FormulaEvaluator(evalSheetData!, _doc.WorkbookPart);
                        var evalResult = evaluator.TryEvaluateFull(value.TrimStart('='));
                        // Cache the freshly-evaluated result (or clear value+type when
                        // unevaluable). Dispatch shared with the L1 cache-refresh path —
                        // see ExcelHandler.FormulaCache.cs.
                        WriteFormulaResultToCell(cell, evalResult);
                    }
                    else
                    {
                        cell.CellValue = null;
                        cell.DataType = null;
                    }
                    // Excel/Sheets recalculate formulas on open via fullCalcOnLoad;
                    // headless cache-trusting readers (openpyxl/pandas) do not. The
                    // cache written here can later go stale if a precedent cell is
                    // mutated after this formula — RefreshStaleFormulaCaches reconciles
                    // that at persist (see ExcelHandler.FormulaCache.cs).
                    EnsureFullCalcOnLoad();
                    break;
                case "type":
                    // CONSISTENCY(cell-type-parity): Add accepts type=richtext;
                    // Set must too. Delegates to ApplyRichTextToCell which builds
                    // a SharedString rich-text entry from `runs=<json>` (or the
                    // legacy run1=… mini-spec).
                    if (value.Equals("richtext", StringComparison.OrdinalIgnoreCase) ||
                        value.Equals("rich", StringComparison.OrdinalIgnoreCase))
                    {
                        ApplyRichTextToCell(cell, properties);
                        break;
                    }
                    cell.DataType = value.ToLowerInvariant() switch
                    {
                        "string" or "str" => new EnumValue<CellValues>(CellValues.String),
                        "number" or "num" => null,
                        "boolean" or "bool" => new EnumValue<CellValues>(CellValues.Boolean),
                        "date" => null, // Dates are stored as numbers; format is applied via numberformat below
                        // CONSISTENCY(cell-type-parity): accept `error`/`err` as in Add.
                        "error" or "err" => new EnumValue<CellValues>(CellValues.Error),
                        _ => throw new ArgumentException($"Invalid cell 'type' value '{value}'. Valid types: string, number, boolean, date, error, richtext.")
                    };
                    // Convert cell value for boolean type
                    if (value.ToLowerInvariant() is "boolean" or "bool" && cell.CellValue != null)
                    {
                        var cv = cell.CellValue.Text.Trim().ToLowerInvariant();
                        if (cv is "true" or "yes" or "1") cell.CellValue = new CellValue("1");
                        else if (cv is "false" or "no" or "0") cell.CellValue = new CellValue("0");
                        else if (!string.IsNullOrEmpty(cv))
                            throw new ArgumentException(
                                $"Cannot store '{cell.CellValue.Text}' as boolean; value must be true/false, yes/no, or 1/0. " +
                                "Use type=string to keep the literal text.");
                    }
                    // For date type, apply a default date number format unless caller already specifies one
                    if (value.Equals("date", StringComparison.OrdinalIgnoreCase)
                        && !properties.ContainsKey("numberformat") && !properties.ContainsKey("numfmt") && !properties.ContainsKey("format"))
                        styleProps["numberformat"] = "m/d/yy";
                    break;
                case "clear":
                    // Per schemas/help/xlsx/cell.json: clear erases value/formula
                    // before applying new content. StyleIndex (font/alignment/
                    // border/numfmt) is independent state and must survive clear,
                    // matching `set`'s overall merge semantics.
                    cell.CellValue = null;
                    cell.CellFormula = null;
                    cell.RemoveAllChildren<InlineString>();
                    cell.DataType = null;
                    break;
                case "arrayformula":
                {
                    // Flag form (arrayformula=true): convert the cell's
                    // existing/companion formula rather than writing the
                    // literal "true" as formula text (silent corruption).
                    var arrText = value;
                    if (arrText.Equals("true", StringComparison.OrdinalIgnoreCase)
                        || arrText == "1"
                        || arrText.Equals("yes", StringComparison.OrdinalIgnoreCase))
                    {
                        arrText = properties.GetValueOrDefault("formula")
                            ?? cell.CellFormula?.Text
                            ?? throw new ArgumentException(
                                "arrayformula=true requires a formula: pass the text directly (arrayformula=\"B1:B3*C1:C3\") or combine with formula=, or set it on a cell that already has a formula.");
                    }
                    RejectCrossWorkbookFormula(arrText);
                    var arrRef = properties.GetValueOrDefault("ref", cellRef);
                    cell.CellFormula = new CellFormula(Core.PivotTableHelper.SanitizeXmlText(Core.ModernFunctionQualifier.Qualify(Core.ModernFunctionQualifier.AutoQuoteSheetRefs(arrText.TrimStart('=')))))
                    {
                        FormulaType = CellFormulaValues.Array,
                        Reference = arrRef
                    };
                    cell.CellValue = null;
                    cell.RemoveAllChildren<InlineString>();
                    break;
                }
                case "ref":
                    // Consumed by the arrayformula case above (spill range via
                    // GetValueOrDefault); without this case the key itself
                    // fell to default → a false "UNSUPPORTED props: ref" on
                    // every dump replay of an array-formula cell.
                    if (!properties.Keys.Any(k => k.Equals("arrayformula", StringComparison.OrdinalIgnoreCase)))
                        unsupported.Add("ref (only valid alongside arrayformula=)");
                    break;
                // CONSISTENCY(xlsx-hyperlink-cell-backed): `query hyperlink` emits
                // the backing cell path with Format["url"]; accept `url` as an
                // alias for the canonical cell `link` so that query result round-
                // trips into set (set "/Sheet1/A1" --prop url=...).
                case "url":
                case "link":
                {
                    var ws = GetSheet(worksheet);
                    var hyperlinksEl = ws.GetFirstChild<Hyperlinks>();
                    if (string.IsNullOrEmpty(value) || value.Equals("none", StringComparison.OrdinalIgnoreCase))
                    {
                        if (hyperlinksEl != null)
                            RemoveCellHyperlinksAndRels(worksheet, hyperlinksEl, cellRef);
                        if (hyperlinksEl != null && !hyperlinksEl.HasChildren)
                            hyperlinksEl.Remove();
                        // Symmetric to H3 above: when removing a hyperlink,
                        // also drop the implicit Hyperlink cellStyle that
                        // Add/Set installed (blue + underline). User-assigned
                        // explicit styles are preserved — we only revert
                        // StyleIndex values that match the Hyperlink xf.
                        if (cell.StyleIndex != null && cell.StyleIndex.Value != 0)
                        {
                            var wbPart = _doc.WorkbookPart;
                            if (wbPart != null)
                            {
                                var styleManager = new ExcelStyleManager(wbPart);
                                if (styleManager.IsHyperlinkCellStyleXf(cell.StyleIndex.Value))
                                {
                                    cell.StyleIndex = null;
                                    _dirtyStylesheet = true;
                                }
                            }
                        }
                    }
                    else
                    {
                        // Validate the target BEFORE creating the <hyperlinks>
                        // container: a rejected scheme used to leave an empty
                        // <x:hyperlinks/> behind — schema-invalid (>=1 child
                        // required), so real Excel refused the file even
                        // though the set itself was correctly rejected.
                        // Reject XML-illegal control chars up front — otherwise
                        // the value is accepted into the DOM and only blows up
                        // at close ("save failed — data may be lost").
                        Core.ParseHelpers.ValidateXmlText(value, "link");
                        var isInternalTarget = ResolveInternalHyperlinkLocation(value) != null;
                        if (!isInternalTarget)
                            Core.HyperlinkUriValidator.RequireSafeScheme(value, "link");
                        if (hyperlinksEl == null)
                        {
                            hyperlinksEl = new Hyperlinks();
                            ws.AppendChild(hyperlinksEl);
                        }
                        // Replacing the link: drop the old <hyperlink> AND its
                        // external relationship (else the .rels accumulates an
                        // orphan target per repoint — a per-file leak).
                        RemoveCellHyperlinksAndRels(worksheet, hyperlinksEl, cellRef);
                        // H2: optional tooltip/screenTip from sibling props.
                        var setHlTip = properties.GetValueOrDefault("tooltip")
                            ?? properties.GetValueOrDefault("screenTip")
                            ?? properties.GetValueOrDefault("screentip");
                        // H2b: optional display= friendly text (OOXML @display).
                        var setHlDisplay = properties.GetValueOrDefault("display");
                        // These land in @tooltip/@display attributes — reject
                        // XML-illegal control chars up front, same as link.
                        Core.ParseHelpers.ValidateXmlText(setHlTip, "tooltip");
                        Core.ParseHelpers.ValidateXmlText(setHlDisplay, "display");
                        // R37-B: also accept bare `SheetName!Cell` (no '#' prefix)
                        // and quoted `'Multi Word'!Cell` as internal targets.
                        // CONSISTENCY(internal-hyperlink): same detection used in Add.Cells.cs.
                        var internalLoc = ResolveInternalHyperlinkLocation(value);
                        if (internalLoc != null)
                        {
                            // Internal target (sheet cell or named range) is
                            // written as an in-document hyperlink via the
                            // `location` attribute, no relationship/target.
                            var hl = new Hyperlink
                            {
                                Reference = cellRef.ToUpperInvariant(),
                                Location = internalLoc
                            };
                            if (!string.IsNullOrEmpty(setHlTip)) hl.Tooltip = setHlTip;
                            if (!string.IsNullOrEmpty(setHlDisplay)) hl.Display = setHlDisplay;
                            hyperlinksEl.AppendChild(hl);
                        }
                        else
                        {
                            // CONSISTENCY(hyperlink-scheme-allowlist): reject
                            // javascript:/file:/data:/vbscript: before the
                            // relationship is created. Internal targets
                            // (Sheet!Cell, named ranges) already routed above
                            // via TryParseInternalHyperlinkLocation.
                            Core.HyperlinkUriValidator.RequireSafeScheme(value, "link");
                            var hlUri = new Uri(value, UriKind.RelativeOrAbsolute);
                            var hlRel = worksheet.AddHyperlinkRelationship(hlUri, isExternal: true);
                            var hl = new Hyperlink { Reference = cellRef.ToUpperInvariant(), Id = hlRel.Id };
                            if (!string.IsNullOrEmpty(setHlTip)) hl.Tooltip = setHlTip;
                            if (!string.IsNullOrEmpty(setHlDisplay)) hl.Display = setHlDisplay;
                            hyperlinksEl.AppendChild(hl);
                        }
                        // H3: apply the built-in "Hyperlink" cellStyle (blue +
                        // underline) if the cell has no user-assigned style.
                        // CONSISTENCY(hyperlink-cellstyle): preserve an
                        // explicit StyleIndex the user already set.
                        if (cell.StyleIndex == null || cell.StyleIndex.Value == 0)
                        {
                            var wbPart = _doc.WorkbookPart
                                ?? throw new InvalidOperationException("Workbook not found");
                            var styleManager = new ExcelStyleManager(wbPart);
                            cell.StyleIndex = styleManager.EnsureHyperlinkCellStyle();
                            _dirtyStylesheet = true;
                        }
                    }
                    break;
                }
                case "merge":
                {
                    // CONSISTENCY(cell-merge): cell Add already accepts
                    // merge=A1:C3 (see ExcelHandler.Add.Cells.cs); cell Set
                    // mirrors it. Empty/false/none/unmerge clears any merge
                    // anchored at this cell.
                    var ws = GetSheet(worksheet);
                    var mergeCellsEl = ws.GetFirstChild<MergeCells>();
                    var clear = string.IsNullOrWhiteSpace(value)
                        || value.Equals("false", StringComparison.OrdinalIgnoreCase)
                        || value.Equals("none", StringComparison.OrdinalIgnoreCase)
                        || value.Equals("unmerge", StringComparison.OrdinalIgnoreCase);
                    if (clear)
                    {
                        // Drop any merge whose top-left equals this cell.
                        if (mergeCellsEl != null)
                        {
                            foreach (var mc in mergeCellsEl.Elements<MergeCell>().ToList())
                            {
                                var refStr = mc.Reference?.Value ?? "";
                                var topLeft = refStr.Split(':')[0];
                                if (string.Equals(topLeft, cellRef, StringComparison.OrdinalIgnoreCase))
                                    mc.Remove();
                            }
                            if (!mergeCellsEl.HasChildren) mergeCellsEl.Remove();
                            else mergeCellsEl.Count = (uint)mergeCellsEl.Elements<MergeCell>().Count();
                        }
                    }
                    else
                    {
                        var refList = value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                        // CONSISTENCY(merge-empty-container): validate every ref before
                        // creating the container so a rejected ref doesn't leave an
                        // empty <mergeCells> in the saved file.
                        foreach (var r in refList) ValidateMergeRefLiteral(r);
                        if (mergeCellsEl == null)
                        {
                            mergeCellsEl = new MergeCells();
                            ws.AppendChild(mergeCellsEl);
                        }
                        // CONSISTENCY(merge-comma): comma in *prop value* is the
                        // supported batch form (here, in cell Add, and in sheet Set)
                        // — split into separate <mergeCell> elements. Comma in
                        // *path* is rejected by InsertMergeCellChecked since path
                        // is a single-target locator.
                        foreach (var rangeRef in refList)
                            InsertMergeCellChecked(mergeCellsEl, rangeRef, worksheet);
                        mergeCellsEl.Count = (uint)mergeCellsEl.Elements<MergeCell>().Count();
                    }
                    break;
                }
                case "tooltip":
                case "screentip":
                {
                    // H2: tooltip may also be applied to an EXISTING hyperlink.
                    var ws = GetSheet(worksheet);
                    var hyperlinksEl = ws.GetFirstChild<Hyperlinks>();
                    var existing = hyperlinksEl?.Elements<Hyperlink>()
                        .FirstOrDefault(h => h.Reference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true);
                    if (existing == null)
                    {
                        unsupported.Add($"tooltip (no hyperlink exists on {cellRef}; add a link first)");
                        break;
                    }
                    existing.Tooltip = string.IsNullOrEmpty(value) ? null : value;
                    break;
                }
                case "display":
                {
                    // H2b: display may also be applied to an EXISTING hyperlink,
                    // or is consumed as a sibling of link= above (idempotent).
                    var ws = GetSheet(worksheet);
                    var hyperlinksEl = ws.GetFirstChild<Hyperlinks>();
                    var existing = hyperlinksEl?.Elements<Hyperlink>()
                        .FirstOrDefault(h => h.Reference?.Value?.Equals(cellRef, StringComparison.OrdinalIgnoreCase) == true);
                    if (existing == null)
                    {
                        unsupported.Add($"display (no hyperlink exists on {cellRef}; add a link first)");
                        break;
                    }
                    existing.Display = string.IsNullOrEmpty(value) ? null : value;
                    break;
                }
                case "runs":
                    // Consumed by ApplyRichTextToCell (the type=richtext case
                    // above); without this case the key falls through to the
                    // unsupported list even though it WAS applied — a false
                    // "UNSUPPORTED props: runs" on every dump→batch replay of a
                    // richtext cell. CE1 parity with Add: runs= without type=
                    // implies richtext.
                    if (!properties.Keys.Any(k => k.Equals("type", StringComparison.OrdinalIgnoreCase)))
                        ApplyRichTextToCell(cell, properties);
                    else if (!properties.Any(p => p.Key.Equals("type", StringComparison.OrdinalIgnoreCase)
                             && (p.Value.Equals("richtext", StringComparison.OrdinalIgnoreCase)
                                 || p.Value.Equals("rich", StringComparison.OrdinalIgnoreCase))))
                        unsupported.Add("runs (only valid with type=richtext)");
                    break;
                default:
                    // Legacy richtext mini-spec keys (run1=, run2=, …) are read
                    // inside ApplyRichTextToCell — same false-unsupported hole
                    // as `runs` above.
                    if (System.Text.RegularExpressions.Regex.IsMatch(key, @"^run\d+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                        && properties.Any(p => p.Key.Equals("type", StringComparison.OrdinalIgnoreCase)
                            && (p.Value.Equals("richtext", StringComparison.OrdinalIgnoreCase)
                                || p.Value.Equals("rich", StringComparison.OrdinalIgnoreCase))))
                        break;
                    // Check for known flat-key misuse first, even before generic
                    // attribute fallback — otherwise user typos like `size=14`
                    // would be silently written as unknown XML attributes.
                    var cellHint = CellPropHints.TryGetHint(key);
                    if (cellHint != null)
                    {
                        unsupported.Add(cellHint);
                    }
                    else if (!GenericXmlQuery.SetGenericAttribute(cell, key, value))
                    {
                        unsupported.Add(unsupported.Count == 0
                            ? $"{key} (valid cell props: value, formula, arrayformula, type, clear, link, bold, italic, strike, underline, superscript, subscript, font.color, font.size, font.name, fill, border.all, alignment.horizontal, numfmt, locked, formulahidden)"
                            : key);
                    }
                    break;
            }
        }

        // Apply style properties if any
        if (styleProps.Count > 0)
        {
            var workbookPart = _doc.WorkbookPart
                ?? throw new InvalidOperationException("Workbook not found");
            var styleManager = new ExcelStyleManager(workbookPart);
            cell.StyleIndex = styleManager.ApplyStyle(cell, styleProps, unsupported);
            _dirtyStylesheet = true;

            // R24-1: numberformat="@" → force text storage. See ExcelHandler.Add.Cells.cs
            // for the matching guard on the Add path.
            if (IsTextNumberFormat(styleProps)
                && cell.DataType?.Value != CellValues.SharedString
                && cell.DataType?.Value != CellValues.InlineString
                && cell.CellFormula == null)
            {
                cell.DataType = new EnumValue<CellValues>(CellValues.String);
            }
        }

        return unsupported;
    }

    // BUG-003: worksheet child order is sheetPr, dimension, sheetViews, ...
    // A blind InsertAt(0) puts sheetViews before an existing sheetPr (e.g. one
    // created by tabColor) and produces schema-invalid OOXML that fails
    // validate (Excel silently repairs it, so it went unnoticed).
}
