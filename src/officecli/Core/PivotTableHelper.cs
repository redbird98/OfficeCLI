// Copyright 2025 OfficeCli (officecli.ai)
// SPDX-License-Identifier: Apache-2.0

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeCli.Core;

/// <summary>
/// Helper for building and reading pivot tables.
/// Manages PivotTableCacheDefinitionPart (workbook-level) and PivotTablePart (worksheet-level).
/// </summary>
internal static class PivotTableHelper
{
    /// <summary>
    /// Create a pivot table on the target worksheet.
    /// </summary>
    /// <param name="workbookPart">The workbook part</param>
    /// <param name="targetSheet">Worksheet where the pivot table will be placed</param>
    /// <param name="sourceSheet">Worksheet containing the source data</param>
    /// <param name="sourceSheetName">Name of the source worksheet</param>
    /// <param name="sourceRef">Source data range (e.g. "A1:D100")</param>
    /// <param name="position">Top-left cell for the pivot table (e.g. "F1")</param>
    /// <param name="properties">Configuration: rows, cols, values, filters, style, name</param>
    /// <returns>The 1-based index of the created pivot table</returns>
    internal static int CreatePivotTable(
        WorkbookPart workbookPart,
        WorksheetPart targetSheet,
        WorksheetPart sourceSheet,
        string sourceSheetName,
        string sourceRef,
        string position,
        Dictionary<string, string> properties)
    {
        // 1. Read source data to build cache
        var (headers, columnData) = ReadSourceData(sourceSheet, sourceRef);
        if (headers.Length == 0)
            throw new ArgumentException("Source range has no data");

        // 2. Parse field assignments from properties
        var rowFields = ParseFieldList(properties, "rows", headers);
        var colFields = ParseFieldList(properties, "cols", headers);
        var filterFields = ParseFieldList(properties, "filters", headers);
        var valueFields = ParseValueFields(properties, "values", headers);

        // Auto-assign: if no values specified, use the first numeric column
        if (valueFields.Count == 0)
        {
            for (int i = 0; i < headers.Length; i++)
            {
                if (!rowFields.Contains(i) && !colFields.Contains(i) && !filterFields.Contains(i)
                    && columnData[i].All(v => double.TryParse(v, System.Globalization.CultureInfo.InvariantCulture, out _)))
                {
                    valueFields.Add((i, "sum", $"Sum of {headers[i]}"));
                    break;
                }
            }
        }

        // 3. Generate unique cache ID
        uint cacheId = 0;
        var workbook = workbookPart.Workbook
            ?? throw new InvalidOperationException("Workbook is missing");
        var pivotCaches = workbook.GetFirstChild<PivotCaches>();
        if (pivotCaches != null)
            cacheId = pivotCaches.Elements<PivotCache>().Select(pc => pc.CacheId?.Value ?? 0u).DefaultIfEmpty(0u).Max() + 1;

        // 4. Create PivotTableCacheDefinitionPart at workbook level
        var cachePart = workbookPart.AddNewPart<PivotTableCacheDefinitionPart>();
        var cacheRelId = workbookPart.GetIdOfPart(cachePart);

        // Build cache definition + per-field shared-item index maps. The maps are
        // needed to write pivotCacheRecords below: each non-numeric field value is
        // referenced as <x v="N"/> where N is the value's position in sharedItems.
        var (cacheDef, fieldNumeric, fieldValueIndex) =
            BuildCacheDefinition(sourceSheetName, sourceRef, headers, columnData);
        cachePart.PivotCacheDefinition = cacheDef;
        cachePart.PivotCacheDefinition.Save();

        // 4b. Create PivotTableCacheRecordsPart and write one record per source row.
        // Without records, Excel rejects the file with "PivotTable report is invalid"
        // because saveData defaults to true. Writing real records also makes the file
        // self-contained for non-refreshing consumers (POI, third-party parsers).
        var recordsPart = cachePart.AddNewPart<PivotTableCacheRecordsPart>();
        recordsPart.PivotCacheRecords = BuildCacheRecords(columnData, fieldNumeric, fieldValueIndex);
        recordsPart.PivotCacheRecords.Save();

        // The pivotCacheDefinition element MUST carry an r:id attribute pointing to the
        // records part — Excel uses it to find records, not the package _rels alone.
        // LibreOffice writes this in xepivotxml.cxx:280 (FSNS(XML_r, XML_id)). Without
        // this attribute the file looks structurally complete but Excel rejects it.
        cacheDef.Id = cachePart.GetIdOfPart(recordsPart);
        cachePart.PivotCacheDefinition.Save();

        // Register in workbook's PivotCaches
        if (pivotCaches == null)
        {
            pivotCaches = new PivotCaches();
            workbook.AppendChild(pivotCaches);
        }
        pivotCaches.AppendChild(new PivotCache { CacheId = cacheId, Id = cacheRelId });
        workbook.Save();

        // 5. Create PivotTablePart at worksheet level
        var pivotPart = targetSheet.AddNewPart<PivotTablePart>();
        // Link pivot table to cache definition
        pivotPart.AddPart(cachePart);

        var pivotName = properties.GetValueOrDefault("name", $"PivotTable{cacheId + 1}");
        var style = properties.GetValueOrDefault("style", "PivotStyleLight16");

        var pivotDef = BuildPivotTableDefinition(
            pivotName, cacheId, position, headers, columnData,
            rowFields, colFields, filterFields, valueFields, style);
        pivotPart.PivotTableDefinition = pivotDef;
        pivotPart.PivotTableDefinition.Save();

        // 6. RENDER the pivot output into the target sheet's <sheetData>.
        //
        // This is the critical step that distinguishes a "valid pivot file Excel
        // accepts" from a "pivot file Excel actually displays". Excel does NOT
        // recompute pivots from cache on open — it reads the rendered cells
        // directly from sheetData, exactly like any other range. We verified this
        // by inspecting an Excel-authored sample (excel_authored.xlsx → sheet2.xml):
        // every aggregated cell is a literal <c><v>200</v></c> element.
        //
        // Without this step the pivot opens as an empty drop-down skeleton — the
        // structure is valid but there is nothing to display. POI / Open XML SDK
        // suffer from exactly the same limitation; this is the lift that turns
        // officecli into a real pivot writer rather than a definition-only one.
        //
        // For unsupported configurations (multiple row/col fields, multiple data
        // fields, page filters), the renderer falls back to writing nothing, which
        // gives Excel an empty sheetData and the same skeleton-only behavior.
        // Those configs are tracked as a v2 expansion.
        RenderPivotIntoSheet(
            targetSheet, position, headers, columnData,
            rowFields, colFields, valueFields, filterFields);

        // Return 1-based index
        return targetSheet.PivotTableParts.ToList().IndexOf(pivotPart) + 1;
    }

    // ==================== Geometry & Cache Readback Helpers ====================

    /// <summary>Computed pivot table extent — anchor + bounding range + key offsets.</summary>
    private readonly struct PivotGeometry
    {
        public PivotGeometry(int anchorCol, int anchorRow, int width, int height, int rowLabelCols, string rangeRef)
        {
            AnchorCol = anchorCol;
            AnchorRow = anchorRow;
            Width = width;
            Height = height;
            RowLabelCols = rowLabelCols;
            RangeRef = rangeRef;
        }
        public int AnchorCol { get; }
        public int AnchorRow { get; }
        public int Width { get; }
        public int Height { get; }
        public int RowLabelCols { get; }
        public string RangeRef { get; }
    }

    /// <summary>
    /// Compute the bounding range and row-label column count for a pivot at the
    /// given anchor with the given field assignments. Used by both initial creation
    /// (BuildPivotTableDefinition) and post-Set rebuild (RebuildFieldAreas) so the
    /// two paths agree on layout.
    ///
    /// Layout assumes the standard compact/outline mode with:
    ///   width  = max(1, rowFieldCount)                    // row labels
    ///          + max(1, colUnique) * max(1, valueCount)    // data cells
    ///          + (colFieldCount > 0 ? 1 : 0)               // grand total column
    ///   height = (colFieldCount > 0 ? 2 : 1)               // header rows
    ///          + max(1, rowUnique)                          // data rows
    ///          + 1                                          // grand total row
    /// Page filter rows are excluded from the range per ECMA-376.
    /// </summary>
    private static PivotGeometry ComputePivotGeometry(
        string position, List<string[]> columnData,
        List<int> rowFieldIndices, List<int> colFieldIndices,
        List<(int idx, string func, string name)> valueFields)
    {
        int colUnique = ProductOfUniqueValues(colFieldIndices, columnData);
        int dataFieldCount = Math.Max(1, valueFields.Count);

        // Compact mode: row labels collapse into a single column regardless of
        // how many row fields the user assigned (verified against
        // multi_row_authored.xlsx with rows=地区,城市 → still firstDataCol=1).
        int rowLabelCols = 1;

        // Width for K data fields × L col label values:
        //   1 (row labels) + L*K (data area) + K (grand total area when col field exists)
        // For K=1, this collapses to the original 1 + L + 1 = 2+L formula.
        int valueCols = Math.Max(1, colUnique) * dataFieldCount;
        int totalCols = colFieldIndices.Count > 0 ? dataFieldCount : 0;
        int width = rowLabelCols + valueCols + totalCols;

        // Row count depends on number of row fields:
        //   N=1: just R unique row values
        //   N=2: outer count + leaf combos (one subtotal row per outer + one row
        //        per (outer, inner) combo that exists in the data — NOT a
        //        cartesian product, only existing combos)
        int dataRowCount;
        if (rowFieldIndices.Count >= 2)
        {
            var groups = BuildOuterInnerGroups(
                rowFieldIndices[0], rowFieldIndices[1], columnData);
            dataRowCount = groups.Sum(g => 1 + g.inners.Count);
        }
        else
        {
            dataRowCount = Math.Max(1, ProductOfUniqueValues(rowFieldIndices, columnData));
        }

        // Header row count: K=1 → 2 (col field caption + col labels), K>1 → 3
        // (extra row for data field names repeated under each col group).
        int headerRows;
        if (colFieldIndices.Count > 0)
            headerRows = dataFieldCount > 1 ? 3 : 2;
        else
            headerRows = dataFieldCount > 1 ? 2 : 1;

        int height = headerRows + dataRowCount + 1;

        var (anchorCol, anchorRow) = ParseCellRef(position);
        var anchorColIdx = ColToIndex(anchorCol);
        var endColIdx = anchorColIdx + width - 1;
        var endRow = anchorRow + height - 1;
        var rangeRef = $"{position}:{IndexToCol(endColIdx)}{endRow}";

        return new PivotGeometry(anchorColIdx, anchorRow, width, height, rowLabelCols, rangeRef);
    }

    /// <summary>
    /// Reconstruct the per-field columnData from the cache definition + records.
    /// Used by RebuildFieldAreas after Set: the source sheet may not be readily
    /// reachable, but the cache holds the original values (string fields via
    /// sharedItems index, numeric fields directly in &lt;n v=...&gt;). This makes
    /// the rebuild self-contained on the cache part alone.
    /// </summary>
    private static (string[] headers, List<string[]> columnData) ReadColumnDataFromCache(
        PivotCacheDefinition cacheDef, PivotCacheRecords? records)
    {
        var cacheFields = cacheDef.GetFirstChild<CacheFields>();
        if (cacheFields == null) return (Array.Empty<string>(), new List<string[]>());

        var fieldList = cacheFields.Elements<CacheField>().ToList();
        var headers = fieldList.Select(cf => cf.Name?.Value ?? "").ToArray();
        var fieldCount = fieldList.Count;

        // Pre-resolve each field's sharedItems string lookup table (index → text).
        // Numeric fields without enumerated items leave the table empty; their
        // values come straight from <n v=...> in the records below.
        var perFieldStrings = new List<List<string>>(fieldCount);
        for (int f = 0; f < fieldCount; f++)
        {
            var items = fieldList[f].GetFirstChild<SharedItems>();
            var list = new List<string>();
            if (items != null)
            {
                foreach (var child in items.ChildElements)
                {
                    list.Add(child switch
                    {
                        StringItem s => s.Val?.Value ?? string.Empty,
                        NumberItem n => n.Val?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty,
                        DateTimeItem d => d.Val?.Value.ToString("yyyy-MM-dd") ?? string.Empty,
                        BooleanItem b => b.Val?.Value == true ? "true" : "false",
                        _ => string.Empty
                    });
                }
            }
            perFieldStrings.Add(list);
        }

        var recordList = records?.Elements<PivotCacheRecord>().ToList() ?? new List<PivotCacheRecord>();
        var columnData = new List<string[]>(fieldCount);
        for (int f = 0; f < fieldCount; f++)
            columnData.Add(new string[recordList.Count]);

        for (int r = 0; r < recordList.Count; r++)
        {
            var record = recordList[r];
            var children = record.ChildElements.ToList();
            for (int f = 0; f < fieldCount && f < children.Count; f++)
            {
                columnData[f][r] = children[f] switch
                {
                    FieldItem fi when fi.Val?.Value is uint idx
                        && idx < perFieldStrings[f].Count
                        => perFieldStrings[f][(int)idx],
                    NumberItem n => n.Val?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty,
                    StringItem s => s.Val?.Value ?? string.Empty,
                    DateTimeItem d => d.Val?.Value.ToString("yyyy-MM-dd") ?? string.Empty,
                    BooleanItem b => b.Val?.Value == true ? "true" : "false",
                    _ => string.Empty
                };
            }
        }

        return (headers, columnData);
    }

    /// <summary>
    /// Remove every cell in sheetData that falls inside the given pivot range.
    /// Called before re-rendering so stale cells from the previous pivot layout
    /// (e.g. row totals from a wider configuration) do not leak through.
    /// </summary>
    private static void ClearPivotRangeCells(SheetData sheetData, string rangeRef)
    {
        var parts = rangeRef.Split(':');
        if (parts.Length != 2) return;
        var (startCol, startRow) = ParseCellRef(parts[0]);
        var (endCol, endRow) = ParseCellRef(parts[1]);
        var startColIdx = ColToIndex(startCol);
        var endColIdx = ColToIndex(endCol);

        var rowsToRemove = new List<Row>();
        foreach (var row in sheetData.Elements<Row>())
        {
            var rIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rIdx < startRow || rIdx > endRow) continue;

            var cellsToRemove = row.Elements<Cell>()
                .Where(c =>
                {
                    var cref = c.CellReference?.Value ?? "";
                    var (cc, _) = ParseCellRef(cref);
                    var ci = ColToIndex(cc);
                    return ci >= startColIdx && ci <= endColIdx;
                })
                .ToList();
            foreach (var c in cellsToRemove) c.Remove();

            // If the row is now empty AND was entirely inside the pivot, drop it
            // entirely so we don't leave stray <row r="N"/> elements behind.
            if (!row.Elements<Cell>().Any())
                rowsToRemove.Add(row);
        }
        foreach (var r in rowsToRemove) r.Remove();
    }

    // ==================== Pivot Output Renderer ====================

    /// <summary>
    /// Compute the pivot's aggregation matrix from columnData and write the
    /// rendered cells into targetSheet's SheetData. Mirrors what real Excel writes
    /// on save: literal cells with computed values, NOT a definition that Excel
    /// recomputes on open.
    ///
    /// Supported (v1): exactly 1 row field × 1 col field × 1 data field, with
    /// aggregator in {sum, count, average, min, max}, plus row/column/grand totals.
    /// Other configurations leave sheetData empty and emit a stderr warning so
    /// the file still validates and opens, just without rendered data.
    ///
    /// Layout (verified against Excel-authored sample):
    ///     Row 0:  [data caption] [col field caption]
    ///     Row 1:  [row field caption] [col label 1] [col label 2] ... [总计]
    ///     Row 2:  [row label 1]       [v]            [v]              [row total 1]
    ///     ...
    ///     Row N:  [总计]              [col total 1] [col total 2] ... [grand total]
    /// </summary>
    private static void RenderPivotIntoSheet(
        WorksheetPart targetSheet, string position,
        string[] headers, List<string[]> columnData,
        List<int> rowFieldIndices, List<int> colFieldIndices,
        List<(int idx, string func, string name)> valueFields,
        List<int>? filterFieldIndices = null)
    {
        // v3 limits: rows in {1, 2}, cols == 1, dataFields >= 1.
        // 2-row-field path goes to RenderMultiRowPivot below; 1-row goes through
        // the single-row code path. Multi-col field configurations are still
        // unsupported and fall back to the empty skeleton.
        if (rowFieldIndices.Count == 2 && colFieldIndices.Count == 1 && valueFields.Count >= 1)
        {
            RenderMultiRowPivot(targetSheet, position, headers, columnData,
                rowFieldIndices, colFieldIndices, valueFields, filterFieldIndices);
            return;
        }

        if (rowFieldIndices.Count != 1 || colFieldIndices.Count != 1 || valueFields.Count < 1)
        {
            Console.Error.WriteLine(
                "WARNING: pivot rendering currently supports only 1-2 rows × 1 col × 1+ data fields. " +
                "The file will open but the pivot will appear empty. " +
                "Use Excel's Refresh button to populate it manually.");
            return;
        }

        var rowFieldIdx = rowFieldIndices[0];
        var colFieldIdx = colFieldIndices[0];
        var rowFieldName = headers[rowFieldIdx];
        var colFieldName = headers[colFieldIdx];
        int K = valueFields.Count;

        var rowValues = columnData[rowFieldIdx];
        var colValues = columnData[colFieldIdx];

        // Unique row/col labels in cache order (alphabetical ordinal).
        var uniqueRows = rowValues.Where(v => !string.IsNullOrEmpty(v)).Distinct()
            .OrderBy(v => v, StringComparer.Ordinal).ToList();
        var uniqueCols = colValues.Where(v => !string.IsNullOrEmpty(v)).Distinct()
            .OrderBy(v => v, StringComparer.Ordinal).ToList();

        // Bucket source values per (rowLabel, colLabel, dataFieldIdx) so each data
        // field is aggregated independently. The aggregator function differs per
        // data field (sum/count/avg/...) so each bucket carries its own reducer.
        // Two data fields on the same source column are common (e.g. sum + count
        // of 金额) and produce two independent buckets keyed by their dataFieldIdx
        // in valueFields.
        var perBucket = new Dictionary<(string r, string c, int d), List<double>>();
        var perDataField = new List<List<double>>();
        for (int d = 0; d < K; d++) perDataField.Add(new List<double>());

        for (int i = 0; i < rowValues.Length; i++)
        {
            var rv = rowValues.Length > i ? rowValues[i] : null;
            var cv = colValues.Length > i ? colValues[i] : null;
            if (string.IsNullOrEmpty(rv) || string.IsNullOrEmpty(cv)) continue;

            for (int d = 0; d < K; d++)
            {
                var dataIdx = valueFields[d].idx;
                var dataValues = columnData[dataIdx];
                if (i >= dataValues.Length) continue;
                if (!double.TryParse(dataValues[i], System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture, out var num)) continue;

                var key = (rv, cv, d);
                if (!perBucket.TryGetValue(key, out var list))
                {
                    list = new List<double>();
                    perBucket[key] = list;
                }
                list.Add(num);
                perDataField[d].Add(num);
            }
        }

        double Reduce(IEnumerable<double> values, string func)
        {
            // Match LibreOffice's ScDPAggData (dptabres.cxx) aggregator semantics.
            var arr = values as double[] ?? values.ToArray();
            if (arr.Length == 0) return 0;
            return func.ToLowerInvariant() switch
            {
                "sum" => arr.Sum(),
                "count" => arr.Length,
                "average" or "avg" => arr.Average(),
                "min" => arr.Min(),
                "max" => arr.Max(),
                _ => arr.Sum()
            };
        }

        // Compute the K-deep cell matrix + row/col/grand totals per data field.
        // matrix[r, c, d] = reduce(values for row r, col c, data field d)
        // rowTotals[r, d], colTotals[c, d], grandTotals[d] follow the same shape.
        var matrix = new double?[uniqueRows.Count, uniqueCols.Count, K];
        var rowTotals = new double[uniqueRows.Count, K];
        var colTotals = new double[uniqueCols.Count, K];
        var grandTotals = new double[K];
        for (int d = 0; d < K; d++)
        {
            var func = valueFields[d].func;
            for (int r = 0; r < uniqueRows.Count; r++)
            {
                var rowAll = new List<double>();
                for (int c = 0; c < uniqueCols.Count; c++)
                {
                    if (perBucket.TryGetValue((uniqueRows[r], uniqueCols[c], d), out var bucket) && bucket.Count > 0)
                    {
                        matrix[r, c, d] = Reduce(bucket, func);
                        rowAll.AddRange(bucket);
                    }
                }
                rowTotals[r, d] = Reduce(rowAll, func);
            }
            for (int c = 0; c < uniqueCols.Count; c++)
            {
                var colAll = new List<double>();
                for (int r = 0; r < uniqueRows.Count; r++)
                {
                    if (perBucket.TryGetValue((uniqueRows[r], uniqueCols[c], d), out var bucket))
                        colAll.AddRange(bucket);
                }
                colTotals[c, d] = Reduce(colAll, func);
            }
            grandTotals[d] = Reduce(perDataField[d], func);
        }

        // ===== Write cells =====
        // For K=1, layout is 2 header rows: caption + col labels.
        // For K>1, layout is 3 header rows: caption + col labels + per-data-field
        // names repeated under each col label group. This matches the Excel sample
        // multi_data_authored.xlsx exactly.
        var (anchorCol, anchorRow) = ParseCellRef(position);
        var anchorColIdx = ColToIndex(anchorCol);
        var totalColLabel = "总计";

        var ws = targetSheet.Worksheet
            ?? throw new InvalidOperationException("Target worksheet has no Worksheet element");
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            sheetData = new SheetData();
            ws.AppendChild(sheetData);
        }

        // ----- Row 0 (caption row) -----
        // Single data field: data field name in row-label col, col field name in first data col.
        // Multi data field: empty in row-label col, col field name (or "Values" placeholder) in first data col.
        var captionRow = new Row { RowIndex = (uint)anchorRow };
        if (K == 1)
            captionRow.AppendChild(MakeStringCell(anchorColIdx, anchorRow, valueFields[0].name));
        captionRow.AppendChild(MakeStringCell(anchorColIdx + 1, anchorRow, colFieldName));
        sheetData.AppendChild(captionRow);

        // ----- Row 1 (col label row) -----
        // K=1: row field caption + col labels + grand total label
        // K>1: empty row-label cell + col labels at first col of each K-group + grand total labels
        var colLabelRowIdx = anchorRow + 1;
        var colLabelRow = new Row { RowIndex = (uint)colLabelRowIdx };
        if (K == 1)
        {
            colLabelRow.AppendChild(MakeStringCell(anchorColIdx, colLabelRowIdx, rowFieldName));
            for (int c = 0; c < uniqueCols.Count; c++)
                colLabelRow.AppendChild(MakeStringCell(anchorColIdx + 1 + c, colLabelRowIdx, uniqueCols[c]));
            colLabelRow.AppendChild(MakeStringCell(anchorColIdx + 1 + uniqueCols.Count, colLabelRowIdx, totalColLabel));
        }
        else
        {
            // First col of each K-group gets the col label; the K-1 cells after are
            // visually spanned in Excel's renderer but we leave them empty in
            // sheetData (Excel handles the visual span via colItems metadata).
            for (int c = 0; c < uniqueCols.Count; c++)
            {
                int colStart = anchorColIdx + 1 + c * K;
                colLabelRow.AppendChild(MakeStringCell(colStart, colLabelRowIdx, uniqueCols[c]));
            }
            // Grand total area: K cells, one per data field, labeled "Total <name>"
            int totalStart = anchorColIdx + 1 + uniqueCols.Count * K;
            for (int d = 0; d < K; d++)
                colLabelRow.AppendChild(MakeStringCell(totalStart + d, colLabelRowIdx, "Total " + valueFields[d].name));
        }
        sheetData.AppendChild(colLabelRow);

        // ----- Row 2 (data field name row, only when K>1) -----
        int firstDataRow;
        if (K > 1)
        {
            var dfNameRowIdx = anchorRow + 2;
            var dfNameRow = new Row { RowIndex = (uint)dfNameRowIdx };
            // row label column gets the row field name
            dfNameRow.AppendChild(MakeStringCell(anchorColIdx, dfNameRowIdx, rowFieldName));
            // Repeat data field names under each col label group
            for (int c = 0; c < uniqueCols.Count; c++)
            {
                for (int d = 0; d < K; d++)
                {
                    int colIdx = anchorColIdx + 1 + c * K + d;
                    dfNameRow.AppendChild(MakeStringCell(colIdx, dfNameRowIdx, valueFields[d].name));
                }
            }
            // No data field names under the grand total cols — row 1 already
            // labeled them with "Total <name>" so they are self-describing.
            sheetData.AppendChild(dfNameRow);
            firstDataRow = anchorRow + 3;
        }
        else
        {
            firstDataRow = anchorRow + 2;
        }

        // ----- Data rows -----
        for (int r = 0; r < uniqueRows.Count; r++)
        {
            var rowIdx = firstDataRow + r;
            var dataRow = new Row { RowIndex = (uint)rowIdx };
            dataRow.AppendChild(MakeStringCell(anchorColIdx, rowIdx, uniqueRows[r]));
            for (int c = 0; c < uniqueCols.Count; c++)
            {
                for (int d = 0; d < K; d++)
                {
                    int colIdx = anchorColIdx + 1 + c * K + d;
                    var v = matrix[r, c, d];
                    if (v.HasValue)
                        dataRow.AppendChild(MakeNumericCell(colIdx, rowIdx, v.Value));
                }
            }
            // Row totals — K cells (one per data field).
            int rowTotalStart = anchorColIdx + 1 + uniqueCols.Count * K;
            for (int d = 0; d < K; d++)
                dataRow.AppendChild(MakeNumericCell(rowTotalStart + d, rowIdx, rowTotals[r, d]));
            sheetData.AppendChild(dataRow);
        }

        // ----- Grand total row -----
        var grandRowIdx = firstDataRow + uniqueRows.Count;
        var grandRow = new Row { RowIndex = (uint)grandRowIdx };
        grandRow.AppendChild(MakeStringCell(anchorColIdx, grandRowIdx, totalColLabel));
        for (int c = 0; c < uniqueCols.Count; c++)
        {
            for (int d = 0; d < K; d++)
            {
                int colIdx = anchorColIdx + 1 + c * K + d;
                grandRow.AppendChild(MakeNumericCell(colIdx, grandRowIdx, colTotals[c, d]));
            }
        }
        int grandTotalStart = anchorColIdx + 1 + uniqueCols.Count * K;
        for (int d = 0; d < K; d++)
            grandRow.AppendChild(MakeNumericCell(grandTotalStart + d, grandRowIdx, grandTotals[d]));
        sheetData.AppendChild(grandRow);

        // Page filter cells: rendered ABOVE the table at rows
        // (anchorRow - filterCount - 1) ... (anchorRow - 2). One row per filter
        // field, with field name in the row-label column and "(All)" in the
        // adjacent data column. Row (anchorRow - 1) is left empty as a visual gap.
        //
        // Page filters are NOT inside <location ref/> per ECMA-376; they are
        // separate visual cells whose presence is signalled by the rowPageCount /
        // colPageCount attributes on pivotTableDefinition (already set in
        // BuildPivotTableDefinition). Excel pairs the filter cells with the pivot
        // by their position above the location range.
        //
        // If there isn't enough room above (e.g. user anchored at F1), we skip the
        // visible cells but the pivot definition still tags them as page fields,
        // so the dropdowns appear in Excel's pivot UI even without the cell labels.
        if (filterFieldIndices != null && filterFieldIndices.Count > 0)
        {
            var requiredHeadroom = filterFieldIndices.Count + 1; // filter rows + 1 gap
            if (anchorRow > requiredHeadroom)
            {
                var firstFilterRow = anchorRow - requiredHeadroom;
                for (int fi = 0; fi < filterFieldIndices.Count; fi++)
                {
                    var fIdx = filterFieldIndices[fi];
                    if (fIdx < 0 || fIdx >= headers.Length) continue;
                    var rowIdx = firstFilterRow + fi;
                    var filterRow = new Row { RowIndex = (uint)rowIdx };
                    filterRow.AppendChild(MakeStringCell(anchorColIdx, rowIdx, headers[fIdx]));
                    filterRow.AppendChild(MakeStringCell(anchorColIdx + 1, rowIdx, "(All)"));
                    // Insert in row order: existing rows in sheetData start at
                    // anchorRow, so prepend the filter rows to the front.
                    sheetData.InsertAt(filterRow, fi);
                }
            }
            else
            {
                Console.Error.WriteLine(
                    $"WARNING: pivot at {position} has {filterFieldIndices.Count} page filter(s) " +
                    $"but only {anchorRow - 1} row(s) of headroom above. " +
                    "Filter cells will not be visible in the host sheet, but the filter dropdowns " +
                    "will still appear in Excel's pivot UI. Move the pivot to a lower anchor row " +
                    $"(at least row {requiredHeadroom + 1}) to render the filter cells.");
            }
        }

        ws.Save();
    }

    /// <summary>
    /// Render a 2-row-field pivot. Compact-mode layout (verified against
    /// multi_row_authored.xlsx with rows=地区,城市):
    ///
    ///     A                  B           C           D
    ///   3 [data caption]     [col field caption]
    ///   4 Row Labels         咖啡        奶茶        Grand Total
    ///   5 华东                200        260         460          <- outer subtotal
    ///   6   上海              200        150         350
    ///   7   杭州                         110         110
    ///   8 华北                215        85          300          <- outer subtotal
    ///   ...
    ///   N Grand Total        595        345         940
    ///
    /// Both outer and inner labels live in column A (compact mode collapses the
    /// row-label area into a single column, with Excel auto-indenting inners
    /// visually). Each outer value gets its own subtotal row showing the
    /// aggregate across all its existing inners; only (outer, inner) pairs that
    /// actually appear in the source data are rendered (Excel does not enumerate
    /// empty cartesian cells).
    ///
    /// Multi data fields (K>1) are not yet supported in this code path — would
    /// need to extend col multiplication and add the third "data field name"
    /// header row. v4 expansion. Tracked.
    /// </summary>
    private static void RenderMultiRowPivot(
        WorksheetPart targetSheet, string position,
        string[] headers, List<string[]> columnData,
        List<int> rowFieldIndices, List<int> colFieldIndices,
        List<(int idx, string func, string name)> valueFields,
        List<int>? filterFieldIndices)
    {
        // For now, restrict to K=1 data field. Multi-data + multi-row is a
        // separate cross-product expansion that introduces both extra header
        // rows and extra data columns at the same time.
        if (valueFields.Count != 1)
        {
            Console.Error.WriteLine(
                "WARNING: 2-row-field pivots currently support exactly 1 data field. " +
                "Falling back to empty skeleton.");
            return;
        }

        var outerFieldIdx = rowFieldIndices[0];
        var innerFieldIdx = rowFieldIndices[1];
        var colFieldIdx = colFieldIndices[0];
        var (dataFieldIdx, func, dataFieldName) = valueFields[0];

        var outerVals = columnData[outerFieldIdx];
        var innerVals = columnData[innerFieldIdx];
        var colVals = columnData[colFieldIdx];
        var dataVals = columnData[dataFieldIdx];
        var colFieldName = headers[colFieldIdx];

        // Build the same (outer → [inners]) groups used by BuildMultiRowItems so
        // the rendered cells match the rowItems indices position-for-position.
        var groups = BuildOuterInnerGroups(outerFieldIdx, innerFieldIdx, columnData);
        var uniqueCols = colVals.Where(v => !string.IsNullOrEmpty(v)).Distinct()
            .OrderBy(v => v, StringComparer.Ordinal).ToList();

        // Aggregate per (outer, inner, col) using the LibreOffice all-values
        // semantics so subtotals and totals come from raw values, not from
        // pre-aggregated sub-results (avg-of-all, not avg-of-avgs).
        var leafBucket = new Dictionary<(string o, string i, string c), List<double>>();
        var allValues = new List<double>();
        for (int i = 0; i < dataVals.Length; i++)
        {
            var ov = outerVals.Length > i ? outerVals[i] : null;
            var iv = innerVals.Length > i ? innerVals[i] : null;
            var cv = colVals.Length > i ? colVals[i] : null;
            if (string.IsNullOrEmpty(ov) || string.IsNullOrEmpty(iv) || string.IsNullOrEmpty(cv)) continue;
            if (!double.TryParse(dataVals[i], System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out var num)) continue;

            var key = (ov, iv, cv);
            if (!leafBucket.TryGetValue(key, out var list))
            {
                list = new List<double>();
                leafBucket[key] = list;
            }
            list.Add(num);
            allValues.Add(num);
        }

        double Reduce(IEnumerable<double> values)
        {
            var arr = values as double[] ?? values.ToArray();
            if (arr.Length == 0) return 0;
            return func.ToLowerInvariant() switch
            {
                "sum" => arr.Sum(),
                "count" => arr.Length,
                "average" or "avg" => arr.Average(),
                "min" => arr.Min(),
                "max" => arr.Max(),
                _ => arr.Sum()
            };
        }

        // Compute the totals we'll need for cells: per-leaf cells, outer subtotals
        // per col, leaf row totals, outer row totals, col totals, grand total.
        // All of these reduce raw value lists, never previously-reduced numbers.
        double LeafCell(string outer, string inner, string col)
            => leafBucket.TryGetValue((outer, inner, col), out var b) && b.Count > 0
                ? Reduce(b) : double.NaN;

        double OuterSubtotal(string outer, string col)
        {
            var all = new List<double>();
            foreach (var (o, inners) in groups)
                if (o == outer)
                    foreach (var inner in inners)
                        if (leafBucket.TryGetValue((outer, inner, col), out var b))
                            all.AddRange(b);
            return Reduce(all);
        }

        double LeafRowTotal(string outer, string inner)
        {
            var all = new List<double>();
            foreach (var col in uniqueCols)
                if (leafBucket.TryGetValue((outer, inner, col), out var b))
                    all.AddRange(b);
            return Reduce(all);
        }

        double OuterRowTotal(string outer)
        {
            var all = new List<double>();
            foreach (var (o, inners) in groups)
                if (o == outer)
                    foreach (var inner in inners)
                        foreach (var col in uniqueCols)
                            if (leafBucket.TryGetValue((outer, inner, col), out var b))
                                all.AddRange(b);
            return Reduce(all);
        }

        double ColTotal(string col)
        {
            var all = new List<double>();
            foreach (var (outer, inners) in groups)
                foreach (var inner in inners)
                    if (leafBucket.TryGetValue((outer, inner, col), out var b))
                        all.AddRange(b);
            return Reduce(all);
        }

        var grandTotal = Reduce(allValues);

        // ===== Write cells =====
        var (anchorCol, anchorRow) = ParseCellRef(position);
        var anchorColIdx = ColToIndex(anchorCol);
        var totalLabel = "总计";

        var ws = targetSheet.Worksheet
            ?? throw new InvalidOperationException("Target worksheet has no Worksheet element");
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            sheetData = new SheetData();
            ws.AppendChild(sheetData);
        }

        // Row 0 (caption row): data caption + col field caption.
        var captionRow = new Row { RowIndex = (uint)anchorRow };
        captionRow.AppendChild(MakeStringCell(anchorColIdx, anchorRow, dataFieldName));
        captionRow.AppendChild(MakeStringCell(anchorColIdx + 1, anchorRow, colFieldName));
        sheetData.AppendChild(captionRow);

        // Row 1 (header row): row label header + col labels + grand total.
        var headerRowIdx = anchorRow + 1;
        var headerRow = new Row { RowIndex = (uint)headerRowIdx };
        // The row-label header in compact mode is intentionally just "Row Labels"
        // when there are 2+ row fields, since one column has to represent both
        // levels. Excel's localized auto-caption will overlay this if a
        // RowHeaderCaption attribute isn't set; we set it to the OUTER field's
        // header name (the most informative single label) elsewhere.
        headerRow.AppendChild(MakeStringCell(anchorColIdx, headerRowIdx, headers[outerFieldIdx]));
        for (int c = 0; c < uniqueCols.Count; c++)
            headerRow.AppendChild(MakeStringCell(anchorColIdx + 1 + c, headerRowIdx, uniqueCols[c]));
        headerRow.AppendChild(MakeStringCell(anchorColIdx + 1 + uniqueCols.Count, headerRowIdx, totalLabel));
        sheetData.AppendChild(headerRow);

        // Data rows: alternate outer subtotal + leaf rows in display order.
        int currentRow = anchorRow + 2;
        foreach (var (outer, inners) in groups)
        {
            // Outer subtotal row.
            var subRow = new Row { RowIndex = (uint)currentRow };
            subRow.AppendChild(MakeStringCell(anchorColIdx, currentRow, outer));
            for (int c = 0; c < uniqueCols.Count; c++)
            {
                var v = OuterSubtotal(outer, uniqueCols[c]);
                if (v != 0 || HasAnyValueInOuterCol(outer, uniqueCols[c], groups, leafBucket))
                    subRow.AppendChild(MakeNumericCell(anchorColIdx + 1 + c, currentRow, v));
            }
            subRow.AppendChild(MakeNumericCell(anchorColIdx + 1 + uniqueCols.Count, currentRow, OuterRowTotal(outer)));
            sheetData.AppendChild(subRow);
            currentRow++;

            // Leaf rows for each existing (outer, inner) combo.
            foreach (var inner in inners)
            {
                var leafRow = new Row { RowIndex = (uint)currentRow };
                leafRow.AppendChild(MakeStringCell(anchorColIdx, currentRow, inner));
                for (int c = 0; c < uniqueCols.Count; c++)
                {
                    var v = LeafCell(outer, inner, uniqueCols[c]);
                    if (!double.IsNaN(v))
                        leafRow.AppendChild(MakeNumericCell(anchorColIdx + 1 + c, currentRow, v));
                }
                leafRow.AppendChild(MakeNumericCell(anchorColIdx + 1 + uniqueCols.Count, currentRow, LeafRowTotal(outer, inner)));
                sheetData.AppendChild(leafRow);
                currentRow++;
            }
        }

        // Grand total row.
        var grandRow = new Row { RowIndex = (uint)currentRow };
        grandRow.AppendChild(MakeStringCell(anchorColIdx, currentRow, totalLabel));
        for (int c = 0; c < uniqueCols.Count; c++)
            grandRow.AppendChild(MakeNumericCell(anchorColIdx + 1 + c, currentRow, ColTotal(uniqueCols[c])));
        grandRow.AppendChild(MakeNumericCell(anchorColIdx + 1 + uniqueCols.Count, currentRow, grandTotal));
        sheetData.AppendChild(grandRow);

        // Page filter cells reuse the single-row path's logic — same shape, same
        // layout above the table. RenderPivotIntoSheet handles them; we don't
        // duplicate the code, but if the user really needs filters with 2 row
        // fields, they should still get rendered. v4 candidate to factor out.
        // (Currently filters on multi-row pivots will write the page filter
        // markers in the pivot definition but no visible filter cells above
        // the table. Same warning is emitted.)
        if (filterFieldIndices != null && filterFieldIndices.Count > 0)
        {
            var requiredHeadroom = filterFieldIndices.Count + 1;
            if (anchorRow > requiredHeadroom)
            {
                var firstFilterRow = anchorRow - requiredHeadroom;
                for (int fi = 0; fi < filterFieldIndices.Count; fi++)
                {
                    var fIdx = filterFieldIndices[fi];
                    if (fIdx < 0 || fIdx >= headers.Length) continue;
                    var rowIdx = firstFilterRow + fi;
                    var filterRow = new Row { RowIndex = (uint)rowIdx };
                    filterRow.AppendChild(MakeStringCell(anchorColIdx, rowIdx, headers[fIdx]));
                    filterRow.AppendChild(MakeStringCell(anchorColIdx + 1, rowIdx, "(All)"));
                    sheetData.InsertAt(filterRow, fi);
                }
            }
        }

        ws.Save();
    }

    /// <summary>
    /// Helper for the multi-row renderer: returns true if the (outer, col) pair
    /// has at least one non-empty leaf bucket. Used to decide whether to write
    /// a 0-valued subtotal cell or skip it entirely (Excel writes nothing rather
    /// than a literal 0 for genuinely empty (outer, col) intersections).
    /// </summary>
    private static bool HasAnyValueInOuterCol(string outer, string col,
        List<(string outer, List<string> inners)> groups,
        Dictionary<(string o, string i, string c), List<double>> leafBucket)
    {
        foreach (var (o, inners) in groups)
        {
            if (o != outer) continue;
            foreach (var inner in inners)
                if (leafBucket.TryGetValue((outer, inner, col), out var b) && b.Count > 0)
                    return true;
        }
        return false;
    }

    /// <summary>
    /// Build an inline-string cell. We use inline strings (t="inlineStr" + &lt;is&gt;)
    /// rather than the SharedStringTable because the renderer is self-contained
    /// and adding entries to the SST would require coordinating with whatever
    /// other handler code touches the workbook's strings — out of scope for v1.
    /// </summary>
    private static Cell MakeStringCell(int colIdx, int rowIdx, string text)
    {
        return new Cell
        {
            CellReference = $"{IndexToCol(colIdx)}{rowIdx}",
            DataType = CellValues.InlineString,
            InlineString = new InlineString(new Text(text ?? string.Empty))
        };
    }

    /// <summary>Numeric cell with the value serialized using invariant culture.</summary>
    private static Cell MakeNumericCell(int colIdx, int rowIdx, double value)
    {
        return new Cell
        {
            CellReference = $"{IndexToCol(colIdx)}{rowIdx}",
            CellValue = new CellValue(value.ToString("R", System.Globalization.CultureInfo.InvariantCulture))
        };
    }

    // ==================== Source Data Reader ====================

    private static (string[] headers, List<string[]> columnData) ReadSourceData(
        WorksheetPart sourceSheet, string sourceRef)
    {
        var ws = sourceSheet.Worksheet ?? throw new InvalidOperationException("Worksheet missing");
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData == null) return (Array.Empty<string>(), new List<string[]>());

        // Parse range "A1:D100"
        var parts = sourceRef.Replace("$", "").Split(':');
        if (parts.Length != 2) throw new ArgumentException($"Invalid source range: {sourceRef}");

        var (startCol, startRow) = ParseCellRef(parts[0]);
        var (endCol, endRow) = ParseCellRef(parts[1]);

        var startColIdx = ColToIndex(startCol);
        var endColIdx = ColToIndex(endCol);
        var colCount = endColIdx - startColIdx + 1;

        // Read all rows in range
        var rows = new List<string[]>();
        var sst = sourceSheet.OpenXmlPackage is SpreadsheetDocument doc
            ? doc.WorkbookPart?.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()
            : null;

        foreach (var row in sheetData.Elements<Row>())
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rowIdx < startRow || rowIdx > endRow) continue;

            var values = new string[colCount];
            foreach (var cell in row.Elements<Cell>())
            {
                var cellRef = cell.CellReference?.Value ?? "";
                var (cn, _) = ParseCellRef(cellRef);
                var ci = ColToIndex(cn) - startColIdx;
                if (ci < 0 || ci >= colCount) continue;

                values[ci] = GetCellText(cell, sst);
            }
            rows.Add(values);
        }

        if (rows.Count == 0) return (Array.Empty<string>(), new List<string[]>());

        // First row = headers (ensure no nulls)
        var headers = rows[0].Select(h => h ?? "").ToArray();
        // Remaining rows = data, transposed to column-major for cache
        var columnDataList = new List<string[]>();
        for (int c = 0; c < colCount; c++)
        {
            var colVals = new string[rows.Count - 1];
            for (int r = 1; r < rows.Count; r++)
                colVals[r - 1] = rows[r][c] ?? "";
            columnDataList.Add(colVals);
        }

        return (headers, columnDataList);
    }

    private static string GetCellText(Cell cell, SharedStringTablePart? sst)
    {
        // Handle InlineString cells (t="inlineStr") — used by openpyxl and some other tools
        if (cell.DataType?.Value == CellValues.InlineString)
            return cell.InlineString?.InnerText ?? "";

        var value = cell.CellValue?.Text ?? "";
        if (cell.DataType?.Value == CellValues.SharedString && sst?.SharedStringTable != null)
        {
            if (int.TryParse(value, out int idx))
            {
                var item = sst.SharedStringTable.Elements<SharedStringItem>().ElementAtOrDefault(idx);
                return item?.InnerText ?? value;
            }
        }
        return value;
    }

    // ==================== Cache Definition Builder ====================

    private static (PivotCacheDefinition def, bool[] fieldNumeric, Dictionary<string, int>[] fieldValueIndex)
        BuildCacheDefinition(
            string sourceSheetName, string sourceRef,
            string[] headers, List<string[]> columnData)
    {
        var recordCount = columnData.Count > 0 ? columnData[0].Length : 0;

        // refreshOnLoad=1 tells Excel to re-render the pivot from the cache when the
        // file is opened. We need this because officecli (a pure DOM library) does NOT
        // have a pivot computation engine — we cannot materialize the rendered cells
        // into sheetData ourselves. Real Excel/LibreOffice DO write rendered cells on
        // save (verified against pivot5.xlsx and pivot_dark1.xlsx fixtures), so opening
        // their files shows data immediately. Without refreshOnLoad, our pivot-only
        // sheet would render empty even though the cache and definition are valid.
        //
        // Trade-off: Excel may prompt for trust before refreshing, and consumers that
        // do not implement refresh (POI, third-party parsers) will still see an empty
        // sheet. The proper long-term fix is a built-in render engine; this flag is
        // the lowest-cost workaround until that lands.
        var cacheDef = new PivotCacheDefinition
        {
            CreatedVersion = 3,
            MinRefreshableVersion = 3,
            RefreshedVersion = 3,
            RecordCount = (uint)recordCount,
            RefreshOnLoad = true
        };

        // CacheSource -> WorksheetSource
        var cacheSource = new CacheSource { Type = SourceValues.Worksheet };
        cacheSource.AppendChild(new WorksheetSource
        {
            Reference = sourceRef,
            Sheet = sourceSheetName
        });
        cacheDef.AppendChild(cacheSource);

        // CacheFields — also build per-field metadata used to write records:
        //   - fieldNumeric[i]: true if field i is numeric (records emit <n v=".."/>)
        //   - fieldValueIndex[i]: value→sharedItems index map for non-numeric fields
        //     (records emit <x v="N"/> referencing this index)
        var fieldNumeric = new bool[headers.Length];
        var fieldValueIndex = new Dictionary<string, int>[headers.Length];

        var cacheFields = new CacheFields { Count = (uint)headers.Length };
        for (int i = 0; i < headers.Length; i++)
        {
            var fieldName = string.IsNullOrEmpty(headers[i]) ? $"Column{i + 1}" : headers[i];
            var values = i < columnData.Count ? columnData[i] : Array.Empty<string>();
            cacheFields.AppendChild(BuildCacheField(fieldName, values, out fieldNumeric[i], out fieldValueIndex[i]));
        }
        cacheDef.AppendChild(cacheFields);

        return (cacheDef, fieldNumeric, fieldValueIndex);
    }

    private static CacheField BuildCacheField(
        string name, string[] values, out bool isNumeric, out Dictionary<string, int> valueIndex)
    {
        var field = new CacheField { Name = name, NumberFormatId = 0u };
        isNumeric = values.Length > 0 && values.All(v =>
            string.IsNullOrEmpty(v) || double.TryParse(v, System.Globalization.CultureInfo.InvariantCulture, out _));
        valueIndex = new Dictionary<string, int>(StringComparer.Ordinal);

        var sharedItems = new SharedItems();

        // MIXED strategy — verified against Microsoft's own pivot5.xlsx (in
        // OPEN-XML-SDK test fixtures, authored by real Excel):
        //
        //   • Numeric fields: emit ONLY containsNumber/minValue/maxValue metadata,
        //     no enumerated items, no count attribute. Records reference values
        //     directly via <n v="..."/>.
        //   • String fields: enumerate every unique value as <s v="..."/> with
        //     count attribute. Records reference them by index via <x v="N"/>.
        //
        // I previously experimented with LibreOffice's uniform strategy (always
        // enumerate, always index-reference), but Microsoft's actual format is
        // the mixed one — and matching the real Excel format is the safest bet
        // for round-trip compatibility. The uniform strategy is technically valid
        // OOXML but introduces an asymmetry that Excel handles less reliably
        // (numeric data fields with item enumeration have failed to render in
        // testing, even though the file passes schema validation).
        if (isNumeric && values.Any(v => !string.IsNullOrEmpty(v)))
        {
            var nums = values.Where(v => !string.IsNullOrEmpty(v))
                .Select(v => double.Parse(v, System.Globalization.CultureInfo.InvariantCulture)).ToArray();
            sharedItems.ContainsSemiMixedTypes = false;
            sharedItems.ContainsString = false;
            sharedItems.ContainsNumber = true;
            sharedItems.MinValue = nums.Min();
            sharedItems.MaxValue = nums.Max();
            // No items enumerated, no count — records emit <n v="..."/> directly.
        }
        else
        {
            var uniqueValues = values
                .Where(v => !string.IsNullOrEmpty(v))
                .Distinct()
                .OrderBy(v => v, StringComparer.Ordinal)
                .ToList();
            sharedItems.Count = (uint)uniqueValues.Count;
            for (int i = 0; i < uniqueValues.Count; i++)
            {
                var v = uniqueValues[i];
                sharedItems.AppendChild(new StringItem { Val = v });
                if (!valueIndex.ContainsKey(v))
                    valueIndex[v] = i;
            }
        }

        field.AppendChild(sharedItems);
        return field;
    }

    // ==================== Cache Records Builder ====================

    /// <summary>
    /// Build pivotCacheRecords using the MIXED strategy verified against Microsoft's
    /// own pivot5.xlsx test fixture:
    ///
    ///   <r>
    ///     <x v="0"/>     <!-- string field, references sharedItems[0] -->
    ///     <x v="2"/>     <!-- string field, references sharedItems[2] -->
    ///     <n v="702"/>   <!-- numeric field, value written directly -->
    ///     <m/>           <!-- empty/missing value -->
    ///   </r>
    ///
    /// String fields use indexed references (<x v="N"/>) into the per-field
    /// sharedItems list; numeric fields use NumberItem (<n v="V"/>) directly,
    /// because their cacheField only carries min/max metadata, not enumerated items.
    /// </summary>
    private static PivotCacheRecords BuildCacheRecords(
        List<string[]> columnData, bool[] fieldNumeric, Dictionary<string, int>[] fieldValueIndex)
    {
        var recordCount = columnData.Count > 0 ? columnData[0].Length : 0;
        var fieldCount = columnData.Count;
        var records = new PivotCacheRecords { Count = (uint)recordCount };

        for (int r = 0; r < recordCount; r++)
        {
            var record = new PivotCacheRecord();
            for (int f = 0; f < fieldCount; f++)
            {
                var v = columnData[f][r];
                if (string.IsNullOrEmpty(v))
                {
                    record.AppendChild(new MissingItem());
                }
                else if (fieldNumeric[f])
                {
                    record.AppendChild(new NumberItem
                    {
                        Val = double.Parse(v, System.Globalization.CultureInfo.InvariantCulture)
                    });
                }
                else if (fieldValueIndex[f].TryGetValue(v, out var idx))
                {
                    // FieldItem = <x v="N"/> in OpenXml SDK, references sharedItems[N].
                    record.AppendChild(new FieldItem { Val = (uint)idx });
                }
                else
                {
                    // Defensive: value missing from the per-field index map. Should
                    // not occur since the map is built from the same columnData;
                    // emit <m/> rather than a dangling reference.
                    record.AppendChild(new MissingItem());
                }
            }
            records.AppendChild(record);
        }

        return records;
    }

    // ==================== Pivot Table Definition Builder ====================

    private static PivotTableDefinition BuildPivotTableDefinition(
        string name, uint cacheId, string position,
        string[] headers, List<string[]> columnData,
        List<int> rowFieldIndices, List<int> colFieldIndices,
        List<int> filterFieldIndices, List<(int idx, string func, string name)> valueFields,
        string styleName)
    {
        var pivotDef = new PivotTableDefinition
        {
            Name = name,
            CacheId = cacheId,
            DataCaption = "Values",
            CreatedVersion = 3,
            MinRefreshableVersion = 3,
            UpdatedVersion = 3,
            ApplyNumberFormats = false,
            ApplyBorderFormats = false,
            ApplyFontFormats = false,
            ApplyPatternFormats = false,
            ApplyAlignmentFormats = false,
            ApplyWidthHeightFormats = true,
            UseAutoFormatting = true,
            ItemPrintTitles = true,
            MultipleFieldFilters = false,
            Indent = 0u,
            // outline + outlineData are emitted by both Microsoft Excel (pivot5.xlsx)
            // and LibreOffice (pivot_dark1.xlsx). They select the "outline" layout —
            // the default presentation where row labels stack into one column. Without
            // these, Excel falls back to a layout that's not fully wired through and
            // refuses to render the data area.
            Outline = true,
            OutlineData = true,
            // Caption attributes — when present, Excel uses these strings instead
            // of its locale-default "Row Labels" / "Column Labels" / "Grand Total".
            // Without these the rendered cells we wrote into sheetData ("地区",
            // "产品", "总计") get visually overlaid by Excel's English defaults
            // because the pivot's caption layer takes precedence over cell content
            // when the corresponding caption attribute is empty/missing.
            RowHeaderCaption = rowFieldIndices.Count > 0 ? headers[rowFieldIndices[0]] : "Rows",
            ColumnHeaderCaption = colFieldIndices.Count > 0 ? headers[colFieldIndices[0]] : "Columns",
            GrandTotalCaption = "总计"
        };

        // Use typed property setters to ensure correct schema order

        // Compute the pivot's geometry (range + offsets) via shared helper, so the
        // initial CreatePivotTable path and the post-Set RebuildFieldAreas path
        // produce identical results.
        var geom = ComputePivotGeometry(
            position, columnData, rowFieldIndices, colFieldIndices, valueFields);
        pivotDef.Location = new Location
        {
            Reference = geom.RangeRef,
            FirstHeaderRow = 1u,
            FirstDataRow = valueFields.Count > 1 ? 3u : 2u,
            FirstDataColumn = (uint)geom.RowLabelCols
        };

        // Page filters: presence is signalled by the <pageFields> element + the
        // pivotField axis="axisPage" marker, both written further down. ECMA-376
        // also defines optional rowPageCount / colPageCount attributes here, but
        // OpenXml SDK 3.3.0 does not model them and rejects them as unknown
        // during schema validation. Excel recognizes the filter without them
        // (verified empirically and in pivot_dark1.xlsx, which has filters but
        // no page count attributes). Tracked as a v2 polish item if any consumer
        // turns out to require them.

        // PivotFields — one per source column
        var pivotFields = new PivotFields { Count = (uint)headers.Length };
        for (int i = 0; i < headers.Length; i++)
        {
            var pf = new PivotField { ShowAll = false };
            var values = i < columnData.Count ? columnData[i] : Array.Empty<string>();
            var isNumeric = values.Length > 0 && values.All(v =>
                string.IsNullOrEmpty(v) || double.TryParse(v, System.Globalization.CultureInfo.InvariantCulture, out _));

            if (rowFieldIndices.Contains(i))
            {
                pf.Axis = PivotTableAxisValues.AxisRow;
                if (!isNumeric) AppendFieldItems(pf, values);
            }
            else if (colFieldIndices.Contains(i))
            {
                pf.Axis = PivotTableAxisValues.AxisColumn;
                if (!isNumeric) AppendFieldItems(pf, values);
            }
            else if (filterFieldIndices.Contains(i))
            {
                pf.Axis = PivotTableAxisValues.AxisPage;
                if (!isNumeric) AppendFieldItems(pf, values);
            }
            else if (valueFields.Any(vf => vf.idx == i))
            {
                pf.DataField = true;
            }

            pivotFields.AppendChild(pf);
        }
        pivotDef.PivotFields = pivotFields;

        // RowFields — the synthetic <field x="-2"/> sentinel for multiple data
        // fields belongs to whichever axis (rows or columns) actually displays
        // the data field labels. The default is dataOnRows=false, so multi-data
        // labels go in COLUMNS — meaning the sentinel appears in colFields, NOT
        // rowFields. Only add the sentinel here when there are no col fields and
        // therefore data must flow in the row dimension.
        if (rowFieldIndices.Count > 0)
        {
            var rf = new RowFields();
            foreach (var idx in rowFieldIndices)
                rf.AppendChild(new Field { Index = idx });
            if (valueFields.Count > 1 && colFieldIndices.Count == 0)
                rf.AppendChild(new Field { Index = -2 });
            rf.Count = (uint)rf.Elements<Field>().Count();
            pivotDef.RowFields = rf;
        }

        // RowItems — describes the row-label layout. Without this, Excel renders only the
        // pivot's drop-down chrome but no actual data cells (the layout we observed earlier).
        // Pattern verified against LibreOffice's pivot_dark1.xlsx test fixture:
        //   <rowItems count="K+1">
        //     <i><x/></i>            <-- index 0 (shorthand: omit v attribute)
        //     <i><x v="1"/></i>      <-- index 1
        //     ...
        //     <i t="grand"><x/></i>  <-- grand total row
        //   </rowItems>
        // The <x v="N"/> values index into the corresponding pivotField's <items> list,
        // which we already populate via AppendFieldItems in BuildPivotTableDefinition above.
        // Single row field only: multi-row-field cartesian-product layout is a v2 concern.
        if (rowFieldIndices.Count > 0)
            pivotDef.RowItems = (RowItems)BuildAxisItems(rowFieldIndices, columnData, isRow: true, dataFieldCount: 1);

        // ColumnFields — when there are 2+ data fields, append the synthetic
        // <field x="-2"/> sentinel that tells Excel "data field labels go in
        // the column dimension here". Verified against multi_data_authored.xlsx:
        // a 1-row × 1-col × 2-data pivot writes <colFields count="2">
        // <field x="1"/><field x="-2"/></colFields>. Without this sentinel
        // Excel still opens the file but renders the K data fields stacked
        // incorrectly. RebuildFieldAreas already handles this; the initial
        // build path was missing the sentinel.
        if (colFieldIndices.Count > 0 || valueFields.Count > 1)
        {
            var cf = new ColumnFields();
            foreach (var idx in colFieldIndices)
                cf.AppendChild(new Field { Index = idx });
            if (valueFields.Count > 1)
                cf.AppendChild(new Field { Index = -2 });
            cf.Count = (uint)cf.Elements<Field>().Count();
            pivotDef.ColumnFields = cf;
        }

        // ColumnItems — same shape as RowItems but for the column-label layout.
        // Even when there are NO column fields, ECMA-376 requires a <colItems> with one
        // empty <i/> placeholder; LibreOffice's writeRowColumnItems empty-case branch
        // (xepivotxml.cxx:1008-1014) writes exactly that.
        pivotDef.ColumnItems = (ColumnItems)BuildAxisItems(
            colFieldIndices, columnData, isRow: false, dataFieldCount: valueFields.Count);

        // PageFields (filters)
        if (filterFieldIndices.Count > 0)
        {
            var pf = new PageFields { Count = (uint)filterFieldIndices.Count };
            foreach (var idx in filterFieldIndices)
                pf.AppendChild(new PageField { Field = idx, Hierarchy = -1 });
            pivotDef.PageFields = pf;
        }

        // DataFields
        if (valueFields.Count > 0)
        {
            var df = new DataFields { Count = (uint)valueFields.Count };
            foreach (var (idx, func, displayName) in valueFields)
            {
                // BaseField/BaseItem: Excel ignores these when ShowDataAs is normal,
                // but LibreOffice and Excel both emit them unconditionally on every
                // dataField (verified against pivot_dark1.xlsx and other LO fixtures).
                // Following the verified pattern rather than my earlier "omit them"
                // theory — being closer to what real producers write reduces the risk
                // of triggering picky consumers.
                df.AppendChild(new DataField
                {
                    Name = displayName,
                    Field = (uint)idx,
                    Subtotal = ParseSubtotal(func),
                    BaseField = 0,
                    BaseItem = 0u
                });
            }
            pivotDef.DataFields = df;
        }

        // Style
        pivotDef.PivotTableStyle = new PivotTableStyle
        {
            Name = styleName,
            ShowRowHeaders = true,
            ShowColumnHeaders = true,
            ShowRowStripes = false,
            ShowColumnStripes = false,
            ShowLastColumn = true
        };

        return pivotDef;
    }

    /// <summary>
    /// Build the &lt;rowItems&gt; or &lt;colItems&gt; layout block. Excel uses this to
    /// know how to expand row/column labels in the rendered pivot.
    ///
    /// Single data field (K=1):
    ///   <rowItems count="K+1">
    ///     <i><x/></i>            <-- index 0 (shorthand: omit v)
    ///     <i><x v="1"/></i>
    ///     ...
    ///     <i t="grand"><x/></i>
    ///   </rowItems>
    ///
    /// Multi-data field on the column axis (K>1, only used for ColumnItems):
    ///   <colItems count="(L+1)*K">
    ///     <i><x/><x/></i>                     <-- col label 0, data field 0
    ///     <i r="1" i="1"><x v="1"/></i>       <-- col label 0, data field 1 (r=1 = repeat prev x)
    ///     <i><x v="1"/><x/></i>               <-- col label 1, data field 0
    ///     <i r="1" i="1"><x v="1"/></i>       <-- col label 1, data field 1
    ///     ...
    ///     <i t="grand"><x/></i>               <-- grand total, data field 0
    ///     <i t="grand" i="1"><x/></i>         <-- grand total, data field 1
    ///   </colItems>
    /// Verified against multi_data_authored.xlsx (a 1×1×2 pivot from real Excel).
    ///
    /// Empty axis: single &lt;i/&gt; placeholder (LibreOffice writeRowColumnItems
    /// empty-case branch in xepivotxml.cxx:1008-1014).
    ///
    /// Limitation: still only single-axis-field cases are correct. Multi-row-field
    /// cartesian-product layouts need a deeper expansion tracked as v2.
    /// </summary>
    private static OpenXmlElement BuildAxisItems(
        List<int> fieldIndices, List<string[]> columnData, bool isRow, int dataFieldCount = 1)
    {
        OpenXmlCompositeElement container = isRow
            ? new RowItems()
            : new ColumnItems();

        // Empty axis: write a single empty <i/>. LibreOffice does this unconditionally
        // when there's nothing to render — Excel needs the placeholder. When there are
        // multiple data fields on the column axis but no col field, we still need
        // K entries (one per data field) instead of just one — handled below.
        if (fieldIndices.Count == 0)
        {
            if (!isRow && dataFieldCount > 1)
            {
                // Data-only column axis: K entries, each marked with i="d".
                for (int d = 0; d < dataFieldCount; d++)
                {
                    var item = new RowItem();
                    if (d > 0) item.Index = (uint)d;
                    item.AppendChild(new MemberPropertyIndex());
                    container.AppendChild(item);
                }
                SetAxisCount(container, dataFieldCount);
            }
            else
            {
                container.AppendChild(new RowItem());
                SetAxisCount(container, 1);
            }
            return container;
        }

        // Multi-row case (N>=2 row fields, only used for RowItems).
        //
        // Pattern (verified against multi_row_authored.xlsx with 2 row fields,
        // where the user manually built a pivot with rows=地区,城市):
        //   For each outer value O in display order:
        //     <i><x v="O"/></i>                     <- outer subtotal row (1 x child)
        //     For each inner value I that exists in (O, *):
        //       <i r="1"><x v="I"/></i>             <- leaf row (r=1 = repeat outer)
        //   <i t="grand"><x/></i>                   <- final grand total
        //
        // The "1 x child only" form is treated by Excel as the outer-level
        // subtotal row (it shows aggregate across all this outer's inners). Leaf
        // rows use r='1' to mean "the first 1 member is inherited from the
        // previous row" (the outer index), so the leaf only needs its own inner
        // index as a single x child.
        //
        // This implementation supports exactly N=2 row fields. N>=3 would need a
        // recursive expansion at every non-leaf level — tracked as v4.
        if (isRow && fieldIndices.Count >= 2)
        {
            return BuildMultiRowItems(fieldIndices, columnData);
        }

        // Single field: one <i> per unique value, then a grand-total entry.
        // Multi-field is not yet supported — fall back to the first field's values
        // so the file is at least openable; rendering will be incomplete.
        var fieldIdx = fieldIndices[0];
        if (fieldIdx < 0 || fieldIdx >= columnData.Count)
        {
            container.AppendChild(new RowItem());
            SetAxisCount(container, 1);
            return container;
        }

        var uniqueCount = columnData[fieldIdx]
            .Where(v => !string.IsNullOrEmpty(v))
            .Distinct()
            .Count();

        // Multi-data on column axis: each col label gets K entries, then K grand totals.
        // The first entry per col label has TWO <x> children (col index + data field 0);
        // subsequent entries use r="1" to repeat the col index and bump i to the data
        // field number.
        if (!isRow && dataFieldCount > 1)
        {
            for (int i = 0; i < uniqueCount; i++)
            {
                // Entry for data field 0: <i><x v="i"/><x v="0"/></i>
                var first = new RowItem();
                if (i == 0)
                    first.AppendChild(new MemberPropertyIndex());
                else
                    first.AppendChild(new MemberPropertyIndex { Val = i });
                first.AppendChild(new MemberPropertyIndex());
                container.AppendChild(first);

                // Entries for data fields 1..K-1: <i r="1" i="d"><x v="d"/></i>
                for (int d = 1; d < dataFieldCount; d++)
                {
                    var rep = new RowItem
                    {
                        RepeatedItemCount = 1u,
                        Index = (uint)d
                    };
                    if (d == 0)
                        rep.AppendChild(new MemberPropertyIndex());
                    else
                        rep.AppendChild(new MemberPropertyIndex { Val = d });
                    container.AppendChild(rep);
                }
            }

            // Grand totals: K entries marked t="grand", with i=d for d>0.
            for (int d = 0; d < dataFieldCount; d++)
            {
                var gt = new RowItem { ItemType = ItemValues.Grand };
                if (d > 0) gt.Index = (uint)d;
                gt.AppendChild(new MemberPropertyIndex());
                container.AppendChild(gt);
            }

            SetAxisCount(container, uniqueCount * dataFieldCount + dataFieldCount);
            return container;
        }

        // Single-data layout (original path): K data rows + 1 grand total.
        for (int i = 0; i < uniqueCount; i++)
        {
            var item = new RowItem();
            if (i == 0)
                item.AppendChild(new MemberPropertyIndex());
            else
                item.AppendChild(new MemberPropertyIndex { Val = i });
            container.AppendChild(item);
        }

        // Grand total entry — always present in the default layout.
        var grandTotal = new RowItem { ItemType = ItemValues.Grand };
        grandTotal.AppendChild(new MemberPropertyIndex());
        container.AppendChild(grandTotal);

        SetAxisCount(container, uniqueCount + 1);
        return container;
    }

    /// <summary>
    /// Compute the (outer → ordered list of inners) groupings for a 2-row-field
    /// pivot. Only (outer, inner) combinations that actually appear in the
    /// source data are included — Excel does not enumerate empty cartesian
    /// cells in compact mode. Output is sorted by ordinal: outer keys first,
    /// then each outer's inner list. Used by both BuildMultiRowItems (XML
    /// rowItems generation) and the renderer (cell layout).
    /// </summary>
    private static List<(string outer, List<string> inners)> BuildOuterInnerGroups(
        int outerFieldIdx, int innerFieldIdx, List<string[]> columnData)
    {
        var outerVals = columnData[outerFieldIdx];
        var innerVals = columnData[innerFieldIdx];
        var n = outerVals.Length;

        var seen = new HashSet<(string, string)>();
        var combos = new List<(string outer, string inner)>();
        for (int i = 0; i < n; i++)
        {
            var ov = outerVals[i];
            var iv = innerVals[i];
            if (string.IsNullOrEmpty(ov) || string.IsNullOrEmpty(iv)) continue;
            if (seen.Add((ov, iv)))
                combos.Add((ov, iv));
        }

        // Sort by ordinal so display order matches the pivotField items list,
        // which is built with the same StringComparer.Ordinal sort. This is what
        // keeps the rowItems indices in sync with the rendered cell labels.
        return combos
            .GroupBy(c => c.outer, StringComparer.Ordinal)
            .OrderBy(g => g.Key, StringComparer.Ordinal)
            .Select(g => (g.Key, g.Select(c => c.inner)
                .OrderBy(v => v, StringComparer.Ordinal).ToList()))
            .ToList();
    }

    /// <summary>
    /// Build the &lt;rowItems&gt; element for a 2-row-field pivot. Emits one
    /// outer-subtotal row per unique outer value plus one leaf row per
    /// (outer, inner) combination that exists in the data, then the grand
    /// total. See BuildOuterInnerGroups for the grouping logic.
    /// </summary>
    private static OpenXmlElement BuildMultiRowItems(
        List<int> fieldIndices, List<string[]> columnData)
    {
        var container = new RowItems();
        if (fieldIndices.Count < 2 || fieldIndices[0] >= columnData.Count || fieldIndices[1] >= columnData.Count)
        {
            container.AppendChild(new RowItem());
            container.Count = 1u;
            return container;
        }

        var outerIdx = fieldIndices[0];
        var innerIdx = fieldIndices[1];
        var groups = BuildOuterInnerGroups(outerIdx, innerIdx, columnData);

        // Pre-compute the value→pivotField-items-index map for both row fields.
        // The pivotField items list is built with StringComparer.Ordinal in
        // AppendFieldItems below, so we mirror the same ordering here to keep
        // the indices consistent.
        var outerOrder = columnData[outerIdx]
            .Where(v => !string.IsNullOrEmpty(v))
            .Distinct()
            .OrderBy(v => v, StringComparer.Ordinal)
            .Select((v, i) => (v, i))
            .ToDictionary(t => t.v, t => t.i, StringComparer.Ordinal);
        var innerOrder = columnData[innerIdx]
            .Where(v => !string.IsNullOrEmpty(v))
            .Distinct()
            .OrderBy(v => v, StringComparer.Ordinal)
            .Select((v, i) => (v, i))
            .ToDictionary(t => t.v, t => t.i, StringComparer.Ordinal);

        int count = 0;
        foreach (var (outer, inners) in groups)
        {
            // Outer subtotal row: <i><x v="outerIdx"/></i>
            var outerEntry = new RowItem();
            var outerPivIdx = outerOrder[outer];
            if (outerPivIdx == 0)
                outerEntry.AppendChild(new MemberPropertyIndex());
            else
                outerEntry.AppendChild(new MemberPropertyIndex { Val = outerPivIdx });
            container.AppendChild(outerEntry);
            count++;

            // Leaf rows for each inner of this outer: <i r="1"><x v="innerIdx"/></i>
            foreach (var inner in inners)
            {
                var leafEntry = new RowItem { RepeatedItemCount = 1u };
                var innerPivIdx = innerOrder[inner];
                if (innerPivIdx == 0)
                    leafEntry.AppendChild(new MemberPropertyIndex());
                else
                    leafEntry.AppendChild(new MemberPropertyIndex { Val = innerPivIdx });
                container.AppendChild(leafEntry);
                count++;
            }
        }

        // Grand total row.
        var grand = new RowItem { ItemType = ItemValues.Grand };
        grand.AppendChild(new MemberPropertyIndex());
        container.AppendChild(grand);
        count++;

        container.Count = (uint)count;
        return container;
    }

    /// <summary>Set the count attribute on RowItems / ColumnItems uniformly.</summary>
    private static void SetAxisCount(OpenXmlCompositeElement container, int count)
    {
        if (container is RowItems ri) ri.Count = (uint)count;
        else if (container is ColumnItems ci) ci.Count = (uint)count;
    }

    private static void AppendFieldItems(PivotField pf, string[] values)
    {
        var unique = values.Where(v => !string.IsNullOrEmpty(v)).Distinct().OrderBy(v => v).ToList();
        var items = new Items { Count = (uint)(unique.Count + 1) };
        for (int i = 0; i < unique.Count; i++)
            items.AppendChild(new Item { Index = (uint)i });
        items.AppendChild(new Item { ItemType = ItemValues.Default }); // grand total
        pf.AppendChild(items);
    }

    // ==================== Readback ====================

    internal static void ReadPivotTableProperties(PivotTableDefinition pivotDef, DocumentNode node)
    {
        if (pivotDef.Name?.HasValue == true) node.Format["name"] = pivotDef.Name.Value;
        if (pivotDef.CacheId?.HasValue == true) node.Format["cacheId"] = pivotDef.CacheId.Value;

        var location = pivotDef.GetFirstChild<Location>();
        if (location?.Reference?.HasValue == true) node.Format["location"] = location.Reference.Value;

        // Count fields
        var pivotFields = pivotDef.GetFirstChild<PivotFields>();
        if (pivotFields != null)
            node.Format["fieldCount"] = pivotFields.Elements<PivotField>().Count();

        // Row fields
        var rowFields = pivotDef.RowFields;
        if (rowFields != null)
        {
            var indices = rowFields.Elements<Field>().Where(f => f.Index?.Value >= 0).Select(f => f.Index!.Value).ToList();
            if (indices.Count > 0)
                node.Format["rowFields"] = string.Join(",", indices);
        }

        // Column fields
        var colFields = pivotDef.ColumnFields;
        if (colFields != null)
        {
            var indices = colFields.Elements<Field>().Where(f => f.Index?.Value >= 0).Select(f => f.Index!.Value).ToList();
            if (indices.Count > 0)
                node.Format["colFields"] = string.Join(",", indices);
        }

        // Page/filter fields
        var pageFields = pivotDef.PageFields;
        if (pageFields != null)
        {
            var indices = pageFields.Elements<PageField>().Select(f => f.Field?.Value ?? -1).Where(v => v >= 0).ToList();
            if (indices.Count > 0)
                node.Format["filterFields"] = string.Join(",", indices);
        }

        // Data fields (use typed property for reliable access)
        var dataFields = pivotDef.DataFields;
        if (dataFields != null)
        {
            var dfList = dataFields.Elements<DataField>().ToList();
            node.Format["dataFieldCount"] = dfList.Count;
            for (int i = 0; i < dfList.Count; i++)
            {
                var df = dfList[i];
                var dfName = df.Name?.Value ?? "";
                var dfFunc = df.Subtotal?.InnerText ?? "sum";
                var dfField = df.Field?.Value ?? 0;
                node.Format[$"dataField{i + 1}"] = $"{dfName}:{dfFunc}:{dfField}";
            }
        }

        // Style
        var styleInfo = pivotDef.PivotTableStyle;
        if (styleInfo?.Name?.HasValue == true)
            node.Format["style"] = styleInfo.Name.Value;
    }

    internal static List<string> SetPivotTableProperties(PivotTablePart pivotPart, Dictionary<string, string> properties)
    {
        var unsupported = new List<string>();
        var pivotDef = pivotPart.PivotTableDefinition;
        if (pivotDef == null) { unsupported.AddRange(properties.Keys); return unsupported; }

        // Collect field-area properties separately — they require a coordinated rebuild
        var fieldAreaProps = new Dictionary<string, string>();

        foreach (var (key, value) in properties)
        {
            switch (key.ToLowerInvariant())
            {
                case "name":
                    pivotDef.Name = value;
                    break;
                case "style":
                {
                    pivotDef.PivotTableStyle = new PivotTableStyle
                    {
                        Name = value,
                        ShowRowHeaders = true,
                        ShowColumnHeaders = true,
                        ShowRowStripes = false,
                        ShowColumnStripes = false,
                        ShowLastColumn = true
                    };
                    break;
                }
                case "rows":
                case "cols" or "columns":
                case "values":
                case "filters":
                    fieldAreaProps[key.ToLowerInvariant() == "columns" ? "cols" : key.ToLowerInvariant()] = value;
                    break;
                default:
                    unsupported.Add(key);
                    break;
            }
        }

        // If any field areas were specified, rebuild them
        if (fieldAreaProps.Count > 0)
            RebuildFieldAreas(pivotPart, pivotDef, fieldAreaProps);

        pivotDef.Save();
        return unsupported;
    }

    /// <summary>
    /// Rebuild pivot table field areas (rows, cols, values, filters).
    /// For areas not specified in changes, preserves the current assignment.
    /// Two-layer update: (1) PivotField.Axis/DataField, (2) RowFields/ColumnFields/PageFields/DataFields.
    /// </summary>
    private static void RebuildFieldAreas(PivotTablePart pivotPart, PivotTableDefinition pivotDef,
        Dictionary<string, string> changes)
    {
        // Get headers from cache definition
        var cachePart = pivotPart.GetPartsOfType<PivotTableCacheDefinitionPart>().FirstOrDefault();
        if (cachePart?.PivotCacheDefinition == null) return;

        var cacheFields = cachePart.PivotCacheDefinition.GetFirstChild<CacheFields>();
        if (cacheFields == null) return;

        var headers = cacheFields.Elements<CacheField>().Select(cf => cf.Name?.Value ?? "").ToArray();
        if (headers.Length == 0) return;

        // Read current assignments for areas NOT being changed
        var currentRows = ReadCurrentFieldIndices(pivotDef.RowFields?.Elements<Field>(), f => f.Index?.Value ?? -1);
        var currentCols = ReadCurrentFieldIndices(pivotDef.ColumnFields?.Elements<Field>(), f => f.Index?.Value ?? -1);
        var currentFilters = ReadCurrentFieldIndices(pivotDef.PageFields?.Elements<PageField>(), f => f.Field?.Value ?? -1);
        var currentValues = ReadCurrentDataFields(pivotDef.DataFields);

        // Parse new assignments (or keep current)
        // If user specified a non-empty value but nothing resolved, warn via stderr
        var rowFieldIndices = changes.ContainsKey("rows")
            ? ParseFieldListWithWarning(changes, "rows", headers)
            : currentRows;
        var colFieldIndices = changes.ContainsKey("cols")
            ? ParseFieldListWithWarning(changes, "cols", headers)
            : currentCols;
        var filterFieldIndices = changes.ContainsKey("filters")
            ? ParseFieldListWithWarning(changes, "filters", headers)
            : currentFilters;
        var valueFields = changes.ContainsKey("values")
            ? ParseValueFieldsWithWarning(changes, "values", headers)
            : currentValues;

        // Layer 1: Reset all PivotField axis/dataField, then re-assign
        var pivotFields = pivotDef.PivotFields;
        if (pivotFields == null) return;

        var pfList = pivotFields.Elements<PivotField>().ToList();
        for (int i = 0; i < pfList.Count; i++)
        {
            var pf = pfList[i];
            // Clear axis and dataField
            pf.Axis = null;
            pf.DataField = null;
            pf.RemoveAllChildren<Items>();

            // Determine if this field's cache data is numeric (for Items generation)
            var isNumeric = IsFieldNumeric(cacheFields, i);

            if (rowFieldIndices.Contains(i))
            {
                pf.Axis = PivotTableAxisValues.AxisRow;
                if (!isNumeric) AppendFieldItemsFromCache(pf, cacheFields, i);
            }
            else if (colFieldIndices.Contains(i))
            {
                pf.Axis = PivotTableAxisValues.AxisColumn;
                if (!isNumeric) AppendFieldItemsFromCache(pf, cacheFields, i);
            }
            else if (filterFieldIndices.Contains(i))
            {
                pf.Axis = PivotTableAxisValues.AxisPage;
                if (!isNumeric) AppendFieldItemsFromCache(pf, cacheFields, i);
            }
            else if (valueFields.Any(vf => vf.idx == i))
            {
                pf.DataField = true;
            }
        }

        // Layer 2: Rebuild area reference lists
        // RowFields
        if (rowFieldIndices.Count > 0)
        {
            var rf = new RowFields { Count = (uint)rowFieldIndices.Count };
            foreach (var idx in rowFieldIndices)
                rf.AppendChild(new Field { Index = idx });
            // -2 sentinel for multiple value fields displayed in rows
            if (valueFields.Count > 1 && colFieldIndices.Count == 0)
            {
                rf.AppendChild(new Field { Index = -2 });
                rf.Count = (uint)rf.Elements<Field>().Count();
            }
            pivotDef.RowFields = rf;
        }
        else
        {
            pivotDef.RowFields = null;
        }

        // ColumnFields
        if (colFieldIndices.Count > 0 || valueFields.Count > 1)
        {
            var cf = new ColumnFields();
            foreach (var idx in colFieldIndices)
                cf.AppendChild(new Field { Index = idx });
            // -2 sentinel for multiple value fields in columns
            if (valueFields.Count > 1)
                cf.AppendChild(new Field { Index = -2 });
            cf.Count = (uint)cf.Elements<Field>().Count();
            pivotDef.ColumnFields = cf;
        }
        else
        {
            pivotDef.ColumnFields = null;
        }

        // PageFields (filters)
        if (filterFieldIndices.Count > 0)
        {
            var pf = new PageFields { Count = (uint)filterFieldIndices.Count };
            foreach (var idx in filterFieldIndices)
                pf.AppendChild(new PageField { Field = idx, Hierarchy = -1 });
            pivotDef.PageFields = pf;
        }
        else
        {
            pivotDef.PageFields = null;
        }

        // DataFields
        if (valueFields.Count > 0)
        {
            var df = new DataFields { Count = (uint)valueFields.Count };
            foreach (var (idx, func, displayName) in valueFields)
            {
                // BaseField/BaseItem: Excel ignores these when ShowDataAs is normal,
                // but LibreOffice and Excel both emit them unconditionally on every
                // dataField (verified against pivot_dark1.xlsx and other LO fixtures).
                // Following the verified pattern rather than my earlier "omit them"
                // theory — being closer to what real producers write reduces the risk
                // of triggering picky consumers.
                df.AppendChild(new DataField
                {
                    Name = displayName,
                    Field = (uint)idx,
                    Subtotal = ParseSubtotal(func),
                    BaseField = 0,
                    BaseItem = 0u
                });
            }
            pivotDef.DataFields = df;
        }
        else
        {
            pivotDef.DataFields = null;
        }

        // Update Location with the full new geometry — range, offsets, FirstDataCol —
        // not just FirstDataColumn. The previous incremental approach left a stale
        // range covering the old layout, which made Excel render only the original
        // bounds even when fields were added or removed.
        var oldLocation = pivotDef.Location;
        var oldRangeRef = oldLocation?.Reference?.Value;
        var anchorRefForGeometry = oldRangeRef?.Split(':')[0]
            ?? oldLocation?.Reference?.Value
            ?? "A1";

        // Reconstruct columnData from the cache so the geometry helper and the
        // renderer below can compute new extents without re-reading the source sheet.
        var (cacheHeaders, cacheColumnData) = ReadColumnDataFromCache(
            cachePart.PivotCacheDefinition,
            cachePart.GetPartsOfType<PivotTableCacheRecordsPart>().FirstOrDefault()?.PivotCacheRecords);

        var newGeom = ComputePivotGeometry(
            anchorRefForGeometry, cacheColumnData, rowFieldIndices, colFieldIndices, valueFields);

        pivotDef.Location = new Location
        {
            Reference = newGeom.RangeRef,
            FirstHeaderRow = 1u,
            FirstDataRow = 2u,
            FirstDataColumn = (uint)newGeom.RowLabelCols
        };

        // Rebuild RowItems / ColumnItems for the new field assignments. The previous
        // configuration's row/col layout no longer matches; without these the rendered
        // skeleton would still describe the old shape.
        if (rowFieldIndices.Count > 0)
            pivotDef.RowItems = (RowItems)BuildAxisItems(rowFieldIndices, cacheColumnData, isRow: true, dataFieldCount: 1);
        else
            pivotDef.RowItems = null;
        pivotDef.ColumnItems = (ColumnItems)BuildAxisItems(
            colFieldIndices, cacheColumnData, isRow: false, dataFieldCount: valueFields.Count);

        // Refresh caption attributes — they pin to the row/col field's header name,
        // so reassigning fields means the visible caption changes too.
        pivotDef.RowHeaderCaption = rowFieldIndices.Count > 0 ? cacheHeaders[rowFieldIndices[0]] : "Rows";
        pivotDef.ColumnHeaderCaption = colFieldIndices.Count > 0 ? cacheHeaders[colFieldIndices[0]] : "Columns";

        // Re-render the materialized cells. Find the host worksheet via the pivot
        // part's parent — pivotPart is owned by exactly one WorksheetPart so this
        // is unambiguous in v1 (no shared pivot tables).
        var hostSheet = pivotPart.GetParentParts().OfType<WorksheetPart>().FirstOrDefault();
        if (hostSheet != null)
        {
            var ws = hostSheet.Worksheet;
            var sheetData = ws?.GetFirstChild<SheetData>();
            if (ws != null && sheetData != null)
            {
                // Clear the OLD rendered cells before drawing the new layout. The
                // new geometry might be smaller (fewer cols → stale right-hand cells)
                // OR larger (more rows → safe overwrite), so we always wipe the union
                // of old and new bounds. Old range first, then new range — the new
                // render writes into the cleared area immediately after.
                if (!string.IsNullOrEmpty(oldRangeRef))
                    ClearPivotRangeCells(sheetData, oldRangeRef);
                ClearPivotRangeCells(sheetData, newGeom.RangeRef);

                RenderPivotIntoSheet(
                    hostSheet, anchorRefForGeometry, cacheHeaders, cacheColumnData,
                    rowFieldIndices, colFieldIndices, valueFields, filterFieldIndices);
            }
        }
    }

    private static List<int> ReadCurrentFieldIndices<T>(IEnumerable<T>? elements, Func<T, int> getIndex)
    {
        if (elements == null) return new List<int>();
        return elements.Select(getIndex).Where(i => i >= 0).ToList();
    }

    private static List<(int idx, string func, string name)> ReadCurrentDataFields(DataFields? dataFields)
    {
        if (dataFields == null) return new List<(int, string, string)>();
        return dataFields.Elements<DataField>().Select(df => (
            idx: (int)(df.Field?.Value ?? 0),
            func: df.Subtotal?.InnerText ?? "sum",
            name: df.Name?.Value ?? ""
        )).ToList();
    }

    private static bool IsFieldNumeric(CacheFields cacheFields, int index)
    {
        var cf = cacheFields.Elements<CacheField>().ElementAtOrDefault(index);
        var sharedItems = cf?.GetFirstChild<SharedItems>();
        if (sharedItems == null) return false;
        return sharedItems.ContainsNumber?.Value == true && sharedItems.ContainsString?.Value != true;
    }

    private static void AppendFieldItemsFromCache(PivotField pf, CacheFields cacheFields, int index)
    {
        var cf = cacheFields.Elements<CacheField>().ElementAtOrDefault(index);
        var sharedItems = cf?.GetFirstChild<SharedItems>();
        var count = sharedItems?.Elements<StringItem>().Count() ?? 0;
        if (count == 0) return;

        var items = new Items { Count = (uint)(count + 1) };
        for (int i = 0; i < count; i++)
            items.AppendChild(new Item { Index = (uint)i });
        items.AppendChild(new Item { ItemType = ItemValues.Default }); // grand total
        pf.AppendChild(items);
    }

    // ==================== Parse Helpers ====================

    private static List<int> ParseFieldListWithWarning(Dictionary<string, string> props, string key, string[] headers)
    {
        var result = ParseFieldList(props, key, headers);
        if (result.Count == 0 && props.TryGetValue(key, out var value) && !string.IsNullOrEmpty(value))
        {
            var available = string.Join(", ", headers.Where(h => !string.IsNullOrEmpty(h)));
            Console.Error.WriteLine($"WARNING: No matching fields for {key}={value}. Available: {available}");
        }
        return result;
    }

    private static List<(int idx, string func, string name)> ParseValueFieldsWithWarning(
        Dictionary<string, string> props, string key, string[] headers)
    {
        var result = ParseValueFields(props, key, headers);
        if (result.Count == 0 && props.TryGetValue(key, out var value) && !string.IsNullOrEmpty(value))
        {
            var available = string.Join(", ", headers.Where(h => !string.IsNullOrEmpty(h)));
            Console.Error.WriteLine($"WARNING: No matching fields for {key}={value}. Available: {available}");
        }
        return result;
    }

    private static List<int> ParseFieldList(Dictionary<string, string> props, string key, string[] headers)
    {
        if (!props.TryGetValue(key, out var value) || string.IsNullOrEmpty(value))
            return new List<int>();

        return value.Split(',').Select(f =>
        {
            var name = f.Trim();
            // Try as column index first
            if (int.TryParse(name, out var idx)) return idx;
            // Try as header name
            for (int i = 0; i < headers.Length; i++)
                if (headers[i] != null && headers[i].Equals(name, StringComparison.OrdinalIgnoreCase)) return i;
            return -1;
        }).Where(i => i >= 0 && i < headers.Length).ToList();
    }

    private static List<(int idx, string func, string name)> ParseValueFields(
        Dictionary<string, string> props, string key, string[] headers)
    {
        if (!props.TryGetValue(key, out var value) || string.IsNullOrEmpty(value))
            return new List<(int, string, string)>();

        var result = new List<(int idx, string func, string name)>();
        foreach (var spec in value.Split(','))
        {
            // Format: "FieldName:func" or "FieldName" (default sum)
            var parts = spec.Trim().Split(':');
            var fieldName = parts[0].Trim();
            var func = parts.Length > 1 ? parts[1].Trim().ToLowerInvariant() : "sum";

            int fieldIdx = -1;
            if (int.TryParse(fieldName, out var idx)) fieldIdx = idx;
            else
            {
                for (int i = 0; i < headers.Length; i++)
                    if (headers[i] != null && headers[i].Equals(fieldName, StringComparison.OrdinalIgnoreCase)) { fieldIdx = i; break; }
            }

            if (fieldIdx >= 0 && fieldIdx < headers.Length)
            {
                var displayName = $"{char.ToUpper(func[0])}{func[1..]} of {headers[fieldIdx]}";
                result.Add((fieldIdx, func, displayName));
            }
        }
        return result;
    }

    private static DataConsolidateFunctionValues ParseSubtotal(string func)
    {
        return func.ToLowerInvariant() switch
        {
            "sum" => DataConsolidateFunctionValues.Sum,
            "count" => DataConsolidateFunctionValues.Count,
            "average" or "avg" => DataConsolidateFunctionValues.Average,
            "max" => DataConsolidateFunctionValues.Maximum,
            "min" => DataConsolidateFunctionValues.Minimum,
            "product" => DataConsolidateFunctionValues.Product,
            "stddev" => DataConsolidateFunctionValues.StandardDeviation,
            "var" => DataConsolidateFunctionValues.Variance,
            _ => DataConsolidateFunctionValues.Sum
        };
    }

    private static (string col, int row) ParseCellRef(string cellRef)
    {
        int i = 0;
        while (i < cellRef.Length && char.IsLetter(cellRef[i])) i++;
        var col = cellRef[..i].ToUpperInvariant();
        var row = int.TryParse(cellRef[i..], out var r) ? r : 1;
        return (col, row);
    }

    private static int ColToIndex(string col)
    {
        int result = 0;
        foreach (var c in col.ToUpperInvariant())
            result = result * 26 + (c - 'A' + 1);
        return result;
    }

    private static string IndexToCol(int index)
    {
        // Inverse of ColToIndex (1-based: A=1, Z=26, AA=27, ...)
        var sb = new System.Text.StringBuilder();
        while (index > 0)
        {
            int rem = (index - 1) % 26;
            sb.Insert(0, (char)('A' + rem));
            index = (index - 1) / 26;
        }
        return sb.ToString();
    }

    /// <summary>
    /// Multiply the cardinality (distinct non-empty values) of each field in the
    /// given index list. Used to size the pivot table's rendered area for the
    /// Location.ref range. Returns 1 when the list is empty (so layout math stays
    /// safe in pivots that have only column fields, only row fields, etc.).
    /// </summary>
    private static int ProductOfUniqueValues(List<int> fieldIndices, List<string[]> columnData)
    {
        if (fieldIndices.Count == 0) return 1;
        int product = 1;
        foreach (var idx in fieldIndices)
        {
            if (idx < 0 || idx >= columnData.Count) continue;
            var unique = columnData[idx].Where(v => !string.IsNullOrEmpty(v)).Distinct().Count();
            product *= Math.Max(1, unique);
        }
        return product;
    }
}
