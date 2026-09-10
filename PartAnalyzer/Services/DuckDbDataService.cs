using System.Data;
using System.Data.Common;
using System.IO;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using DuckDB.NET.Data;
using PartAnalyzer.Models;

namespace PartAnalyzer.Services;

public sealed class DuckDbDataService : IDisposable
{
    private const string SourceTableName = "source_rows";
    private const string PartNumberComparisonColumnName = "part_number_comparison_value";
    private const string CategoryComparisonColumnName = "category_comparison_value";
    private const string ManufacturerComparisonColumnName = "manufacturer_comparison_value";
    private const string PartSummaryViewName = "part_summary";
    private const string PartMetadataTableName = "part_metadata";
    private static readonly Regex SafeColumnNamePattern = new(@"^col_\d{3,}$", RegexOptions.Compiled);
    private readonly string _databasePath;
    private DuckDBConnection? _connection;
    private DataLoadResult? _currentDataset;
    private PartAnalysisMapping? _currentAnalysisMapping;

    public DuckDbDataService(string? databasePath = null)
    {
        _databasePath = string.IsNullOrWhiteSpace(databasePath)
            ? Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "PartAnalyzer",
                "Temp",
                "session.duckdb")
            : databasePath;
        CleanupStaleSessionFiles();
    }

    public string DatabasePath => _databasePath;

    public DataLoadResult? CurrentDataset => _currentDataset;

    public Task<DataLoadResult> LoadWorksheetAsync(
        WorksheetProcessingSession session,
        CancellationToken cancellationToken = default)
    {
        return Task.Run(() => LoadWorksheet(session, cancellationToken), cancellationToken);
    }

    public Task<DataPageResult> GetPageAsync(
        int pageNumber,
        int pageSize,
        IReadOnlyList<FilterCriteria>? filters = null,
        QuerySort? sort = null,
        CancellationToken cancellationToken = default)
    {
        return Task.Run(() => GetPage(pageNumber, pageSize, filters ?? Array.Empty<FilterCriteria>(), sort, cancellationToken), cancellationToken);
    }

    public Task<IReadOnlyList<FilterDefinition>> DiscoverSourceFiltersAsync(CancellationToken cancellationToken = default)
    {
        return Task.Run(() => DiscoverSourceFilters(cancellationToken), cancellationToken);
    }

    public Task<IReadOnlyList<FilterValueOption>> GetFilterValueOptionsAsync(
        FilterDefinition definition,
        string? searchText,
        int limit = 100,
        CancellationToken cancellationToken = default)
    {
        return Task.Run(() => GetFilterValueOptions(definition, searchText, limit, cancellationToken), cancellationToken);
    }

    public Task<IReadOnlyList<FilterValueOption>> GetSessionFilterValuesAsync(SessionFilterColumn column, string? searchText, int limit = 500, CancellationToken cancellationToken = default)
    {
        return Task.Run(() => GetSessionFilterValues(column, searchText, limit, cancellationToken), cancellationToken);
    }

    public Task<IReadOnlyList<FilterValueOption>> GetManufacturerPreviewAsync(CancellationToken cancellationToken = default)
    {
        return Task.Run(() => GetSessionFilterValues(SessionFilterColumn.Manufacturer, null, int.MaxValue, cancellationToken), cancellationToken);
    }

    public Task<IReadOnlyList<int>> GetMatchingExcelRowNumbersAsync(FilterSelectionSnapshot selections, CancellationToken cancellationToken = default)
    {
        return Task.Run(() => GetMatchingExcelRowNumbers(selections, cancellationToken), cancellationToken);
    }

    public Task<IReadOnlyList<UnavailableFilterSelection>> GetUnavailableSelectionsAsync(FilterSelectionSnapshot selections, CancellationToken cancellationToken = default)
    {
        return Task.Run(() => GetUnavailableSelections(selections, cancellationToken), cancellationToken);
    }

    public void ClearDataset()
    {
        if (_connection is null)
        {
            _currentDataset = null;
            return;
        }

        using var command = _connection.CreateCommand();
        command.CommandText = $"""
            DROP VIEW IF EXISTS {PartSummaryViewName};
            DROP TABLE IF EXISTS {PartMetadataTableName};
            DROP TABLE IF EXISTS {SourceTableName};
            """;
        command.ExecuteNonQuery();
        _currentDataset = null;
        _currentAnalysisMapping = null;
    }

    public Task<PartAnalysisResult> AnalyzePartsAsync(
        PartAnalysisMapping mapping,
        CancellationToken cancellationToken = default)
    {
        return Task.Run(() => AnalyzeParts(mapping, cancellationToken), cancellationToken);
    }

    public Task<DataPageResult> GetPartSummaryPageAsync(
        int pageNumber,
        int pageSize,
        IReadOnlyList<FilterCriteria>? filters = null,
        QuerySort? sort = null,
        CancellationToken cancellationToken = default)
    {
        return Task.Run(() => GetPartSummaryPage(pageNumber, pageSize, filters ?? Array.Empty<FilterCriteria>(), sort, cancellationToken), cancellationToken);
    }

    public Task<DataTable> GetPartDetailRowsAsync(
        string partIdentifier,
        CancellationToken cancellationToken = default)
    {
        return Task.Run(() => GetPartDetailRows(partIdentifier, cancellationToken), cancellationToken);
    }

    public Task SetPartReviewedAsync(
        string partIdentifier,
        bool reviewed,
        CancellationToken cancellationToken = default)
    {
        return Task.Run(() => SetPartReviewed(partIdentifier, reviewed, cancellationToken), cancellationToken);
    }

    public Task<ExportResult> ExportWorkbookAsync(
        ExportRequest request,
        CancellationToken cancellationToken = default)
    {
        return Task.Run(() => ExportWorkbook(request, cancellationToken), cancellationToken);
    }

    private DataLoadResult LoadWorksheet(
        WorksheetProcessingSession session,
        CancellationToken cancellationToken)
    {
        var workbookPath = session.WorkbookPath;
        var worksheetName = session.WorksheetName;
        using var measurement = PerformanceLogger.Measure("Worksheet import", $"{Path.GetFileName(workbookPath)}::{worksheetName}");
        if (string.IsNullOrWhiteSpace(workbookPath) || !File.Exists(workbookPath))
        {
            throw new DuckDbDataException("The source workbook is no longer available.");
        }

        if (session.EffectiveColumns.Count == 0)
        {
            throw new DuckDbDataException("The selected worksheet does not contain importable columns.");
        }

        try
        {
            EnsureOpenConnection(resetDatabase: true);

            using var stream = File.Open(workbookPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == worksheetName)
                ?? throw new DuckDbDataException("The selected worksheet could not be found in the workbook.");

            var columns = CreateImportedColumns(session.EffectiveColumns);

            using var transaction = _connection!.BeginTransaction();
            try
            {
                using (PerformanceLogger.Measure("DuckDB table creation/load setup", $"{columns.Count} columns"))
                {
                    DropAndCreateSourceTable(columns, transaction);
                }

                (long ImportedRows, int SkippedBlankRows) importCounts;
                using (PerformanceLogger.Measure("DuckDB row insertion", $"{session.EligibleRowCount} eligible rows"))
                {
                    importCounts = InsertRows(worksheet, session, columns, transaction, cancellationToken);
                }
                transaction.Commit();

                var importedRowCount = GetImportedRowCount();
                if (importedRowCount != importCounts.ImportedRows)
                {
                    throw new DuckDbDataException("The imported row count could not be verified.");
                }

                _currentDataset = new DataLoadResult
                {
                    WorkbookPath = workbookPath,
                    WorksheetName = worksheetName,
                    HeaderRowNumber = session.HeaderRowNumber,
                    ImportedRowCount = importedRowCount,
                    ImportedColumnCount = columns.Count,
                    SkippedBlankRowCount = importCounts.SkippedBlankRows,
                    Columns = columns,
                    ProcessingSession = session
                };
                _currentAnalysisMapping = null;

                return _currentDataset;
            }
            catch
            {
                transaction.Rollback();
                _currentDataset = null;
                throw;
            }
        }
        catch (DuckDbDataException)
        {
            throw;
        }
        catch (OperationCanceledException)
        {
            _currentDataset = null;
            throw;
        }
        catch (Exception ex)
        {
            _currentDataset = null;
            throw new DuckDbDataException("The worksheet data could not be loaded into the local session database.", ex);
        }
    }

    private DataPageResult GetPage(int pageNumber, int pageSize, IReadOnlyList<FilterCriteria> filters, QuerySort? sort, CancellationToken cancellationToken)
    {
        using var measurement = PerformanceLogger.Measure("Raw page query", $"page={pageNumber}; filters={filters.Count(filter => filter.IsActive)}");
        if (_connection is null || _currentDataset is null)
        {
            return new DataPageResult { PageNumber = 1, PageSize = pageSize };
        }

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            var filterSql = BuildRawWhereClause(filters);
            var totalRows = GetCount($"SELECT COUNT(*) FROM {SourceTableName}{filterSql.WhereClause};", filterSql.Parameters);
            var totalPages = totalRows == 0 ? 1 : (int)Math.Ceiling(totalRows / (double)pageSize);
            var safePageNumber = Math.Clamp(pageNumber, 1, totalPages);
            var offset = (safePageNumber - 1) * pageSize;
            var orderBy = BuildRawOrderBy(sort);

            using var command = _connection.CreateCommand();
            command.CommandText = CreatePageQuery(_currentDataset.Columns, filterSql.WhereClause, orderBy);
            foreach (var value in filterSql.Parameters)
            {
                AddParameter(command, value ?? DBNull.Value);
            }

            AddParameter(command, pageSize);
            AddParameter(command, offset);

            using var reader = command.ExecuteReader();
            var table = new DataTable();
            foreach (var column in _currentDataset.Columns)
            {
                table.Columns.Add(column.InternalColumnName, typeof(string));
            }

            while (reader.Read())
            {
                var row = table.NewRow();
                for (var columnIndex = 0; columnIndex < _currentDataset.Columns.Count; columnIndex++)
                {
                    row[columnIndex] = reader.IsDBNull(columnIndex) ? DBNull.Value : reader.GetString(columnIndex);
                }

                table.Rows.Add(row);
            }

            return new DataPageResult
            {
                Rows = table,
                PageNumber = safePageNumber,
                PageSize = pageSize,
                TotalRows = totalRows,
                TotalPages = totalPages
            };
        }
        catch (Exception ex)
        {
            throw new DuckDbDataException("The current data page could not be read from the local session database.", ex);
        }
    }

    private IReadOnlyList<FilterDefinition> DiscoverSourceFilters(CancellationToken cancellationToken)
    {
        using var measurement = PerformanceLogger.Measure("Filter metadata/profile");
        if (_connection is null || _currentDataset is null)
        {
            return Array.Empty<FilterDefinition>();
        }

        var definitions = new List<FilterDefinition>();
        foreach (var column in _currentDataset.Columns)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ValidateMappedColumn(column);
            var distinctCount = ExecuteLong($"""
                SELECT COUNT(DISTINCT {column.InternalColumnName})
                FROM {SourceTableName}
                WHERE {column.InternalColumnName} IS NOT NULL;
                """);
            var blankCount = ExecuteLong($"""
                SELECT COUNT(*)
                FROM {SourceTableName}
                WHERE {column.InternalColumnName} IS NULL OR NULLIF(TRIM({column.InternalColumnName}), '') IS NULL;
                """);
            var nonBlankCount = ExecuteLong($"""
                SELECT COUNT(*)
                FROM {SourceTableName}
                WHERE NULLIF(TRIM({column.InternalColumnName}), '') IS NOT NULL;
                """);
            var numericCount = ExecuteLong($"""
                SELECT COUNT(*)
                FROM {SourceTableName}
                WHERE NULLIF(TRIM({column.InternalColumnName}), '') IS NOT NULL
                  AND TRY_CAST({column.InternalColumnName} AS DOUBLE) IS NOT NULL;
                """);
            var leadingZeroCount = ExecuteLong($"""
                SELECT COUNT(*)
                FROM {SourceTableName}
                WHERE LENGTH(TRIM({column.InternalColumnName})) > 1
                  AND TRIM({column.InternalColumnName}) LIKE '0%'
                  AND TRY_CAST({column.InternalColumnName} AS DOUBLE) IS NOT NULL;
                """);

            var kind = ChooseFilterKind(distinctCount, nonBlankCount, numericCount, leadingZeroCount);
            definitions.Add(new FilterDefinition
            {
                Id = column.InternalColumnName,
                DisplayName = column.GridHeader,
                InternalColumnName = column.InternalColumnName,
                DisplayOrder = column.DisplayOrder,
                Kind = kind,
                Target = FilterTarget.SourceRows,
                DistinctNonBlankCount = distinctCount,
                BlankCount = blankCount
            });
        }

        return definitions;
    }

    private IReadOnlyList<FilterValueOption> GetFilterValueOptions(
        FilterDefinition definition,
        string? searchText,
        int limit,
        CancellationToken cancellationToken)
    {
        using var measurement = PerformanceLogger.Measure("Distinct-value query", $"{definition.DisplayName}; search={!string.IsNullOrWhiteSpace(searchText)}; limit={limit}");
        if (_connection is null || _currentDataset is null)
        {
            return Array.Empty<FilterValueOption>();
        }

        ValidateFilterDefinition(definition);
        cancellationToken.ThrowIfCancellationRequested();

        using var command = _connection.CreateCommand();
        var hasSearch = !string.IsNullOrWhiteSpace(searchText);
        command.CommandText = $"""
            SELECT {definition.InternalColumnName} AS value, COUNT(*) AS value_count
            FROM {SourceTableName}
            WHERE {(hasSearch ? $"LOWER({definition.InternalColumnName}) LIKE ? ESCAPE '\\'" : "1 = 1")}
            GROUP BY {definition.InternalColumnName}
            ORDER BY value IS NULL, value
            LIMIT ?;
            """;
        if (hasSearch)
        {
            AddParameter(command, $"%{EscapeLikeValue(searchText!.Trim().ToLowerInvariant())}%");
        }

        AddParameter(command, limit);

        using var reader = command.ExecuteReader();
        var values = new List<FilterValueOption>();
        while (reader.Read())
        {
            var value = reader.IsDBNull(0) ? null : reader.GetString(0);
            values.Add(new FilterValueOption
            {
                Value = value,
                DisplayValue = string.IsNullOrWhiteSpace(value) ? "(Blank)" : value,
                Count = Convert.ToInt64(reader.GetValue(1))
            });
        }

        return values;
    }

    private IReadOnlyList<FilterValueOption> GetSessionFilterValues(SessionFilterColumn column, string? searchText, int limit, CancellationToken cancellationToken)
    {
        if (_connection is null || _currentDataset?.ProcessingSession is null) return Array.Empty<FilterValueOption>();
        cancellationToken.ThrowIfCancellationRequested();
        var expression = GetSessionFilterValueExpression(column);
        using var command = _connection.CreateCommand();
        var hasSearch = !string.IsNullOrWhiteSpace(searchText);
        command.CommandText = $"SELECT {expression} AS value, COUNT(*) AS value_count FROM {SourceTableName} WHERE " +
            (hasSearch ? $"LOWER({expression}) LIKE ? ESCAPE '\\'" : "1 = 1") +
            $" GROUP BY {expression} ORDER BY value LIMIT ?;";
        if (hasSearch) AddParameter(command, $"%{EscapeLikeValue(searchText!.ToLowerInvariant())}%");
        AddParameter(command, limit);
        using var reader = command.ExecuteReader();
        var results = new List<FilterValueOption>();
        while (reader.Read())
        {
            var value = reader.IsDBNull(0) ? null : reader.GetString(0);
            results.Add(new FilterValueOption
            {
                Value = value,
                DisplayValue = value is null || value == FilterSelectionValues.BlankCategory ? "(Blank)" : value,
                Count = Convert.ToInt64(reader.GetValue(1))
            });
        }
        return results;
    }

    private IReadOnlyList<int> GetMatchingExcelRowNumbers(FilterSelectionSnapshot selections, CancellationToken cancellationToken)
    {
        if (_connection is null || _currentDataset?.ProcessingSession is null) return Array.Empty<int>();
        cancellationToken.ThrowIfCancellationRequested();
        var conditions = new List<string>();
        using var command = _connection.CreateCommand();
        foreach (var column in Enum.GetValues<SessionFilterColumn>())
        {
            var values = selections.ValuesFor(column);
            if (values.Count == 0) continue;
            var expression = GetSessionFilterExpression(column);
            var nonBlankValues = column == SessionFilterColumn.Category
                ? values.Where(value => value != FilterSelectionValues.BlankCategory).ToList()
                : values.ToList();
            var columnConditions = new List<string>();
            if (nonBlankValues.Count > 0)
            {
                columnConditions.Add($"{expression} IN ({string.Join(", ", nonBlankValues.Select(_ => "?"))})");
                foreach (var value in nonBlankValues) AddParameter(command, value);
            }
            if (column == SessionFilterColumn.Category && values.Contains(FilterSelectionValues.BlankCategory, StringComparer.Ordinal))
            {
                columnConditions.Add($"({expression} IS NULL OR {expression} = '')");
            }
            conditions.Add(columnConditions.Count == 1 ? columnConditions[0] : $"({string.Join(" OR ", columnConditions)})");
        }
        command.CommandText = $"SELECT _excel_row_number FROM {SourceTableName}" + (conditions.Count == 0 ? string.Empty : $" WHERE {string.Join(" AND ", conditions)}") + " ORDER BY _excel_row_number;";
        using var reader = command.ExecuteReader();
        var rows = new List<int>();
        while (reader.Read()) rows.Add(reader.GetInt32(0));
        return rows;
    }

    private IReadOnlyList<UnavailableFilterSelection> GetUnavailableSelections(FilterSelectionSnapshot selections, CancellationToken cancellationToken)
    {
        var unavailable = new List<UnavailableFilterSelection>();
        foreach (var column in Enum.GetValues<SessionFilterColumn>())
        {
            var values = selections.ValuesFor(column);
            if (values.Count == 0) continue;
            cancellationToken.ThrowIfCancellationRequested();
            var available = new HashSet<string>(StringComparer.Ordinal);
            using var command = _connection!.CreateCommand();
            var expression = GetSessionFilterValueExpression(column);
            command.CommandText = $"SELECT DISTINCT {expression} FROM {SourceTableName} WHERE {expression} IN ({string.Join(", ", values.Select(_ => "?"))});";
            foreach (var value in values) AddParameter(command, value);
            using var reader = command.ExecuteReader();
            while (reader.Read()) if (!reader.IsDBNull(0)) available.Add(reader.GetString(0));
            var missing = values.Where(value => !available.Contains(value)).ToList();
            if (missing.Count > 0) unavailable.Add(new UnavailableFilterSelection { Column = column, Values = missing });
        }
        return unavailable;
    }

    private string GetSessionFilterExpression(SessionFilterColumn column)
    {
        if (_currentDataset?.ProcessingSession is null) throw new DuckDbDataException("Load a validated worksheet before filtering.");
        return column switch
        {
            SessionFilterColumn.PartNumber => PartNumberComparisonColumnName,
            SessionFilterColumn.Category => CategoryComparisonColumnName,
            SessionFilterColumn.Manufacturer => ManufacturerComparisonColumnName,
            _ => throw new DuckDbDataException("The selected filter column is not supported.")
        };
    }

    private string GetSessionFilterValueExpression(SessionFilterColumn column)
    {
        var expression = GetSessionFilterExpression(column);
        return column == SessionFilterColumn.Category
            ? $"COALESCE(NULLIF({expression}, ''), '{FilterSelectionValues.BlankCategory}')"
            : expression;
    }

    private PartAnalysisResult AnalyzeParts(PartAnalysisMapping mapping, CancellationToken cancellationToken)
    {
        using var measurement = PerformanceLogger.Measure("Part analysis", mapping.PrimaryPartIdentifier.DisplayHeader);
        if (_connection is null || _currentDataset is null)
        {
            throw new DuckDbDataException("Load worksheet data before analyzing parts.");
        }

        ValidateMappedColumn(mapping.PrimaryPartIdentifier);
        ValidateMappedColumn(mapping.Manufacturer);
        if (mapping.ManufacturerPartNumber is not null)
        {
            ValidateMappedColumn(mapping.ManufacturerPartNumber);
        }

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            using var command = _connection.CreateCommand();
            command.CommandText = $"""
                CREATE OR REPLACE TEMP VIEW {PartSummaryViewName} AS
                WITH normalized_rows AS (
                    SELECT
                        NULLIF(TRIM({mapping.PrimaryPartIdentifier.InternalColumnName}), '') AS part_identifier,
                        CASE
                            WHEN NULLIF(TRIM({mapping.Manufacturer.InternalColumnName}), '') IS NULL THEN NULL
                            ELSE LOWER(TRIM({mapping.Manufacturer.InternalColumnName}))
                        END AS normalized_manufacturer,
                        _excel_row_number
                    FROM {SourceTableName}
                )
                SELECT
                    part_identifier,
                    COUNT(*) AS part_row_count,
                    COUNT(DISTINCT normalized_manufacturer) AS manufacturer_count,
                    MIN(_excel_row_number) AS first_excel_row_number
                FROM normalized_rows
                WHERE part_identifier IS NOT NULL
                GROUP BY part_identifier;
                """;
            command.ExecuteNonQuery();
            EnsurePartMetadataTable();

            _currentAnalysisMapping = mapping;
            return new PartAnalysisResult
            {
                UniquePartCount = ExecuteLong($"SELECT COUNT(*) FROM {PartSummaryViewName};"),
                DuplicatePartCount = ExecuteLong($"SELECT COUNT(*) FROM {PartSummaryViewName} WHERE part_row_count > 1;"),
                MultipleManufacturerPartCount = ExecuteLong($"SELECT COUNT(*) FROM {PartSummaryViewName} WHERE manufacturer_count > 1;"),
                BlankPrimaryIdentifierRowCount = ExecuteLong($"""
                    SELECT COUNT(*)
                    FROM {SourceTableName}
                    WHERE NULLIF(TRIM({mapping.PrimaryPartIdentifier.InternalColumnName}), '') IS NULL;
                    """)
            };
        }
        catch (DuckDbDataException)
        {
            throw;
        }
        catch (Exception ex)
        {
            _currentAnalysisMapping = null;
            throw new DuckDbDataException("Part analysis could not be completed.", ex);
        }
    }

    private DataPageResult GetPartSummaryPage(int pageNumber, int pageSize, IReadOnlyList<FilterCriteria> filters, QuerySort? sort, CancellationToken cancellationToken)
    {
        using var measurement = PerformanceLogger.Measure("Grouped-page query", $"page={pageNumber}; filters={filters.Count(filter => filter.IsActive)}");
        if (_connection is null || _currentAnalysisMapping is null)
        {
            return new DataPageResult { PageNumber = 1, PageSize = pageSize };
        }

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            var filterSql = BuildGroupedWhereClause(filters);
            var totalRows = GetCount($"""
                SELECT COUNT(*)
                FROM {PartSummaryViewName} ps
                LEFT JOIN {PartMetadataTableName} pm ON pm.part_identifier = ps.part_identifier
                {filterSql.WhereClause};
                """, filterSql.Parameters);
            var totalPages = totalRows == 0 ? 1 : (int)Math.Ceiling(totalRows / (double)pageSize);
            var safePageNumber = Math.Clamp(pageNumber, 1, totalPages);
            var offset = (safePageNumber - 1) * pageSize;
            var orderBy = BuildGroupedOrderBy(sort);

            using var command = _connection.CreateCommand();
            command.CommandText = $"""
                SELECT
                    ps.part_identifier AS PartIdentifier,
                    ps.part_row_count AS PartRowCount,
                    ps.manufacturer_count AS ManufacturerCount,
                    CASE WHEN ps.part_row_count > 1 THEN 'Yes' ELSE 'No' END AS IsDuplicatePart,
                    CASE WHEN ps.manufacturer_count > 1 THEN 'Yes' ELSE 'No' END AS HasMultipleManufacturers,
                    COALESCE(pm.reviewed, false) AS Reviewed
                FROM {PartSummaryViewName} ps
                LEFT JOIN {PartMetadataTableName} pm ON pm.part_identifier = ps.part_identifier
                {filterSql.WhereClause}
                {orderBy}
                LIMIT ?
                OFFSET ?;
            """;
            foreach (var value in filterSql.Parameters)
            {
                AddParameter(command, value ?? DBNull.Value);
            }

            AddParameter(command, pageSize);
            AddParameter(command, offset);

            using var reader = command.ExecuteReader();
            return new DataPageResult
            {
                Rows = ReadPartSummaryTable(reader),
                PageNumber = safePageNumber,
                PageSize = pageSize,
                TotalRows = totalRows,
                TotalPages = totalPages
            };
        }
        catch (Exception ex)
        {
            throw new DuckDbDataException("The grouped parts page could not be read.", ex);
        }
    }

    private void SetPartReviewed(string partIdentifier, bool reviewed, CancellationToken cancellationToken)
    {
        if (_connection is null || _currentAnalysisMapping is null)
        {
            throw new DuckDbDataException("Analyze parts before marking a part reviewed.");
        }

        if (string.IsNullOrWhiteSpace(partIdentifier))
        {
            return;
        }

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            EnsurePartMetadataTable();
            using var command = _connection.CreateCommand();
            command.CommandText = $"""
                INSERT INTO {PartMetadataTableName} (part_identifier, reviewed)
                VALUES (?, ?)
                ON CONFLICT (part_identifier) DO UPDATE SET reviewed = excluded.reviewed;
                """;
            AddParameter(command, partIdentifier);
            AddParameter(command, reviewed);
            command.ExecuteNonQuery();
        }
        catch (Exception ex)
        {
            throw new DuckDbDataException("The reviewed state could not be saved in the local session.", ex);
        }
    }

    private ExportResult ExportWorkbook(ExportRequest request, CancellationToken cancellationToken)
    {
        using var measurement = PerformanceLogger.Measure("Excel export", $"{request.Mode}; {Path.GetFileName(request.DestinationPath)}");
        string? tempPath = null;
        if (_connection is null || _currentDataset is null)
        {
            throw new DuckDbDataException("Load worksheet data before exporting.");
        }

        if (string.IsNullOrWhiteSpace(request.DestinationPath))
        {
            throw new DuckDbDataException("Choose an export filename.");
        }

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            tempPath = CreateTemporaryExportPath(request.DestinationPath);
            using var workbook = new XLWorkbook();

            long sourceRows = 0;
            long groupedRows = 0;
            if (request.Mode is ExportMode.ProcessedWorkbook or ExportMode.AllSourceRows or ExportMode.CurrentFilteredRows)
            {
                var sourceFilters = request.Mode == ExportMode.CurrentFilteredRows
                    ? request.Filters
                    : Array.Empty<FilterCriteria>();
                sourceRows = AddProcessedDataSheet(workbook, sourceFilters, cancellationToken);
            }

            if (request.Mode is ExportMode.ProcessedWorkbook or ExportMode.GroupedPartSummary or ExportMode.CurrentFilteredGroupedParts)
            {
                var groupedFilters = request.Mode == ExportMode.CurrentFilteredGroupedParts
                    ? request.Filters
                    : Array.Empty<FilterCriteria>();
                groupedRows = AddPartSummarySheet(workbook, groupedFilters, cancellationToken);
            }

            workbook.SaveAs(tempPath);
            File.Move(tempPath, request.DestinationPath, overwrite: true);
            return new ExportResult
            {
                DestinationPath = request.DestinationPath,
                SourceRowCount = sourceRows,
                GroupedRowCount = groupedRows
            };
        }
        catch (DuckDbDataException)
        {
            TryDeleteFile(tempPath);
            throw;
        }
        catch (UnauthorizedAccessException ex)
        {
            TryDeleteFile(tempPath);
            throw new DuckDbDataException("The export file could not be written because access was denied.", ex);
        }
        catch (IOException ex)
        {
            TryDeleteFile(tempPath);
            throw new DuckDbDataException("The export file could not be written. Close it in Excel and try again.", ex);
        }
        catch (Exception ex)
        {
            TryDeleteFile(tempPath);
            throw new DuckDbDataException("The Excel export could not be completed.", ex);
        }
    }

    private long AddProcessedDataSheet(XLWorkbook workbook, IReadOnlyList<FilterCriteria> filters, CancellationToken cancellationToken)
    {
        var worksheet = workbook.Worksheets.Add("Processed Data");
        var headers = _currentDataset!.Columns
            .OrderBy(column => column.DisplayOrder)
            .Select(column => column.DisplayHeader)
            .Concat(new[]
            {
                "Duplicate?",
                "Multiple Manufacturer?",
                "Part Row Count",
                "Manufacturer Count",
                "Reviewed"
            })
            .ToList();
        WriteHeaders(worksheet, headers);

        using var command = _connection!.CreateCommand();
        var filterSql = BuildRawWhereClause(filters);
        command.CommandText = CreateProcessedDataExportQuery(filterSql.WhereClause);
        foreach (var value in filterSql.Parameters)
        {
            AddParameter(command, value ?? DBNull.Value);
        }

        using var reader = command.ExecuteReader();
        var rowNumber = 2;
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            for (var index = 0; index < _currentDataset.Columns.Count; index++)
            {
                worksheet.Cell(rowNumber, index + 1).SetValue(reader.IsDBNull(index) ? string.Empty : Convert.ToString(reader.GetValue(index)) ?? string.Empty);
                worksheet.Cell(rowNumber, index + 1).Style.NumberFormat.Format = "@";
            }

            var generatedStart = _currentDataset.Columns.Count + 1;
            WriteGeneratedBoolean(worksheet.Cell(rowNumber, generatedStart), reader, _currentDataset.Columns.Count);
            WriteGeneratedBoolean(worksheet.Cell(rowNumber, generatedStart + 1), reader, _currentDataset.Columns.Count + 1);
            WriteGeneratedNumber(worksheet.Cell(rowNumber, generatedStart + 2), reader, _currentDataset.Columns.Count + 2);
            WriteGeneratedNumber(worksheet.Cell(rowNumber, generatedStart + 3), reader, _currentDataset.Columns.Count + 3);
            WriteGeneratedBoolean(worksheet.Cell(rowNumber, generatedStart + 4), reader, _currentDataset.Columns.Count + 4);
            rowNumber++;
        }

        FormatWorksheet(worksheet, headers.Count, rowNumber - 1);
        return rowNumber - 2;
    }

    private long AddPartSummarySheet(XLWorkbook workbook, IReadOnlyList<FilterCriteria> filters, CancellationToken cancellationToken)
    {
        if (_currentAnalysisMapping is null)
        {
            throw new DuckDbDataException("Analyze parts before exporting a part summary.");
        }

        var worksheet = workbook.Worksheets.Add("Part Summary");
        var headers = new[]
        {
            "Part Identifier",
            "Part Row Count",
            "Manufacturer Count",
            "Duplicate?",
            "Multiple Manufacturer?",
            "Reviewed"
        };
        WriteHeaders(worksheet, headers);

        using var command = _connection!.CreateCommand();
        var filterSql = BuildGroupedWhereClause(filters);
        command.CommandText = CreatePartSummaryExportQuery(filterSql.WhereClause);
        foreach (var value in filterSql.Parameters)
        {
            AddParameter(command, value ?? DBNull.Value);
        }

        using var reader = command.ExecuteReader();
        var rowNumber = 2;
        while (reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();
            worksheet.Cell(rowNumber, 1).SetValue(reader.IsDBNull(0) ? string.Empty : reader.GetString(0));
            worksheet.Cell(rowNumber, 1).Style.NumberFormat.Format = "@";
            worksheet.Cell(rowNumber, 2).SetValue(Convert.ToInt64(reader.GetValue(1)));
            worksheet.Cell(rowNumber, 3).SetValue(Convert.ToInt64(reader.GetValue(2)));
            worksheet.Cell(rowNumber, 4).SetValue(Convert.ToBoolean(reader.GetValue(3)) ? "Yes" : "No");
            worksheet.Cell(rowNumber, 5).SetValue(Convert.ToBoolean(reader.GetValue(4)) ? "Yes" : "No");
            worksheet.Cell(rowNumber, 6).SetValue(Convert.ToBoolean(reader.GetValue(5)) ? "Yes" : "No");
            rowNumber++;
        }

        FormatWorksheet(worksheet, headers.Length, rowNumber - 1);
        return rowNumber - 2;
    }

    private DataTable GetPartDetailRows(string partIdentifier, CancellationToken cancellationToken)
    {
        using var measurement = PerformanceLogger.Measure("Selected-part detail query", partIdentifier);
        if (_connection is null || _currentAnalysisMapping is null)
        {
            return new DataTable();
        }

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            var mapping = _currentAnalysisMapping;
            var selectColumns = new List<string>
            {
                "_excel_row_number AS ExcelRowNumber",
                $"{mapping.Manufacturer.InternalColumnName} AS Manufacturer"
            };

            if (mapping.ManufacturerPartNumber is not null)
            {
                selectColumns.Add($"{mapping.ManufacturerPartNumber.InternalColumnName} AS ManufacturerPartNumber");
            }

            using var command = _connection.CreateCommand();
            command.CommandText = $"""
                SELECT {string.Join(", ", selectColumns)}
                FROM {SourceTableName}
                WHERE NULLIF(TRIM({mapping.PrimaryPartIdentifier.InternalColumnName}), '') = ?
                ORDER BY _source_row_id;
                """;
            AddParameter(command, partIdentifier);

            using var reader = command.ExecuteReader();
            return ReadToDataTable(reader);
        }
        catch (Exception ex)
        {
            throw new DuckDbDataException("The selected part detail rows could not be read.", ex);
        }
    }

    private void EnsureOpenConnection(bool resetDatabase)
    {
        if (resetDatabase)
        {
            DisposeConnection();
            Directory.CreateDirectory(Path.GetDirectoryName(_databasePath)!);
            DeleteDatabaseFiles();
        }

        if (_connection is not null)
        {
            return;
        }

        _connection = new DuckDBConnection($"Data Source={_databasePath}");
        _connection.Open();
    }

    private void EnsurePartMetadataTable()
    {
        using var command = _connection!.CreateCommand();
        command.CommandText = $"""
            CREATE TABLE IF NOT EXISTS {PartMetadataTableName} (
                part_identifier VARCHAR PRIMARY KEY,
                reviewed BOOLEAN NOT NULL DEFAULT false
            );
            """;
        command.ExecuteNonQuery();
    }

    private void DropAndCreateSourceTable(IReadOnlyList<ImportedColumnInfo> columns, DuckDBTransaction transaction)
    {
        using var dropCommand = _connection!.CreateCommand();
        dropCommand.Transaction = transaction;
        dropCommand.CommandText = $"""
            DROP VIEW IF EXISTS {PartSummaryViewName};
            DROP TABLE IF EXISTS {PartMetadataTableName};
            DROP TABLE IF EXISTS {SourceTableName};
            """;
        dropCommand.ExecuteNonQuery();

        var columnSql = string.Join(
            ", ",
            columns.Select(column => $"{column.InternalColumnName} VARCHAR"));

        using var createCommand = _connection.CreateCommand();
        createCommand.Transaction = transaction;
        createCommand.CommandText = $"""
            CREATE TABLE {SourceTableName} (
                _source_row_id BIGINT NOT NULL,
                _excel_row_number INTEGER NOT NULL,
                {columnSql},
                {PartNumberComparisonColumnName} VARCHAR,
                {CategoryComparisonColumnName} VARCHAR,
                {ManufacturerComparisonColumnName} VARCHAR
            );
            """;
        createCommand.ExecuteNonQuery();
    }

    private (long ImportedRows, int SkippedBlankRows) InsertRows(
        IXLWorksheet worksheet,
        WorksheetProcessingSession session,
        IReadOnlyList<ImportedColumnInfo> columns,
        DuckDBTransaction transaction,
        CancellationToken cancellationToken)
    {
        long sourceRowId = 0;
        using var appender = _connection!.CreateAppender(SourceTableName);

        foreach (var sessionRow in session.EligibleRows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var values = columns
                .Select(column => GetCellText(worksheet.Cell(sessionRow.ExcelRowNumber, column.ExcelColumnNumber)))
                .ToList();

            sourceRowId++;
            appender.AppendRow(row =>
            {
                row.AppendValue(sourceRowId);
                row.AppendValue(sessionRow.ExcelRowNumber);
                foreach (var value in values)
                {
                    row.AppendValue(value);
                }

                row.AppendValue(sessionRow.PartNumberComparisonValue);
                row.AppendValue(sessionRow.CategoryComparisonValue);
                row.AppendValue(sessionRow.ManufacturerComparisonValue);
            });
        }

        appender.Close();
        return (sourceRowId, 0);
    }

    private long GetImportedRowCount()
    {
        using var command = _connection!.CreateCommand();
        command.CommandText = $"SELECT COUNT(*) FROM {SourceTableName};";
        return Convert.ToInt64(command.ExecuteScalar());
    }

    private long ExecuteLong(string commandText)
    {
        using var command = _connection!.CreateCommand();
        command.CommandText = commandText;
        return Convert.ToInt64(command.ExecuteScalar());
    }

    private long GetCount(string commandText, IReadOnlyList<object?> parameters)
    {
        using var command = _connection!.CreateCommand();
        command.CommandText = commandText;
        foreach (var value in parameters)
        {
            AddParameter(command, value ?? DBNull.Value);
        }

        return Convert.ToInt64(command.ExecuteScalar());
    }

    private static DataTable ReadToDataTable(IDataReader reader)
    {
        var table = new DataTable();
        for (var index = 0; index < reader.FieldCount; index++)
        {
            table.Columns.Add(reader.GetName(index), typeof(string));
        }

        while (reader.Read())
        {
            var row = table.NewRow();
            for (var index = 0; index < reader.FieldCount; index++)
            {
                row[index] = reader.IsDBNull(index) ? DBNull.Value : Convert.ToString(reader.GetValue(index));
            }

            table.Rows.Add(row);
        }

        return table;
    }

    private static DataTable ReadPartSummaryTable(IDataReader reader)
    {
        var table = new DataTable();
        table.Columns.Add("PartIdentifier", typeof(string));
        table.Columns.Add("PartRowCount", typeof(long));
        table.Columns.Add("ManufacturerCount", typeof(long));
        table.Columns.Add("IsDuplicatePart", typeof(string));
        table.Columns.Add("HasMultipleManufacturers", typeof(string));
        table.Columns.Add("Reviewed", typeof(bool));

        while (reader.Read())
        {
            var row = table.NewRow();
            row["PartIdentifier"] = reader.IsDBNull(0) ? DBNull.Value : reader.GetString(0);
            row["PartRowCount"] = Convert.ToInt64(reader.GetValue(1));
            row["ManufacturerCount"] = Convert.ToInt64(reader.GetValue(2));
            row["IsDuplicatePart"] = reader.IsDBNull(3) ? DBNull.Value : Convert.ToString(reader.GetValue(3));
            row["HasMultipleManufacturers"] = reader.IsDBNull(4) ? DBNull.Value : Convert.ToString(reader.GetValue(4));
            row["Reviewed"] = !reader.IsDBNull(5) && Convert.ToBoolean(reader.GetValue(5));
            table.Rows.Add(row);
        }

        return table;
    }

    private void ValidateMappedColumn(ImportedColumnInfo column)
    {
        if (!SafeColumnNamePattern.IsMatch(column.InternalColumnName)
            || _currentDataset?.Columns.Any(existing => existing.InternalColumnName == column.InternalColumnName) != true)
        {
            throw new DuckDbDataException("A selected mapping no longer exists in the loaded dataset.");
        }
    }

    private static string CreatePageQuery(IReadOnlyList<ImportedColumnInfo> columns, string whereClause, string orderBy)
    {
        var selectColumns = string.Join(", ", columns.Select(column => column.InternalColumnName));
        return $"""
            SELECT {selectColumns}
            FROM {SourceTableName}
            {whereClause}
            {orderBy}
            LIMIT ?
            OFFSET ?;
            """;
    }

    private string BuildRawOrderBy(QuerySort? sort)
    {
        if (sort is null || string.IsNullOrWhiteSpace(sort.ColumnKey))
        {
            return "ORDER BY _source_row_id ASC";
        }

        if (!SafeColumnNamePattern.IsMatch(sort.ColumnKey)
            || _currentDataset?.Columns.Any(column => column.InternalColumnName == sort.ColumnKey) != true)
        {
            throw new DuckDbDataException("The selected sort column no longer exists in the loaded dataset.");
        }

        return $"ORDER BY {sort.ColumnKey} {ToSqlDirection(sort.Direction)}, _source_row_id ASC";
    }

    private static string BuildGroupedOrderBy(QuerySort? sort)
    {
        if (sort is null || string.IsNullOrWhiteSpace(sort.ColumnKey))
        {
            return "ORDER BY ps.part_identifier ASC";
        }

        var expression = sort.ColumnKey switch
        {
            "PartIdentifier" => "ps.part_identifier",
            "PartRowCount" => "ps.part_row_count",
            "ManufacturerCount" => "ps.manufacturer_count",
            "IsDuplicatePart" => "ps.part_row_count > 1",
            "HasMultipleManufacturers" => "ps.manufacturer_count > 1",
            "Reviewed" => "COALESCE(pm.reviewed, false)",
            _ => throw new DuckDbDataException("The selected grouped sort column is not supported.")
        };

        return $"ORDER BY {expression} {ToSqlDirection(sort.Direction)}, ps.part_identifier ASC";
    }

    private static string ToSqlDirection(SortDirection direction)
    {
        return direction == SortDirection.Descending ? "DESC" : "ASC";
    }

    private static FilterKind ChooseFilterKind(long distinctCount, long nonBlankCount, long numericCount, long leadingZeroCount)
    {
        if (leadingZeroCount > 0)
        {
            return FilterKind.Text;
        }

        if (leadingZeroCount == 0
            && nonBlankCount > 0
            && numericCount == nonBlankCount)
        {
            return FilterKind.NumericRange;
        }

        if (distinctCount <= 50)
        {
            return FilterKind.ValueList;
        }

        if (distinctCount <= 1000)
        {
            return FilterKind.SearchableValueList;
        }

        return FilterKind.Text;
    }

    private FilterSql BuildRawWhereClause(IReadOnlyList<FilterCriteria> filters)
    {
        return BuildSourceConditions(filters.Where(filter => filter.Definition.Target == FilterTarget.SourceRows), "source_rows");
    }

    private FilterSql BuildGroupedWhereClause(IReadOnlyList<FilterCriteria> filters)
    {
        var parameters = new List<object?>();
        var conditions = new List<string>();
        var computedFilters = filters.Where(filter => filter.IsActive && filter.Definition.Target == FilterTarget.GroupedPartsComputed);
        foreach (var filter in computedFilters)
        {
            var condition = BuildComputedCondition(filter, parameters);
            if (!string.IsNullOrWhiteSpace(condition))
            {
                conditions.Add(condition);
            }
        }

        var sourceFilters = filters.Where(filter => filter.IsActive && filter.Definition.Target == FilterTarget.SourceRows).ToList();
        if (sourceFilters.Count > 0 && _currentAnalysisMapping is not null)
        {
            var sourceSql = BuildSourceConditions(sourceFilters, "sr");
            conditions.Add($"""
                EXISTS (
                    SELECT 1
                    FROM {SourceTableName} sr
                    WHERE NULLIF(TRIM(sr.{_currentAnalysisMapping.PrimaryPartIdentifier.InternalColumnName}), '') = ps.part_identifier
                    {sourceSql.AndClause}
                )
                """);
            parameters.AddRange(sourceSql.Parameters);
        }

        return new FilterSql(conditions.Count == 0 ? string.Empty : $" WHERE {string.Join(" AND ", conditions)}", parameters);
    }

    private FilterSql BuildSourceConditions(IEnumerable<FilterCriteria> filters, string tableAlias)
    {
        var parameters = new List<object?>();
        var conditions = new List<string>();

        foreach (var filter in filters.Where(filter => filter.IsActive))
        {
            ValidateFilterDefinition(filter.Definition);
            var column = $"{tableAlias}.{filter.Definition.InternalColumnName}";
            switch (filter.Definition.Kind)
            {
                case FilterKind.ValueList:
                case FilterKind.SearchableValueList:
                    var selectedConditions = new List<string>();
                    foreach (var value in filter.SelectedValues)
                    {
                        if (value is null)
                        {
                            selectedConditions.Add($"({column} IS NULL OR NULLIF(TRIM({column}), '') IS NULL)");
                        }
                        else
                        {
                            selectedConditions.Add($"LOWER(TRIM({column})) = ?");
                            parameters.Add(value.Trim().ToLowerInvariant());
                        }
                    }

                    if (selectedConditions.Count > 0)
                    {
                        conditions.Add($"({string.Join(" OR ", selectedConditions)})");
                    }
                    break;
                case FilterKind.Text:
                    if (!string.IsNullOrWhiteSpace(filter.Text))
                    {
                        var text = filter.Text.Trim().ToLowerInvariant();
                        var pattern = filter.TextMode switch
                        {
                            TextFilterMode.Exact => text,
                            TextFilterMode.StartsWith => $"{EscapeLikeValue(text)}%",
                            _ => $"%{EscapeLikeValue(text)}%"
                        };
                        var op = filter.TextMode == TextFilterMode.Exact ? "=" : "LIKE";
                        conditions.Add(filter.TextMode == TextFilterMode.Exact
                            ? $"LOWER(TRIM({column})) = ?"
                            : $"LOWER(TRIM({column})) LIKE ? ESCAPE '\\'");
                        parameters.Add(pattern);
                    }
                    break;
                case FilterKind.NumericRange:
                    if (filter.Minimum is not null)
                    {
                        conditions.Add($"TRY_CAST({column} AS DOUBLE) >= ?");
                        parameters.Add(filter.Minimum.Value);
                    }

                    if (filter.Maximum is not null)
                    {
                        conditions.Add($"TRY_CAST({column} AS DOUBLE) <= ?");
                        parameters.Add(filter.Maximum.Value);
                    }
                    break;
            }
        }

        var where = conditions.Count == 0 ? string.Empty : $" WHERE {string.Join(" AND ", conditions)}";
        var and = conditions.Count == 0 ? string.Empty : $" AND {string.Join(" AND ", conditions)}";
        return new FilterSql(where, parameters, and);
    }

    private static string BuildComputedCondition(FilterCriteria filter, List<object?> parameters)
    {
        return filter.Definition.Id switch
        {
            "computed_duplicate_part" => BuildComputedBooleanCondition("ps.part_row_count > 1", filter),
            "computed_multiple_manufacturer" => BuildComputedBooleanCondition("ps.manufacturer_count > 1", filter),
            "computed_reviewed" => BuildComputedBooleanCondition("COALESCE(pm.reviewed, false)", filter),
            "computed_part_row_count" => BuildComputedRangeCondition("ps.part_row_count", filter, parameters),
            "computed_manufacturer_count" => BuildComputedRangeCondition("ps.manufacturer_count", filter, parameters),
            _ => string.Empty
        };
    }

    private static string BuildComputedBooleanCondition(string expression, FilterCriteria filter)
    {
        var selected = filter.SelectedValues.Where(value => value is not null).Select(value => value!.ToLowerInvariant()).ToList();
        if (selected.Count == 0 || selected.Count == 2)
        {
            return string.Empty;
        }

        return selected[0] == "yes" ? expression : $"NOT ({expression})";
    }

    private static string BuildComputedRangeCondition(string expression, FilterCriteria filter, List<object?> parameters)
    {
        var conditions = new List<string>();
        if (filter.Minimum is not null)
        {
            conditions.Add($"{expression} >= ?");
            parameters.Add(filter.Minimum.Value);
        }

        if (filter.Maximum is not null)
        {
            conditions.Add($"{expression} <= ?");
            parameters.Add(filter.Maximum.Value);
        }

        return string.Join(" AND ", conditions);
    }

    private void ValidateFilterDefinition(FilterDefinition definition)
    {
        if (definition.Target == FilterTarget.GroupedPartsComputed)
        {
            return;
        }

        if (!SafeColumnNamePattern.IsMatch(definition.InternalColumnName)
            || _currentDataset?.Columns.Any(existing => existing.InternalColumnName == definition.InternalColumnName) != true)
        {
            throw new DuckDbDataException("A selected filter column no longer exists in the loaded dataset.");
        }
    }

    private static string EscapeLikeValue(string value)
    {
        return value.Replace(@"\", @"\\").Replace("%", @"\%").Replace("_", @"\_");
    }

    private string CreateProcessedDataExportQuery(string whereClause)
    {
        var sourceColumns = string.Join(", ", _currentDataset!.Columns
            .OrderBy(column => column.DisplayOrder)
            .Select(column => $"source_rows.{column.InternalColumnName}"));

        if (_currentAnalysisMapping is null)
        {
            return $"""
                SELECT
                    {sourceColumns},
                    NULL AS duplicate_part,
                    NULL AS multiple_manufacturer,
                    NULL AS part_row_count,
                    NULL AS manufacturer_count,
                    NULL AS reviewed
                FROM {SourceTableName}
                {whereClause}
                ORDER BY _source_row_id;
                """;
        }

        var partColumn = _currentAnalysisMapping.PrimaryPartIdentifier.InternalColumnName;
        return $"""
            SELECT
                {sourceColumns},
                CASE WHEN ps.part_identifier IS NULL THEN NULL ELSE ps.part_row_count > 1 END AS duplicate_part,
                CASE WHEN ps.part_identifier IS NULL THEN NULL ELSE ps.manufacturer_count > 1 END AS multiple_manufacturer,
                ps.part_row_count,
                ps.manufacturer_count,
                CASE WHEN ps.part_identifier IS NULL THEN NULL ELSE COALESCE(pm.reviewed, false) END AS reviewed
            FROM {SourceTableName}
            LEFT JOIN {PartSummaryViewName} ps ON ps.part_identifier = NULLIF(TRIM(source_rows.{partColumn}), '')
            LEFT JOIN {PartMetadataTableName} pm ON pm.part_identifier = ps.part_identifier
            {whereClause}
            ORDER BY _source_row_id;
            """;
    }

    private static string CreatePartSummaryExportQuery(string whereClause)
    {
        return $"""
            SELECT
                ps.part_identifier,
                ps.part_row_count,
                ps.manufacturer_count,
                ps.part_row_count > 1 AS duplicate_part,
                ps.manufacturer_count > 1 AS multiple_manufacturer,
                COALESCE(pm.reviewed, false) AS reviewed
            FROM {PartSummaryViewName} ps
            LEFT JOIN {PartMetadataTableName} pm ON pm.part_identifier = ps.part_identifier
            {whereClause}
            ORDER BY ps.part_identifier;
            """;
    }

    private static void WriteHeaders(IXLWorksheet worksheet, IReadOnlyList<string> headers)
    {
        for (var index = 0; index < headers.Count; index++)
        {
            worksheet.Cell(1, index + 1).SetValue(headers[index]);
        }
    }

    private static void WriteGeneratedBoolean(IXLCell cell, IDataRecord reader, int index)
    {
        if (reader.IsDBNull(index))
        {
            cell.SetValue(string.Empty);
            return;
        }

        cell.SetValue(Convert.ToBoolean(reader.GetValue(index)) ? "Yes" : "No");
    }

    private static void WriteGeneratedNumber(IXLCell cell, IDataRecord reader, int index)
    {
        if (reader.IsDBNull(index))
        {
            cell.SetValue(string.Empty);
            return;
        }

        cell.SetValue(Convert.ToInt64(reader.GetValue(index)));
    }

    private static void FormatWorksheet(IXLWorksheet worksheet, int columnCount, long lastRow)
    {
        if (columnCount == 0)
        {
            return;
        }

        var headerRange = worksheet.Range(1, 1, 1, columnCount);
        headerRange.Style.Font.Bold = true;
        worksheet.SheetView.FreezeRows(1);

        if (lastRow >= 1)
        {
            worksheet.Range(1, 1, (int)Math.Min(lastRow, int.MaxValue), columnCount).SetAutoFilter();
        }

        worksheet.Columns(1, columnCount).AdjustToContents();
        foreach (var column in worksheet.Columns(1, columnCount))
        {
            if (column.Width > 40)
            {
                column.Width = 40;
            }
        }
    }

    private static string CreateTemporaryExportPath(string destinationPath)
    {
        var directory = Path.GetDirectoryName(destinationPath);
        var fileName = $"{Path.GetFileNameWithoutExtension(destinationPath)}.{Guid.NewGuid():N}.tmp.xlsx";
        return Path.Combine(string.IsNullOrWhiteSpace(directory) ? Environment.CurrentDirectory : directory, fileName);
    }

    private sealed record FilterSql(string WhereClause, IReadOnlyList<object?> Parameters, string AndClause = "");

    private static IReadOnlyList<ImportedColumnInfo> CreateImportedColumns(IReadOnlyList<WorksheetSessionColumn> columns)
    {
        var duplicateCounts = columns
            .GroupBy(column => DisplayHeader(column), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);

        return columns
            .Select((column, index) =>
            {
                var displayHeader = DisplayHeader(column);
                var hasDuplicateDisplayName = duplicateCounts[displayHeader] > 1;
                return new ImportedColumnInfo
                {
                    DisplayOrder = index + 1,
                    ExcelColumnNumber = column.ExcelColumnNumber,
                    ExcelColumnLetter = XLHelper.GetColumnLetterFromNumber(column.ExcelColumnNumber),
                    OriginalHeader = column.OriginalHeaderText,
                    DisplayHeader = displayHeader,
                    GridHeader = hasDuplicateDisplayName ? $"{displayHeader} ({XLHelper.GetColumnLetterFromNumber(column.ExcelColumnNumber)})" : displayHeader,
                    InternalColumnName = $"col_{index + 1:000}"
                };
            })
            .ToList();
    }

    private static string DisplayHeader(WorksheetSessionColumn column)
    {
        return string.IsNullOrWhiteSpace(column.OriginalHeaderText)
            ? $"Column {column.ExcelColumnNumber}"
            : column.OriginalHeaderText;
    }

    private static string? GetCellText(IXLCell cell)
    {
        return cell.GetFormattedString();
    }

    private static void AddParameter(DbCommand command, object value)
    {
        var parameter = command.CreateParameter();
        parameter.Value = value;
        command.Parameters.Add(parameter);
    }

    private void DeleteDatabaseFiles()
    {
        var searchPattern = $"{Path.GetFileName(_databasePath)}*";
        foreach (var path in Directory.GetFiles(Path.GetDirectoryName(_databasePath)!, searchPattern))
        {
            File.Delete(path);
        }
    }

    private void CleanupStaleSessionFiles()
    {
        try
        {
            var directory = Path.GetDirectoryName(_databasePath);
            if (directory is null || !Directory.Exists(directory))
            {
                return;
            }

            DeleteDatabaseFiles();
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            PerformanceLogger.Write($"TEMP_CLEANUP_SKIPPED message=\"{ex.Message}\"");
        }
    }

    private static void TryDeleteFile(string? path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return;
        }

        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            PerformanceLogger.Write($"TEMP_EXPORT_CLEANUP_SKIPPED path=\"{path}\" message=\"{ex.Message}\"");
        }
    }

    private void DisposeConnection()
    {
        _connection?.Dispose();
        _connection = null;
    }

    public void Dispose()
    {
        DisposeConnection();
        try
        {
            if (Directory.Exists(Path.GetDirectoryName(_databasePath)))
            {
                DeleteDatabaseFiles();
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}
