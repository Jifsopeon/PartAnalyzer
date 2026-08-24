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
    private const string PartSummaryViewName = "part_summary";
    private static readonly Regex SafeColumnNamePattern = new(@"^col_\d{3,}$", RegexOptions.Compiled);
    private readonly string _databasePath;
    private DuckDBConnection? _connection;
    private DataLoadResult? _currentDataset;
    private PartAnalysisMapping? _currentAnalysisMapping;

    public DuckDbDataService()
    {
        var tempDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "PartAnalyzer",
            "Temp");

        _databasePath = Path.Combine(tempDirectory, "session.duckdb");
    }

    public string DatabasePath => _databasePath;

    public DataLoadResult? CurrentDataset => _currentDataset;

    public Task<DataLoadResult> LoadWorksheetAsync(
        string workbookPath,
        WorksheetInfo worksheetInfo,
        bool ignoreHiddenRows,
        CancellationToken cancellationToken = default)
    {
        return Task.Run(() => LoadWorksheet(workbookPath, worksheetInfo, ignoreHiddenRows, cancellationToken), cancellationToken);
    }

    public Task<DataPageResult> GetPageAsync(
        int pageNumber,
        int pageSize,
        IReadOnlyList<FilterCriteria>? filters = null,
        CancellationToken cancellationToken = default)
    {
        return Task.Run(() => GetPage(pageNumber, pageSize, filters ?? Array.Empty<FilterCriteria>(), cancellationToken), cancellationToken);
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
        CancellationToken cancellationToken = default)
    {
        return Task.Run(() => GetPartSummaryPage(pageNumber, pageSize, filters ?? Array.Empty<FilterCriteria>(), cancellationToken), cancellationToken);
    }

    public Task<DataTable> GetPartDetailRowsAsync(
        string partIdentifier,
        CancellationToken cancellationToken = default)
    {
        return Task.Run(() => GetPartDetailRows(partIdentifier, cancellationToken), cancellationToken);
    }

    private DataLoadResult LoadWorksheet(
        string workbookPath,
        WorksheetInfo worksheetInfo,
        bool ignoreHiddenRows,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(workbookPath) || !File.Exists(workbookPath))
        {
            throw new DuckDbDataException("The source workbook is no longer available.");
        }

        if (worksheetInfo.Headers.Count == 0)
        {
            throw new DuckDbDataException("The selected worksheet does not contain importable columns.");
        }

        try
        {
            EnsureOpenConnection(resetDatabase: true);

            using var stream = File.Open(workbookPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == worksheetInfo.Name)
                ?? throw new DuckDbDataException("The selected worksheet could not be found in the workbook.");

            var columns = CreateImportedColumns(worksheetInfo.Headers);

            using var transaction = _connection!.BeginTransaction();
            try
            {
                DropAndCreateSourceTable(columns, transaction);
                var importCounts = InsertRows(worksheet, worksheetInfo, columns, ignoreHiddenRows, transaction, cancellationToken);
                transaction.Commit();

                var importedRowCount = GetImportedRowCount();
                if (importedRowCount != importCounts.ImportedRows)
                {
                    throw new DuckDbDataException("The imported row count could not be verified.");
                }

                _currentDataset = new DataLoadResult
                {
                    WorkbookPath = workbookPath,
                    WorksheetName = worksheetInfo.Name,
                    HeaderRowNumber = worksheetInfo.HeaderRowNumber,
                    ImportedRowCount = importedRowCount,
                    ImportedColumnCount = columns.Count,
                    SkippedBlankRowCount = importCounts.SkippedBlankRows,
                    Columns = columns
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

    private DataPageResult GetPage(int pageNumber, int pageSize, IReadOnlyList<FilterCriteria> filters, CancellationToken cancellationToken)
    {
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

            using var command = _connection.CreateCommand();
            command.CommandText = CreatePageQuery(_currentDataset.Columns, filterSql.WhereClause);
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

    private PartAnalysisResult AnalyzeParts(PartAnalysisMapping mapping, CancellationToken cancellationToken)
    {
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

    private DataPageResult GetPartSummaryPage(int pageNumber, int pageSize, IReadOnlyList<FilterCriteria> filters, CancellationToken cancellationToken)
    {
        if (_connection is null || _currentAnalysisMapping is null)
        {
            return new DataPageResult { PageNumber = 1, PageSize = pageSize };
        }

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            var filterSql = BuildGroupedWhereClause(filters);
            var totalRows = GetCount($"SELECT COUNT(*) FROM {PartSummaryViewName} ps{filterSql.WhereClause};", filterSql.Parameters);
            var totalPages = totalRows == 0 ? 1 : (int)Math.Ceiling(totalRows / (double)pageSize);
            var safePageNumber = Math.Clamp(pageNumber, 1, totalPages);
            var offset = (safePageNumber - 1) * pageSize;

            using var command = _connection.CreateCommand();
            command.CommandText = $"""
                SELECT
                    part_identifier AS PartIdentifier,
                    part_row_count AS PartRowCount,
                    manufacturer_count AS ManufacturerCount,
                    CASE WHEN part_row_count > 1 THEN 'Yes' ELSE 'No' END AS IsDuplicatePart,
                    CASE WHEN manufacturer_count > 1 THEN 'Yes' ELSE 'No' END AS HasMultipleManufacturers
                FROM {PartSummaryViewName} ps
                {filterSql.WhereClause}
                ORDER BY part_identifier
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
                Rows = ReadToDataTable(reader),
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

    private DataTable GetPartDetailRows(string partIdentifier, CancellationToken cancellationToken)
    {
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

    private void DropAndCreateSourceTable(IReadOnlyList<ImportedColumnInfo> columns, DuckDBTransaction transaction)
    {
        using var dropCommand = _connection!.CreateCommand();
        dropCommand.Transaction = transaction;
        dropCommand.CommandText = $"""
            DROP VIEW IF EXISTS {PartSummaryViewName};
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
                {columnSql}
            );
            """;
        createCommand.ExecuteNonQuery();
    }

    private (long ImportedRows, int SkippedBlankRows) InsertRows(
        IXLWorksheet worksheet,
        WorksheetInfo worksheetInfo,
        IReadOnlyList<ImportedColumnInfo> columns,
        bool ignoreHiddenRows,
        DuckDBTransaction transaction,
        CancellationToken cancellationToken)
    {
        var placeholders = string.Join(", ", Enumerable.Repeat("?", columns.Count + 2));
        var insertColumns = string.Join(", ", columns.Select(column => column.InternalColumnName));
        using var command = _connection!.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = $"""
            INSERT INTO {SourceTableName} (_source_row_id, _excel_row_number, {insertColumns})
            VALUES ({placeholders});
            """;

        for (var parameterIndex = 0; parameterIndex < columns.Count + 2; parameterIndex++)
        {
            command.Parameters.Add(command.CreateParameter());
        }

        var lastRowNumber = worksheet.RangeUsed()?.LastRow().RowNumber() ?? worksheetInfo.HeaderRowNumber;
        long sourceRowId = 0;
        var skippedBlankRows = 0;

        for (var rowNumber = worksheetInfo.HeaderRowNumber + 1; rowNumber <= lastRowNumber; rowNumber++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (ignoreHiddenRows && worksheet.Row(rowNumber).IsHidden)
            {
                continue;
            }

            var values = columns
                .Select(column => GetCellText(worksheet.Cell(rowNumber, column.ExcelColumnNumber)))
                .ToList();

            if (values.All(string.IsNullOrWhiteSpace))
            {
                skippedBlankRows++;
                continue;
            }

            sourceRowId++;
            command.Parameters[0].Value = sourceRowId;
            command.Parameters[1].Value = rowNumber;
            for (var valueIndex = 0; valueIndex < values.Count; valueIndex++)
            {
                command.Parameters[valueIndex + 2].Value = string.IsNullOrWhiteSpace(values[valueIndex])
                    ? DBNull.Value
                    : values[valueIndex];
            }

            command.ExecuteNonQuery();
        }

        return (sourceRowId, skippedBlankRows);
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

    private void ValidateMappedColumn(ImportedColumnInfo column)
    {
        if (!SafeColumnNamePattern.IsMatch(column.InternalColumnName)
            || _currentDataset?.Columns.Any(existing => existing.InternalColumnName == column.InternalColumnName) != true)
        {
            throw new DuckDbDataException("A selected mapping no longer exists in the loaded dataset.");
        }
    }

    private static string CreatePageQuery(IReadOnlyList<ImportedColumnInfo> columns, string whereClause)
    {
        var selectColumns = string.Join(", ", columns.Select(column => column.InternalColumnName));
        return $"""
            SELECT {selectColumns}
            FROM {SourceTableName}
            {whereClause}
            ORDER BY _source_row_id
            LIMIT ?
            OFFSET ?;
            """;
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

    private sealed record FilterSql(string WhereClause, IReadOnlyList<object?> Parameters, string AndClause = "");

    private static IReadOnlyList<ImportedColumnInfo> CreateImportedColumns(IReadOnlyList<HeaderInfo> headers)
    {
        var duplicateCounts = headers
            .GroupBy(header => header.DisplayName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);

        return headers
            .Select((header, index) =>
            {
                var displayHeader = header.DisplayName;
                var hasDuplicateDisplayName = duplicateCounts[displayHeader] > 1;
                return new ImportedColumnInfo
                {
                    DisplayOrder = index + 1,
                    ExcelColumnNumber = header.ColumnIndex,
                    ExcelColumnLetter = header.ColumnLetter,
                    OriginalHeader = header.Name ?? string.Empty,
                    DisplayHeader = displayHeader,
                    GridHeader = hasDuplicateDisplayName ? $"{displayHeader} ({header.ColumnLetter})" : displayHeader,
                    InternalColumnName = $"col_{index + 1:000}"
                };
            })
            .ToList();
    }

    private static string? GetCellText(IXLCell cell)
    {
        var text = cell.GetFormattedString();
        return string.IsNullOrWhiteSpace(text) ? null : text.Trim();
    }

    private static void AddParameter(DbCommand command, object value)
    {
        var parameter = command.CreateParameter();
        parameter.Value = value;
        command.Parameters.Add(parameter);
    }

    private void DeleteDatabaseFiles()
    {
        foreach (var path in Directory.GetFiles(Path.GetDirectoryName(_databasePath)!, "session.duckdb*"))
        {
            File.Delete(path);
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
