using System.Data.Common;
using System.IO;
using ClosedXML.Excel;
using DuckDB.NET.Data;
using ReportExtract.Models;

namespace ReportExtract.Services;

public sealed class DuckDbDataService : IDisposable
{
    private const string SourceTableName = "source_rows";
    private const string PartNumberComparisonColumnName = "part_number_comparison_value";
    private const string CategoryComparisonColumnName = "category_comparison_value";
    private const string ManufacturerComparisonColumnName = "manufacturer_comparison_value";
    private readonly string _databasePath;
    private DuckDBConnection? _connection;
    private DataLoadResult? _currentDataset;

    public DuckDbDataService(string? databasePath = null)
    {
        _databasePath = string.IsNullOrWhiteSpace(databasePath)
            ? Path.Combine(PortableApplicationPaths.Current.TempDirectory, "session.duckdb")
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

    public Task<MavlResult> CalculateMavlAsync(CancellationToken cancellationToken = default)
    {
        return Task.Run(() => CalculateMavl(cancellationToken), cancellationToken);
    }

    public void ClearDataset()
    {
        if (_connection is null)
        {
            _currentDataset = null;
            return;
        }

        using var command = _connection.CreateCommand();
        command.CommandText = $"DROP TABLE IF EXISTS {SourceTableName};";
        command.ExecuteNonQuery();
        _currentDataset = null;
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

    private MavlResult CalculateMavl(CancellationToken cancellationToken)
    {
        using var measurement = PerformanceLogger.Measure("MAVL calculation", $"rows={_currentDataset?.ImportedRowCount ?? 0}");
        if (_connection is null || _currentDataset?.ProcessingSession is null) throw new DuckDbDataException("Load a validated worksheet before calculating MAVL.");
        cancellationToken.ThrowIfCancellationRequested();
        using var command = _connection.CreateCommand();
        command.CommandText = $"SELECT COUNT(*) FROM {SourceTableName} WHERE {PartNumberComparisonColumnName} IS NULL OR {PartNumberComparisonColumnName} = '' OR {ManufacturerComparisonColumnName} IS NULL OR {ManufacturerComparisonColumnName} = '';";
        if (Convert.ToInt64(command.ExecuteScalar()) != 0) throw new DuckDbDataException("The validated session contains a blank P+F part number or Manufacturer identity.");
        command.CommandText = $"""
            WITH part_mavl AS (
                SELECT {PartNumberComparisonColumnName} AS part_identity,
                       COUNT(DISTINCT {ManufacturerComparisonColumnName}) > 1 AS is_mavl
                FROM {SourceTableName}
                GROUP BY {PartNumberComparisonColumnName}
            )
            SELECT sr._excel_row_number, pm.is_mavl
            FROM {SourceTableName} sr
            INNER JOIN part_mavl pm ON pm.part_identity = sr.{PartNumberComparisonColumnName}
            ORDER BY sr._excel_row_number;
            """;
        using var reader = command.ExecuteReader();
        var values = new Dictionary<int, MavlClassification>();
        while (reader.Read()) values.Add(reader.GetInt32(0), Convert.ToBoolean(reader.GetValue(1)) ? MavlClassification.Yes : MavlClassification.No);
        if (values.Count != _currentDataset.ImportedRowCount) throw new DuckDbDataException("MAVL row mapping could not be verified.");
        return new MavlResult { ByExcelRowNumber = values };
    }

    private string GetSessionFilterValueExpression(SessionFilterColumn column)
    {
        var expression = GetSessionFilterExpression(column);
        return column == SessionFilterColumn.Category
            ? $"COALESCE(NULLIF({expression}, ''), '{FilterSelectionValues.BlankCategory}')"
            : expression;
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
        dropCommand.CommandText = $"DROP TABLE IF EXISTS {SourceTableName};";
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

    private static string EscapeLikeValue(string value)
    {
        return value.Replace(@"\", @"\\").Replace("%", @"\%").Replace("_", @"\_");
    }

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
