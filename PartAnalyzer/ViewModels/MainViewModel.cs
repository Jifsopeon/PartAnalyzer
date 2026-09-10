using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Windows.Threading;
using PartAnalyzer.Models;
using PartAnalyzer.Services;
using PartAnalyzer.Utilities;

namespace PartAnalyzer.ViewModels;

public sealed class MainViewModel : ObservableObject, IDisposable
{
    private const int DefaultPageSize = 200;
    private readonly SettingsService _settingsService;
    private readonly ExcelWorkbookService _excelWorkbookService;
    private readonly WorksheetProcessingSessionService _worksheetProcessingSessionService;
    private readonly FileDialogService _fileDialogService;
    private readonly DuckDbDataService _duckDbDataService;
    private readonly AppSettings _settings;
    private readonly AsyncRelayCommand _importWorkbookCommand;
    private readonly AsyncRelayCommand _loadDataCommand;
    private readonly AsyncRelayCommand _firstPageCommand;
    private readonly AsyncRelayCommand _previousPageCommand;
    private readonly AsyncRelayCommand _nextPageCommand;
    private readonly AsyncRelayCommand _lastPageCommand;
    private readonly AsyncRelayCommand _analyzePartsCommand;
    private readonly AsyncRelayCommand _firstGroupedPageCommand;
    private readonly AsyncRelayCommand _previousGroupedPageCommand;
    private readonly AsyncRelayCommand _nextGroupedPageCommand;
    private readonly AsyncRelayCommand _lastGroupedPageCommand;
    private readonly AsyncRelayCommand _exportWorkbookCommand;
    private readonly AsyncRelayCommand _addFilterCommand;
    private readonly AsyncRelayCommand _clearAllFiltersCommand;
    private readonly AsyncRelayCommand _removeAllFiltersCommand;
    private readonly RelayCommand _clearSelectedFilterCommand;
    private readonly RelayCommand _removeSelectedFilterCommand;
    private readonly DispatcherTimer _filterRefreshTimer;
    private readonly DispatcherTimer _filterValueSearchTimer;

    private WorkbookInfo? _currentWorkbook;
    private WorksheetInfo? _selectedWorksheet;
    private bool _isInspecting;
    private bool _isLoadingData;
    private bool _isAnalyzingParts;
    private bool _isExporting;
    private bool _isFiltering;
    private bool _hasUnexportedChanges;
    private string _statusMessage = "Ready.";
    private string _dataStateMessage = "No workbook selected.";
    private string _groupedPartsStateMessage = "Load worksheet data before analyzing parts.";
    private string _analysisSummary = "No part analysis has been run.";
    private int _selectedTabIndex;
    private DataView? _dataRowsView;
    private DataView? _groupedPartsRowsView;
    private DataView? _partDetailRowsView;
    private DataRowView? _selectedGroupedPart;
    private DataLoadResult? _loadedDataset;
    private WorksheetProcessingSession? _processingSession;
    private MavlResult? _mavlResult;
    private IReadOnlyList<ImportedColumnInfo> _importedColumns = Array.Empty<ImportedColumnInfo>();
    private IReadOnlyList<ImportedColumnInfo> _mappingColumns = Array.Empty<ImportedColumnInfo>();
    private IReadOnlyList<PartDetailColumn> _partDetailColumns = Array.Empty<PartDetailColumn>();
    private ImportedColumnInfo? _selectedPrimaryPartColumn;
    private ImportedColumnInfo? _selectedManufacturerColumn;
    private ImportedColumnInfo? _selectedManufacturerPartNumberColumn;
    private PartAnalysisResult? _partAnalysisResult;
    private ExportModeOption? _selectedExportModeOption;
    private FilterDefinition? _selectedAvailableFilter;
    private ActiveFilterViewModel? _selectedActiveFilter;
    private QuerySort? _rawSort;
    private QuerySort? _groupedSort;
    private int _filterRefreshVersion;
    private int _rawPageRequestVersion;
    private int _groupedPageRequestVersion;
    private int _partDetailRequestVersion;
    private int _constrainedFilterRequestVersion;
    private string? _presetName;
    private IReadOnlyList<UnavailableFilterSelection> _unavailableSelections = Array.Empty<UnavailableFilterSelection>();
    private int _currentPageNumber = 1;
    private int _selectedRawPageNumber = 1;
    private int _totalPages = 1;
    private long _totalImportedRows;
    private long _firstDisplayRow;
    private long _lastDisplayRow;
    private int _groupedPageNumber = 1;
    private int _selectedGroupedPageNumber = 1;
    private int _groupedTotalPages = 1;
    private long _totalGroupedRows;
    private long _firstGroupedDisplayRow;
    private long _lastGroupedDisplayRow;

    public MainViewModel(
        SettingsService settingsService,
        ExcelWorkbookService excelWorkbookService,
        WorksheetProcessingSessionService worksheetProcessingSessionService,
        FileDialogService fileDialogService,
        DuckDbDataService duckDbDataService)
    {
        _settingsService = settingsService;
        _excelWorkbookService = excelWorkbookService;
        _worksheetProcessingSessionService = worksheetProcessingSessionService;
        _fileDialogService = fileDialogService;
        _duckDbDataService = duckDbDataService;
        _settings = _settingsService.Load();
        _importWorkbookCommand = new AsyncRelayCommand(ImportWorkbookAsync, () => CanStartInspection);
        _loadDataCommand = new AsyncRelayCommand(LoadDataAsync, () => CanLoadData);
        _firstPageCommand = new AsyncRelayCommand(() => LoadPageAsync(1), () => CanMovePreviousPage);
        _previousPageCommand = new AsyncRelayCommand(() => LoadPageAsync(CurrentPageNumber - 1), () => CanMovePreviousPage);
        _nextPageCommand = new AsyncRelayCommand(() => LoadPageAsync(CurrentPageNumber + 1), () => CanMoveNextPage);
        _lastPageCommand = new AsyncRelayCommand(() => LoadPageAsync(TotalPages), () => CanMoveNextPage);
        _analyzePartsCommand = new AsyncRelayCommand(AnalyzePartsAsync, () => CanAnalyzeParts);
        _firstGroupedPageCommand = new AsyncRelayCommand(() => LoadGroupedPageAsync(1), () => CanMovePreviousGroupedPage);
        _previousGroupedPageCommand = new AsyncRelayCommand(() => LoadGroupedPageAsync(GroupedPageNumber - 1), () => CanMovePreviousGroupedPage);
        _nextGroupedPageCommand = new AsyncRelayCommand(() => LoadGroupedPageAsync(GroupedPageNumber + 1), () => CanMoveNextGroupedPage);
        _lastGroupedPageCommand = new AsyncRelayCommand(() => LoadGroupedPageAsync(GroupedTotalPages), () => CanMoveNextGroupedPage);
        _exportWorkbookCommand = new AsyncRelayCommand(ExportWorkbookAsync, () => CanExport);
        _addFilterCommand = new AsyncRelayCommand(AddSelectedFilterAsync, () => CanAddFilter);
        _clearAllFiltersCommand = new AsyncRelayCommand(ClearAllFiltersAsync, () => ActiveFilters.Count > 0);
        _removeAllFiltersCommand = new AsyncRelayCommand(RemoveAllFiltersAsync, () => ActiveFilters.Count > 0);
        _clearSelectedFilterCommand = new RelayCommand(ClearSelectedFilter, () => SelectedActiveFilter is not null);
        _removeSelectedFilterCommand = new RelayCommand(RemoveSelectedFilter, () => SelectedActiveFilter is not null);
        _filterRefreshTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
        _filterRefreshTimer.Tick += async (_, _) => await RefreshFilteredViewsFromTimerAsync();
        _filterValueSearchTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(300) };
        _filterValueSearchTimer.Tick += async (_, _) => await RefreshSelectedFilterValuesFromTimerAsync();
        ActiveFilters.CollectionChanged += ActiveFiltersCollectionChanged;
        RebuildPageNumbers(RawPageNumbers, 1);
        RebuildPageNumbers(GroupedPageNumbers, 1);
        SelectedExportModeOption = ExportModes[0];
        foreach (var filter in ConstrainedFilters) filter.PropertyChanged += ConstrainedFilterPropertyChanged;
        RefreshPresetNames();
    }

    public ObservableCollection<WorksheetInfo> Worksheets { get; } = new();

    public AsyncRelayCommand ImportWorkbookCommand => _importWorkbookCommand;

    public AsyncRelayCommand LoadDataCommand => _loadDataCommand;

    public AsyncRelayCommand FirstPageCommand => _firstPageCommand;

    public AsyncRelayCommand PreviousPageCommand => _previousPageCommand;

    public AsyncRelayCommand NextPageCommand => _nextPageCommand;

    public AsyncRelayCommand LastPageCommand => _lastPageCommand;

    public AsyncRelayCommand AnalyzePartsCommand => _analyzePartsCommand;

    public AsyncRelayCommand FirstGroupedPageCommand => _firstGroupedPageCommand;

    public AsyncRelayCommand PreviousGroupedPageCommand => _previousGroupedPageCommand;

    public AsyncRelayCommand NextGroupedPageCommand => _nextGroupedPageCommand;

    public AsyncRelayCommand LastGroupedPageCommand => _lastGroupedPageCommand;

    public AsyncRelayCommand ExportWorkbookCommand => _exportWorkbookCommand;

    public AsyncRelayCommand AddFilterCommand => _addFilterCommand;

    public AsyncRelayCommand ClearAllFiltersCommand => _clearAllFiltersCommand;

    public AsyncRelayCommand RemoveAllFiltersCommand => _removeAllFiltersCommand;

    public RelayCommand ClearSelectedFilterCommand => _clearSelectedFilterCommand;

    public RelayCommand RemoveSelectedFilterCommand => _removeSelectedFilterCommand;

    public ObservableCollection<FilterDefinition> AvailableFilters { get; } = new();

    public ObservableCollection<ActiveFilterViewModel> ActiveFilters { get; } = new();

    public ObservableCollection<ConstrainedFilterViewModel> ConstrainedFilters { get; } = new()
    {
        new(SessionFilterColumn.PartNumber, "P+F part number"),
        new(SessionFilterColumn.Category, "Category"),
        new(SessionFilterColumn.Manufacturer, "Manufacturer")
    };

    public ObservableCollection<FilterValueOption> ManufacturerPreview { get; } = new();

    public ObservableCollection<string> PresetNames { get; } = new();

    public string? PresetName { get => _presetName; set => SetProperty(ref _presetName, value); }

    public bool HasUnavailableSelectedValues => _unavailableSelections.Count > 0;

    public IReadOnlyList<UnavailableFilterSelection> UnavailableSelections => _unavailableSelections;

    public ObservableCollection<int> RawPageNumbers { get; } = new();

    public ObservableCollection<int> GroupedPageNumbers { get; } = new();

    public IReadOnlyList<ExportModeOption> ExportModes { get; } = new[]
    {
        new ExportModeOption { Mode = ExportMode.ProcessedWorkbook, DisplayName = "Processed Workbook", Description = "Processed Data and Part Summary sheets." },
        new ExportModeOption { Mode = ExportMode.AllSourceRows, DisplayName = "All Source Rows", Description = "All loaded rows with generated fields appended." },
        new ExportModeOption { Mode = ExportMode.CurrentFilteredRows, DisplayName = "Current Filtered Rows", Description = "All source rows matching the current filters." },
        new ExportModeOption { Mode = ExportMode.GroupedPartSummary, DisplayName = "Grouped Part Summary", Description = "One row per analyzed part." },
        new ExportModeOption { Mode = ExportMode.CurrentFilteredGroupedParts, DisplayName = "Current Filtered Grouped Parts", Description = "All grouped parts matching the current filters." }
    };

    public WorkbookInfo? CurrentWorkbook
    {
        get => _currentWorkbook;
        private set
        {
            if (SetProperty(ref _currentWorkbook, value))
            {
                OnPropertyChanged(nameof(CanLoadData));
                OnPropertyChanged(nameof(CanAnalyzeParts));
                RaiseCommandStates();
            }
        }
    }

    public WorksheetInfo? SelectedWorksheet
    {
        get => _selectedWorksheet;
        set
        {
            if (!Equals(_selectedWorksheet, value) && HasUnexportedChanges && !ConfirmDiscardUnsavedChanges())
            {
                OnPropertyChanged();
                return;
            }

            if (!SetProperty(ref _selectedWorksheet, value))
            {
                return;
            }

            PreferredWorksheet = value?.Name;
            ClearLoadedData(value is null ? "No workbook selected." : "Worksheet inspected. Load data to view records.");
            MappingColumns = CreateMappingColumns(value);
            RestoreMappings();
            OnPropertyChanged(nameof(CanLoadData));
            OnPropertyChanged(nameof(CanAnalyzeParts));
            OnPropertyChanged(nameof(HiddenRowPolicySummary));
            OnPropertyChanged(nameof(SelectedWorksheetEligibleRowCount));
            SaveSettings();
        }
    }

    public bool IncludeHiddenRowsAndColumns
    {
        get => _settings.IncludeHiddenRowsAndColumns;
        set
        {
            if (_settings.IncludeHiddenRowsAndColumns == value)
            {
                return;
            }

            if (HasUnexportedChanges && !ConfirmDiscardUnsavedChanges())
            {
                OnPropertyChanged();
                return;
            }

            _settings.IncludeHiddenRowsAndColumns = value;
            ClearLoadedData(SelectedWorksheet is null
                ? "No workbook selected."
                : "Import settings changed. Reload data to apply.");
            OnPropertyChanged();
            OnPropertyChanged(nameof(HiddenRowPolicySummary));
            OnPropertyChanged(nameof(SelectedWorksheetEligibleRowCount));
            SaveSettings();
        }
    }

    public string? LastImportDirectory
    {
        get => _settings.LastImportDirectory;
        private set
        {
            if (_settings.LastImportDirectory == value)
            {
                return;
            }

            _settings.LastImportDirectory = value;
            OnPropertyChanged();
        }
    }

    public string? PreferredWorksheet
    {
        get => _settings.PreferredWorksheet;
        private set
        {
            if (_settings.PreferredWorksheet == value)
            {
                return;
            }

            _settings.PreferredWorksheet = value;
            OnPropertyChanged();
        }
    }

    public bool IsInspecting
    {
        get => _isInspecting;
        private set
        {
            if (!SetProperty(ref _isInspecting, value))
            {
                return;
            }

            OnPropertyChanged(nameof(CanStartInspection));
            OnPropertyChanged(nameof(CanLoadData));
            OnPropertyChanged(nameof(CanAnalyzeParts));
            OnPropertyChanged(nameof(CanExport));
            OnPropertyChanged(nameof(IsBusy));
            RaiseCommandStates();
        }
    }

    public bool IsLoadingData
    {
        get => _isLoadingData;
        private set
        {
            if (!SetProperty(ref _isLoadingData, value))
            {
                return;
            }

            OnPropertyChanged(nameof(IsBusy));
            OnPropertyChanged(nameof(CanLoadData));
            OnPropertyChanged(nameof(CanAnalyzeParts));
            OnPropertyChanged(nameof(CanExport));
            RaiseCommandStates();
        }
    }

    public bool IsAnalyzingParts
    {
        get => _isAnalyzingParts;
        private set
        {
            if (!SetProperty(ref _isAnalyzingParts, value))
            {
                return;
            }

            OnPropertyChanged(nameof(IsBusy));
            OnPropertyChanged(nameof(CanAnalyzeParts));
            OnPropertyChanged(nameof(CanExport));
            RaiseCommandStates();
        }
    }

    public bool IsExporting
    {
        get => _isExporting;
        private set
        {
            if (!SetProperty(ref _isExporting, value))
            {
                return;
            }

            OnPropertyChanged(nameof(IsBusy));
            OnPropertyChanged(nameof(CanExport));
            RaiseCommandStates();
        }
    }

    public bool IsFiltering
    {
        get => _isFiltering;
        private set
        {
            if (!SetProperty(ref _isFiltering, value))
            {
                return;
            }

            OnPropertyChanged(nameof(IsBusy));
            RaiseCommandStates();
        }
    }

    public bool HasUnexportedChanges
    {
        get => _hasUnexportedChanges;
        private set
        {
            if (SetProperty(ref _hasUnexportedChanges, value))
            {
                OnPropertyChanged(nameof(DirtyStateSummary));
            }
        }
    }

    public string DirtyStateSummary => HasUnexportedChanges
        ? "Reviewed changes have not been exported."
        : "No unexported Reviewed changes.";

    public bool IsBusy => IsInspecting || IsLoadingData || IsAnalyzingParts || IsExporting || IsFiltering;

    public bool CanStartInspection => !IsBusy;

    public bool CanLoadData => !IsBusy && CurrentWorkbook is not null && SelectedWorksheet is not null && !SelectedWorksheet.IsHidden && SelectedWorksheet.Headers.Count > 0;

    public WorksheetProcessingSession? ProcessingSession => _processingSession;

    public bool HasValidatedProcessingSession => _processingSession is not null;

    public MavlResult? MavlResult => _mavlResult;

    public bool CanAnalyzeParts => !IsBusy
        && _loadedDataset is not null
        && SelectedPrimaryPartColumn is not null
        && SelectedManufacturerColumn is not null
        && SelectedPrimaryPartColumn.InternalColumnName != SelectedManufacturerColumn.InternalColumnName;

    public bool CanExport => !IsBusy
        && _loadedDataset is not null
        && SelectedExportModeOption is not null
        && (SelectedExportModeOption.Mode is ExportMode.ProcessedWorkbook or ExportMode.GroupedPartSummary or ExportMode.CurrentFilteredGroupedParts
            ? _partAnalysisResult is not null
            : true);

    public ExportModeOption? SelectedExportModeOption
    {
        get => _selectedExportModeOption;
        set
        {
            if (SetProperty(ref _selectedExportModeOption, value))
            {
                OnPropertyChanged(nameof(CanExport));
                _exportWorkbookCommand.RaiseCanExecuteChanged();
            }
        }
    }

    public bool CanAddFilter => SelectedAvailableFilter is not null
        && ActiveFilters.All(filter => filter.Definition.Id != SelectedAvailableFilter.Id);

    public FilterDefinition? SelectedAvailableFilter
    {
        get => _selectedAvailableFilter;
        set
        {
            if (SetProperty(ref _selectedAvailableFilter, value))
            {
                OnPropertyChanged(nameof(CanAddFilter));
                _addFilterCommand.RaiseCanExecuteChanged();
            }
        }
    }

    public ActiveFilterViewModel? SelectedActiveFilter
    {
        get => _selectedActiveFilter;
        set
        {
            if (!SetProperty(ref _selectedActiveFilter, value))
            {
                return;
            }

            OnPropertyChanged(nameof(SelectedFilterUsesValues));
            OnPropertyChanged(nameof(SelectedFilterUsesValueSearch));
            OnPropertyChanged(nameof(SelectedFilterUsesText));
            OnPropertyChanged(nameof(SelectedFilterUsesRange));
            _clearSelectedFilterCommand.RaiseCanExecuteChanged();
            _removeSelectedFilterCommand.RaiseCanExecuteChanged();
        }
    }

    public bool SelectedFilterUsesValues => SelectedActiveFilter?.UsesValues == true;

    public bool SelectedFilterUsesValueSearch => SelectedActiveFilter?.UsesValueSearch == true;

    public bool SelectedFilterUsesText => SelectedActiveFilter?.UsesText == true;

    public bool SelectedFilterUsesRange => SelectedActiveFilter?.UsesRange == true;

    public int ActiveFilterCount => ActiveFilters.Count(filter => filter.IsActive);

    public string ActiveFilterSummary => ActiveFilterCount == 0
        ? "No active filters"
        : $"{ActiveFilterCount} active filters";

    public string StatusMessage
    {
        get => _statusMessage;
        private set => SetProperty(ref _statusMessage, value);
    }

    public int SelectedTabIndex
    {
        get => _selectedTabIndex;
        set => SetProperty(ref _selectedTabIndex, value);
    }

    public DataView? DataRowsView
    {
        get => _dataRowsView;
        private set => SetProperty(ref _dataRowsView, value);
    }

    public IReadOnlyList<ImportedColumnInfo> ImportedColumns
    {
        get => _importedColumns;
        private set => SetProperty(ref _importedColumns, value);
    }

    public IReadOnlyList<ImportedColumnInfo> MappingColumns
    {
        get => _mappingColumns;
        private set => SetProperty(ref _mappingColumns, value);
    }

    public ImportedColumnInfo? SelectedPrimaryPartColumn
    {
        get => _selectedPrimaryPartColumn;
        set
        {
            if (!SetProperty(ref _selectedPrimaryPartColumn, value))
            {
                return;
            }

            _settings.PrimaryPartIdentifierMapping = CreateMapping(value);
            ClearPartAnalysis("Mapping changed. Analyze parts to refresh grouped results.");
            if (ApplyPrimaryPartIdentifierFilterOverride())
            {
                ScheduleFilterRefresh();
            }

            SaveSettings();
        }
    }

    public ImportedColumnInfo? SelectedManufacturerColumn
    {
        get => _selectedManufacturerColumn;
        set
        {
            if (!SetProperty(ref _selectedManufacturerColumn, value))
            {
                return;
            }

            _settings.ManufacturerMapping = CreateMapping(value);
            ClearPartAnalysis("Mapping changed. Analyze parts to refresh grouped results.");
            SaveSettings();
        }
    }

    public ImportedColumnInfo? SelectedManufacturerPartNumberColumn
    {
        get => _selectedManufacturerPartNumberColumn;
        set
        {
            if (!SetProperty(ref _selectedManufacturerPartNumberColumn, value))
            {
                return;
            }

            _settings.ManufacturerPartNumberMapping = CreateMapping(value);
            ClearPartAnalysis("Mapping changed. Analyze parts to refresh grouped results.");
            SaveSettings();
        }
    }

    public string DataStateMessage
    {
        get => _dataStateMessage;
        private set => SetProperty(ref _dataStateMessage, value);
    }

    public string GroupedPartsStateMessage
    {
        get => _groupedPartsStateMessage;
        private set => SetProperty(ref _groupedPartsStateMessage, value);
    }

    public string AnalysisSummary
    {
        get => _analysisSummary;
        private set => SetProperty(ref _analysisSummary, value);
    }

    public DataView? GroupedPartsRowsView
    {
        get => _groupedPartsRowsView;
        private set
        {
            if (_groupedPartsRowsView?.Table is not null)
            {
                _groupedPartsRowsView.Table.ColumnChanged -= GroupedPartsColumnChanged;
            }

            if (SetProperty(ref _groupedPartsRowsView, value) && _groupedPartsRowsView?.Table is not null)
            {
                _groupedPartsRowsView.Table.ColumnChanged += GroupedPartsColumnChanged;
            }
        }
    }

    public DataView? PartDetailRowsView
    {
        get => _partDetailRowsView;
        private set => SetProperty(ref _partDetailRowsView, value);
    }

    public IReadOnlyList<PartDetailColumn> PartDetailColumns
    {
        get => _partDetailColumns;
        private set => SetProperty(ref _partDetailColumns, value);
    }

    public DataRowView? SelectedGroupedPart
    {
        get => _selectedGroupedPart;
        set
        {
            if (!SetProperty(ref _selectedGroupedPart, value))
            {
                return;
            }

            _ = LoadSelectedPartDetailsAsync();
        }
    }

    public int CurrentPageNumber
    {
        get => _currentPageNumber;
        private set
        {
            if (SetProperty(ref _currentPageNumber, value))
            {
                SetSelectedRawPageNumber(value);
                OnPropertyChanged(nameof(PageNumberSummary));
                OnPropertyChanged(nameof(CanMovePreviousPage));
                OnPropertyChanged(nameof(CanMoveNextPage));
            }
        }
    }

    public int TotalPages
    {
        get => _totalPages;
        private set
        {
            if (SetProperty(ref _totalPages, value))
            {
                RebuildPageNumbers(RawPageNumbers, value);
                if (SelectedRawPageNumber > value)
                {
                    SetSelectedRawPageNumber(value);
                }

                OnPropertyChanged(nameof(PageNumberSummary));
                OnPropertyChanged(nameof(CanMovePreviousPage));
                OnPropertyChanged(nameof(CanMoveNextPage));
            }
        }
    }

    public int SelectedRawPageNumber
    {
        get => _selectedRawPageNumber;
        set
        {
            var safeValue = Math.Clamp(value, 1, Math.Max(1, TotalPages));
            if (!SetProperty(ref _selectedRawPageNumber, safeValue))
            {
                return;
            }

            if (_loadedDataset is not null && safeValue != CurrentPageNumber)
            {
                _ = LoadPageAsync(safeValue);
            }
        }
    }

    public bool CanMovePreviousPage => !IsBusy && _loadedDataset is not null && CurrentPageNumber > 1;

    public bool CanMoveNextPage => !IsBusy && _loadedDataset is not null && CurrentPageNumber < TotalPages;

    public string PageNumberSummary => _loadedDataset is null
        ? "Page 0 of 0"
        : $"Page {CurrentPageNumber} of {TotalPages}";

    public string PageSummary
    {
        get
        {
            if (_loadedDataset is null)
            {
                return "Rows 0-0 of 0";
            }

            return $"Rows {_firstDisplayRow}-{_lastDisplayRow} of {_totalImportedRows}";
        }
    }

    public int GroupedPageNumber
    {
        get => _groupedPageNumber;
        private set
        {
            if (SetProperty(ref _groupedPageNumber, value))
            {
                SetSelectedGroupedPageNumber(value);
                OnPropertyChanged(nameof(GroupedPageNumberSummary));
                OnPropertyChanged(nameof(CanMovePreviousGroupedPage));
                OnPropertyChanged(nameof(CanMoveNextGroupedPage));
            }
        }
    }

    public int GroupedTotalPages
    {
        get => _groupedTotalPages;
        private set
        {
            if (SetProperty(ref _groupedTotalPages, value))
            {
                RebuildPageNumbers(GroupedPageNumbers, value);
                if (SelectedGroupedPageNumber > value)
                {
                    SetSelectedGroupedPageNumber(value);
                }

                OnPropertyChanged(nameof(GroupedPageNumberSummary));
                OnPropertyChanged(nameof(CanMovePreviousGroupedPage));
                OnPropertyChanged(nameof(CanMoveNextGroupedPage));
            }
        }
    }

    public int SelectedGroupedPageNumber
    {
        get => _selectedGroupedPageNumber;
        set
        {
            var safeValue = Math.Clamp(value, 1, Math.Max(1, GroupedTotalPages));
            if (!SetProperty(ref _selectedGroupedPageNumber, safeValue))
            {
                return;
            }

            if (_partAnalysisResult is not null && safeValue != GroupedPageNumber)
            {
                _ = LoadGroupedPageAsync(safeValue);
            }
        }
    }

    public bool CanMovePreviousGroupedPage => !IsBusy && _partAnalysisResult is not null && GroupedPageNumber > 1;

    public bool CanMoveNextGroupedPage => !IsBusy && _partAnalysisResult is not null && GroupedPageNumber < GroupedTotalPages;

    public string GroupedPageNumberSummary => _partAnalysisResult is null
        ? "Page 0 of 0"
        : $"Page {GroupedPageNumber} of {GroupedTotalPages}";

    public string GroupedPageSummary
    {
        get
        {
            if (_partAnalysisResult is null)
            {
                return "Parts 0-0 of 0";
            }

            return $"Parts {_firstGroupedDisplayRow}-{_lastGroupedDisplayRow} of {_totalGroupedRows}";
        }
    }

    public string HiddenRowPolicySummary
    {
        get
        {
            if (SelectedWorksheet is null)
            {
                return IncludeHiddenRowsAndColumns
                    ? "Hidden rows and columns will be included in future processing."
                    : "Hidden rows and columns will be excluded from future processing.";
            }

            return IncludeHiddenRowsAndColumns
                ? $"All {SelectedWorksheet.TotalDataRowCount} data rows are currently eligible for processing."
                : $"{SelectedWorksheetEligibleRowCount} of {SelectedWorksheet.TotalDataRowCount} data rows are currently eligible before blank-row validation.";
        }
    }

    public int? SelectedWorksheetEligibleRowCount
    {
        get
        {
            if (SelectedWorksheet is null)
            {
                return null;
            }

            return IncludeHiddenRowsAndColumns
                ? SelectedWorksheet.TotalDataRowCount
                : SelectedWorksheet.TotalDataRowCount - SelectedWorksheet.HiddenDataRowCount;
        }
    }

    private async Task ImportWorkbookAsync()
    {
        if (HasUnexportedChanges && !ConfirmDiscardUnsavedChanges())
        {
            StatusMessage = "Import cancelled.";
            return;
        }

        var filePath = _fileDialogService.SelectExcelWorkbook(LastImportDirectory);
        if (filePath is null)
        {
            StatusMessage = "Import cancelled.";
            return;
        }

        ClearLoadedData("No dataset loaded.");
        LastImportDirectory = Path.GetDirectoryName(filePath);
        SaveSettings();

        IsInspecting = true;
        StatusMessage = "Inspecting workbook...";

        try
        {
            var workbook = await _excelWorkbookService.InspectWorkbookAsync(filePath, IncludeHiddenRowsAndColumns);
            CurrentWorkbook = workbook;
            Worksheets.Clear();

            foreach (var worksheet in workbook.Worksheets)
            {
                Worksheets.Add(worksheet);
            }

            SelectedWorksheet = SelectInitialWorksheet(workbook);
            StatusMessage = $"Inspected {workbook.FileName}.";
            DataStateMessage = "Worksheet inspected. Load data to view records.";
        }
        catch (WorkbookInspectionException ex)
        {
            StatusMessage = ex.Message;
            DataStateMessage = "Workbook inspection failed.";
            CurrentWorkbook = null;
            Worksheets.Clear();
            SelectedWorksheet = null;
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Workbook inspection was cancelled.";
        }
        finally
        {
            IsInspecting = false;
        }
    }

    private async Task LoadDataAsync()
    {
        if (CurrentWorkbook is null || SelectedWorksheet is null)
        {
            DataStateMessage = "No worksheet selected.";
            return;
        }

        if (HasUnexportedChanges && !ConfirmDiscardUnsavedChanges())
        {
            StatusMessage = "Data load cancelled.";
            return;
        }

        IsLoadingData = true;
        DataRowsView = null;
        DataStateMessage = "Loading worksheet data...";
        StatusMessage = $"Loading {SelectedWorksheet.Name}...";

        try
        {
            _processingSession = await _worksheetProcessingSessionService.CreateAsync(
                CurrentWorkbook.FullPath,
                SelectedWorksheet,
                IncludeHiddenRowsAndColumns);
            OnPropertyChanged(nameof(ProcessingSession));
            OnPropertyChanged(nameof(HasValidatedProcessingSession));
            _loadedDataset = await _duckDbDataService.LoadWorksheetAsync(_processingSession);
            _mavlResult = await _duckDbDataService.CalculateMavlAsync();
            OnPropertyChanged(nameof(MavlResult));
            ImportedColumns = _loadedDataset.Columns;
            MappingColumns = _loadedDataset.Columns;
            RestoreMappings();
            await InitializeConstrainedFiltersAsync();

            await LoadPageAsync(1);
            SelectedTabIndex = 1;
            DataStateMessage = $"Validated and loaded {_loadedDataset.ImportedRowCount} eligible rows and {_loadedDataset.ImportedColumnCount} effective columns.";
            StatusMessage = $"Validated {SelectedWorksheet.Name}. {_loadedDataset.ImportedRowCount} eligible rows and {_loadedDataset.ImportedColumnCount} effective columns are available for later processing.";
        }
        catch (DuckDbDataException ex)
        {
            _loadedDataset = null;
            DataRowsView = null;
            DataStateMessage = ex.Message;
            StatusMessage = ex.Message;
        }
        catch (WorksheetValidationException ex)
        {
            _processingSession = null;
            OnPropertyChanged(nameof(ProcessingSession));
            OnPropertyChanged(nameof(HasValidatedProcessingSession));
            _loadedDataset = null;
            DataRowsView = null;
            DataStateMessage = ex.Message;
            StatusMessage = ex.Message;
        }
        catch (OperationCanceledException)
        {
            _loadedDataset = null;
            DataRowsView = null;
            DataStateMessage = "Data load was cancelled.";
            StatusMessage = "Data load was cancelled.";
        }
        finally
        {
            IsLoadingData = false;
            RefreshPagingProperties();
        }
    }

    private async Task AnalyzePartsAsync()
    {
        if (SelectedPrimaryPartColumn is null || SelectedManufacturerColumn is null)
        {
            GroupedPartsStateMessage = "Select a Primary Part Identifier and Manufacturer column before analyzing parts.";
            return;
        }

        if (SelectedPrimaryPartColumn.InternalColumnName == SelectedManufacturerColumn.InternalColumnName)
        {
            GroupedPartsStateMessage = "Primary Part Identifier and Manufacturer must use different columns.";
            return;
        }

        IsAnalyzingParts = true;
        GroupedPartsStateMessage = "Analyzing parts...";
        StatusMessage = "Analyzing parts...";

        try
        {
            var mapping = new PartAnalysisMapping
            {
                PrimaryPartIdentifier = SelectedPrimaryPartColumn,
                Manufacturer = SelectedManufacturerColumn,
                ManufacturerPartNumber = SelectedManufacturerPartNumberColumn
            };

            _partAnalysisResult = await _duckDbDataService.AnalyzePartsAsync(mapping);
            PartDetailColumns = CreatePartDetailColumns(mapping);
            AddComputedFilterDefinitions();
            await LoadGroupedPageAsync(1);
            SelectedTabIndex = 2;
            AnalysisSummary = $"{_partAnalysisResult.UniquePartCount} unique parts; {_partAnalysisResult.DuplicatePartCount} duplicate parts; {_partAnalysisResult.MultipleManufacturerPartCount} multiple-manufacturer parts; {_partAnalysisResult.BlankPrimaryIdentifierRowCount} rows excluded because part identifier was blank.";
            GroupedPartsStateMessage = "Part analysis complete.";
            StatusMessage = "Part analysis complete.";
        }
        catch (DuckDbDataException ex)
        {
            ClearPartAnalysis(ex.Message);
            StatusMessage = ex.Message;
        }
        catch (OperationCanceledException)
        {
            ClearPartAnalysis("Part analysis was cancelled.");
            StatusMessage = "Part analysis was cancelled.";
        }
        finally
        {
            IsAnalyzingParts = false;
            RefreshPagingProperties();
        }
    }

    private async Task ExportWorkbookAsync()
    {
        if (_loadedDataset is null || SelectedExportModeOption is null)
        {
            StatusMessage = "Load data before exporting.";
            return;
        }

        if (!CanExport)
        {
            StatusMessage = "Run part analysis before exporting this mode.";
            return;
        }

        var destinationPath = _fileDialogService.SelectExportWorkbook(
            CreateSuggestedExportFileName(_loadedDataset.WorkbookPath),
            Path.GetDirectoryName(_loadedDataset.WorkbookPath));
        if (destinationPath is null)
        {
            StatusMessage = "Export cancelled.";
            return;
        }

        if (PathsReferToSameFile(destinationPath, _loadedDataset.WorkbookPath))
        {
            StatusMessage = "The source workbook cannot be overwritten. Choose a different export filename.";
            return;
        }

        IsExporting = true;
        StatusMessage = "Exporting...";
        try
        {
            var result = await _duckDbDataService.ExportWorkbookAsync(new ExportRequest
            {
                Mode = SelectedExportModeOption.Mode,
                DestinationPath = destinationPath,
                Filters = GetActiveCriteria()
            });

            HasUnexportedChanges = false;
            var rowSummary = result.GroupedRowCount > 0 && result.SourceRowCount > 0
                ? $"{result.SourceRowCount} source rows and {result.GroupedRowCount} grouped rows"
                : $"{result.TotalRowCount} rows";
            StatusMessage = $"Export complete. {rowSummary} written to: {result.DestinationPath}";
        }
        catch (DuckDbDataException ex)
        {
            StatusMessage = ex.Message;
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "Export cancelled.";
        }
        finally
        {
            IsExporting = false;
        }
    }

    private async Task LoadGroupedPageAsync(int pageNumber)
    {
        if (_partAnalysisResult is null)
        {
            return;
        }

        var requestVersion = ++_groupedPageRequestVersion;
        var page = await _duckDbDataService.GetPartSummaryPageAsync(pageNumber, DefaultPageSize, GetActiveCriteria(), _groupedSort);
        if (requestVersion != _groupedPageRequestVersion)
        {
            return;
        }

        SelectedGroupedPart = null;
        PartDetailRowsView = null;
        GroupedPartsRowsView = page.Rows.DefaultView;
        _totalGroupedRows = page.TotalRows;
        _firstGroupedDisplayRow = page.FirstDisplayRow;
        _lastGroupedDisplayRow = page.LastDisplayRow;
        GroupedPageNumber = page.PageNumber;
        GroupedTotalPages = page.TotalPages;
        OnPropertyChanged(nameof(GroupedPageSummary));
        RaiseCommandStates();
    }

    private async Task LoadSelectedPartDetailsAsync()
    {
        var requestVersion = ++_partDetailRequestVersion;
        if (SelectedGroupedPart is null)
        {
            PartDetailRowsView = null;
            return;
        }

        try
        {
            var partIdentifier = Convert.ToString(SelectedGroupedPart["PartIdentifier"]);
            if (string.IsNullOrWhiteSpace(partIdentifier))
            {
                PartDetailRowsView = null;
                return;
            }

            var table = await _duckDbDataService.GetPartDetailRowsAsync(partIdentifier);
            if (requestVersion != _partDetailRequestVersion)
            {
                return;
            }

            PartDetailRowsView = table.DefaultView;
        }
        catch (DuckDbDataException ex)
        {
            StatusMessage = ex.Message;
            PartDetailRowsView = null;
        }
    }

    private async void GroupedPartsColumnChanged(object sender, DataColumnChangeEventArgs e)
    {
        if (e.Column?.ColumnName != "Reviewed")
        {
            return;
        }

        var partIdentifier = Convert.ToString(e.Row["PartIdentifier"]);
        if (string.IsNullOrWhiteSpace(partIdentifier))
        {
            return;
        }

        try
        {
            var reviewed = e.ProposedValue is not DBNull && Convert.ToBoolean(e.ProposedValue);
            await _duckDbDataService.SetPartReviewedAsync(partIdentifier, reviewed);
            HasUnexportedChanges = true;
            StatusMessage = reviewed
                ? $"Marked {partIdentifier} reviewed."
                : $"Marked {partIdentifier} not reviewed.";
        }
        catch (DuckDbDataException ex)
        {
            StatusMessage = ex.Message;
        }
    }

    private async Task LoadPageAsync(int pageNumber)
    {
        if (_loadedDataset is null)
        {
            return;
        }

        var requestVersion = ++_rawPageRequestVersion;
        var page = await _duckDbDataService.GetPageAsync(pageNumber, DefaultPageSize, GetActiveCriteria(), _rawSort);
        if (requestVersion != _rawPageRequestVersion)
        {
            return;
        }

        DataRowsView = page.Rows.DefaultView;
        _totalImportedRows = page.TotalRows;
        _firstDisplayRow = page.FirstDisplayRow;
        _lastDisplayRow = page.LastDisplayRow;
        CurrentPageNumber = page.PageNumber;
        TotalPages = page.TotalPages;
        OnPropertyChanged(nameof(PageSummary));
        StatusMessage = $"Showing rows {_firstDisplayRow}-{_lastDisplayRow} of {_totalImportedRows}.";
        RaiseCommandStates();
    }

    private WorksheetInfo? SelectInitialWorksheet(WorkbookInfo workbook)
    {
        if (!string.IsNullOrWhiteSpace(PreferredWorksheet))
        {
            var preferred = workbook.Worksheets.FirstOrDefault(
                sheet => string.Equals(sheet.Name, PreferredWorksheet, StringComparison.OrdinalIgnoreCase));

            if (preferred is not null && !preferred.IsHidden)
            {
                return preferred;
            }
        }

        var visibleSheetsWithData = workbook.Worksheets
            .Where(sheet => !sheet.IsHidden && sheet.UsedRowCount > 0 && sheet.UsedColumnCount > 0)
            .ToList();

        return visibleSheetsWithData.Count == 1
            ? visibleSheetsWithData[0]
            : visibleSheetsWithData.FirstOrDefault()
              ?? workbook.Worksheets.FirstOrDefault(sheet => !sheet.IsHidden)
              ?? workbook.Worksheets.FirstOrDefault();
    }

    private void SaveSettings()
    {
        try
        {
            _settingsService.Save(_settings);
        }
        catch (IOException)
        {
            StatusMessage = "Settings could not be saved.";
        }
        catch (UnauthorizedAccessException)
        {
            StatusMessage = "Settings could not be saved because access was denied.";
        }
    }

    public bool ConfirmCloseWithUnsavedChanges()
    {
        return !HasUnexportedChanges || ConfirmDiscardUnsavedChanges("You have changes that have not been exported. Exit and discard them?");
    }

    private bool ConfirmDiscardUnsavedChanges()
    {
        return ConfirmDiscardUnsavedChanges("The current dataset contains changes that have not been exported. Continue and discard them?");
    }

    private bool ConfirmDiscardUnsavedChanges(string message)
    {
        return _fileDialogService.ConfirmDiscardUnsavedChanges(message);
    }

    private static string CreateSuggestedExportFileName(string sourcePath)
    {
        var name = Path.GetFileNameWithoutExtension(sourcePath);
        return $"{name}_Processed.xlsx";
    }

    private static bool PathsReferToSameFile(string firstPath, string secondPath)
    {
        return string.Equals(
            Path.GetFullPath(firstPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
            Path.GetFullPath(secondPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
            StringComparison.OrdinalIgnoreCase);
    }

    private void ClearLoadedData(string message)
    {
        _rawPageRequestVersion++;
        _groupedPageRequestVersion++;
        _partDetailRequestVersion++;
        _rawSort = null;
        _groupedSort = null;
        _duckDbDataService.ClearDataset();
        _mavlResult = null;
        OnPropertyChanged(nameof(MavlResult));
        ClearConstrainedFilterSessionState();
        _processingSession = null;
        OnPropertyChanged(nameof(ProcessingSession));
        OnPropertyChanged(nameof(HasValidatedProcessingSession));
        _loadedDataset = null;
        ImportedColumns = Array.Empty<ImportedColumnInfo>();
        ClearFilters(clearDefinitions: true);
        ClearPartAnalysis("Load worksheet data before analyzing parts.");
        DataRowsView = null;
        _totalImportedRows = 0;
        _firstDisplayRow = 0;
        _lastDisplayRow = 0;
        CurrentPageNumber = 1;
        TotalPages = 1;
        SetSelectedRawPageNumber(1);
        RebuildPageNumbers(RawPageNumbers, 1);
        HasUnexportedChanges = false;
        DataStateMessage = message;
        OnPropertyChanged(nameof(PageSummary));
        RaiseCommandStates();
    }

    private void RefreshPagingProperties()
    {
        OnPropertyChanged(nameof(CanStartInspection));
        OnPropertyChanged(nameof(CanLoadData));
        OnPropertyChanged(nameof(CanAnalyzeParts));
        OnPropertyChanged(nameof(CanExport));
        OnPropertyChanged(nameof(CanMovePreviousPage));
        OnPropertyChanged(nameof(CanMoveNextPage));
        OnPropertyChanged(nameof(CanMovePreviousGroupedPage));
        OnPropertyChanged(nameof(CanMoveNextGroupedPage));
        OnPropertyChanged(nameof(PageNumberSummary));
        OnPropertyChanged(nameof(PageSummary));
        OnPropertyChanged(nameof(GroupedPageNumberSummary));
        OnPropertyChanged(nameof(GroupedPageSummary));
        RaiseCommandStates();
    }

    private void SetSelectedRawPageNumber(int value)
    {
        if (SetProperty(ref _selectedRawPageNumber, value, nameof(SelectedRawPageNumber)))
        {
            OnPropertyChanged(nameof(PageNumberSummary));
        }
    }

    private void SetSelectedGroupedPageNumber(int value)
    {
        if (SetProperty(ref _selectedGroupedPageNumber, value, nameof(SelectedGroupedPageNumber)))
        {
            OnPropertyChanged(nameof(GroupedPageNumberSummary));
        }
    }

    private static void RebuildPageNumbers(ObservableCollection<int> target, int totalPages)
    {
        var safeTotal = Math.Max(1, totalPages);
        if (target.Count == safeTotal
            && target.Count > 0
            && target[0] == 1
            && target[^1] == safeTotal)
        {
            return;
        }

        target.Clear();
        for (var page = 1; page <= safeTotal; page++)
        {
            target.Add(page);
        }
    }

    private void RaiseCommandStates()
    {
        _importWorkbookCommand.RaiseCanExecuteChanged();
        _loadDataCommand.RaiseCanExecuteChanged();
        _firstPageCommand.RaiseCanExecuteChanged();
        _previousPageCommand.RaiseCanExecuteChanged();
        _nextPageCommand.RaiseCanExecuteChanged();
        _lastPageCommand.RaiseCanExecuteChanged();
        _analyzePartsCommand.RaiseCanExecuteChanged();
        _firstGroupedPageCommand.RaiseCanExecuteChanged();
        _previousGroupedPageCommand.RaiseCanExecuteChanged();
        _nextGroupedPageCommand.RaiseCanExecuteChanged();
        _lastGroupedPageCommand.RaiseCanExecuteChanged();
        _exportWorkbookCommand.RaiseCanExecuteChanged();
        _addFilterCommand.RaiseCanExecuteChanged();
        _clearAllFiltersCommand.RaiseCanExecuteChanged();
        _removeAllFiltersCommand.RaiseCanExecuteChanged();
        _clearSelectedFilterCommand.RaiseCanExecuteChanged();
        _removeSelectedFilterCommand.RaiseCanExecuteChanged();
    }

    private async Task AddSelectedFilterAsync()
    {
        if (SelectedAvailableFilter is null || !CanAddFilter)
        {
            return;
        }

        var activeFilter = new ActiveFilterViewModel(SelectedAvailableFilter);
        activeFilter.PropertyChanged += ActiveFilterPropertyChanged;
        ActiveFilters.Add(activeFilter);
        SelectedActiveFilter = activeFilter;
        await LoadFilterOptionsAsync(activeFilter);
        OnPropertyChanged(nameof(CanAddFilter));
        _addFilterCommand.RaiseCanExecuteChanged();
    }

    public async Task ApplyRawSortAsync(string columnKey, SortDirection direction)
    {
        if (_loadedDataset is null)
        {
            return;
        }

        _rawSort = new QuerySort { ColumnKey = columnKey, Direction = direction };
        await LoadPageAsync(1);
    }

    public async Task ApplyGroupedSortAsync(string columnKey, SortDirection direction)
    {
        if (_partAnalysisResult is null)
        {
            return;
        }

        _groupedSort = new QuerySort { ColumnKey = columnKey, Direction = direction };
        await LoadGroupedPageAsync(1);
    }

    private async Task ClearAllFiltersAsync()
    {
        foreach (var filter in ActiveFilters)
        {
            filter.Clear();
        }

        await RefreshFilteredViewsAsync();
    }

    private async Task RemoveAllFiltersAsync()
    {
        foreach (var filter in ActiveFilters)
        {
            filter.PropertyChanged -= ActiveFilterPropertyChanged;
            filter.Clear();
        }

        ActiveFilters.Clear();
        SelectedActiveFilter = null;
        SelectedAvailableFilter = null;
        OnPropertyChanged(nameof(ActiveFilterCount));
        OnPropertyChanged(nameof(ActiveFilterSummary));
        OnPropertyChanged(nameof(CanAddFilter));
        RaiseCommandStates();
        await RefreshFilteredViewsAsync();
    }

    private void ClearSelectedFilter()
    {
        SelectedActiveFilter?.Clear();
        ScheduleFilterRefresh();
    }

    private void RemoveSelectedFilter()
    {
        if (SelectedActiveFilter is null)
        {
            return;
        }

        SelectedActiveFilter.PropertyChanged -= ActiveFilterPropertyChanged;
        ActiveFilters.Remove(SelectedActiveFilter);
        SelectedActiveFilter = ActiveFilters.FirstOrDefault();
        ScheduleFilterRefresh();
    }

    private void ActiveFiltersCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        OnPropertyChanged(nameof(ActiveFilterCount));
        OnPropertyChanged(nameof(ActiveFilterSummary));
        OnPropertyChanged(nameof(CanAddFilter));
        RaiseCommandStates();
    }

    private void ActiveFilterPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (sender is not ActiveFilterViewModel filter)
        {
            return;
        }

        if (e.PropertyName == nameof(ActiveFilterViewModel.IsActive)
            || e.PropertyName == nameof(ActiveFilterViewModel.Text)
            || e.PropertyName == nameof(ActiveFilterViewModel.MinimumText)
            || e.PropertyName == nameof(ActiveFilterViewModel.MaximumText))
        {
            OnPropertyChanged(nameof(ActiveFilterCount));
            OnPropertyChanged(nameof(ActiveFilterSummary));
            ScheduleFilterRefresh();
        }
        else if (e.PropertyName == nameof(ActiveFilterViewModel.ValueSearchText))
        {
            SelectedActiveFilter = filter;
            _filterValueSearchTimer.Stop();
            _filterValueSearchTimer.Start();
        }
    }

    private void ScheduleFilterRefresh()
    {
        _filterRefreshVersion++;
        _filterRefreshTimer.Stop();
        _filterRefreshTimer.Start();
    }

    private async Task RefreshFilteredViewsFromTimerAsync()
    {
        _filterRefreshTimer.Stop();
        await RefreshFilteredViewsAsync();
    }

    private async Task RefreshFilteredViewsAsync()
    {
        var version = ++_filterRefreshVersion;
        IsFiltering = true;
        try
        {
            if (_loadedDataset is not null)
            {
                await LoadPageAsync(1);
            }

            if (_partAnalysisResult is not null)
            {
                await LoadGroupedPageAsync(1);
            }

            if (version == _filterRefreshVersion)
            {
                StatusMessage = ActiveFilterCount == 0 ? "Filters cleared." : "Filters applied.";
            }
        }
        catch (DuckDbDataException ex)
        {
            StatusMessage = ex.Message;
        }
        finally
        {
            IsFiltering = false;
        }
    }

    private async Task RefreshSelectedFilterValuesFromTimerAsync()
    {
        _filterValueSearchTimer.Stop();
        if (SelectedActiveFilter is not null)
        {
            await LoadFilterOptionsAsync(SelectedActiveFilter);
        }
    }

    private async Task LoadFilterOptionsAsync(ActiveFilterViewModel filter)
    {
        if (!filter.UsesValues)
        {
            return;
        }

        if (filter.Definition.Target == FilterTarget.GroupedPartsComputed)
        {
            filter.ReplaceOptions(new[]
            {
                new FilterValueOption { DisplayValue = "Yes", Value = "yes", Count = 0 },
                new FilterValueOption { DisplayValue = "No", Value = "no", Count = 0 }
            });
            return;
        }

        var options = await _duckDbDataService.GetFilterValueOptionsAsync(filter.Definition, filter.ValueSearchText);
        filter.ReplaceOptions(options);
    }

    private IReadOnlyList<FilterCriteria> GetActiveCriteria()
    {
        return ActiveFilters
            .Select(filter => filter.ToCriteria())
            .Where(criteria => criteria?.IsActive == true)
            .Cast<FilterCriteria>()
            .ToList();
    }

    private async Task RefreshFilterDefinitionsAsync()
    {
        AvailableFilters.Clear();
        var definitions = await _duckDbDataService.DiscoverSourceFiltersAsync();
        foreach (var definition in definitions.OrderBy(definition => definition.DisplayOrder))
        {
            AvailableFilters.Add(definition);
        }

        ApplyPrimaryPartIdentifierFilterOverride();
        OnPropertyChanged(nameof(CanAddFilter));
        _addFilterCommand.RaiseCanExecuteChanged();
    }

    private bool ApplyPrimaryPartIdentifierFilterOverride()
    {
        if (SelectedPrimaryPartColumn is null)
        {
            return false;
        }

        var primaryColumnName = SelectedPrimaryPartColumn.InternalColumnName;
        var changed = false;
        for (var index = 0; index < AvailableFilters.Count; index++)
        {
            var definition = AvailableFilters[index];
            if (definition.Target != FilterTarget.SourceRows
                || definition.InternalColumnName != primaryColumnName
                || definition.Kind == FilterKind.Text)
            {
                continue;
            }

            AvailableFilters[index] = CopyFilterDefinition(definition, FilterKind.Text);
            changed = true;
        }

        foreach (var filter in ActiveFilters
                     .Where(filter => filter.Definition.Target == FilterTarget.SourceRows
                         && filter.Definition.InternalColumnName == primaryColumnName
                         && filter.Definition.Kind != FilterKind.Text)
                     .ToList())
        {
            filter.PropertyChanged -= ActiveFilterPropertyChanged;
            ActiveFilters.Remove(filter);
            changed = true;
        }

        if (changed)
        {
            SelectedActiveFilter = ActiveFilters.FirstOrDefault();
            SelectedAvailableFilter = AvailableFilters.FirstOrDefault(definition => definition.InternalColumnName == primaryColumnName)
                ?? SelectedAvailableFilter;
            OnPropertyChanged(nameof(CanAddFilter));
            _addFilterCommand.RaiseCanExecuteChanged();
        }

        return changed;
    }

    private static FilterDefinition CopyFilterDefinition(FilterDefinition source, FilterKind kind)
    {
        return new FilterDefinition
        {
            Id = source.Id,
            DisplayName = source.DisplayName,
            InternalColumnName = source.InternalColumnName,
            DisplayOrder = source.DisplayOrder,
            Kind = kind,
            Target = source.Target,
            DistinctNonBlankCount = source.DistinctNonBlankCount,
            BlankCount = source.BlankCount
        };
    }

    private void AddComputedFilterDefinitions()
    {
        var computed = new[]
        {
            new FilterDefinition { Id = "computed_duplicate_part", DisplayName = "Duplicate Part", Kind = FilterKind.ComputedBoolean, Target = FilterTarget.GroupedPartsComputed },
            new FilterDefinition { Id = "computed_multiple_manufacturer", DisplayName = "Multiple Manufacturer", Kind = FilterKind.ComputedBoolean, Target = FilterTarget.GroupedPartsComputed },
            new FilterDefinition { Id = "computed_reviewed", DisplayName = "Reviewed", Kind = FilterKind.ComputedBoolean, Target = FilterTarget.GroupedPartsComputed },
            new FilterDefinition { Id = "computed_part_row_count", DisplayName = "Part Row Count", Kind = FilterKind.ComputedRange, Target = FilterTarget.GroupedPartsComputed },
            new FilterDefinition { Id = "computed_manufacturer_count", DisplayName = "Manufacturer Count", Kind = FilterKind.ComputedRange, Target = FilterTarget.GroupedPartsComputed }
        };

        foreach (var definition in computed)
        {
            if (AvailableFilters.All(existing => existing.Id != definition.Id))
            {
                AvailableFilters.Add(definition);
            }
        }
    }

    private void RemoveComputedFilters()
    {
        foreach (var filter in ActiveFilters.Where(filter => filter.Definition.Target == FilterTarget.GroupedPartsComputed).ToList())
        {
            filter.PropertyChanged -= ActiveFilterPropertyChanged;
            ActiveFilters.Remove(filter);
        }

        foreach (var definition in AvailableFilters.Where(definition => definition.Target == FilterTarget.GroupedPartsComputed).ToList())
        {
            AvailableFilters.Remove(definition);
        }
    }

    private void ClearFilters(bool clearDefinitions)
    {
        foreach (var filter in ActiveFilters)
        {
            filter.PropertyChanged -= ActiveFilterPropertyChanged;
        }

        ActiveFilters.Clear();
        SelectedActiveFilter = null;
        SelectedAvailableFilter = null;
        if (clearDefinitions)
        {
            AvailableFilters.Clear();
        }

        OnPropertyChanged(nameof(ActiveFilterCount));
        OnPropertyChanged(nameof(ActiveFilterSummary));
        OnPropertyChanged(nameof(CanAddFilter));
    }

    private void ClearPartAnalysis(string message)
    {
        _groupedPageRequestVersion++;
        _partDetailRequestVersion++;
        _groupedSort = null;
        RemoveComputedFilters();
        _partAnalysisResult = null;
        GroupedPartsRowsView = null;
        PartDetailRowsView = null;
        SelectedGroupedPart = null;
        GroupedPageNumber = 1;
        GroupedTotalPages = 1;
        SetSelectedGroupedPageNumber(1);
        RebuildPageNumbers(GroupedPageNumbers, 1);
        _totalGroupedRows = 0;
        _firstGroupedDisplayRow = 0;
        _lastGroupedDisplayRow = 0;
        AnalysisSummary = "No current part analysis.";
        GroupedPartsStateMessage = message;
        OnPropertyChanged(nameof(GroupedPageSummary));
        RefreshPagingProperties();
    }

    private static IReadOnlyList<ImportedColumnInfo> CreateMappingColumns(WorksheetInfo? worksheet)
    {
        if (worksheet is null)
        {
            return Array.Empty<ImportedColumnInfo>();
        }

        var duplicateCounts = worksheet.Headers
            .GroupBy(header => header.DisplayName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);

        return worksheet.Headers
            .Select((header, index) =>
            {
                var hasDuplicateDisplayName = duplicateCounts[header.DisplayName] > 1;
                return new ImportedColumnInfo
                {
                    DisplayOrder = index + 1,
                    ExcelColumnNumber = header.ColumnIndex,
                    ExcelColumnLetter = header.ColumnLetter,
                    OriginalHeader = header.Name ?? string.Empty,
                    DisplayHeader = header.DisplayName,
                    GridHeader = hasDuplicateDisplayName ? $"{header.DisplayName} ({header.ColumnLetter})" : header.DisplayName,
                    InternalColumnName = header.InternalColumnName
                };
            })
            .ToList();
    }

    private void RestoreMappings()
    {
        SelectedPrimaryPartColumn = ResolveMapping(_settings.PrimaryPartIdentifierMapping);
        SelectedManufacturerColumn = ResolveMapping(_settings.ManufacturerMapping);
        SelectedManufacturerPartNumberColumn = ResolveMapping(_settings.ManufacturerPartNumberMapping);
    }

    private ImportedColumnInfo? ResolveMapping(ColumnMapping? mapping)
    {
        if (mapping is null)
        {
            return null;
        }

        var resolved = MappingColumns.FirstOrDefault(column =>
            column.InternalColumnName == mapping.InternalColumnName
            && column.ExcelColumnNumber == mapping.ExcelColumnNumber
            && column.ExcelColumnLetter == mapping.ExcelColumnLetter
            && column.DisplayHeader == mapping.DisplayHeader);

        return resolved;
    }

    private static ColumnMapping? CreateMapping(ImportedColumnInfo? column)
    {
        if (column is null)
        {
            return null;
        }

        return new ColumnMapping
        {
            InternalColumnName = column.InternalColumnName,
            ExcelColumnNumber = column.ExcelColumnNumber,
            ExcelColumnLetter = column.ExcelColumnLetter,
            DisplayHeader = column.DisplayHeader
        };
    }

    private static IReadOnlyList<PartDetailColumn> CreatePartDetailColumns(PartAnalysisMapping mapping)
    {
        var columns = new List<PartDetailColumn>
        {
            new() { DisplayOrder = 1, Header = "Excel Row", BindingPath = "ExcelRowNumber" },
            new() { DisplayOrder = 2, Header = mapping.Manufacturer.GridHeader, BindingPath = "Manufacturer" }
        };

        if (mapping.ManufacturerPartNumber is not null)
        {
            columns.Add(new PartDetailColumn
            {
                DisplayOrder = 3,
                Header = mapping.ManufacturerPartNumber.GridHeader,
                BindingPath = "ManufacturerPartNumber"
            });
        }

        return columns;
    }

    public FilterSelectionSnapshot GetConstrainedFilterSnapshot()
    {
        return new FilterSelectionSnapshot
        {
            PartNumbers = Filter(SessionFilterColumn.PartNumber).SelectedValues.ToList(),
            Categories = Filter(SessionFilterColumn.Category).SelectedValues.ToList(),
            Manufacturers = Filter(SessionFilterColumn.Manufacturer).SelectedValues.ToList()
        };
    }

    public async Task<IReadOnlyList<int>> GetMatchingExcelRowNumbersAsync()
    {
        return _processingSession is null ? Array.Empty<int>() : await _duckDbDataService.GetMatchingExcelRowNumbersAsync(GetConstrainedFilterSnapshot());
    }

    public void SavePreset(string name, bool overwrite)
    {
        var normalized = name.Trim();
        if (string.IsNullOrWhiteSpace(normalized)) throw new InvalidOperationException("Preset name cannot be blank.");
        var existing = _settings.FilterPresets.FirstOrDefault(preset => string.Equals(preset.Name, normalized, StringComparison.OrdinalIgnoreCase));
        if (existing is not null && !overwrite) throw new InvalidOperationException("A preset with that name already exists.");
        var snapshot = GetConstrainedFilterSnapshot();
        if (existing is null) _settings.FilterPresets.Add(new FilterPreset { Name = normalized, Selections = snapshot });
        else { existing.Name = normalized; existing.Selections = snapshot; }
        SaveSettings();
        RefreshPresetNames();
    }

    public async Task<string?> LoadPresetAsync(string name)
    {
        var preset = _settings.FilterPresets.FirstOrDefault(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
        if (preset is null) return null;
        ApplyConstrainedSelections(preset.Selections);
        await RefreshConstrainedFiltersAsync();
        return HasUnavailableSelectedValues ? FormatUnavailableSelections() : null;
    }

    public void DeletePreset(string name)
    {
        var preset = _settings.FilterPresets.FirstOrDefault(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
        if (preset is null) return;
        _settings.FilterPresets.Remove(preset);
        SaveSettings();
        RefreshPresetNames();
    }

    private async Task InitializeConstrainedFiltersAsync()
    {
        ApplyConstrainedSelections(_settings.LastUsedFilterSelections);
        await RefreshConstrainedFiltersAsync();
    }

    private async Task RefreshConstrainedFiltersAsync()
    {
        if (_processingSession is null) return;
        var version = ++_constrainedFilterRequestVersion;
        foreach (var filter in ConstrainedFilters)
        {
            var options = await _duckDbDataService.GetSessionFilterValuesAsync(filter.Column, filter.SearchText);
            if (version != _constrainedFilterRequestVersion) return;
            filter.ReplaceOptions(options);
        }
        var preview = await _duckDbDataService.GetManufacturerPreviewAsync();
        if (version != _constrainedFilterRequestVersion) return;
        ManufacturerPreview.Clear();
        foreach (var item in preview) ManufacturerPreview.Add(item);
        _unavailableSelections = await _duckDbDataService.GetUnavailableSelectionsAsync(GetConstrainedFilterSnapshot());
        OnPropertyChanged(nameof(HasUnavailableSelectedValues));
        OnPropertyChanged(nameof(UnavailableSelections));
    }

    private void ApplyConstrainedSelections(FilterSelectionSnapshot snapshot)
    {
        Filter(SessionFilterColumn.PartNumber).SetSelections(snapshot.PartNumbers);
        Filter(SessionFilterColumn.Category).SetSelections(snapshot.Categories);
        Filter(SessionFilterColumn.Manufacturer).SetSelections(snapshot.Manufacturers);
    }

    private void ConstrainedFilterPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(ConstrainedFilterViewModel.SearchText)) _ = RefreshConstrainedFiltersAsync();
        if (e.PropertyName == nameof(ConstrainedFilterViewModel.SelectedValues))
        {
            _settings.LastUsedFilterSelections = GetConstrainedFilterSnapshot();
            SaveSettings();
            _ = RefreshUnavailableSelectionsAsync();
        }
    }

    private async Task RefreshUnavailableSelectionsAsync()
    {
        if (_processingSession is null) return;
        _unavailableSelections = await _duckDbDataService.GetUnavailableSelectionsAsync(GetConstrainedFilterSnapshot());
        OnPropertyChanged(nameof(HasUnavailableSelectedValues));
        OnPropertyChanged(nameof(UnavailableSelections));
    }

    private void ClearConstrainedFilterSessionState()
    {
        _constrainedFilterRequestVersion++;
        ManufacturerPreview.Clear();
        _unavailableSelections = Array.Empty<UnavailableFilterSelection>();
        foreach (var filter in ConstrainedFilters) filter.ReplaceOptions(Array.Empty<FilterValueOption>());
        OnPropertyChanged(nameof(HasUnavailableSelectedValues));
        OnPropertyChanged(nameof(UnavailableSelections));
    }

    private void RefreshPresetNames()
    {
        PresetNames.Clear();
        foreach (var preset in _settings.FilterPresets.OrderBy(preset => preset.Name, StringComparer.OrdinalIgnoreCase)) PresetNames.Add(preset.Name);
    }

    private ConstrainedFilterViewModel Filter(SessionFilterColumn column) => ConstrainedFilters.Single(filter => filter.Column == column);

    private string FormatUnavailableSelections() => string.Join("; ", _unavailableSelections.Select(item => $"{item.Column}: {string.Join(", ", item.Values.Take(5))}"));

    public void Dispose()
    {
        _duckDbDataService.Dispose();
    }
}
