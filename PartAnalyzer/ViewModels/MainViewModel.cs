using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using PartAnalyzer.Models;
using PartAnalyzer.Services;
using PartAnalyzer.Utilities;

namespace PartAnalyzer.ViewModels;

public sealed class MainViewModel : ObservableObject, IDisposable
{
    private readonly SettingsService _settingsService;
    private readonly ExcelWorkbookService _excelWorkbookService;
    private readonly WorksheetProcessingSessionService _worksheetProcessingSessionService;
    private readonly FileDialogService _fileDialogService;
    private readonly DuckDbDataService _duckDbDataService;
    private readonly FidelityExportService _fidelityExportService;
    private readonly AppSettings _settings;
    private readonly AsyncRelayCommand _importWorkbookCommand;
    private readonly AsyncRelayCommand _loadDataCommand;
    private readonly AsyncRelayCommand _exportFilteredWorkbookCommand;
    private WorkbookInfo? _currentWorkbook;
    private WorksheetInfo? _selectedWorksheet;
    private DataLoadResult? _loadedDataset;
    private WorksheetProcessingSession? _processingSession;
    private MavlResult? _mavlResult;
    private bool _isInspecting;
    private bool _isLoadingData;
    private bool _isExporting;
    private bool _isFiltering;
    private string _statusMessage = "Ready.";
    private string _dataStateMessage = "No workbook selected.";
    private string? _presetName;
    private int _constrainedFilterRequestVersion;
    private IReadOnlyList<UnavailableFilterSelection> _unavailableSelections = Array.Empty<UnavailableFilterSelection>();

    public MainViewModel(SettingsService settingsService, ExcelWorkbookService excelWorkbookService, WorksheetProcessingSessionService worksheetProcessingSessionService, FileDialogService fileDialogService, DuckDbDataService duckDbDataService, FidelityExportService fidelityExportService)
    {
        _settingsService = settingsService;
        _excelWorkbookService = excelWorkbookService;
        _worksheetProcessingSessionService = worksheetProcessingSessionService;
        _fileDialogService = fileDialogService;
        _duckDbDataService = duckDbDataService;
        _fidelityExportService = fidelityExportService;
        _settings = _settingsService.Load();
        _importWorkbookCommand = new AsyncRelayCommand(ImportWorkbookAsync, () => CanStartInspection);
        _loadDataCommand = new AsyncRelayCommand(LoadDataAsync, () => CanLoadData);
        _exportFilteredWorkbookCommand = new AsyncRelayCommand(ExportFilteredWorkbookAsync, () => CanExportFilteredWorkbook);
        foreach (var filter in ConstrainedFilters) filter.PropertyChanged += ConstrainedFilterPropertyChanged;
        RefreshPresetNames();
        PresetName = "Generic";
    }

    public ObservableCollection<WorksheetInfo> Worksheets { get; } = new();
    public ObservableCollection<ConstrainedFilterViewModel> ConstrainedFilters { get; } = new()
    {
        new(SessionFilterColumn.PartNumber, "P+F part number"),
        new(SessionFilterColumn.Category, "Category"),
        new(SessionFilterColumn.Manufacturer, "Manufacturer")
    };
    public ObservableCollection<FilterValueOption> ManufacturerPreview { get; } = new();
    public ObservableCollection<string> PresetNames { get; } = new();
    public AsyncRelayCommand ImportWorkbookCommand => _importWorkbookCommand;
    public AsyncRelayCommand LoadDataCommand => _loadDataCommand;
    public AsyncRelayCommand ExportFilteredWorkbookCommand => _exportFilteredWorkbookCommand;

    public WorkbookInfo? CurrentWorkbook
    {
        get => _currentWorkbook;
        private set
        {
            if (SetProperty(ref _currentWorkbook, value))
            {
                OnPropertyChanged(nameof(CanLoadData));
                RaiseCommandStates();
            }
        }
    }

    public WorksheetInfo? SelectedWorksheet
    {
        get => _selectedWorksheet;
        set
        {
            if (!SetProperty(ref _selectedWorksheet, value)) return;
            PreferredWorksheet = value?.Name;
            ClearLoadedData(value is null ? "No workbook selected." : "Worksheet inspected. Load data to view records.");
            OnPropertyChanged(nameof(CanLoadData));
            SaveSettings();
        }
    }

    public bool IncludeHiddenRowsAndColumns
    {
        get => _settings.IncludeHiddenRowsAndColumns;
        set
        {
            if (_settings.IncludeHiddenRowsAndColumns == value) return;
            _settings.IncludeHiddenRowsAndColumns = value;
            ClearLoadedData(SelectedWorksheet is null ? "No workbook selected." : "Import settings changed. Reload data to apply.");
            OnPropertyChanged();
            SaveSettings();
        }
    }

    public string? LastImportDirectory
    {
        get => _settings.LastImportDirectory;
        private set
        {
            if (_settings.LastImportDirectory == value) return;
            _settings.LastImportDirectory = value;
            OnPropertyChanged();
        }
    }

    public string? PreferredWorksheet
    {
        get => _settings.PreferredWorksheet;
        private set
        {
            if (_settings.PreferredWorksheet == value) return;
            _settings.PreferredWorksheet = value;
            OnPropertyChanged();
        }
    }

    public bool IsInspecting { get => _isInspecting; private set { if (SetProperty(ref _isInspecting, value)) RefreshBusyState(); } }
    public bool IsLoadingData { get => _isLoadingData; private set { if (SetProperty(ref _isLoadingData, value)) RefreshBusyState(); } }
    public bool IsExporting { get => _isExporting; private set { if (SetProperty(ref _isExporting, value)) RefreshBusyState(); } }
    public bool IsFiltering { get => _isFiltering; private set { if (SetProperty(ref _isFiltering, value)) RefreshBusyState(); } }
    public bool IsBusy => IsInspecting || IsLoadingData || IsExporting || IsFiltering;
    public bool CanStartInspection => !IsBusy;
    public bool CanLoadData => !IsBusy && CurrentWorkbook is not null && SelectedWorksheet is not null && !SelectedWorksheet.IsHidden && SelectedWorksheet.Headers.Count > 0;
    public WorksheetProcessingSession? ProcessingSession => _processingSession;
    public bool HasValidatedProcessingSession => _processingSession is not null;
    public MavlResult? MavlResult => _mavlResult;
    public bool CanExportFilteredWorkbook => !IsBusy && _processingSession is not null && _mavlResult is not null && _loadedDataset is not null;
    public string StatusMessage { get => _statusMessage; private set => SetProperty(ref _statusMessage, value); }
    public string DataStateMessage { get => _dataStateMessage; private set => SetProperty(ref _dataStateMessage, value); }
    public string? PresetName { get => _presetName; set => SetProperty(ref _presetName, value); }
    public bool HasUnavailableSelectedValues => _unavailableSelections.Count > 0;
    public IReadOnlyList<UnavailableFilterSelection> UnavailableSelections => _unavailableSelections;
    public string ConstrainedFilterSummary => string.Join("   ", ConstrainedFilters.Select(filter => $"{filter.DisplayName}: {(filter.SelectedValues.Count == 0 ? "none" : $"{filter.SelectedValues.Count} selected")}"));

    private async Task ImportWorkbookAsync()
    {
        var filePath = _fileDialogService.SelectExcelWorkbook(LastImportDirectory);
        if (filePath is null) { StatusMessage = "Import cancelled."; return; }
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
            foreach (var worksheet in workbook.Worksheets) Worksheets.Add(worksheet);
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
        catch (OperationCanceledException) { StatusMessage = "Workbook inspection was cancelled."; }
        finally { IsInspecting = false; }
    }

    private async Task LoadDataAsync()
    {
        if (CurrentWorkbook is null || SelectedWorksheet is null) { DataStateMessage = "No worksheet selected."; return; }
        IsLoadingData = true;
        DataStateMessage = "Loading worksheet data...";
        StatusMessage = $"Loading {SelectedWorksheet.Name}...";
        try
        {
            _processingSession = await _worksheetProcessingSessionService.CreateAsync(CurrentWorkbook.FullPath, SelectedWorksheet, IncludeHiddenRowsAndColumns);
            OnPropertyChanged(nameof(ProcessingSession));
            OnPropertyChanged(nameof(HasValidatedProcessingSession));
            _loadedDataset = await _duckDbDataService.LoadWorksheetAsync(_processingSession);
            _mavlResult = await _duckDbDataService.CalculateMavlAsync();
            OnPropertyChanged(nameof(MavlResult));
            await InitializeConstrainedFiltersAsync();
            DataStateMessage = $"Validated and loaded {_loadedDataset.ImportedRowCount} eligible rows and {_loadedDataset.ImportedColumnCount} effective columns.";
            StatusMessage = $"Validated {SelectedWorksheet.Name}. {_loadedDataset.ImportedRowCount} eligible rows and {_loadedDataset.ImportedColumnCount} effective columns are available for later processing.";
        }
        catch (DuckDbDataException ex) { ClearLoadedData(ex.Message); StatusMessage = ex.Message; }
        catch (WorksheetValidationException ex) { ClearLoadedData(ex.Message); StatusMessage = ex.Message; }
        catch (OperationCanceledException) { ClearLoadedData("Data load was cancelled."); StatusMessage = "Data load was cancelled."; }
        finally { IsLoadingData = false; RaiseCommandStates(); }
    }

    private async Task ExportFilteredWorkbookAsync()
    {
        if (!CanExportFilteredWorkbook || _processingSession is null || _mavlResult is null) return;
        var unavailable = await _duckDbDataService.GetUnavailableSelectionsAsync(GetConstrainedFilterSnapshot());
        if (unavailable.Count > 0 && !_fileDialogService.ConfirmUnavailableSelections("Some selected values are unavailable and will match zero rows. Continue with export?")) return;
        var rows = await GetMatchingExcelRowNumbersAsync();
        if (rows.Count == 0) { StatusMessage = "No matching rows."; return; }
        var destination = _fileDialogService.SelectExportWorkbook(CreateSuggestedExportFileName(_processingSession.WorkbookPath), Path.GetDirectoryName(_processingSession.WorkbookPath));
        if (destination is null) return;
        IsExporting = true;
        StatusMessage = "Exporting...";
        try
        {
            var result = await _fidelityExportService.ExportAsync(_processingSession, _mavlResult, rows, destination);
            StatusMessage = $"Export completed: {result.ExportedDataRowCount} rows to {result.DestinationPath}";
        }
        catch (DuckDbDataException ex) { StatusMessage = ex.Message; }
        finally { IsExporting = false; }
    }

    public FilterSelectionSnapshot GetConstrainedFilterSnapshot() => new()
    {
        PartNumbers = Filter(SessionFilterColumn.PartNumber).SelectedValues.ToList(),
        Categories = Filter(SessionFilterColumn.Category).SelectedValues.ToList(),
        Manufacturers = Filter(SessionFilterColumn.Manufacturer).SelectedValues.ToList()
    };

    public async Task<IReadOnlyList<int>> GetMatchingExcelRowNumbersAsync() => _processingSession is null ? Array.Empty<int>() : await _duckDbDataService.GetMatchingExcelRowNumbersAsync(GetConstrainedFilterSnapshot());

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
        IsFiltering = true;
        try
        {
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
        finally { IsFiltering = false; }
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
        if (e.PropertyName != nameof(ConstrainedFilterViewModel.SelectedValues)) return;
        OnPropertyChanged(nameof(ConstrainedFilterSummary));
        _settings.LastUsedFilterSelections = GetConstrainedFilterSnapshot();
        SaveSettings();
        _ = RefreshUnavailableSelectionsAsync();
    }

    private async Task RefreshUnavailableSelectionsAsync()
    {
        if (_processingSession is null) return;
        _unavailableSelections = await _duckDbDataService.GetUnavailableSelectionsAsync(GetConstrainedFilterSnapshot());
        OnPropertyChanged(nameof(HasUnavailableSelectedValues));
        OnPropertyChanged(nameof(UnavailableSelections));
    }

    private void ClearLoadedData(string message)
    {
        _constrainedFilterRequestVersion++;
        _duckDbDataService.ClearDataset();
        _loadedDataset = null;
        _processingSession = null;
        _mavlResult = null;
        ClearConstrainedFilterSessionState();
        DataStateMessage = message;
        OnPropertyChanged(nameof(ProcessingSession));
        OnPropertyChanged(nameof(HasValidatedProcessingSession));
        OnPropertyChanged(nameof(MavlResult));
        RaiseCommandStates();
    }

    private void ClearConstrainedFilterSessionState()
    {
        ManufacturerPreview.Clear();
        _unavailableSelections = Array.Empty<UnavailableFilterSelection>();
        foreach (var filter in ConstrainedFilters) filter.ReplaceOptions(Array.Empty<FilterValueOption>());
        OnPropertyChanged(nameof(HasUnavailableSelectedValues));
        OnPropertyChanged(nameof(UnavailableSelections));
    }

    private WorksheetInfo? SelectInitialWorksheet(WorkbookInfo workbook)
    {
        if (!string.IsNullOrWhiteSpace(PreferredWorksheet))
        {
            var preferred = workbook.Worksheets.FirstOrDefault(sheet => string.Equals(sheet.Name, PreferredWorksheet, StringComparison.OrdinalIgnoreCase));
            if (preferred is not null && !preferred.IsHidden) return preferred;
        }
        var visibleSheetsWithData = workbook.Worksheets.Where(sheet => !sheet.IsHidden && sheet.UsedRowCount > 0 && sheet.UsedColumnCount > 0).ToList();
        return visibleSheetsWithData.Count == 1 ? visibleSheetsWithData[0] : visibleSheetsWithData.FirstOrDefault() ?? workbook.Worksheets.FirstOrDefault(sheet => !sheet.IsHidden) ?? workbook.Worksheets.FirstOrDefault();
    }

    private void RefreshBusyState()
    {
        OnPropertyChanged(nameof(IsBusy));
        OnPropertyChanged(nameof(CanStartInspection));
        OnPropertyChanged(nameof(CanLoadData));
        RaiseCommandStates();
    }

    private void RaiseCommandStates()
    {
        _importWorkbookCommand.RaiseCanExecuteChanged();
        _loadDataCommand.RaiseCanExecuteChanged();
        _exportFilteredWorkbookCommand.RaiseCanExecuteChanged();
    }

    private void RefreshPresetNames()
    {
        PresetNames.Clear();
        foreach (var preset in _settings.FilterPresets.OrderBy(preset => preset.Name, StringComparer.OrdinalIgnoreCase)) PresetNames.Add(preset.Name);
    }

    private void SaveSettings()
    {
        try { _settingsService.Save(_settings); }
        catch (IOException) { StatusMessage = "Settings could not be saved."; }
        catch (UnauthorizedAccessException) { StatusMessage = "Settings could not be saved because access was denied."; }
    }

    private ConstrainedFilterViewModel Filter(SessionFilterColumn column) => ConstrainedFilters.Single(filter => filter.Column == column);
    private static string CreateSuggestedExportFileName(string sourcePath) => $"{Path.GetFileNameWithoutExtension(sourcePath)}_Processed.xlsx";
    private string FormatUnavailableSelections() => string.Join("; ", _unavailableSelections.Select(item => $"{item.Column}: {string.Join(", ", item.Values.Take(5))}"));
    public void Dispose() => _duckDbDataService.Dispose();
}
