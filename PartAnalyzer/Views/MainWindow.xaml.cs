using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.ComponentModel;
using PartAnalyzer.Models;
using PartAnalyzer.Services;
using PartAnalyzer.ViewModels;

namespace PartAnalyzer.Views;

public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel;

    public MainWindow()
    {
        InitializeComponent();

        var settingsService = new SettingsService();
        _viewModel = new MainViewModel(
            settingsService,
            new ExcelWorkbookService(),
            new WorksheetProcessingSessionService(),
            new FileDialogService(),
            new DuckDbDataService());
        DataContext = _viewModel;
        _viewModel.PropertyChanged += ViewModelPropertyChanged;
        Closing += MainWindowClosing;
        Closed += (_, _) => _viewModel.Dispose();
    }

    private void ViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(MainViewModel.ImportedColumns))
        {
            RebuildRawDataColumns(_viewModel.ImportedColumns);
        }
        else if (e.PropertyName == nameof(MainViewModel.PartDetailColumns))
        {
            RebuildPartDetailColumns(_viewModel.PartDetailColumns);
        }
    }

    private void OpenFilters_Click(object sender, RoutedEventArgs e)
    {
        var window = new FilterWindow(_viewModel)
        {
            Owner = this
        };
        window.ShowDialog();
    }

    private void MainWindowClosing(object? sender, CancelEventArgs e)
    {
        if (!_viewModel.ConfirmCloseWithUnsavedChanges())
        {
            e.Cancel = true;
        }
    }

    private async void RawDataGrid_Sorting(object sender, DataGridSortingEventArgs e)
    {
        e.Handled = true;
        if (string.IsNullOrWhiteSpace(e.Column.SortMemberPath))
        {
            return;
        }

        var direction = GetNextSortDirection(e.Column);
        ApplySortIndicator(RawDataGrid, e.Column, direction);
        await _viewModel.ApplyRawSortAsync(e.Column.SortMemberPath, ToQueryDirection(direction));
    }

    private async void GroupedPartsGrid_Sorting(object sender, DataGridSortingEventArgs e)
    {
        e.Handled = true;
        if (string.IsNullOrWhiteSpace(e.Column.SortMemberPath))
        {
            return;
        }

        var direction = GetNextSortDirection(e.Column);
        ApplySortIndicator(GroupedPartsGrid, e.Column, direction);
        await _viewModel.ApplyGroupedSortAsync(e.Column.SortMemberPath, ToQueryDirection(direction));
    }

    private static ListSortDirection GetNextSortDirection(DataGridColumn column)
    {
        return column.SortDirection == ListSortDirection.Ascending
            ? ListSortDirection.Descending
            : ListSortDirection.Ascending;
    }

    private static void ApplySortIndicator(DataGrid grid, DataGridColumn sortedColumn, ListSortDirection direction)
    {
        foreach (var column in grid.Columns)
        {
            column.SortDirection = column == sortedColumn ? direction : null;
        }
    }

    private static SortDirection ToQueryDirection(ListSortDirection direction)
    {
        return direction == ListSortDirection.Descending
            ? SortDirection.Descending
            : SortDirection.Ascending;
    }

    private void RebuildRawDataColumns(IReadOnlyList<ImportedColumnInfo> columns)
    {
        RawDataGrid.Columns.Clear();

        foreach (var column in columns.OrderBy(column => column.DisplayOrder))
        {
            RawDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = column.GridHeader,
                Binding = new Binding(column.InternalColumnName),
                SortMemberPath = column.InternalColumnName,
                Width = new DataGridLength(160)
            });
        }
    }

    private void RebuildPartDetailColumns(IReadOnlyList<PartDetailColumn> columns)
    {
        PartDetailGrid.Columns.Clear();

        foreach (var column in columns.OrderBy(column => column.DisplayOrder))
        {
            PartDetailGrid.Columns.Add(new DataGridTextColumn
            {
                Header = column.Header,
                Binding = new Binding(column.BindingPath),
                Width = new DataGridLength(160)
            });
        }
    }
}
