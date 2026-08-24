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
            new FileDialogService(),
            new DuckDbDataService());
        DataContext = _viewModel;
        _viewModel.PropertyChanged += ViewModelPropertyChanged;
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

    private void RebuildRawDataColumns(IReadOnlyList<ImportedColumnInfo> columns)
    {
        RawDataGrid.Columns.Clear();

        foreach (var column in columns.OrderBy(column => column.DisplayOrder))
        {
            RawDataGrid.Columns.Add(new DataGridTextColumn
            {
                Header = column.GridHeader,
                Binding = new Binding(column.InternalColumnName),
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
