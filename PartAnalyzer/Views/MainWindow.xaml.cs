using System.ComponentModel;
using System.Windows;
using PartAnalyzer.Services;
using PartAnalyzer.ViewModels;

namespace PartAnalyzer.Views;

public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel;
    public MainWindow()
    {
        InitializeComponent();
        _viewModel = new MainViewModel(new SettingsService(), new ExcelWorkbookService(), new WorksheetProcessingSessionService(), new FileDialogService(), new DuckDbDataService(), new FidelityExportService());
        DataContext = _viewModel;
        Closing += MainWindowClosing;
        Closed += (_, _) => _viewModel.Dispose();
    }
    private void OpenFilters_Click(object sender, RoutedEventArgs e) => new FilterWindow(_viewModel) { Owner = this }.ShowDialog();
    private void MainWindowClosing(object? sender, CancelEventArgs e) { if (!_viewModel.ConfirmCloseWithUnsavedChanges()) e.Cancel = true; }
}
