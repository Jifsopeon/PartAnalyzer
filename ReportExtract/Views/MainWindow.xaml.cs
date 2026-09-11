using System.Windows;
using ReportExtract.Services;
using ReportExtract.ViewModels;

namespace ReportExtract.Views;

public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel;
    public MainWindow()
    {
        InitializeComponent();
        _viewModel = new MainViewModel(new SettingsService(), new ExcelWorkbookService(), new WorksheetProcessingSessionService(), new FileDialogService(), new DuckDbDataService(), new FidelityExportService());
        DataContext = _viewModel;
        Closed += (_, _) => _viewModel.Dispose();
    }
    private void OpenFilters_Click(object sender, RoutedEventArgs e) => new FilterWindow(_viewModel) { Owner = this }.ShowDialog();
}
