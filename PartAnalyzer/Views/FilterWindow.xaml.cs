using System.Windows;
using PartAnalyzer.ViewModels;

namespace PartAnalyzer.Views;

public partial class FilterWindow : Window
{
    public FilterWindow(MainViewModel viewModel)
    {
        InitializeComponent();
        DataContext = viewModel;
    }

    private void Close_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private async void LoadPreset_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not MainViewModel viewModel || string.IsNullOrWhiteSpace(viewModel.PresetName)) return;
        var warning = await viewModel.LoadPresetAsync(viewModel.PresetName);
        if (!string.IsNullOrWhiteSpace(warning)) MessageBox.Show($"Unavailable preset selections: {warning}", "Preset warning", MessageBoxButton.OK, MessageBoxImage.Warning);
    }

    private void SavePreset_Click(object sender, RoutedEventArgs e) => SavePreset(false);
    private void UpdatePreset_Click(object sender, RoutedEventArgs e) => SavePreset(true);

    private void SavePreset(bool overwrite)
    {
        try
        {
            if (DataContext is not MainViewModel viewModel || string.IsNullOrWhiteSpace(viewModel.PresetName)) return;
            viewModel.SavePreset(viewModel.PresetName, overwrite);
            MessageBox.Show(overwrite ? "Preset updated." : "Preset saved.", "Preset", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (InvalidOperationException error) { MessageBox.Show(error.Message, "Preset", MessageBoxButton.OK, MessageBoxImage.Warning); }
    }

    private void DeletePreset_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not MainViewModel viewModel || string.IsNullOrWhiteSpace(viewModel.PresetName)) return;
        if (!viewModel.PresetNames.Contains(viewModel.PresetName, StringComparer.OrdinalIgnoreCase)) return;
        viewModel.DeletePreset(viewModel.PresetName);
        viewModel.PresetName = "Generic";
        MessageBox.Show("Preset deleted.", "Preset", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void ClearFilter_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.Tag is ConstrainedFilterViewModel filter) filter.ClearSelections();
    }
}
