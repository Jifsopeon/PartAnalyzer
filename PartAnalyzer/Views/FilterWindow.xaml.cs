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
}
