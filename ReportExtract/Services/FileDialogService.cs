using System.IO;
using System.Windows;
using Microsoft.Win32;

namespace ReportExtract.Services;

public sealed class FileDialogService
{
    public string? SelectExcelWorkbook(string? initialDirectory)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Import Excel File",
            Filter = "Excel workbooks (*.xlsx)|*.xlsx",
            CheckFileExists = true,
            Multiselect = false
        };

        if (!string.IsNullOrWhiteSpace(initialDirectory) && Directory.Exists(initialDirectory))
        {
            dialog.InitialDirectory = initialDirectory;
        }

        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }

    public string? SelectExportWorkbook(string suggestedFileName, string? initialDirectory)
    {
        var dialog = new SaveFileDialog
        {
            Title = "Export Excel File",
            Filter = "Excel workbook (*.xlsx)|*.xlsx",
            DefaultExt = ".xlsx",
            AddExtension = true,
            OverwritePrompt = true,
            FileName = suggestedFileName
        };

        if (!string.IsNullOrWhiteSpace(initialDirectory) && Directory.Exists(initialDirectory))
        {
            dialog.InitialDirectory = initialDirectory;
        }

        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }

    public bool ConfirmUnavailableSelections(string message)
    {
        return MessageBox.Show(message, "Unavailable filter selections", MessageBoxButton.OKCancel, MessageBoxImage.Warning, MessageBoxResult.Cancel) == MessageBoxResult.OK;
    }

    public void ShowNoMatchingRows()
    {
        MessageBox.Show(
            "No matching rows were found for the current filter selections.",
            "No Matching Rows",
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }
}
