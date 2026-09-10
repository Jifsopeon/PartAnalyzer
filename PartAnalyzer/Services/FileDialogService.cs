using System.IO;
using System.Windows;
using Microsoft.Win32;

namespace PartAnalyzer.Services;

public sealed class FileDialogService
{
    public string? SelectExcelWorkbook(string? initialDirectory)
    {
        var dialog = new OpenFileDialog
        {
            Title = "Import Excel File",
            Filter = "Excel workbooks (*.xlsx;*.xlsm)|*.xlsx;*.xlsm",
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

    public bool ConfirmDiscardUnsavedChanges(string message)
    {
        return MessageBox.Show(
            message,
            "Unsaved Changes",
            MessageBoxButton.OKCancel,
            MessageBoxImage.Warning,
            MessageBoxResult.Cancel) == MessageBoxResult.OK;
    }
}
