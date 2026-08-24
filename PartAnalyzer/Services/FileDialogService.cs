using System.IO;
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
}
