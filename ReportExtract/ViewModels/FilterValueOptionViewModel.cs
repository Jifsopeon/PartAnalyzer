using ReportExtract.Models;
using ReportExtract.Utilities;

namespace ReportExtract.ViewModels;

public sealed class FilterValueOptionViewModel : ObservableObject
{
    private bool _isSelected;

    public FilterValueOptionViewModel(FilterValueOption option)
    {
        Option = option;
    }

    public FilterValueOption Option { get; }

    public string DisplayWithCount => Option.DisplayWithCount;

    public string? Value => Option.Value;

    public bool IsSelected
    {
        get => _isSelected;
        set => SetProperty(ref _isSelected, value);
    }
}
