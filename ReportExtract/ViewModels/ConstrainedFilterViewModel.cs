using System.Collections.ObjectModel;
using System.ComponentModel;
using ReportExtract.Models;
using ReportExtract.Utilities;

namespace ReportExtract.ViewModels;

public sealed class ConstrainedFilterViewModel : ObservableObject
{
    private readonly HashSet<string> _selectedValues = new(StringComparer.Ordinal);
    private string? _searchText;

    public ConstrainedFilterViewModel(SessionFilterColumn column, string displayName)
    {
        Column = column;
        DisplayName = displayName;
    }

    public SessionFilterColumn Column { get; }
    public string DisplayName { get; }
    public ObservableCollection<FilterValueOptionViewModel> Options { get; } = new();
    public IReadOnlyCollection<string> SelectedValues => _selectedValues;

    public string? SearchText
    {
        get => _searchText;
        set => SetProperty(ref _searchText, value);
    }

    public bool IsActive => _selectedValues.Count > 0;

    public void SetSelections(IEnumerable<string> values)
    {
        _selectedValues.Clear();
        foreach (var value in values) _selectedValues.Add(value);
        foreach (var option in Options) option.IsSelected = _selectedValues.Contains(option.Value ?? string.Empty);
        OnPropertyChanged(nameof(IsActive));
        OnPropertyChanged(nameof(SelectedValues));
    }

    public void ReplaceOptions(IEnumerable<FilterValueOption> options)
    {
        foreach (var existing in Options) existing.PropertyChanged -= OptionPropertyChanged;
        Options.Clear();
        foreach (var option in options)
        {
            var vm = new FilterValueOptionViewModel(option) { IsSelected = option.Value is not null && _selectedValues.Contains(option.Value) };
            vm.PropertyChanged += OptionPropertyChanged;
            Options.Add(vm);
        }
    }

    public void ClearSelections() => SetSelections(Array.Empty<string>());

    private void OptionPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName != nameof(FilterValueOptionViewModel.IsSelected) || sender is not FilterValueOptionViewModel option || option.Value is null) return;
        if (option.IsSelected) _selectedValues.Add(option.Value); else _selectedValues.Remove(option.Value);
        OnPropertyChanged(nameof(IsActive));
        OnPropertyChanged(nameof(SelectedValues));
    }
}
