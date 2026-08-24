using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using PartAnalyzer.Models;
using PartAnalyzer.Utilities;

namespace PartAnalyzer.ViewModels;

public sealed class ActiveFilterViewModel : ObservableObject
{
    private readonly HashSet<string?> _selectedValues = new();
    private string? _text;
    private string? _minimumText;
    private string? _maximumText;
    private string? _valueSearchText;
    private TextFilterMode _textMode = TextFilterMode.Contains;

    public ActiveFilterViewModel(FilterDefinition definition)
    {
        Definition = definition;
        ValueOptions.CollectionChanged += (_, _) => OnPropertyChanged(nameof(IsActive));
    }

    public FilterDefinition Definition { get; }

    public ObservableCollection<FilterValueOptionViewModel> ValueOptions { get; } = new();

    public string DisplayName => Definition.DisplayName;

    public string FilterKindLabel => Definition.Kind switch
    {
        FilterKind.ValueList => "Values",
        FilterKind.SearchableValueList => "Searchable values",
        FilterKind.NumericRange => "Numeric range",
        FilterKind.ComputedBoolean => "Yes/No",
        FilterKind.ComputedRange => "Range",
        _ => "Text"
    };

    public bool UsesValues => Definition.Kind is FilterKind.ValueList or FilterKind.SearchableValueList or FilterKind.ComputedBoolean;

    public bool UsesValueSearch => UsesValues;

    public bool UsesText => Definition.Kind == FilterKind.Text;

    public bool UsesRange => Definition.Kind is FilterKind.NumericRange or FilterKind.ComputedRange;

    public bool HasValidationError => !string.IsNullOrWhiteSpace(ValidationMessage);

    public string? ValidationMessage { get; private set; }

    public string? Text
    {
        get => _text;
        set
        {
            if (SetProperty(ref _text, value))
            {
                OnPropertyChanged(nameof(IsActive));
            }
        }
    }

    public TextFilterMode TextMode
    {
        get => _textMode;
        set => SetProperty(ref _textMode, value);
    }

    public string? MinimumText
    {
        get => _minimumText;
        set
        {
            if (SetProperty(ref _minimumText, value))
            {
                OnPropertyChanged(nameof(IsActive));
            }
        }
    }

    public string? MaximumText
    {
        get => _maximumText;
        set
        {
            if (SetProperty(ref _maximumText, value))
            {
                OnPropertyChanged(nameof(IsActive));
            }
        }
    }

    public string? ValueSearchText
    {
        get => _valueSearchText;
        set => SetProperty(ref _valueSearchText, value);
    }

    public bool IsActive
    {
        get
        {
            if (UsesValues)
            {
                return _selectedValues.Count > 0;
            }

            if (UsesText)
            {
                return !string.IsNullOrWhiteSpace(Text);
            }

            return !string.IsNullOrWhiteSpace(MinimumText) || !string.IsNullOrWhiteSpace(MaximumText);
        }
    }

    public FilterCriteria? ToCriteria()
    {
        ValidationMessage = null;
        OnPropertyChanged(nameof(ValidationMessage));
        OnPropertyChanged(nameof(HasValidationError));

        double? minimum = null;
        double? maximum = null;
        if (UsesRange)
        {
            if (!string.IsNullOrWhiteSpace(MinimumText) && !double.TryParse(MinimumText, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
            {
                SetValidation("Minimum must be a valid number.");
                return null;
            }

            if (!string.IsNullOrWhiteSpace(MaximumText) && !double.TryParse(MaximumText, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
            {
                SetValidation("Maximum must be a valid number.");
                return null;
            }

            minimum = string.IsNullOrWhiteSpace(MinimumText) ? null : double.Parse(MinimumText, CultureInfo.InvariantCulture);
            maximum = string.IsNullOrWhiteSpace(MaximumText) ? null : double.Parse(MaximumText, CultureInfo.InvariantCulture);
        }

        return new FilterCriteria
        {
            Definition = Definition,
            SelectedValues = _selectedValues.ToList(),
            TextMode = TextMode,
            Text = Text,
            Minimum = minimum,
            Maximum = maximum,
            IsActive = IsActive
        };
    }

    public void Clear()
    {
        Text = null;
        MinimumText = null;
        MaximumText = null;
        ValueSearchText = null;
        _selectedValues.Clear();
        foreach (var option in ValueOptions)
        {
            option.IsSelected = false;
        }

        OnPropertyChanged(nameof(IsActive));
    }

    public void ReplaceOptions(IEnumerable<FilterValueOption> options)
    {
        foreach (var existing in ValueOptions)
        {
            existing.PropertyChanged -= OptionPropertyChanged;
        }

        ValueOptions.Clear();
        foreach (var option in options)
        {
            var viewModel = new FilterValueOptionViewModel(option)
            {
                IsSelected = _selectedValues.Contains(option.Value)
            };
            viewModel.PropertyChanged += OptionPropertyChanged;
            ValueOptions.Add(viewModel);
        }

        OnPropertyChanged(nameof(IsActive));
    }

    private void OptionPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(FilterValueOptionViewModel.IsSelected))
        {
            if (sender is FilterValueOptionViewModel option)
            {
                if (option.IsSelected)
                {
                    _selectedValues.Add(option.Value);
                }
                else
                {
                    _selectedValues.Remove(option.Value);
                }
            }

            OnPropertyChanged(nameof(IsActive));
        }
    }

    private void SetValidation(string message)
    {
        ValidationMessage = message;
        OnPropertyChanged(nameof(ValidationMessage));
        OnPropertyChanged(nameof(HasValidationError));
    }
}
