using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using Compass.Serialization;
using Microsoft.Win32;

namespace Compass.ViewModels;

public class DrillManagerViewModel : INotifyPropertyChanged
{
    public const int MinimumDrills = 1;
    public const int MaximumDrills = 20;

    private int _drillCount = 12;
    private int _selectedDrillIndex = 1;
    private string _selectedDrillName = string.Empty;
    private string _setAllName = string.Empty;
    private string _delimitedNames = string.Empty;
    private string _delimitedSeparator = ",";
    private string _cornerTopLeft = "NW";
    private string _cornerTopRight = "NE";
    private string _cornerBottomRight = "SE";
    private string _cornerBottomLeft = "SW";
    private string _autoFillPrefix = "DRILL ";
    private int _autoFillStartIndex = 1;
    private string _jsonStatusMessage = string.Empty;

    private readonly RelayCommand _setSelectedCommand;
    private readonly RelayCommand _clearSelectedCommand;
    private readonly RelayCommand _setAllCommand;
    private readonly RelayCommand _clearAllCommand;
    private readonly RelayCommand _applyDelimitedCommand;
    private readonly RelayCommand _copyDelimitedCommand;
    private readonly RelayCommand _applyCornersCommand;
    private readonly RelayCommand _autoFillCommand;
    private readonly RelayCommand _loadJsonCommand;
    private readonly RelayCommand _saveJsonCommand;

    public DrillManagerViewModel()
    {
        Drills = new ObservableCollection<DrillSlotViewModel>();
        DrillCountOptions = Enumerable.Range(MinimumDrills, MaximumDrills - MinimumDrills + 1).ToList();
        UpdateDrillSlots();
        DrillProps = new DrillPropsAccessor(this);

        _setSelectedCommand = new RelayCommand(_ => SetSelectedDrill(), _ => CanMutateSelectedDrill());
        _clearSelectedCommand = new RelayCommand(_ => ClearSelectedDrill(), _ => CanMutateSelectedDrill());
        _setAllCommand = new RelayCommand(_ => SetAllDrills(), _ => DrillCount > 0);
        _clearAllCommand = new RelayCommand(_ => ClearAllDrills(), _ => DrillCount > 0);
        _applyDelimitedCommand = new RelayCommand(_ => ApplyDelimitedNames(), _ => true);
        _copyDelimitedCommand = new RelayCommand(_ => CopyDelimitedNames(), _ => DrillCount > 0);
        _applyCornersCommand = new RelayCommand(_ => ApplyCornerTemplate(), _ => DrillCount > 0);
        _autoFillCommand = new RelayCommand(_ => AutoFillEmpty(), _ => DrillCount > 0);
        _loadJsonCommand = new RelayCommand(_ => LoadFromJsonDialog(), _ => true);
        _saveJsonCommand = new RelayCommand(_ => SaveToJsonDialog(), _ => DrillCount > 0);

        SelectedDrillIndex = 1;
        JsonStatusMessage = "Load or save drill definitions as JSON files.";
    }

    public ObservableCollection<DrillSlotViewModel> Drills { get; }

    public IReadOnlyList<int> DrillCountOptions { get; }

    public DrillPropsAccessor DrillProps { get; }

    public ICommand SetSelectedCommand => _setSelectedCommand;
    public ICommand ClearSelectedCommand => _clearSelectedCommand;
    public ICommand SetAllCommand => _setAllCommand;
    public ICommand ClearAllCommand => _clearAllCommand;
    public ICommand ApplyDelimitedCommand => _applyDelimitedCommand;
    public ICommand CopyDelimitedCommand => _copyDelimitedCommand;
    public ICommand ApplyCornersCommand => _applyCornersCommand;
    public ICommand AutoFillCommand => _autoFillCommand;
    public ICommand LoadJsonCommand => _loadJsonCommand;
    public ICommand SaveJsonCommand => _saveJsonCommand;

    public int DrillCount
    {
        get => _drillCount;
        set
        {
            var newValue = Math.Max(MinimumDrills, Math.Min(MaximumDrills, value));
            if (_drillCount != newValue)
            {
                _drillCount = newValue;
                OnPropertyChanged();
                UpdateDrillSlots();
                EnsureSelectedIndexInRange();
                RefreshCommandStates();
            }
        }
    }

    public int SelectedDrillIndex
    {
        get => _selectedDrillIndex;
        set
        {
            var maxAllowed = Math.Max(MinimumDrills, Math.Min(MaximumDrills, DrillCount));
            var newValue = Math.Max(MinimumDrills, Math.Min(maxAllowed, value));
            if (_selectedDrillIndex != newValue)
            {
                _selectedDrillIndex = newValue;
                OnPropertyChanged();
                RefreshSelectedName();
                RefreshCommandStates();
            }
        }
    }

    public string SelectedDrillName
    {
        get => _selectedDrillName;
        set
        {
            if (_selectedDrillName != value)
            {
                _selectedDrillName = value ?? string.Empty;
                OnPropertyChanged();
            }
        }
    }

    public string SetAllName
    {
        get => _setAllName;
        set
        {
            if (_setAllName != value)
            {
                _setAllName = value ?? string.Empty;
                OnPropertyChanged();
            }
        }
    }

    public string DelimitedNames
    {
        get => _delimitedNames;
        set
        {
            if (_delimitedNames != value)
            {
                _delimitedNames = value ?? string.Empty;
                OnPropertyChanged();
            }
        }
    }

    public string DelimitedSeparator
    {
        get => _delimitedSeparator;
        set
        {
            var normalized = string.IsNullOrEmpty(value) ? "," : value.Substring(0, 1);
            if (_delimitedSeparator != normalized)
            {
                _delimitedSeparator = normalized;
                OnPropertyChanged();
            }
        }
    }

    public string CornerTopLeft
    {
        get => _cornerTopLeft;
        set
        {
            if (_cornerTopLeft != value)
            {
                _cornerTopLeft = value ?? string.Empty;
                OnPropertyChanged();
            }
        }
    }

    public string CornerTopRight
    {
        get => _cornerTopRight;
        set
        {
            if (_cornerTopRight != value)
            {
                _cornerTopRight = value ?? string.Empty;
                OnPropertyChanged();
            }
        }
    }

    public string CornerBottomRight
    {
        get => _cornerBottomRight;
        set
        {
            if (_cornerBottomRight != value)
            {
                _cornerBottomRight = value ?? string.Empty;
                OnPropertyChanged();
            }
        }
    }

    public string CornerBottomLeft
    {
        get => _cornerBottomLeft;
        set
        {
            if (_cornerBottomLeft != value)
            {
                _cornerBottomLeft = value ?? string.Empty;
                OnPropertyChanged();
            }
        }
    }

    public string AutoFillPrefix
    {
        get => _autoFillPrefix;
        set
        {
            if (_autoFillPrefix != value)
            {
                _autoFillPrefix = value ?? string.Empty;
                OnPropertyChanged();
            }
        }
    }

    public string JsonStatusMessage
    {
        get => _jsonStatusMessage;
        set
        {
            if (_jsonStatusMessage != value)
            {
                _jsonStatusMessage = value ?? string.Empty;
                OnPropertyChanged();
            }
        }
    }

    public int AutoFillStartIndex
    {
        get => _autoFillStartIndex;
        set
        {
            var newValue = Math.Max(1, value);
            if (_autoFillStartIndex != newValue)
            {
                _autoFillStartIndex = newValue;
                OnPropertyChanged();
            }
        }
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public void LoadExistingNames(IEnumerable<string> names)
    {
        var nameList = names?.Where(name => !string.IsNullOrWhiteSpace(name)).ToList() ?? new List<string>();
        DrillCount = Math.Max(MinimumDrills, Math.Min(MaximumDrills, nameList.Count));
        for (var i = 0; i < Drills.Count; i++)
        {
            Drills[i].Name = i < nameList.Count ? nameList[i] : string.Empty;
        }

        RefreshSelectedName();
        RefreshCommandStates();
    }

    public int LoadFromJson(string path)
    {
        var names = DrillPropsJsonStore.Load(path);
        DrillProps.SetDrillProps(names);
        JsonStatusMessage = $"Loaded {names.Count} drills from '{Path.GetFileName(path)}'.";
        return names.Count;
    }

    public void SaveToJson(string path)
    {
        DrillPropsJsonStore.Save(path, DrillProps.GetDrillProps());
        JsonStatusMessage = $"Saved drills to '{Path.GetFileName(path)}'.";
    }

    internal DrillSlotViewModel EnsureSlot(int index)
    {
        if (index < MinimumDrills || index > MaximumDrills)
        {
            throw new ArgumentOutOfRangeException(nameof(index), index, $"Index must be between {MinimumDrills} and {MaximumDrills}.");
        }

        if (_drillCount < index)
        {
            DrillCount = index;
        }

        return Drills[index - 1];
    }

    internal DrillSlotViewModel? TryGetSlot(int index)
    {
        if (index < MinimumDrills || index > MaximumDrills)
        {
            return null;
        }

        if (index > Drills.Count)
        {
            return null;
        }

        return Drills[index - 1];
    }

    private void UpdateDrillSlots()
    {
        if (Drills.Count < _drillCount)
        {
            for (var i = Drills.Count + 1; i <= _drillCount; i++)
            {
                Drills.Add(new DrillSlotViewModel(i));
            }
        }
        else if (Drills.Count > _drillCount)
        {
            while (Drills.Count > _drillCount)
            {
                Drills.RemoveAt(Drills.Count - 1);
            }
        }

        OnPropertyChanged(nameof(Drills));
    }

    private bool CanMutateSelectedDrill()
    {
        return SelectedDrillIndex >= MinimumDrills && SelectedDrillIndex <= DrillCount;
    }

    private void SetSelectedDrill()
    {
        DrillProps.SetDrillProp(SelectedDrillIndex, SelectedDrillName);
        RefreshSelectedName();
    }

    private void ClearSelectedDrill()
    {
        DrillProps.ClearDrillProp(SelectedDrillIndex);
        RefreshSelectedName();
    }

    public void SetAllDrills()
    {
        SetAllDrills(SetAllName);
    }

    public void SetAllDrills(string? value)
    {
        var normalized = value ?? string.Empty;
        var values = Enumerable.Repeat(normalized, DrillCount);
        DrillProps.SetDrillProps(values);
        RefreshSelectedName();
    }

    public void ClearAllDrills()
    {
        DrillProps.ClearAllDrillProps();
        RefreshSelectedName();
    }

    private void ApplyDelimitedNames()
    {
        var separator = string.IsNullOrEmpty(DelimitedSeparator) ? ',' : DelimitedSeparator[0];
        DrillProps.LoadFromDelimitedList(DelimitedNames, separator);
        RefreshSelectedName();
    }

    private void CopyDelimitedNames()
    {
        var separator = string.IsNullOrEmpty(DelimitedSeparator) ? ',' : DelimitedSeparator[0];
        DelimitedNames = DrillProps.ToDelimitedList(separator);
    }

    public void ApplyCornerTemplate()
    {
        const int required = 4;
        if (DrillCount < required)
        {
            DrillCount = required;
        }

        DrillProps.SetDrillProp(1, CornerTopLeft);
        DrillProps.SetDrillProp(2, CornerTopRight);
        DrillProps.SetDrillProp(3, CornerBottomRight);
        DrillProps.SetDrillProp(4, CornerBottomLeft);
        RefreshSelectedName();
    }

    public void ApplyCornerTemplate(string topLeft, string topRight, string bottomRight, string bottomLeft)
    {
        CornerTopLeft = topLeft ?? string.Empty;
        CornerTopRight = topRight ?? string.Empty;
        CornerBottomRight = bottomRight ?? string.Empty;
        CornerBottomLeft = bottomLeft ?? string.Empty;
        ApplyCornerTemplate();
    }

    public void AutoFillEmpty()
    {
        var prefix = AutoFillPrefix ?? string.Empty;
        var start = AutoFillStartIndex;
        DrillProps.FillEmptyWith(i =>
        {
            var value = start + i - 1;
            return $"{prefix}{value}".Trim();
        });
        RefreshSelectedName();
    }

    public void AutoFillEmpty(string prefix, int startIndex)
    {
        AutoFillPrefix = prefix ?? string.Empty;
        AutoFillStartIndex = Math.Max(1, startIndex);
        AutoFillEmpty();
    }

    private void LoadFromJsonDialog()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"
        };

        if (dialog.ShowDialog() == true)
        {
            try
            {
                LoadFromJson(dialog.FileName);
            }
            catch (Exception ex)
            {
                JsonStatusMessage = $"Failed to load JSON: {ex.Message}";
            }
        }
    }

    private void SaveToJsonDialog()
    {
        var dialog = new SaveFileDialog
        {
            Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
            DefaultExt = ".json"
        };

        if (dialog.ShowDialog() == true)
        {
            try
            {
                SaveToJson(dialog.FileName);
            }
            catch (Exception ex)
            {
                JsonStatusMessage = $"Failed to save JSON: {ex.Message}";
            }
        }
    }

    private void RefreshSelectedName()
    {
        if (SelectedDrillIndex >= MinimumDrills && SelectedDrillIndex <= MaximumDrills)
        {
            SelectedDrillName = DrillProps.GetDrillProp(SelectedDrillIndex);
        }
    }

    private void EnsureSelectedIndexInRange()
    {
        if (_selectedDrillIndex > _drillCount)
        {
            SelectedDrillIndex = _drillCount;
        }
        else if (_selectedDrillIndex < MinimumDrills)
        {
            SelectedDrillIndex = MinimumDrills;
        }
        else
        {
            RefreshSelectedName();
        }
    }

    private void RefreshCommandStates()
    {
        _setSelectedCommand.RaiseCanExecuteChanged();
        _clearSelectedCommand.RaiseCanExecuteChanged();
        _setAllCommand.RaiseCanExecuteChanged();
        _clearAllCommand.RaiseCanExecuteChanged();
        _copyDelimitedCommand.RaiseCanExecuteChanged();
        _applyCornersCommand.RaiseCanExecuteChanged();
        _autoFillCommand.RaiseCanExecuteChanged();
        _saveJsonCommand.RaiseCanExecuteChanged();
    }

    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
