using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace Compass.ViewModels;

public class DrillManagerViewModel : INotifyPropertyChanged
{
    public const int MinimumDrills = 1;
    public const int MaximumDrills = 20;

    private int _drillCount = 12;

    public DrillManagerViewModel()
    {
        Drills = new ObservableCollection<DrillSlotViewModel>();
        DrillCountOptions = Enumerable.Range(MinimumDrills, MaximumDrills - MinimumDrills + 1).ToList();
        UpdateDrillSlots();
        DrillProps = new DrillPropsAccessor(this);
    }

    public ObservableCollection<DrillSlotViewModel> Drills { get; }

    public IReadOnlyList<int> DrillCountOptions { get; }

    public DrillPropsAccessor DrillProps { get; }

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

    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
