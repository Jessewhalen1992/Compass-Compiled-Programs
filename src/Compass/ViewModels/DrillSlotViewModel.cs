using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Compass.ViewModels;

public class DrillSlotViewModel : INotifyPropertyChanged
{
    private string _name = string.Empty;
    private string _committedName = string.Empty;

    public DrillSlotViewModel(int index)
    {
        Index = index;
    }

    public int Index { get; }

    public string DisplayLabel
    {
        get
        {
            var name = Name?.Trim();
            if (string.IsNullOrEmpty(name))
            {
                return $"Drill {Index}";
            }

            return $"Drill {Index} - {name}";
        }
    }

    public string Name
    {
        get => _name;
        set
        {
            if (_name != value)
            {
                _name = value;
                OnPropertyChanged();
            }
        }
    }

    public string CommittedName
    {
        get => _committedName;
        private set
        {
            if (_committedName != value)
            {
                _committedName = value;
                OnPropertyChanged();
            }
        }
    }

    public void Commit()
    {
        CommittedName = _name;
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        if (propertyName == nameof(Name))
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(DisplayLabel)));
        }
    }
}
