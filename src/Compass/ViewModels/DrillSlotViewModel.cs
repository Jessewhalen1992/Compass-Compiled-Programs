using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Compass.ViewModels;

public class DrillSlotViewModel : INotifyPropertyChanged
{
    private string _name = string.Empty;

    public DrillSlotViewModel(int index)
    {
        Index = index;
    }

    public int Index { get; }

    public string DisplayLabel => $"Drill {Index}";

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
