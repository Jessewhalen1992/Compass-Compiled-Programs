using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Compass.ViewModels;

namespace Compass.UI;

public partial class CompassControl : UserControl
{
    private readonly ObservableCollection<CompassModuleDefinition> _modules = new();

    public CompassControl()
    {
        InitializeComponent();
        ModulesList.ItemsSource = _modules;
    }

    public event EventHandler<string>? ModuleRequested;

    public void LoadModules(params CompassModuleDefinition[] modules)
    {
        LoadModules(modules.AsEnumerable());
    }

    public void LoadModules(System.Collections.Generic.IEnumerable<CompassModuleDefinition> modules)
    {
        _modules.Clear();
        foreach (var module in modules.OrderBy(m => m.DisplayOrder))
        {
            _modules.Add(module);
        }
    }

    private void OnModuleButtonClick(object sender, RoutedEventArgs e)
    {
        if (sender is Button button && button.Tag is string id)
        {
            ModuleRequested?.Invoke(this, id);
        }
    }
}
