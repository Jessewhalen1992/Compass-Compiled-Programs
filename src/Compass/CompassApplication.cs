using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using Compass.Infrastructure;
using Compass.Modules;
using Compass.UI;
using Compass.ViewModels;

[assembly: CommandClass(typeof(Compass.CompassApplication))]

namespace Compass;

public class CompassApplication : IExtensionApplication
{
    private static readonly Dictionary<string, ICompassModule> Modules = new(StringComparer.OrdinalIgnoreCase);
    private static readonly List<ICompassModule> ModuleList = new();
    private static PaletteSet? _compassPalette;
    private static CompassControl? _compassControl;

    public void Initialize()
    {
        CompassEnvironment.Initialize();
        EnsureModules();
    }

    public void Terminate()
    {
        try
        {
            GetDrillManagerModule().SaveState();
        }
        catch (System.Exception)
        {
            // ignore shutdown failures
        }

        if (_compassPalette != null)
        {
            _compassPalette.Visible = false;
            _compassPalette.Dispose();
            _compassPalette = null;
        }
    }

    [CommandMethod("Compass", CommandFlags.Modal | CommandFlags.Session)]
    public static void ShowCompass()
    {
        EnsureModules();
        EnsureCompassPalette();
        if (_compassPalette != null)
        {
            _compassPalette.Visible = true;
            _compassPalette.Activate(0);
        }
    }

    private static void EnsureModules()
    {
        if (Modules.Count > 0)
        {
            return;
        }

        RegisterModule(new DrillManagerModule());
        RegisterModule(new ProfileManagerModule());
        RegisterModule(new SurfaceDevelopmentModule());
        RegisterModule(new CrossingManagerModule());
        RegisterModule(new HybridManagerModule());
        RegisterModule(new WorkspaceManagerModule());
        RegisterModule(new OnestopManagerModule());
    }

    public static void RegisterModule(ICompassModule module)
    {
        if (Modules.TryGetValue(module.Id, out var existing))
        {
            ModuleList.Remove(existing);
        }

        Modules[module.Id] = module;
        ModuleList.Add(module);

        if (_compassControl != null)
        {
            var definitions = ModuleList
                .Select((m, index) => new CompassModuleDefinition(m.Id, m.DisplayName, m.Description, index))
                .ToArray();

            _compassControl.LoadModules(definitions);
        }
    }

    private static void EnsureCompassPalette()
    {
        if (_compassPalette != null)
        {
            return;
        }

        _compassControl = new CompassControl();
        _compassControl.ModuleRequested += OnModuleRequested;

        var definitions = ModuleList
            .Select((module, index) => new CompassModuleDefinition(module.Id, module.DisplayName, module.Description, index))
            .ToArray();

        _compassControl.LoadModules(definitions);

        _compassPalette = new PaletteSet("Compass", new Guid("c833fa27-9db1-4d67-85d4-45115ac0a2c2"))
        {
            Style = PaletteSetStyles.ShowCloseButton | PaletteSetStyles.ShowPropertiesMenu | PaletteSetStyles.Snappable,
            DockEnabled = DockSides.Left | DockSides.Right | DockSides.Top | DockSides.Bottom,
            Visible = false
        };

        _compassPalette.MinimumSize = new System.Drawing.Size(320, 240);
        _compassPalette.AddVisual("Programs", _compassControl);
    }

    private static void OnModuleRequested(object? sender, string moduleId)
    {
        if (Modules.TryGetValue(moduleId, out var module))
        {
            module.Show();
        }
    }

    internal static DrillManagerModule GetDrillManagerModule()
    {
        EnsureModules();

        var module = ModuleList.OfType<DrillManagerModule>().FirstOrDefault();
        if (module == null)
        {
            module = new DrillManagerModule();
            RegisterModule(module);
        }

        return module;
    }
}
