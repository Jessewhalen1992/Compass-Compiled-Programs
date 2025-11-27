using Compass;
using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.Runtime;
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
    private static CompassControl? _compassControl;

    public void Initialize()
    {
        CompassEnvironment.Initialize();
        EnsureModules();
        EnsureCompassPaletteTab();

        // Explicitly initialize other modules so they add their tabs
        new FormatTablesApplication().Initialize();
        new CompassToolsApplication().Initialize();
        new CompassLegalApplication().Initialize();
        new CogoApplication().Initialize();
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

        if (_compassControl != null)
        {
            _compassControl.ModuleRequested -= OnModuleRequested;
            _compassControl = null;
        }
    }

    [CommandMethod("Compass", CommandFlags.Modal | CommandFlags.Session)]
    public static void ShowCompass()
    {
        EnsureModules();
        EnsureCompassPaletteTab();
        UnifiedPaletteHost.ShowPalette("Compass");
    }

    private static void EnsureModules()
    {
        if (Modules.Count > 0)
        {
            return;
        }

        RegisterModule(new DrillManagerModule());
        RegisterModule(new ProfileManagerModule());
        RegisterModule(new SectionGeneratorModule());
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

    private static void EnsureCompassPaletteTab()
    {
        if (_compassControl != null)
        {
            return;
        }

        _compassControl = new CompassControl();
        _compassControl.ModuleRequested += OnModuleRequested;

        var definitions = ModuleList
            .Select((module, index) => new CompassModuleDefinition(module.Id, module.DisplayName, module.Description, index))
            .ToArray();

        _compassControl.LoadModules(definitions);

        UnifiedPaletteHost.EnsurePalette();
        UnifiedPaletteHost.AddTab("Compass", _compassControl);
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
