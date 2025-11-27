using System;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Compass.UI;
using Compass.ViewModels;

[assembly: CommandClass(typeof(Compass.CompassLegalApplication))]

namespace Compass;

public class CompassLegalApplication : IExtensionApplication
{
    private static readonly MacroToolDefinition[] LegalTools =
    {
        new(
            "plan-to-legal",
            "Plan to Legal",
            "Convert Plan to Legal Plan",
            // note the doubled quotes inside the verbatim string
            @"^C^C(IF (NOT C:leglayers) (LOAD ""leglayers""));legf"),
        new(
            "legal-layer-convert",
            "Legal Layer Convert",
            "Convert Layers to Legal Layers",
            @"^C^C(IF (NOT C:PLTO) (LOAD ""PLTO""));PLTO"),
        new(
            "lto-check",
            "LTO Check",
            "Check LTO Layers (Use in Paper Space)",
            @"^C^C(IF (NOT C:LTO_Check) (LOAD ""LTO_Check""));check")
    };

    private static CompassControl? _control;

    public void Initialize()
    {
        // Only create and register the Legal tab once
        if (_control == null)
        {
            _control = new CompassControl
            {
                TitleText = "Compass Legal",
                SubtitleText = "Select a legal tool to launch."
            };
            _control.ModuleRequested += OnToolRequested;
            _control.LoadModules(LegalTools.Select((tool, index) =>
                new CompassModuleDefinition(tool.Id, tool.DisplayName, tool.Description, index)));

            // Preload the Legal tab in the unified palette
            UnifiedPaletteHost.EnsurePalette();
            UnifiedPaletteHost.AddTab("Legal", _control);
        }
    }

    public void Terminate()
    {
        if (_control != null)
        {
            _control.ModuleRequested -= OnToolRequested;
            _control = null;
        }
    }

    [CommandMethod("Clegal", CommandFlags.Modal | CommandFlags.Session)]
    public static void ShowCompassLegal()
    {
        EnsurePalette();
        UnifiedPaletteHost.ShowPalette("Legal");
    }

    private static void EnsurePalette()
    {
        if (_control != null)
        {
            return;
        }

        _control = new CompassControl
        {
            TitleText = "Compass Legal",
            SubtitleText = "Select a legal tool to launch."
        };

        _control.ModuleRequested += OnToolRequested;
        _control.LoadModules(LegalTools.Select((tool, index) =>
            new CompassModuleDefinition(tool.Id, tool.DisplayName, tool.Description, index)));

        UnifiedPaletteHost.EnsurePalette();
        UnifiedPaletteHost.AddTab("Legal", _control);
    }

    private static void OnToolRequested(object? sender, string toolId)
    {
        var tool = LegalTools.FirstOrDefault(t => t.Id.Equals(toolId, StringComparison.OrdinalIgnoreCase));
        if (tool == null)
        {
            return;
        }

        LaunchTool(tool);
    }

    private static void LaunchTool(MacroToolDefinition tool)
    {
        var document = Application.DocumentManager.MdiActiveDocument;
        if (document == null)
        {
            Application.ShowAlertDialog("Open a drawing first.");
            return;
        }

        document.SendStringToExecute($"{tool.Macro}\n", true, false, false);
    }

    private record MacroToolDefinition(string Id, string DisplayName, string Description, string Macro);
}
