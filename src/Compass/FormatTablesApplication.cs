using System;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Compass.UI;
using Compass.ViewModels;

[assembly: CommandClass(typeof(Compass.FormatTablesApplication))]

namespace Compass;

public class FormatTablesApplication : IExtensionApplication
{
    private static readonly MacroToolDefinition[] FormatTableTools =
    {
        new(
            "format-surface-impact-box",
            "Format Surface Impact Box",
            "Fix Surface Impact Cell and Text Sizes",
            @"^C^Csift"),
        new(
            "format-hybrid-table",
            "Format Hybrid Table",
            "Fix Hybrid Cell and Text Sizes",
            @"^C^CHFT"),
        new(
            "format-well-coordinate-table",
            "Format Well Coordinate Table",
            "Fix Well Coordinate Cell and Text Sizes",
            @"^C^CWFT"),
        new(
            "format-workspace-table",
            "Format Workspace Table",
            "Fix Workspace Cell and Text Sizes",
            @"^C^CWSFT")
    };

    private static CompassControl? _control;

    public void Initialize()
    {
        // Preload the Format Tables palette during initialization so the tab is ready when needed
        EnsurePalette();
    }

    public void Terminate()
    {
        if (_control != null)
        {
            _control.ModuleRequested -= OnToolRequested;
            _control = null;
        }
    }

    [CommandMethod("Cformat", CommandFlags.Modal | CommandFlags.Session)]
    public static void ShowFormatTables()
    {
        EnsurePalette();
        UnifiedPaletteHost.ShowPalette("Format Tables");
    }

    private static void EnsurePalette()
    {
        if (_control != null)
        {
            return;
        }

        _control = new CompassControl
        {
            TitleText = "FORMAT TABLES",
            SubtitleText = "Select a formatting tool to launch."
        };

        _control.ModuleRequested += OnToolRequested;
        _control.LoadModules(FormatTableTools.Select((tool, index) =>
            new CompassModuleDefinition(tool.Id, tool.DisplayName, tool.Description, index)));

        UnifiedPaletteHost.EnsurePalette();
        UnifiedPaletteHost.AddTab("Format Tables", _control);
    }

    private static void OnToolRequested(object? sender, string toolId)
    {
        var tool = FormatTableTools.FirstOrDefault(t => t.Id.Equals(toolId, StringComparison.OrdinalIgnoreCase));
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

        document.SendStringToExecute($"{tool.Macro}\\n", true, false, false);
    }

    private record MacroToolDefinition(string Id, string DisplayName, string Description, string Macro);
}
