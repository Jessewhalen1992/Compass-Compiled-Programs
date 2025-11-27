using System;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
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

    private static PaletteSet? _palette;
    private static CompassControl? _control;

    public void Initialize()
    {
    }

    public void Terminate()
    {
        if (_palette != null)
        {
            _palette.Visible = false;
            _palette.Dispose();
            _palette = null;
        }
    }

    [CommandMethod("Cformat", CommandFlags.Modal | CommandFlags.Session)]
    public static void ShowFormatTables()
    {
        EnsurePalette();
        if (_palette != null)
        {
            _palette.Visible = true;
            _palette.Activate(0);
        }
    }

    private static void EnsurePalette()
    {
        if (_palette != null)
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

        _palette = new PaletteSet("Format Tables", new Guid("e2af6cf7-3c8c-4de1-bb36-4434726fa2f7"))
        {
            Style = PaletteSetStyles.ShowCloseButton | PaletteSetStyles.ShowPropertiesMenu | PaletteSetStyles.Snappable,
            DockEnabled = DockSides.Left | DockSides.Right | DockSides.Top | DockSides.Bottom,
            Visible = false
        };

        _palette.MinimumSize = new System.Drawing.Size(320, 240);
        _palette.AddVisual("Format Tables", _control);
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
