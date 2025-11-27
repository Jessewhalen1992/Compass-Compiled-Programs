using System;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
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

    [CommandMethod("Clegal", CommandFlags.Modal | CommandFlags.Session)]
    public static void ShowCompassLegal()
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
            TitleText = "Compass Legal",
            SubtitleText = "Select a legal tool to launch."
        };

        _control.ModuleRequested += OnToolRequested;
        _control.LoadModules(LegalTools.Select((tool, index) =>
            new CompassModuleDefinition(tool.Id, tool.DisplayName, tool.Description, index)));

        _palette = new PaletteSet("Compass Legal", new Guid("5b8e35b2-66cd-4878-b328-6e275c35ff6d"))
        {
            Style = PaletteSetStyles.ShowCloseButton | PaletteSetStyles.ShowPropertiesMenu | PaletteSetStyles.Snappable,
            DockEnabled = DockSides.Left | DockSides.Right | DockSides.Top | DockSides.Bottom,
            Visible = false
        };

        _palette.MinimumSize = new System.Drawing.Size(320, 240);
        _palette.AddVisual("Legal", _control);
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
