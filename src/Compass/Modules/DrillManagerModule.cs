using System;
using System.Drawing;
using Autodesk.AutoCAD.Windows;
using Compass.UI;

namespace Compass.Modules;

public class DrillManagerModule : ICompassModule
{
    private PaletteSet? _paletteSet;
    private DrillManagerControl? _control;

    public string Id => "drill-manager";
    public string DisplayName => "Drill Manager";
    public string Description => "Manage drill definitions with support for up to 20 drills.";

    public void Show()
    {
        if (_paletteSet == null)
        {
            _control = new DrillManagerControl();
            _paletteSet = CreatePalette();
            _paletteSet.AddVisual("Drill Manager", _control);
        }

        _paletteSet.Visible = true;
        _paletteSet.Activate(0);
    }

    private PaletteSet CreatePalette()
    {
        var palette = new PaletteSet(DisplayName, new Guid("879b5f68-8aa6-4b67-86f0-744f30c58f7b"))
        {
            Style = PaletteSetStyles.ShowCloseButton | PaletteSetStyles.ShowPropertiesMenu | PaletteSetStyles.Snappable,
            DockEnabled = DockSides.Left | DockSides.Right | DockSides.Top | DockSides.Bottom,
            MinimumSize = new Size(360, 480)
        };

        return palette;
    }
}
