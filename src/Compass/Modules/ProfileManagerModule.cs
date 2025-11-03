using System;
using System.Drawing;
using Autodesk.AutoCAD.Windows;
using Compass.UI;

namespace Compass.Modules;

public class ProfileManagerModule : ICompassModule
{
    private PaletteSet? _paletteSet;
    private ProfileManagerControl? _control;

    public string Id => "profile-manager";
    public string DisplayName => "3D Profile Manager";
    public string Description => "Launch the Profile-Xing-Gen tooling.";

    public void Show()
    {
        if (_paletteSet == null)
        {
            _control = new ProfileManagerControl();
            _paletteSet = CreatePalette();
            _paletteSet.AddVisual("3D Profile Manager", _control);
        }

        _paletteSet.Visible = true;
        _paletteSet.Activate(0);
    }

    private PaletteSet CreatePalette()
    {
        var palette = new PaletteSet(DisplayName, new Guid("1c1df9d7-0a89-43b0-b52c-24b7955aa78f"))
        {
            Style = PaletteSetStyles.ShowCloseButton | PaletteSetStyles.ShowPropertiesMenu | PaletteSetStyles.Snappable,
            DockEnabled = DockSides.Left | DockSides.Right | DockSides.Top | DockSides.Bottom,
            MinimumSize = new Size(360, 320)
        };

        return palette;
    }
}
