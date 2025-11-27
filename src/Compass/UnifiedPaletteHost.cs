using System;
using System.Collections.Generic;
using System.Windows.Controls;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;

[assembly: CommandClass(typeof(Compass.UnifiedPaletteHost))]

namespace Compass
{
    public class UnifiedPaletteHost : IExtensionApplication
    {
        private static readonly Dictionary<string, int> TabIndices = new(StringComparer.OrdinalIgnoreCase);
        private static PaletteSet? _mainPalette;

        public void Initialize() => EnsurePalette();
        public void Terminate()
        {
            if (_mainPalette != null)
            {
                _mainPalette.Visible = false;
                _mainPalette.Dispose();
                _mainPalette = null;
            }
            TabIndices.Clear();
        }

        public static void EnsurePalette()
        {
            if (_mainPalette != null) return;

            _mainPalette = new PaletteSet("Compass", new Guid("c833fa27-9db1-4d67-85d4-45115ac0a2c2"))
            {
                Style = PaletteSetStyles.ShowCloseButton | PaletteSetStyles.ShowPropertiesMenu | PaletteSetStyles.Snappable,
                DockEnabled = DockSides.Left | DockSides.Right | DockSides.Top | DockSides.Bottom,
                Visible = false,
                MinimumSize = new System.Drawing.Size(320, 240)
            };
        }

        public static bool AddTab(string name, UserControl control)
        {
            EnsurePalette();
            if (_mainPalette == null || TabIndices.ContainsKey(name)) return false;

            _mainPalette.AddVisual(name, control);
            // Use Count instead of PaletteCount; Count is the number of palettes in the set
            TabIndices[name] = _mainPalette.Count - 1;
            return true;
        }

        public static void ShowPalette(string? tabName = null)
        {
            EnsurePalette();
            if (_mainPalette == null) return;

            _mainPalette.Visible = true;
            if (!string.IsNullOrWhiteSpace(tabName) && TabIndices.TryGetValue(tabName, out var index))
            {
                _mainPalette.Activate(index);
            }
            else if (_mainPalette.Count > 0)
            {
                _mainPalette.Activate(0);
            }
        }

        [CommandMethod("CompassPalette", CommandFlags.Modal | CommandFlags.Session)]
        public static void ShowUnifiedPalette() => ShowPalette();
    }
}
