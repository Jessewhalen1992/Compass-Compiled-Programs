using System;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;

namespace Compass.Modules;

public class ProfileManagerModule : ICompassModule
{
    private static readonly string[] AssemblySearchPaths =
    {
        @"C:\\AUTOCAD-SETUP CG\\CG_LISP\\COMPASS\\PROFILE PROGRAM\\ProfileCrossings.dll",
        @"C:\\AUTOCAD-SETUP\\Lisp_2000\\COMPASS\\PROFILE PROGRAM\\ProfileCrossings.dll"
    };

    public string Id => "profile-manager";
    public string DisplayName => "3D Profile Manager";
    public string Description => "Launch the Profile-Xing-Gen tooling.";

    public void Show()
    {
        try
        {
            var document = Application.DocumentManager.MdiActiveDocument;
            if (document == null)
            {
                Application.ShowAlertDialog("No active AutoCAD document.");
                return;
            }

            string? assemblyPath = null;
            foreach (var path in AssemblySearchPaths)
            {
                if (File.Exists(path))
                {
                    assemblyPath = path;
                    break;
                }
            }

            if (string.IsNullOrEmpty(assemblyPath))
            {
                Application.ShowAlertDialog("ProfileCrossings.dll was not found in the expected locations.");
                return;
            }

            var netloadCommand = $@"_.NETLOAD ""{assemblyPath}"" ";
            document.SendStringToExecute(netloadCommand, true, false, false);
            document.SendStringToExecute("profilemanager ", true, false, false);
        }
        catch (Exception ex)
        {
            Application.ShowAlertDialog($"Failed to launch 3D Profile Manager: {ex.Message}");
        }
    }
}
