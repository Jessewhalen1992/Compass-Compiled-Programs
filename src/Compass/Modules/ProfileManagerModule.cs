using System;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;

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

            try
            {
                SystemObjects.DynamicLinker.LoadModule(assemblyPath, false, false);
            }
            catch (FileLoadException)
            {
                // The assembly has already been loaded in this AutoCAD session.
            }
            catch (Autodesk.AutoCAD.Runtime.Exception)
            {
                // Ignore AutoCAD runtime exceptions triggered by repeated loading attempts.
            }
            document.SendStringToExecute("profilemanager\n", true, false, false);
        }
        catch (System.Exception ex)
        {
            Application.ShowAlertDialog($"Failed to launch 3D Profile Manager: {ex.Message}");
        }
    }
}
