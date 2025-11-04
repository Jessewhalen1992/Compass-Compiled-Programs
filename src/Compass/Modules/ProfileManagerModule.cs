using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices;

namespace Compass.Modules
{
    /// <summary>
    /// Launches the Profile‑Xing‑Gen tooling from the Compass palette.
    /// This implementation programmatically loads the managed ProfileCrossings.dll
    /// (replicating NETLOAD) and then posts the "profilemanager" command to AutoCAD.
    /// </summary>
    public class ProfileManagerModule : ICompassModule
    {
        // Paths to the ProfileCrossings.dll – adjust to match your installation.
        private static readonly string[] CandidatePaths =
        {
            @"C:\AUTOCAD-SETUP CG\CG_LISP\COMPASS\PROFILE PROGRAM\ProfileCrossings.dll",
            @"C:\AUTOCAD-SETUP\Lisp_2000\COMPASS\PROFILE PROGRAM\ProfileCrossings.dll"
        };

        // The command exported by ProfileCrossings.dll.  Lower-case matches what you type.
        private const string ProfileCommand = "profilemanager";

        public string Id => "profile-manager";
        public string DisplayName => "3D Profile Manager";
        public string Description => "Launch the Profile‑Xing‑Gen tooling.";

        public void Show()
        {
            try
            {
                // Defer to Idle to avoid context/state issues.
                void OnIdle(object? sender, EventArgs e)
                {
                    Application.Idle -= OnIdle;
                    Launch();
                }
                Application.Idle += OnIdle;
            }
            catch (System.Exception ex)
            {
                Application.ShowAlertDialog($"Failed to launch 3D Profile Manager: {ex.Message}");
            }
        }

        private static void Launch()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                Application.ShowAlertDialog("Open a drawing first.");
                return;
            }

            // Find the DLL in one of the known locations.
            var dllPath = CandidatePaths.FirstOrDefault(File.Exists);
            if (dllPath == null)
            {
                Application.ShowAlertDialog(
                    "ProfileCrossings.dll was not found in the expected locations.");
                return;
            }

            // Add the folder to TRUSTEDPATHS if SECURELOAD is on.
            TrustPathIfNeeded(Path.GetDirectoryName(dllPath)!);

            // Load the managed plug‑in; this is the equivalent of NETLOAD for .NET assemblies.
            try
            {
                Assembly.LoadFrom(dllPath);
            }
            catch (Exception ex)
            {
                Application.ShowAlertDialog($"Could not load ProfileCrossings.dll:\n{ex.Message}");
                return;
            }

            // Cancel any running command (ESC ESC), then invoke the plug‑in’s command.
            doc.SendStringToExecute("\u001B\u001B", true, false, false);
            doc.SendStringToExecute($"{ProfileCommand}\n", true, false, false);
        }

        /// <summary>
        /// Adds a folder to TRUSTEDPATHS if SECURELOAD is enabled and the folder isn’t already present.
        /// Uses system variables so no COM/dynamic code is required.
        /// </summary>
        private static void TrustPathIfNeeded(string folder)
        {
            if (string.IsNullOrWhiteSpace(folder))
                return;

            short secureLoad = 0;
            try
            {
                object val = Application.GetSystemVariable("SECURELOAD");
                secureLoad = Convert.ToInt16(val);
            }
            catch { /* ignore */ }

            if (secureLoad == 0) return; // No restrictions; nothing to do.

            try
            {
                var current = Convert.ToString(
                    Application.GetSystemVariable("TRUSTEDPATHS")) ?? string.Empty;
                var normalized = folder.EndsWith("\\") ? folder : folder + "\\";
                var parts = current.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                                   .Select(p => p.Trim());
                if (!parts.Contains(normalized, StringComparer.OrdinalIgnoreCase))
                {
                    var updated = string.IsNullOrEmpty(current)
                        ? normalized
                        : $"{current};{normalized}";
                    Application.SetSystemVariable("TRUSTEDPATHS", updated);
                }
            }
            catch
            {
                // Non‑fatal; if it fails, the user can add the path manually.
            }
        }
    }
}
