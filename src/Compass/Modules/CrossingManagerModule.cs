using System.Collections.Generic;

namespace Compass.Modules;

/// <summary>
/// Launches the Crossing Manager tooling by loading XingManager.dll and invoking its command.
/// </summary>
public class CrossingManagerModule : ManagedPluginModuleBase
{
    private static readonly string[] CandidatePaths =
    {
        @"C:\AUTOCAD-SETUP CG\CG_LISP\XING MANAGER\XingManager.DLL"
    };

    public override string Id => "crossing-manager";
    public override string DisplayName => "Crossing Manager";
    public override string Description => "Launch the Xing Manager program.";

    protected override IReadOnlyList<string> CandidateDllPaths => CandidatePaths;
    protected override string CommandName => "crossingmanager";
}
