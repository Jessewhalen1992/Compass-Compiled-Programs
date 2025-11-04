using System.Collections.Generic;

namespace Compass.Modules;

/// <summary>
/// Launches the Hybrid Manager tooling by loading HybridProgram_One.dll and invoking its command.
/// </summary>
public class HybridManagerModule : ManagedPluginModuleBase
{
    private static readonly string[] CandidatePaths =
    {
        @"C:\AUTOCAD-SETUP CG\CG_LISP\HYBRID PROGRAM\HybridProgram_One.DLL"
    };

    public override string Id => "hybrid-manager";
    public override string DisplayName => "Hybrid Manager";
    public override string Description => "Launch the Hybrid Program tools.";

    protected override IReadOnlyList<string> CandidateDllPaths => CandidatePaths;
    protected override string CommandName => "hybridmanager";
}
