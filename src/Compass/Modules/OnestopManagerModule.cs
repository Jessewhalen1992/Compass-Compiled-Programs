using System.Collections.Generic;

namespace Compass.Modules;

/// <summary>
/// Launches the OneStop Manager tooling by loading the appropriate WellShapeProgram DLL and invoking its command.
/// </summary>
public class OnestopManagerModule : ManagedPluginModuleBase
{
    private static readonly string[] CandidatePaths =
    {
        @"C:\AUTOCAD-SETUP CG\CG_LISP\WELL LICENSE FINALS\WellShapeProgram_Modern.DLL",
        @"C:\AUTOCAD-SETUP CG\CG_LISP\WELL LICENSE FINALS (2015)\WellShapeProgram_Legacy.DLL",
        @"C:\AUTOCAD-SETUP\Lisp_2000\COMPASS\WELL LICENSE FINALS (2015)\WellShapeProgram_Legacy.DLL"
    };

    public override string Id => "onestop-manager";
    public override string DisplayName => "OneStop Manager";
    public override string Description => "Launch the Well License OneStop workflow.";

    protected override IReadOnlyList<string> CandidateDllPaths => CandidatePaths;
    protected override string CommandName => "onestopmanager";
}
