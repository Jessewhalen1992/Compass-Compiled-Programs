# Compass Compiled Programs

Compass is a consolidated AutoCAD palette set that surfaces multiple Compass tooling experiences behind a single `Compass` command. The initial implementation introduces the Drill Manager (Drill-Namer-2025) and the 3D Profile Manager (Profile-Xing-Gen) as dockable palettes that behave like standard AutoCAD property windows.

## Solution Layout

```
Compass.sln            # Visual Studio solution targeting .NET Framework 4.8
src/Compass/           # AutoCAD plug-in entry point and WPF UI
lib/AutoCAD2022/       # Place Autodesk AutoCAD 2022 managed DLL dependencies here
```

The `Compass` project is a class library that references the AutoCAD 2022 managed assemblies. The build output should be loaded into AutoCAD MAP3D 2015/2025 via `NETLOAD`.

## AutoCAD Dependencies

Copy the following files from an AutoCAD 2022 installation into `lib/AutoCAD2022` before building the solution:

- `AcDbMgd.dll`
- `AcMgd.dll`
- `AcCoreMgd.dll`
- `AdWindows.dll`

Set their **Copy Local** property to `False` to ensure the plug-in binds to the host AutoCAD installation.

## Commands

Once the assembly has been NETLOADed into AutoCAD, run the `Compass` command to display the main launcher palette. Each button opens its respective module in a dockable palette:

- **Drill Manager** – Hosts the Drill-Namer-2025 workflow with support for up to 20 drills.
- **3D Profile Manager** – Placeholder for the Profile-Xing-Gen workflow integration.

## Extending Modules

Modules implement the `ICompassModule` interface found in `src/Compass/Modules/ICompassModule.cs`. Additional programs can be registered from `CompassApplication.EnsureModules` to expose future tooling through the launcher UI.
