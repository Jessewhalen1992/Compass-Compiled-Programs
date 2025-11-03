# AutoCAD 2022 References

Place the Autodesk AutoCAD 2022 managed assemblies required for compilation in this directory.

Recommended files:
- AcDbMgd.dll
- AcMgd.dll
- AdWindows.dll
- AcCoreMgd.dll

Ensure the file versions match the target AutoCAD installation (2022) and set their "Copy Local" property to `false` in Visual Studio so they are resolved from the host AutoCAD installation at runtime.
