using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Compass.Modules;
using Compass.UI;
using Compass.ViewModels;
using Microsoft.Win32;

[assembly: CommandClass(typeof(Compass.Programs.DrillProps))]

namespace Compass.Programs;

public static class DrillProps
{
    private static DrillManagerModule Module => CompassApplication.GetDrillManagerModule();

    public static DrillManagerControl Control => Module.Control;

    public static DrillManagerViewModel ViewModel => Control.ViewModel;

    public static DrillPropsAccessor Accessor => ViewModel.DrillProps;

    public static int Minimum => Control.Minimum;

    public static int Maximum => Control.Maximum;

    public static int DrillCount
    {
        get => Control.DrillCount;
        set => Control.DrillCount = value;
    }

    public static IReadOnlyList<string> GetDrillProps()
    {
        return Control.GetDrillProps();
    }

    public static IReadOnlyDictionary<int, string> GetIndexedDrillProps()
    {
        return Control.GetIndexedDrillProps();
    }

    public static string GetDrillProp(int index)
    {
        return Control.GetDrillProp(index);
    }

    public static void SetDrillProp(int index, string? name)
    {
        Control.SetDrillProp(index, name);
    }

    public static void SetDrillProps(IEnumerable<string?>? names)
    {
        Control.SetDrillProps(names);
    }

    public static void ClearDrillProp(int index)
    {
        Control.ClearDrillProp(index);
    }

    public static void ClearAllDrillProps()
    {
        Control.ClearAllDrillProps();
    }

    public static void EnsureCapacity(int desiredCount)
    {
        Control.EnsureCapacity(desiredCount);
    }

    public static void LoadFromDelimitedList(string? delimitedNames, char separator = ',')
    {
        Control.LoadFromDelimitedList(delimitedNames, separator);
    }

    public static string ToDelimitedList(char separator = ',')
    {
        return Control.ToDelimitedList(separator);
    }

    public static void FillEmptyWith(Func<int, string> nameFactory)
    {
        Control.FillEmptyWith(nameFactory);
    }

    public static void Apply(Func<int, string, string?> mutator)
    {
        Control.Apply(mutator);
    }

    public static void SetAllDrillProps(string? name)
    {
        Control.SetAllDrillProps(name);
    }

    public static void ApplyCornerTemplate(string topLeft, string topRight, string bottomRight, string bottomLeft)
    {
        Control.ApplyCornerTemplate(topLeft, topRight, bottomRight, bottomLeft);
    }

    public static void AutoFillEmpty(string prefix, int startIndex)
    {
        Control.AutoFillEmpty(prefix, startIndex);
    }

    public static int LoadFromJson(string path)
    {
        return Control.LoadFromJson(path);
    }

    public static void SaveToJson(string path)
    {
        Control.SaveToJson(path);
    }

    [CommandMethod("DRILLPROPS", CommandFlags.Modal | CommandFlags.Session)]
    public static void ShowPalette()
    {
        Module.Show();
    }

    public static string DrillProp1
    {
        get => Control.DrillProp1;
        set => Control.DrillProp1 = value;
    }

    public static string DrillProp2
    {
        get => Control.DrillProp2;
        set => Control.DrillProp2 = value;
    }

    public static string DrillProp3
    {
        get => Control.DrillProp3;
        set => Control.DrillProp3 = value;
    }

    public static string DrillProp4
    {
        get => Control.DrillProp4;
        set => Control.DrillProp4 = value;
    }

    public static string DrillProp5
    {
        get => Control.DrillProp5;
        set => Control.DrillProp5 = value;
    }

    public static string DrillProp6
    {
        get => Control.DrillProp6;
        set => Control.DrillProp6 = value;
    }

    public static string DrillProp7
    {
        get => Control.DrillProp7;
        set => Control.DrillProp7 = value;
    }

    public static string DrillProp8
    {
        get => Control.DrillProp8;
        set => Control.DrillProp8 = value;
    }

    public static string DrillProp9
    {
        get => Control.DrillProp9;
        set => Control.DrillProp9 = value;
    }

    public static string DrillProp10
    {
        get => Control.DrillProp10;
        set => Control.DrillProp10 = value;
    }

    public static string DrillProp11
    {
        get => Control.DrillProp11;
        set => Control.DrillProp11 = value;
    }

    public static string DrillProp12
    {
        get => Control.DrillProp12;
        set => Control.DrillProp12 = value;
    }

    public static string DrillProp13
    {
        get => Control.DrillProp13;
        set => Control.DrillProp13 = value;
    }

    public static string DrillProp14
    {
        get => Control.DrillProp14;
        set => Control.DrillProp14 = value;
    }

    public static string DrillProp15
    {
        get => Control.DrillProp15;
        set => Control.DrillProp15 = value;
    }

    public static string DrillProp16
    {
        get => Control.DrillProp16;
        set => Control.DrillProp16 = value;
    }

    public static string DrillProp17
    {
        get => Control.DrillProp17;
        set => Control.DrillProp17 = value;
    }

    public static string DrillProp18
    {
        get => Control.DrillProp18;
        set => Control.DrillProp18 = value;
    }

    public static string DrillProp19
    {
        get => Control.DrillProp19;
        set => Control.DrillProp19 = value;
    }

    public static string DrillProp20
    {
        get => Control.DrillProp20;
        set => Control.DrillProp20 = value;
    }

    public static string GetDrillProp1()
    {
        return Control.GetDrillProp1();
    }

    public static void SetDrillProp1(string? value)
    {
        Control.SetDrillProp1(value);
    }

    public static string GetDrillProp2()
    {
        return Control.GetDrillProp2();
    }

    public static void SetDrillProp2(string? value)
    {
        Control.SetDrillProp2(value);
    }

    public static string GetDrillProp3()
    {
        return Control.GetDrillProp3();
    }

    public static void SetDrillProp3(string? value)
    {
        Control.SetDrillProp3(value);
    }

    public static string GetDrillProp4()
    {
        return Control.GetDrillProp4();
    }

    public static void SetDrillProp4(string? value)
    {
        Control.SetDrillProp4(value);
    }

    public static string GetDrillProp5()
    {
        return Control.GetDrillProp5();
    }

    public static void SetDrillProp5(string? value)
    {
        Control.SetDrillProp5(value);
    }

    public static string GetDrillProp6()
    {
        return Control.GetDrillProp6();
    }

    public static void SetDrillProp6(string? value)
    {
        Control.SetDrillProp6(value);
    }

    public static string GetDrillProp7()
    {
        return Control.GetDrillProp7();
    }

    public static void SetDrillProp7(string? value)
    {
        Control.SetDrillProp7(value);
    }

    public static string GetDrillProp8()
    {
        return Control.GetDrillProp8();
    }

    public static void SetDrillProp8(string? value)
    {
        Control.SetDrillProp8(value);
    }

    public static string GetDrillProp9()
    {
        return Control.GetDrillProp9();
    }

    public static void SetDrillProp9(string? value)
    {
        Control.SetDrillProp9(value);
    }

    public static string GetDrillProp10()
    {
        return Control.GetDrillProp10();
    }

    public static void SetDrillProp10(string? value)
    {
        Control.SetDrillProp10(value);
    }

    public static string GetDrillProp11()
    {
        return Control.GetDrillProp11();
    }

    public static void SetDrillProp11(string? value)
    {
        Control.SetDrillProp11(value);
    }

    public static string GetDrillProp12()
    {
        return Control.GetDrillProp12();
    }

    public static void SetDrillProp12(string? value)
    {
        Control.SetDrillProp12(value);
    }

    public static string GetDrillProp13()
    {
        return Control.GetDrillProp13();
    }

    public static void SetDrillProp13(string? value)
    {
        Control.SetDrillProp13(value);
    }

    public static string GetDrillProp14()
    {
        return Control.GetDrillProp14();
    }

    public static void SetDrillProp14(string? value)
    {
        Control.SetDrillProp14(value);
    }

    public static string GetDrillProp15()
    {
        return Control.GetDrillProp15();
    }

    public static void SetDrillProp15(string? value)
    {
        Control.SetDrillProp15(value);
    }

    public static string GetDrillProp16()
    {
        return Control.GetDrillProp16();
    }

    public static void SetDrillProp16(string? value)
    {
        Control.SetDrillProp16(value);
    }

    public static string GetDrillProp17()
    {
        return Control.GetDrillProp17();
    }

    public static void SetDrillProp17(string? value)
    {
        Control.SetDrillProp17(value);
    }

    public static string GetDrillProp18()
    {
        return Control.GetDrillProp18();
    }

    public static void SetDrillProp18(string? value)
    {
        Control.SetDrillProp18(value);
    }

    public static string GetDrillProp19()
    {
        return Control.GetDrillProp19();
    }

    public static void SetDrillProp19(string? value)
    {
        Control.SetDrillProp19(value);
    }

    public static string GetDrillProp20()
    {
        return Control.GetDrillProp20();
    }

    public static void SetDrillProp20(string? value)
    {
        Control.SetDrillProp20(value);
    }

    [CommandMethod("DRILLPROPSSET", CommandFlags.Modal | CommandFlags.Session)]
    public static void CommandSetDrill()
    {
        var editor = TryGetEditor();
        if (editor == null)
        {
            return;
        }

        var index = PromptForIndex(editor, "\nEnter drill slot number");
        if (!index.HasValue)
        {
            return;
        }

        var name = PromptForName(editor, "\nEnter drill name", GetDrillProp(index.Value));
        if (name == null)
        {
            return;
        }

        SetDrillProp(index.Value, name);
        ViewModel.SelectedDrillIndex = index.Value;
        ViewModel.SelectedDrillName = GetDrillProp(index.Value);
        editor.WriteMessage($"\nDrill {index.Value} set to \"{GetDrillProp(index.Value)}\".");
    }

    [CommandMethod("DRILLPROPSGET", CommandFlags.Modal | CommandFlags.Session)]
    public static void CommandGetDrill()
    {
        var editor = TryGetEditor();
        if (editor == null)
        {
            return;
        }

        var index = PromptForIndex(editor, "\nEnter drill slot number");
        if (!index.HasValue)
        {
            return;
        }

        editor.WriteMessage($"\nDrill {index.Value}: {GetDrillProp(index.Value)}");
    }

    [CommandMethod("DRILLPROPSSETALL", CommandFlags.Modal | CommandFlags.Session)]
    public static void CommandSetAll()
    {
        var editor = TryGetEditor();
        if (editor == null)
        {
            return;
        }

        var name = PromptForName(editor, "\nEnter name to apply to all drills", ViewModel.SetAllName);
        if (name == null)
        {
            return;
        }

        ViewModel.SetAllName = name;
        ViewModel.SetAllDrills();
        editor.WriteMessage("\nAll drill names updated.");
    }

    [CommandMethod("DRILLPROPSCLEAR", CommandFlags.Modal | CommandFlags.Session)]
    public static void CommandClearDrill()
    {
        var editor = TryGetEditor();
        if (editor == null)
        {
            return;
        }

        var index = PromptForIndex(editor, "\nEnter drill slot number to clear");
        if (!index.HasValue)
        {
            return;
        }

        ClearDrillProp(index.Value);
        ViewModel.SelectedDrillIndex = index.Value;
        ViewModel.SelectedDrillName = GetDrillProp(index.Value);
        editor.WriteMessage($"\nCleared drill {index.Value}.");
    }

    [CommandMethod("DRILLPROPSCLEARALL", CommandFlags.Modal | CommandFlags.Session)]
    public static void CommandClearAll()
    {
        var editor = TryGetEditor();
        if (editor == null)
        {
            return;
        }

        var options = new PromptKeywordOptions("\nClear all drill names? [Yes/No]", "Yes No")
        {
            DefaultKeyword = "No"
        };

        var result = editor.GetKeywords(options);
        if (result.Status != PromptStatus.OK || !string.Equals(result.StringResult, "Yes", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        ClearAllDrillProps();
        editor.WriteMessage("\nAll drill names cleared.");
    }

    [CommandMethod("DRILLPROPSCORNERS", CommandFlags.Modal | CommandFlags.Session)]
    public static void CommandApplyCorners()
    {
        var editor = TryGetEditor();
        if (editor == null)
        {
            return;
        }

        var topLeft = PromptForName(editor, "\nTop left corner name", ViewModel.CornerTopLeft);
        if (topLeft == null)
        {
            return;
        }

        var topRight = PromptForName(editor, "\nTop right corner name", ViewModel.CornerTopRight);
        if (topRight == null)
        {
            return;
        }

        var bottomRight = PromptForName(editor, "\nBottom right corner name", ViewModel.CornerBottomRight);
        if (bottomRight == null)
        {
            return;
        }

        var bottomLeft = PromptForName(editor, "\nBottom left corner name", ViewModel.CornerBottomLeft);
        if (bottomLeft == null)
        {
            return;
        }

        ApplyCornerTemplate(topLeft, topRight, bottomRight, bottomLeft);
        editor.WriteMessage("\nCorner template applied.");
    }

    [CommandMethod("DRILLPROPSAUTO", CommandFlags.Modal | CommandFlags.Session)]
    public static void CommandAutoFill()
    {
        var editor = TryGetEditor();
        if (editor == null)
        {
            return;
        }

        var prefix = PromptForName(editor, "\nPrefix for empty drills", ViewModel.AutoFillPrefix);
        if (prefix == null)
        {
            return;
        }

        var startIndex = PromptForStartIndex(editor, "\nStarting number", ViewModel.AutoFillStartIndex);
        if (!startIndex.HasValue)
        {
            return;
        }

        AutoFillEmpty(prefix, startIndex.Value);
        editor.WriteMessage("\nEmpty drills filled.");
    }

    [CommandMethod("DRILLPROPSJSONLOAD", CommandFlags.Modal | CommandFlags.Session)]
    public static void CommandLoadJson()
    {
        var editor = TryGetEditor();
        if (editor == null)
        {
            return;
        }

        var fileName = PromptForJsonFileToOpen();
        if (string.IsNullOrEmpty(fileName))
        {
            return;
        }

        try
        {
            var count = LoadFromJson(fileName);
            editor.WriteMessage($"\nLoaded {count} drill names from '{Path.GetFileName(fileName)}'.");
        }
        catch (Exception ex)
        {
            editor.WriteMessage($"\nFailed to load drill JSON: {ex.Message}");
        }
    }

    [CommandMethod("DRILLPROPSJSONSAVE", CommandFlags.Modal | CommandFlags.Session)]
    public static void CommandSaveJson()
    {
        var editor = TryGetEditor();
        if (editor == null)
        {
            return;
        }

        var fileName = PromptForJsonFileToSave();
        if (string.IsNullOrEmpty(fileName))
        {
            return;
        }

        try
        {
            SaveToJson(fileName);
            editor.WriteMessage($"\nSaved drill names to '{Path.GetFileName(fileName)}'.");
        }
        catch (Exception ex)
        {
            editor.WriteMessage($"\nFailed to save drill JSON: {ex.Message}");
        }
    }

    private static Editor? TryGetEditor()
    {
        return Application.DocumentManager.MdiActiveDocument?.Editor;
    }

    private static int? PromptForIndex(Editor editor, string message)
    {
        var options = new PromptIntegerOptions(message)
        {
            AllowNegative = false,
            AllowZero = false,
            LowerLimit = Minimum,
            UpperLimit = Maximum,
            DefaultValue = Math.Max(Minimum, Math.Min(Maximum, ViewModel.SelectedDrillIndex))
        };

        var result = editor.GetInteger(options);
        return result.Status == PromptStatus.OK ? result.Value : null;
    }

    private static int? PromptForStartIndex(Editor editor, string message, int defaultValue)
    {
        var options = new PromptIntegerOptions(message)
        {
            AllowNegative = false,
            AllowZero = false,
            LowerLimit = 1,
            DefaultValue = Math.Max(1, defaultValue)
        };

        var result = editor.GetInteger(options);
        return result.Status == PromptStatus.OK ? result.Value : null;
    }

    private static string? PromptForName(Editor editor, string message, string? defaultValue)
    {
        var options = new PromptStringOptions(message)
        {
            AllowSpaces = true,
            DefaultValue = defaultValue ?? string.Empty
        };

        var result = editor.GetString(options);
        return result.Status == PromptStatus.OK ? result.StringResult : null;
    }

    private static string? PromptForJsonFileToOpen()
    {
        var dialog = new OpenFileDialog
        {
            Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*"
        };

        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }

    private static string? PromptForJsonFileToSave()
    {
        var dialog = new SaveFileDialog
        {
            Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
            DefaultExt = ".json"
        };

        return dialog.ShowDialog() == true ? dialog.FileName : null;
    }
}
