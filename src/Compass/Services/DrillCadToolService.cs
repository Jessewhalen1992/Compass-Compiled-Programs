using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Compass.Infrastructure.Logging;
using Compass.Models;
using OfficeOpenXml;
using Microsoft.Win32;

using AutoCADApplication = Autodesk.AutoCAD.ApplicationServices.Application;

namespace Compass.Services;

public class DrillCadToolService
{
    private const string DrillPointsLayer = "Z-DRILL-POINT";
    private const string HeadingBlockName = "DRILL HEADING";
    private const string DrillBlockName = "DRILL";
    private const string OffsetsLayer = "P-Drill-Offset";
    private const string HorizontalLayer = "L-SEC-HB";
    private const string CordsDirectory = @"C:\\CORDS";
    private static readonly string[] CordsExecutableSearchPaths =
    {
        @"C:\\AUTOCAD-SETUP CG\\CG_LISP\\COMPASS\\cords.exe",
        @"C:\\AUTOCAD-SETUP\\Lisp_2000\\DRILL PROPERTIES\\cords.exe"
    };

    private readonly ILog _log;
    private readonly LayerService _layerService;
    private readonly NaturalStringComparer _naturalComparer = new();

    public DrillCadToolService(ILog log, LayerService layerService)
    {
        _log = log ?? throw new ArgumentNullException(nameof(log));
        _layerService = layerService ?? throw new ArgumentNullException(nameof(layerService));
    }

    public DrillCheckSummary Check(IReadOnlyList<string> drillNames)
    {
        var document = AutoCADApplication.DocumentManager.MdiActiveDocument;
        if (document == null)
        {
            MessageBox.Show("No active AutoCAD document is available.", "Check", MessageBoxButton.OK, MessageBoxImage.Information);
            return new DrillCheckSummary(completed: false, Array.Empty<DrillCheckResult>(), reportPath: null);
        }

        var editor = document.Editor;
        var database = document.Database;

        try
        {
            var tableOptions = new PromptEntityOptions("\nSelect the data-linked table:")
            {
                AllowObjectOnLockedLayer = true
            };
            tableOptions.SetRejectMessage("\nOnly table entities are allowed.");
            tableOptions.AddAllowedClass(typeof(Table), exactMatch: true);

            var tableResult = editor.GetEntity(tableOptions);
            if (tableResult.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nTable selection cancelled.");
                return new DrillCheckSummary(completed: false, Array.Empty<DrillCheckResult>(), reportPath: null);
            }

            List<string> tableValues;
            using (var transaction = database.TransactionManager.StartTransaction())
            {
                if (transaction.GetObject(tableResult.ObjectId, OpenMode.ForRead) is not Table table)
                {
                    MessageBox.Show("Selected entity is not a valid table.", "Check", MessageBoxButton.OK, MessageBoxImage.Error);
                    return new DrillCheckSummary(completed: false, Array.Empty<DrillCheckResult>(), reportPath: null);
                }

                tableValues = ExtractBottomHoleValues(table);
                transaction.Commit();
            }

            if (tableValues.Count == 0)
            {
                return new DrillCheckSummary(completed: false, Array.Empty<DrillCheckResult>(), reportPath: null);
            }

            var blockData = SelectBlocksWithDrillName(document);
            if (blockData == null || blockData.Count == 0)
            {
                MessageBox.Show("No blocks selected or no blocks with DRILLNAME attribute found.", "Check", MessageBoxButton.OK, MessageBoxImage.Warning);
                return new DrillCheckSummary(completed: false, Array.Empty<DrillCheckResult>(), reportPath: null);
            }

            return CompareDrillNamesWithTable(drillNames, tableValues, blockData);
        }
        catch (System.Exception ex)
        {
            _log.Error($"Error during check operation: {ex.Message}", ex);
            MessageBox.Show($"An error occurred during the check operation: {ex.Message}", "Check", MessageBoxButton.OK, MessageBoxImage.Error);
            return new DrillCheckSummary(completed: false, Array.Empty<DrillCheckResult>(), reportPath: null);
        }
    }

    public void HeadingsAll(IReadOnlyList<string> drillNames, string heading)
    {
        var confirm = MessageBox.Show(
            "Are you sure you want to insert heading blocks (and DRILL blocks) for all non-default drills?",
            "Confirm HEADING ALL",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes)
        {
            return;
        }

        var document = AutoCADApplication.DocumentManager.MdiActiveDocument;
        if (document == null)
        {
            MessageBox.Show("No active AutoCAD document is available.", "Headings All", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var editor = document.Editor;
        var database = document.Database;

        try
        {
            var headingLabel = string.Equals(heading, "HEEL", StringComparison.OrdinalIgnoreCase) ? "HEEL" : "ICP";
            var surfaceOptions = new PromptEntityOptions("\nSelect the data-linked table containing 'SURFACE':")
            {
                AllowObjectOnLockedLayer = true
            };
            surfaceOptions.SetRejectMessage("\nOnly table entities are allowed.");
            surfaceOptions.AddAllowedClass(typeof(Table), exactMatch: true);

            var surfaceResult = editor.GetEntity(surfaceOptions);
            if (surfaceResult.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nTable selection cancelled.");
                return;
            }

            List<(int Row, int Column)> surfaceCells;
            using (var transaction = database.TransactionManager.StartTransaction())
            {
                if (transaction.GetObject(surfaceResult.ObjectId, OpenMode.ForRead) is not Table surfaceTable)
                {
                    MessageBox.Show("Selected entity is not a valid table.", "Headings All", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                surfaceCells = FindSurfaceCells(surfaceTable);
                transaction.Commit();
            }

            if (surfaceCells.Count == 0)
            {
                MessageBox.Show("No 'SURFACE' cells found in the selected table.", "Headings All", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var nonDefault = GetNonDefaultDrills(drillNames);
            if (nonDefault.Count == 0)
            {
                MessageBox.Show("No non-default drills to insert blocks for.", "Headings All", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (nonDefault.Count > surfaceCells.Count)
            {
                MessageBox.Show("Not enough SURFACE cells for all non-default drills.", "Headings All", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            using (var transaction = AutoCADHelper.StartTransaction())
            {
                for (var i = 0; i < nonDefault.Count; i++)
                {
                    var drillName = nonDefault[i];
                    var (row, column) = surfaceCells[i];
                    if (transaction.GetObject(surfaceResult.ObjectId, OpenMode.ForRead) is not Table surfaceTable)
                    {
                        continue;
                    }

                    var nwCorner = GetCellNorthWest(surfaceTable, row, column);
                    var headingAttributes = new Dictionary<string, string> { { "DRILLNAME", drillName } };
                    var headingId = AutoCADHelper.InsertBlock(HeadingBlockName, nwCorner, headingAttributes, 1.0);
                    if (headingId == ObjectId.Null)
                    {
                        MessageBox.Show($"Failed to insert {HeadingBlockName} for {drillName}.", "Headings All", MessageBoxButton.OK, MessageBoxImage.Error);
                        continue;
                    }

                    var drillInsertion = new Point3d(nwCorner.X - 50.0, nwCorner.Y, nwCorner.Z);
                    var drillAttributes = new Dictionary<string, string> { { "DRILL", drillName } };
                    var drillId = AutoCADHelper.InsertBlock(DrillBlockName, drillInsertion, drillAttributes, 2.0);
                    if (drillId == ObjectId.Null)
                    {
                        MessageBox.Show($"Failed to insert DRILL block for {drillName}.", "Headings All", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }

                transaction.Commit();
            }

            MessageBox.Show($"Successfully created DRILL + {headingLabel} blocks for all non-default drills.", "Headings All", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (System.Exception ex)
        {
            _log.Error($"Exception in HeadingsAll: {ex.Message}", ex);
            MessageBox.Show($"An error occurred while inserting heading blocks: {ex.Message}", "Headings All", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    public void CreateXlsFromTable()
    {
        var document = AutoCADApplication.DocumentManager.MdiActiveDocument;
        if (document == null)
        {
            MessageBox.Show("No active AutoCAD document is available.", "Create XLS", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var editor = document.Editor;
        var database = document.Database;

        try
        {
            var options = new PromptEntityOptions("\nSelect the table to export to XLS:")
            {
                AllowObjectOnLockedLayer = true
            };
            options.SetRejectMessage("\nOnly table entities are allowed.");
            options.AddAllowedClass(typeof(Table), exactMatch: true);

            var result = editor.GetEntity(options);
            if (result.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nXLS save cancelled.");
                return;
            }

            Table table;
            using (var transaction = database.TransactionManager.StartTransaction())
            {
                table = transaction.GetObject(result.ObjectId, OpenMode.ForRead) as Table ?? throw new InvalidOperationException("Selected entity is not a table.");
                transaction.Commit();
            }

            var saveDialog = new SaveFileDialog
            {
                Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                FileName = "ExportedTable.xlsx",
                Title = "Save XLS File"
            };

            if (saveDialog.ShowDialog() != true)
            {
                editor.WriteMessage("\nXLS save cancelled.");
                return;
            }

            var rows = table.Rows.Count;
            var cols = table.Columns.Count;

            EpplusCompat.EnsureLicense();
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("ExportedTable");
                for (var row = 0; row < rows; row++)
                {
                    for (var col = 0; col < cols; col++)
                    {
                        var value = table.Cells[row, col].TextString.Trim();
                        worksheet.Cells[row + 1, col + 1].Value = value;
                    }
                }

                if (cols >= 3)
                {
                    worksheet.Column(1).Width = 15;
                    worksheet.Column(2).Width = 12;
                    worksheet.Column(3).Width = 12;
                    worksheet.Column(2).Style.Numberformat.Format = "0.00";
                    worksheet.Column(3).Style.Numberformat.Format = "0.00";
                }

                package.SaveAs(new FileInfo(saveDialog.FileName));
            }

            MessageBox.Show($"XLS file created successfully at:\n{saveDialog.FileName}", "Create XLS", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (System.Exception ex)
        {
            MessageBox.Show($"An error occurred while creating XLS:\n{ex.Message}", "Create XLS", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    public void CompleteCords(IReadOnlyList<string> drillNames, string heading)
    {
        var document = AutoCADApplication.DocumentManager.MdiActiveDocument;
        if (document == null)
        {
            ShowAlert("No active AutoCAD document.");
            return;
        }

        var confirm = MessageBox.Show(
            "ARE YOU IN UTM?",
            "Complete CORDS",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            EpplusCompat.EnsureLicense();

            var csvPath = DrillCsvPipeline(document.Database);
            if (string.IsNullOrEmpty(csvPath))
            {
                return;
            }

            var excelPath = RunCordsExecutable(csvPath, heading);
            if (string.IsNullOrEmpty(excelPath))
            {
                return;
            }

            var tableData = ReadExcel(excelPath);
            if (tableData == null)
            {
                return;
            }

            AdjustTableForClient(tableData, heading);

            if (!InsertTablePipeline(document, tableData))
            {
                return;
            }

            MessageBox.Show("Coordinate table created!", "Complete CORDS", MessageBoxButton.OK, MessageBoxImage.Information);
            _log.Info("COMPLETE CORDS succeeded.");

            HeadingsAll(drillNames, heading);
        }
        catch (System.Exception ex)
        {
            _log.Error($"Error in COMPLETE CORDS: {ex.Message}", ex);
            ShowAlert($"Error in COMPLETE CORDS: {ex.Message}");
        }
    }

    public void GetUtms()
    {
        var document = AutoCADApplication.DocumentManager.MdiActiveDocument;
        if (document == null)
        {
            ShowAlert("No active AutoCAD document.");
            return;
        }

        var confirm = MessageBox.Show(
            "ARE YOU IN UTM?",
            "Get UTMs",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            var database = document.Database;
            var points = ReadGridPoints(database);
            points = points.OrderBy(p => p.Label, _naturalComparer).ToList();

            Directory.CreateDirectory(CordsDirectory);
            var csvPath = Path.Combine(CordsDirectory, "CORDS.csv");
            using (var writer = new StreamWriter(csvPath, false))
            {
                writer.WriteLine("Label,Northing,Easting");
                foreach (var point in points)
                {
                    writer.WriteLine($"{point.Label},{point.Northing},{point.Easting}");
                }
            }

            MessageBox.Show($"UTM CSV created successfully at:\n{csvPath}", "Get UTMs", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (System.Exception ex)
        {
            _log.Error($"Error generating UTMs CSV: {ex.Message}", ex);
            MessageBox.Show($"Error generating UTMs CSV: {ex.Message}", "Get UTMs", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    public void AddDrillPoints()
    {
        var document = AutoCADApplication.DocumentManager.MdiActiveDocument;
        if (document == null)
        {
            ShowAlert("No active AutoCAD document.");
            return;
        }

        var editor = document.Editor;
        var promptOptions = new PromptStringOptions("\nEnter letter for drill points:")
        {
            AllowSpaces = false,
            DefaultValue = "A"
        };
        var promptResult = editor.GetString(promptOptions);
        if (promptResult.Status != PromptStatus.OK)
        {
            editor.WriteMessage("\nOperation cancelled.");
            return;
        }

        var letter = promptResult.StringResult.Trim();
        if (string.IsNullOrWhiteSpace(letter))
        {
            MessageBox.Show("Letter cannot be empty.", "Add Drill Points", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        var database = document.Database;

        var options = new PromptEntityOptions("\nSelect a polyline:");
        options.SetRejectMessage("\nOnly polylines are allowed.");
        options.AddAllowedClass(typeof(Polyline), false);

        var result = editor.GetEntity(options);
        if (result.Status != PromptStatus.OK)
        {
            editor.WriteMessage("\nSelection cancelled.");
            return;
        }

        try
        {
            using (document.LockDocument())
            using (var transaction = database.TransactionManager.StartTransaction())
            {
                _layerService.EnsureLayer(database, DrillPointsLayer);

                var polyline = transaction.GetObject(result.ObjectId, OpenMode.ForRead) as Polyline;
                if (polyline == null)
                {
                    MessageBox.Show("Selected entity is not a polyline.", "Add Drill Points", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var count = Math.Min(polyline.NumberOfVertices, 150);
                var blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
                var modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                for (var i = 0; i < count; i++)
                {
                    var point2d = polyline.GetPoint2dAt(i);
                    var point3d = new Point3d(point2d.X, point2d.Y, 0);
                    var label = $"{letter.Trim()}{i + 1}";

                    var text = new DBText
                    {
                        Position = point3d,
                        Height = 2.0,
                        TextString = label,
                        Layer = DrillPointsLayer,
                        ColorIndex = 7
                    };

                    modelSpace.AppendEntity(text);
                    transaction.AddNewlyCreatedDBObject(text, true);
                }

                transaction.Commit();
            }

            MessageBox.Show("Drill points added successfully.", "Add Drill Points", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (System.Exception ex)
        {
            _log.Error($"Error in AddDrillPoints: {ex.Message}", ex);
            MessageBox.Show($"Error: {ex.Message}", "Add Drill Points", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    public void AddOffsets()
    {
        var document = AutoCADApplication.DocumentManager.MdiActiveDocument;
        if (document == null)
        {
            ShowAlert("No active AutoCAD document.");
            return;
        }

        var confirm = MessageBox.Show(
            "ARE YOU IN GROUND?",
            "Add Offsets",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (confirm != MessageBoxResult.Yes)
        {
            return;
        }

        try
        {
            var database = document.Database;
            var textEntities = GetEntitiesOnLayer(database, DrillPointsLayer, typeof(DBText), typeof(MText)).ToList();
            var gridPoints = new List<(string Label, Point3d Point)>();
            var gridRegex = new Regex("^[A-Z][1-9][0-9]{0,2}$", RegexOptions.IgnoreCase);

            foreach (var entity in textEntities)
            {
                switch (entity)
                {
                    case DBText dbText when gridRegex.IsMatch(dbText.TextString.Trim()):
                        gridPoints.Add((dbText.TextString.Trim(), dbText.Position));
                        break;
                    case MText mText when gridRegex.IsMatch(mText.Contents.Trim()):
                        gridPoints.Add((mText.Contents.Trim(), mText.Location));
                        break;
                }
            }

            if (gridPoints.Count == 0)
            {
                ShowAlert("No Z-DRILL-POINT labels found.");
                return;
            }

            var curves = GetEntitiesOnLayer(database, HorizontalLayer, typeof(Line), typeof(Polyline), typeof(Polyline2d), typeof(Polyline3d))
                .OfType<Curve>()
                .ToList();
            if (curves.Count == 0)
            {
                ShowAlert("No L-SEC-HB polylines/lines found.");
                return;
            }

            var tolerance = new Tolerance(1e-3, 1e-3);
            var warnings = new List<string>();

            using (document.LockDocument())
            using (var transaction = database.TransactionManager.StartTransaction())
            {
                _layerService.EnsureLayer(database, OffsetsLayer);

                var blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
                var modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                bool OffsetExists(Point3d first, Point3d second)
                {
                    foreach (ObjectId id in modelSpace)
                    {
                        if (transaction.GetObject(id, OpenMode.ForRead) is Line line &&
                            line.Layer.Equals(OffsetsLayer, StringComparison.OrdinalIgnoreCase))
                        {
                            if ((line.StartPoint.IsEqualTo(first, tolerance) && line.EndPoint.IsEqualTo(second, tolerance)) ||
                                (line.StartPoint.IsEqualTo(second, tolerance) && line.EndPoint.IsEqualTo(first, tolerance)))
                            {
                                return true;
                            }
                        }
                    }

                    return false;
                }

                void DrawOffset(Point3d from, Point3d to)
                {
                    if (OffsetExists(from, to))
                    {
                        return;
                    }

                    var distance = from.DistanceTo(to);
                    var distanceText = distance.ToString("0.0", CultureInfo.InvariantCulture);

                    var line = new Line(from, to)
                    {
                        Layer = OffsetsLayer
                    };
                    modelSpace.AppendEntity(line);
                    transaction.AddNewlyCreatedDBObject(line, true);

                    var midPoint = new Point3d((from.X + to.X) / 2.0, (from.Y + to.Y) / 2.0, (from.Z + to.Z) / 2.0);
                    var text = new MText
                    {
                        Location = midPoint,
                        TextHeight = 2.5,
                        Contents = $"{{\\C1;{distanceText}}}",
                        Layer = OffsetsLayer
                    };
                    modelSpace.AppendEntity(text);
                    transaction.AddNewlyCreatedDBObject(text, true);

                    document.SendStringToExecute("DIMPERP ", true, false, false);
                }

                foreach (var (label, point) in gridPoints)
                {
                    Curve? northSouth = null;
                    Curve? eastWest = null;
                    Point3d northSouthClosest = Point3d.Origin;
                    Point3d eastWestClosest = Point3d.Origin;
                    double northSouthDelta = double.MaxValue;
                    double eastWestDelta = double.MaxValue;

                    foreach (var curve in curves)
                    {
                        var closest = curve.GetClosestPointTo(point, false);
                        var distance = point.DistanceTo(closest);
                        if (distance > 830.0)
                        {
                            continue;
                        }

                        var dx = Math.Abs(point.X - closest.X);
                        var dy = Math.Abs(point.Y - closest.Y);

                        if (dx < northSouthDelta)
                        {
                            northSouthDelta = dx;
                            northSouth = curve;
                            northSouthClosest = closest;
                        }

                        if (dy < eastWestDelta)
                        {
                            eastWestDelta = dy;
                            eastWest = curve;
                            eastWestClosest = closest;
                        }
                    }

                    var nsMade = false;
                    var ewMade = false;

                    if (northSouth != null)
                    {
                        DrawOffset(point, northSouthClosest);
                        nsMade = true;
                    }

                    if (eastWest != null)
                    {
                        DrawOffset(point, eastWestClosest);
                        ewMade = true;
                    }

                    if (!nsMade)
                    {
                        warnings.Add($"{label} (N-S)");
                    }

                    if (!ewMade)
                    {
                        warnings.Add($"{label} (E-W)");
                    }
                }

                transaction.Commit();
            }

            if (warnings.Count > 0)
            {
                ShowAlert("Unable to find L-SEC-HB for:\n  • " + string.Join("\n  • ", warnings));
            }
            else
            {
                ShowAlert("Add Offsets complete.");
            }
        }
        catch (System.Exception ex)
        {
            _log.Error($"AddOffsets: {ex.Message}", ex);
            ShowAlert($"Error: {ex.Message}");
        }
    }

    public void UpdateOffsets()
    {
        var document = AutoCADApplication.DocumentManager.MdiActiveDocument;
        if (document == null)
        {
            MessageBox.Show("No active AutoCAD document.", "Update Offsets", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        var editor = document.Editor;
        var database = document.Database;

        try
        {
            var options = new PromptEntityOptions("\nSelect the table to update offsets:")
            {
                AllowObjectOnLockedLayer = true
            };
            options.SetRejectMessage("\nOnly table entities are allowed.");
            options.AddAllowedClass(typeof(Table), exactMatch: true);

            var result = editor.GetEntity(options);
            if (result.Status != PromptStatus.OK)
            {
                editor.WriteMessage("\nUpdate Offsets canceled by user.\n");
                return;
            }

            using (document.LockDocument())
            using (var transaction = database.TransactionManager.StartTransaction())
            {
                if (transaction.GetObject(result.ObjectId, OpenMode.ForRead) is not Table table)
                {
                    MessageBox.Show("The selected entity is not a table.", "Update Offsets", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                var columnsToCheck = new[] { 2, 3 };
                var dataCells = new List<(int Row, int Column, double Value, string Direction, string OriginalText)>();
                var offsetRegex = new Regex(@"^\s*([+-]?\d+(\.\d+)?)\s*([NnSsEeWw])\s*$");

                for (var row = 0; row < table.Rows.Count; row++)
                {
                    foreach (var column in columnsToCheck)
                    {
                        if (column >= table.Columns.Count)
                        {
                            continue;
                        }

                        var text = table.Cells[row, column].TextString?.Trim() ?? string.Empty;
                        if (string.IsNullOrWhiteSpace(text))
                        {
                            continue;
                        }

                        var match = offsetRegex.Match(text);
                        if (!match.Success)
                        {
                            continue;
                        }

                        var numeric = match.Groups[1].Value;
                        var direction = match.Groups[3].Value.ToUpperInvariant();
                        if (double.TryParse(numeric, NumberStyles.Any, CultureInfo.InvariantCulture, out var value))
                        {
                            dataCells.Add((row, column, value, direction, text));
                        }
                    }
                }

                if (dataCells.Count == 0)
                {
                    MessageBox.Show("No numeric offset values found in columns C or D of the selected table.", "Update Offsets", MessageBoxButton.OK, MessageBoxImage.Warning);
                    _log.Info("Update Offsets: No data found in columns C/D.");
                    return;
                }

                var blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
                var modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                var offsetTextIds = new List<ObjectId>();
                foreach (ObjectId entityId in modelSpace)
                {
                    if (transaction.GetObject(entityId, OpenMode.ForRead) is not Entity entity)
                    {
                        continue;
                    }

                    if (!entity.Layer.Equals(OffsetsLayer, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (entity is DBText or MText)
                    {
                        offsetTextIds.Add(entityId);
                    }
                }

                if (offsetTextIds.Count == 0)
                {
                    MessageBox.Show($"No offset text found on layer \"{OffsetsLayer}\".", "Update Offsets", MessageBoxButton.OK, MessageBoxImage.Information);
                    _log.Info("Update Offsets: No text objects on P-Drill-Offset layer.");
                    return;
                }

                if (offsetTextIds.Count > dataCells.Count)
                {
                    editor.SetImpliedSelection(offsetTextIds.ToArray());
                    var warning = $"Found {offsetTextIds.Count} offset text objects for only {dataCells.Count} table entries. Please review and try again.";
                    editor.WriteMessage($"\n{warning}\n");
                    MessageBox.Show(warning, "Update Offsets", MessageBoxButton.OK, MessageBoxImage.Warning);
                    _log.Warn($"Update Offsets aborted: {offsetTextIds.Count} offset texts for {dataCells.Count} table cells.");
                    return;
                }

                var offsetValues = new List<(double Value, ObjectId Id)>();
                foreach (var entityId in offsetTextIds)
                {
                    var entity = transaction.GetObject(entityId, OpenMode.ForRead);
                    switch (entity)
                    {
                        case DBText dbText:
                            var text = dbText.TextString.Trim();
                            _log.Info($"DBText => '{text}'");
                            if (double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var dbValue))
                            {
                                offsetValues.Add((dbValue, entityId));
                            }
                            break;
                        case MText mText:
                            var raw = mText.Contents;
                            _log.Info($"Raw MText.Contents => '{raw}'");
                            var plain = Regex.Replace(raw, @"\{\\[^\}]+;([^\}]+)\}", "$1");
                            plain = Regex.Replace(plain, @"\\[a-zA-Z]+\s*", " ");
                            plain = plain.Trim();
                            _log.Info($"Cleaned MText => '{plain}'");

                            var lines = plain.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                            var foundNumeric = false;
                            foreach (var line in lines)
                            {
                                var candidate = line.Trim();
                                if (double.TryParse(candidate, NumberStyles.Any, CultureInfo.InvariantCulture, out var parsed))
                                {
                                    offsetValues.Add((parsed, entityId));
                                    foundNumeric = true;
                                    break;
                                }

                                var matchNumber = Regex.Match(candidate, @"([+-]?\d+(\.\d+)?)");
                                if (matchNumber.Success && double.TryParse(matchNumber.Groups[1].Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var extracted))
                                {
                                    offsetValues.Add((extracted, entityId));
                                    foundNumeric = true;
                                    break;
                                }
                            }

                            if (!foundNumeric)
                            {
                                _log.Info("No numeric value found in MText => skipping.");
                            }

                            break;
                    }
                }

                if (offsetValues.Count == 0)
                {
                    MessageBox.Show($"Text found on \"{OffsetsLayer}\" layer, but none contained valid numeric values.", "Update Offsets", MessageBoxButton.OK, MessageBoxImage.Warning);
                    _log.Info("Update Offsets: No numeric values found in offset text objects.");
                    return;
                }

                const double tolerance = 10.0;
                var updatedCount = 0;
                var noMatchCount = 0;

                table.UpgradeOpen();

                foreach (var cell in dataCells)
                {
                    var bestDiff = double.MaxValue;
                    var bestIndex = -1;

                    for (var index = 0; index < offsetValues.Count; index++)
                    {
                        var diff = Math.Abs(offsetValues[index].Value - cell.Value);
                        if (diff < bestDiff)
                        {
                            bestDiff = diff;
                            bestIndex = index;
                        }
                    }

                    if (bestIndex != -1 && bestDiff <= tolerance)
                    {
                        var matched = offsetValues[bestIndex].Value;
                        offsetValues.RemoveAt(bestIndex);

                        var formatted = matched.ToString("F1", CultureInfo.InvariantCulture) + " " + cell.Direction;
                        table.Cells[cell.Row, cell.Column].TextString = formatted;
                        updatedCount++;
                    }
                    else
                    {
                        table.Cells[cell.Row, cell.Column].TextString = cell.OriginalText;
                        table.Cells[cell.Row, cell.Column].BackgroundColor = Color.FromRgb(255, 0, 0);
                        noMatchCount++;
                    }
                }

                table.GenerateLayout();
                transaction.Commit();

                editor.WriteMessage($"\nUpdate Offsets completed: {updatedCount} cells updated, {noMatchCount} cells with no match.\n");
                MessageBox.Show($"Update Offsets completed: {updatedCount} cells updated, {noMatchCount} cells with no match.", "Update Offsets", MessageBoxButton.OK, MessageBoxImage.Information);
                _log.Info($"Update Offsets done => {updatedCount} updated, {noMatchCount} no-match cells.");
            }
        }
        catch (System.Exception ex)
        {
            _log.Error($"UpdateOffsets: {ex.Message}", ex);
            MessageBox.Show($"Error updating offsets: {ex.Message}", "Update Offsets", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static List<string> ExtractBottomHoleValues(Table table)
    {
        var results = new List<string>();
        for (var row = 0; row < table.Rows.Count; row++)
        {
            var columnValue = table.Cells[row, 0].TextString.Trim();
            var normalized = columnValue.ToUpperInvariant().Replace(" ", string.Empty);
            if (normalized.Contains("BOTTOMHOLE"))
            {
                var value = table.Cells[row, 1].TextString.Trim();
                results.Add(value);
            }
        }

        if (results.Count == 0)
        {
            MessageBox.Show("No 'BOTTOM HOLE' entries found in the selected table.", "Check", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        return results;
    }

    private List<BlockAttributeData>? SelectBlocksWithDrillName(Document document)
    {
        try
        {
            var editor = document.Editor;
            var database = document.Database;

            var options = new PromptSelectionOptions
            {
                MessageForAdding = "\nSelect blocks with 'DRILLNAME' attribute:",
                AllowDuplicates = false
            };

            var filter = new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "INSERT") });
            var selectionResult = editor.GetSelection(options, filter);
            if (selectionResult.Status != PromptStatus.OK)
            {
                return null;
            }

            var results = new List<BlockAttributeData>();
            using (var transaction = database.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject selected in selectionResult.Value)
                {
                    if (selected == null)
                    {
                        continue;
                    }

                    if (transaction.GetObject(selected.ObjectId, OpenMode.ForRead) is not BlockReference block)
                    {
                        continue;
                    }

                    var drillName = GetAttributeValue(block, "DRILLNAME", transaction);
                    if (string.IsNullOrEmpty(drillName))
                    {
                        continue;
                    }

                    results.Add(new BlockAttributeData
                    {
                        BlockReference = null,
                        DrillName = drillName,
                        YCoordinate = block.Position.Y
                    });
                }

                transaction.Commit();
            }

            return results;
        }
        catch (System.Exception ex)
        {
            _log.Error($"Error in SelectBlocksWithDrillName: {ex.Message}", ex);
            return null;
        }
    }

    private static string GetAttributeValue(BlockReference blockReference, string tag, Transaction transaction)
    {
        foreach (ObjectId id in blockReference.AttributeCollection)
        {
            if (transaction.GetObject(id, OpenMode.ForRead) is AttributeReference attribute &&
                AutoCADBlockService.TagMatches(attribute.Tag, tag))
            {
                return attribute.TextString.Trim();
            }
        }

        return string.Empty;
    }

    private DrillCheckSummary CompareDrillNamesWithTable(IReadOnlyList<string> drillNames, IReadOnlyList<string> tableValues, List<BlockAttributeData> blockData)
    {
        blockData.Sort((a, b) => b.YCoordinate.CompareTo(a.YCoordinate));

        var comparisons = Math.Min(Math.Min(tableValues.Count, drillNames.Count), blockData.Count);
        var discrepancies = new List<string>();
        var reportLines = new List<string>();
        var results = new List<DrillCheckResult>();

        for (var i = 0; i < comparisons; i++)
        {
            var tableValue = tableValues[i];
            var drillName = (drillNames[i] ?? string.Empty).Trim();
            var blockDrillName = blockData[i].DrillName;

            var normalizedDrillName = DrillParsers.NormalizeDrillName(drillName);
            var normalizedTableValue = DrillParsers.NormalizeTableValue(tableValue);
            var normalizedBlockName = DrillParsers.NormalizeDrillName(blockDrillName);

            var mismatch = false;
            var details = new List<string>();

            if (!string.Equals(normalizedDrillName, normalizedTableValue, StringComparison.OrdinalIgnoreCase))
            {
                details.Add("Drill name does not match table value.");
                mismatch = true;
            }

            if (!string.Equals(normalizedBlockName, normalizedDrillName, StringComparison.OrdinalIgnoreCase))
            {
                details.Add("Block DRILLNAME does not match drill name.");
                mismatch = true;
            }

            if (!string.Equals(normalizedBlockName, normalizedTableValue, StringComparison.OrdinalIgnoreCase))
            {
                details.Add("Block DRILLNAME does not match table value.");
                mismatch = true;
            }

            reportLines.Add($"DRILL_{i + 1} NAME: {drillName}");
            reportLines.Add($"TABLE RESULT: {tableValue}");
            reportLines.Add($"BLOCK DRILLNAME: {blockDrillName}");
            reportLines.Add($"Normalized Drill Name: {normalizedDrillName}");
            reportLines.Add($"Normalized Table Value: {normalizedTableValue}");
            reportLines.Add($"Normalized Block DrillName: {normalizedBlockName}");
            reportLines.Add($"STATUS: {(mismatch ? "FAIL" : "PASS")}");

            if (mismatch)
            {
                reportLines.Add("Discrepancies:");
                foreach (var detail in details)
                {
                    reportLines.Add($"- {detail}");
                }

                discrepancies.Add($"DRILL_{i + 1}: {string.Join("; ", details)}");
            }

            results.Add(new DrillCheckResult(i + 1, drillName, tableValue, blockDrillName, details));
            reportLines.Add(string.Empty);
        }

        var reportPath = GetReportFilePath();
        try
        {
            File.WriteAllLines(reportPath, reportLines);
        }
        catch (System.Exception ex)
        {
            _log.Error($"Error writing report file: {ex.Message}", ex);
            MessageBox.Show($"An error occurred while writing the report file: {ex.Message}", "Check", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        if (discrepancies.Count > 0)
        {
            var message = "Discrepancies found:\n" + string.Join("\n", discrepancies);
            MessageBox.Show($"{message}\n\nDetailed report saved at:\n{reportPath}", "Check Results", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
        else
        {
            MessageBox.Show($"All drill names match the table values and block DRILLNAME attributes.\n\nDetailed report saved at:\n{reportPath}", "Check Results", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        return new DrillCheckSummary(completed: true, results, reportPath);
    }
    private static string GetReportFilePath()
    {
        var document = AutoCADApplication.DocumentManager.MdiActiveDocument;
        if (document == null || string.IsNullOrEmpty(document.Name))
        {
            throw new InvalidOperationException("No active AutoCAD document found.");
        }

        var directory = Path.GetDirectoryName(document.Name) ?? CordsDirectory;
        var drawingName = Path.GetFileNameWithoutExtension(document.Name);
        return Path.Combine(directory, $"{drawingName}_CheckReport.txt");
    }

    private List<(string Label, double Northing, double Easting)> ReadGridPoints(Database database)
    {
        var results = new List<(string Label, double Northing, double Easting)>();

        using (var transaction = database.TransactionManager.StartTransaction())
        {
            var blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
            var modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            foreach (ObjectId id in modelSpace)
            {
                if (transaction.GetObject(id, OpenMode.ForRead) is Entity entity &&
                    entity.Layer.Equals(DrillPointsLayer, StringComparison.OrdinalIgnoreCase))
                {
                    switch (entity)
                    {
                        case DBText text when DrillParsers.IsGridLabel(text.TextString):
                            results.Add((text.TextString.Trim(), text.Position.Y, text.Position.X));
                            break;
                        case MText mText when DrillParsers.IsGridLabel(mText.Contents):
                            results.Add((mText.Contents.Trim(), mText.Location.Y, mText.Location.X));
                            break;
                    }
                }
            }

            transaction.Commit();
        }

        return results;
    }

    private static List<(int Row, int Column)> FindSurfaceCells(Table table)
    {
        var cells = new List<(int Row, int Column)>();
        for (var row = 0; row < table.Rows.Count; row++)
        {
            for (var column = 0; column < table.Columns.Count; column++)
            {
                var text = table.Cells[row, column].TextString.Trim().ToUpperInvariant();
                if (text.Contains("SURFACE"))
                {
                    cells.Add((row, column));
                }
            }
        }

        return cells;
    }

    private static List<string> GetNonDefaultDrills(IReadOnlyList<string> drillNames)
    {
        var results = new List<string>();
        for (var i = 0; i < drillNames.Count; i++)
        {
            var name = (drillNames[i] ?? string.Empty).Trim();
            var defaultName = $"DRILL_{i + 1}";
            if (!string.IsNullOrWhiteSpace(name) &&
                !name.Equals(defaultName, StringComparison.OrdinalIgnoreCase))
            {
                results.Add(name);
            }
        }

        return results;
    }

    private static Point3d GetCellNorthWest(Table table, int row, int column)
    {
        try
        {
            // The table API exposes precise cell bounds that account for rotation, scaling, and
            // title/header rows. Using the extents ensures the heading blocks are anchored to
            // the actual north-west (top-left) corner of the target cell even for data-linked
            // tables whose insertion point is the lower-left corner of the overall grid.
            var points = new Point3dCollection();
            table.GetCellExtents(row, column, isOuterCell: true, points);
            if (points.Count > 0)
            {
                var minX = points.Cast<Point3d>().Min(point => point.X);
                var maxY = points.Cast<Point3d>().Max(point => point.Y);
                var minZ = points.Cast<Point3d>().Min(point => point.Z);
                return new Point3d(minX, maxY, minZ);
            }
        }
        catch (Autodesk.AutoCAD.Runtime.Exception)
        {
            // Older drawing versions may not support resolving cell extents. Fall back to the
            // manual calculation so we still place a block, even if alignment is approximate.
        }

        var x = table.Position.X;
        var y = table.Position.Y;

        for (var index = 0; index < column; index++)
        {
            x += table.Columns[index].Width;
        }

        // Table.Position is the lower-left of the table. To emulate a north-west corner we
        // need to walk upward from the bottom of the table to the requested row. Summing the
        // height of the target row and every row beneath it moves us to the correct vertical
        // offset.
        for (var index = row; index < table.Rows.Count; index++)
        {
            y += table.Rows[index].Height;
        }

        return new Point3d(x, y, 0.0);
    }

    private string DrillCsvPipeline(Database database)
    {
        if (!LayerExists(database, DrillPointsLayer))
        {
            ShowAlert("Layer 'Z-DRILL-POINT' not found.");
            return string.Empty;
        }

        var gridData = ReadGridPoints(database).OrderBy(p => p.Label, _naturalComparer).ToList();

        Directory.CreateDirectory(CordsDirectory);
        var csvPath = Path.Combine(CordsDirectory, "cords.csv");
        using (var writer = new StreamWriter(csvPath, false))
        {
            writer.WriteLine("Label,Northing,Easting");
            foreach (var point in gridData)
            {
                writer.WriteLine($"{point.Label},{point.Northing},{point.Easting}");
            }
        }

        ShowAlert("DONT TOUCH, WAIT FOR INSTRUCTION");
        return csvPath;
    }

    private string RunCordsExecutable(string csvPath, string heading)
    {
        if (string.IsNullOrEmpty(csvPath))
        {
            return string.Empty;
        }

        var processExe = FindCordsExecutable();
        if (!string.IsNullOrEmpty(processExe))
        {
            _log.Info($"Running: {processExe} \"{csvPath}\" \"{heading}\"");
            var startInfo = new ProcessStartInfo(processExe, $"\"{csvPath}\" \"{heading}\"")
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = Encoding.GetEncoding(1252),
                StandardErrorEncoding = Encoding.GetEncoding(1252)
            };

            using (var process = Process.Start(startInfo))
            {
                if (process == null)
                {
                    ShowAlert("Failed to start cord.exe.");
                    return string.Empty;
                }

                if (!process.WaitForExit(180_000))
                {
                    try
                    {
                        process.Kill();
                    }
                    catch
                    {
                        // ignore
                    }

                    _log.Error("cord.exe timeout");
                    ShowAlert("cord.exe did not exit in time.");
                    return string.Empty;
                }

                var output = process.StandardOutput.ReadToEnd();
                var error = process.StandardError.ReadToEnd();
                if (!string.IsNullOrEmpty(output))
                {
                    _log.Info(output);
                }

                if (!string.IsNullOrEmpty(error))
                {
                    _log.Error(error);
                }

                if (process.ExitCode != 0)
                {
                    _log.Error($"cord.exe exited with {process.ExitCode}");
                    ShowAlert($"cord.exe exited with {process.ExitCode}");
                    return string.Empty;
                }
            }
        }
        else
        {
            _log.Warn("cord.exe not found in any expected location, skipping.");
            ShowAlert("cord.exe not found in the expected locations.");
        }

        var excelPath = Path.Combine(Path.GetDirectoryName(csvPath) ?? CordsDirectory, "ExportedCoordsFormatted.xlsx");
        if (!File.Exists(excelPath))
        {
            ShowAlert("ExportedCoordsFormatted.xlsx not found.");
            return string.Empty;
        }

        return excelPath;
    }

    private static string? FindCordsExecutable()
    {
        foreach (var path in CordsExecutableSearchPaths)
        {
            if (File.Exists(path))
            {
                return path;
            }
        }

        return null;
    }

    private static string[,]? ReadExcel(string excelFilePath)
    {
        using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
        {
            var worksheet = package.Workbook.Worksheets.First();
            var dimension = GetWorksheetDimension(worksheet);
            if (dimension == null)
            {
                ShowAlert("The Excel file is empty.");
                return null;
            }

            var (startRow, startColumn, endRow, endColumn) = dimension.Value;
            var lastRow = endRow;
            for (var row = endRow; row >= startRow; row--)
            {
                var blank = true;
                for (var column = startColumn; column <= endColumn; column++)
                {
                    if (!string.IsNullOrWhiteSpace(worksheet.Cells[row, column].Text))
                    {
                        blank = false;
                        break;
                    }
                }

                if (!blank)
                {
                    lastRow = row;
                    break;
                }
            }

            var rowCount = lastRow - startRow + 1;
            var columnCount = endColumn - startColumn + 1;
            var data = new string[rowCount, columnCount];
            for (var row = 0; row < rowCount; row++)
            {
                for (var column = 0; column < columnCount; column++)
                {
                    data[row, column] = worksheet.Cells[startRow + row, startColumn + column].Text;
                }
            }

            return data;
        }
    }

    private static (int StartRow, int StartColumn, int EndRow, int EndColumn)? GetWorksheetDimension(ExcelWorksheet worksheet)
    {
        if (worksheet == null)
        {
            return null;
        }

        var dimensionProperty = worksheet.GetType().GetProperty("Dimension");
        if (dimensionProperty?.GetValue(worksheet) is ExcelAddressBase dimension && dimension.Start != null && dimension.End != null)
        {
            return (dimension.Start.Row, dimension.Start.Column, dimension.End.Row, dimension.End.Column);
        }

        var hasValues = false;
        var startRow = int.MaxValue;
        var startColumn = int.MaxValue;
        var endRow = 0;
        var endColumn = 0;

        foreach (var cell in worksheet.Cells)
        {
            if (cell == null)
            {
                continue;
            }

            var hasContent = cell.Value != null || !string.IsNullOrEmpty(cell.Text) || !string.IsNullOrEmpty(cell.Formula);
            if (!hasContent)
            {
                continue;
            }

            hasValues = true;
            startRow = Math.Min(startRow, cell.Start.Row);
            startColumn = Math.Min(startColumn, cell.Start.Column);
            endRow = Math.Max(endRow, cell.End.Row);
            endColumn = Math.Max(endColumn, cell.End.Column);
        }

        if (!hasValues)
        {
            return null;
        }

        return (startRow, startColumn, endRow, endColumn);
    }

    private void AdjustTableForClient(string[,] tableData, string heading)
    {
        if (tableData == null || tableData.GetLength(0) == 0 || tableData.GetLength(1) == 0)
        {
            return;
        }

        var headingValue = string.Equals(heading, "HEEL", StringComparison.OrdinalIgnoreCase) ? "HEEL" : "ICP";
        for (var row = 0; row < tableData.GetLength(0); row++)
        {
            for (var column = 0; column < tableData.GetLength(1); column++)
            {
                var cellText = tableData[row, column];
                if (string.IsNullOrEmpty(cellText))
                {
                    continue;
                }

                if (headingValue == "HEEL" && cellText.IndexOf("ICP", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    tableData[row, column] = Regex.Replace(cellText, "ICP", "HEEL", RegexOptions.IgnoreCase);
                }
                else if (headingValue == "ICP" && cellText.IndexOf("HEEL", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    tableData[row, column] = Regex.Replace(cellText, "HEEL", "ICP", RegexOptions.IgnoreCase);
                }
            }
        }

        var groups = new Dictionary<char, List<int>>();
        var tagRegex = new Regex("^[A-Z][0-9]+$", RegexOptions.IgnoreCase);

        for (var row = 0; row < tableData.GetLength(0); row++)
        {
            char? letter = null;
            for (var column = 0; column < tableData.GetLength(1); column++)
            {
                var value = tableData[row, column];
                if (string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }

                var trimmed = value.Trim();
                if (tagRegex.IsMatch(trimmed))
                {
                    letter = char.ToUpperInvariant(trimmed[0]);
                    break;
                }
            }

            if (!letter.HasValue)
            {
                continue;
            }

            if (!groups.TryGetValue(letter.Value, out var rows))
            {
                rows = new List<int>();
                groups[letter.Value] = rows;
            }

            rows.Add(row);
        }

        foreach (var entry in groups)
        {
            var rows = entry.Value;
            rows.Sort();
            var count = rows.Count;
            if (count == 0)
            {
                continue;
            }

            tableData[rows[0], 0] = "SURFACE";

            if (count == 1)
            {
                continue;
            }

            tableData[rows[count - 1], 0] = "BOTTOM HOLE";
            tableData[rows[1], 0] = headingValue;

            for (var index = 2; index < count - 1; index++)
            {
                tableData[rows[index], 0] = $"TURN #{index - 1}";
            }
        }

        var nonEmptyRows = 0;
        for (var row = 0; row < tableData.GetLength(0); row++)
        {
            var hasData = false;
            for (var column = 0; column < tableData.GetLength(1); column++)
            {
                if (!string.IsNullOrWhiteSpace(tableData[row, column]))
                {
                    hasData = true;
                    break;
                }
            }

            if (hasData)
            {
                nonEmptyRows++;
            }
        }

        if (nonEmptyRows == 1)
        {
            var cellValue = tableData[0, 0] ?? string.Empty;
            var normalized = cellValue.ToUpperInvariant().Replace(" ", string.Empty);
            if (normalized.Contains("BOTTOMHOLE"))
            {
                tableData[0, 0] = "SURFACE";
            }
        }

        var hasSurfaceInFirstColumn = false;
        var bottomHoleRows = new List<int>();
        for (var row = 0; row < tableData.GetLength(0); row++)
        {
            var value = tableData[row, 0] ?? string.Empty;
            var normalized = value.ToUpperInvariant().Replace(" ", string.Empty);
            if (normalized.Contains("SURFACE"))
            {
                hasSurfaceInFirstColumn = true;
                break;
            }

            if (normalized.Contains("BOTTOMHOLE"))
            {
                bottomHoleRows.Add(row);
            }
        }

        if (!hasSurfaceInFirstColumn)
        {
            foreach (var row in bottomHoleRows)
            {
                tableData[row, 0] = "SURFACE";
            }
        }
    }

    private bool InsertTablePipeline(Document document, string[,] tableData)
    {
        var editor = document.Editor;
        ShowAlert("BACK TO CAD, PICK A POINT");
        var pointResult = editor.GetPoint("\nSelect insertion point:");
        if (pointResult.Status != PromptStatus.OK)
        {
            editor.WriteMessage("\nCancelled.");
            return false;
        }

        using (document.LockDocument())
        {
            var database = document.Database;
            _layerService.EnsureLayer(database, "CG-NOTES");
            InsertAndFormatTable(database, pointResult.Value, tableData, "induction Bend");
        }

        return true;
    }

    private void InsertAndFormatTable(Database database, Point3d insertionPoint, string[,] cellData, string tableStyleName)
    {
        using (var transaction = database.TransactionManager.StartTransaction())
        {
            var blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
            var modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

            var rows = cellData.GetLength(0);
            var columns = cellData.GetLength(1);

            double[] columnWidths;
            if (columns == 12)
            {
                columnWidths = new[]
                {
                    100.0, 100.0, 60.0, 60.0,
                    80.0, 80.0, 80.0, 80.0,
                    80.0, 80.0, 80.0, 80.0
                };
            }
            else
            {
                columnWidths = Enumerable.Repeat(80.0, columns).ToArray();
            }

            var table = new Table
            {
                TableStyle = GetTableStyleId(database, transaction, tableStyleName),
                Position = insertionPoint,
                Layer = "CG-NOTES"
            };
            table.SetSize(rows, columns);

            const double defaultRowHeight = 25.0;
            const double emptyRowHeight = 125.0;

            for (var row = 0; row < rows; row++)
            {
                var hasEmpty = false;
                for (var column = 0; column < columns; column++)
                {
                    var value = cellData[row, column] ?? string.Empty;
                    table.Cells[row, column].TextString = value;
                    if (string.IsNullOrWhiteSpace(value))
                    {
                        hasEmpty = true;
                    }
                }

                table.Rows[row].Height = hasEmpty ? emptyRowHeight : defaultRowHeight;
            }

            for (var column = 0; column < columns; column++)
            {
                table.Columns[column].Width = columnWidths[column];
            }

            modelSpace.AppendEntity(table);
            transaction.AddNewlyCreatedDBObject(table, true);

            table.GenerateLayout();
            UnmergeAllCells(table);
            table.GenerateLayout();
            RemoveAllCellBorders(table);
            AddBordersForDataCells(table);
            table.RecomputeTableBlock(true);

            transaction.Commit();
        }
    }

    private static ObjectId GetTableStyleId(Database database, Transaction transaction, string styleName)
    {
        var dictionary = (DBDictionary)transaction.GetObject(database.TableStyleDictionaryId, OpenMode.ForRead);
        foreach (var entry in dictionary)
        {
            if (entry.Key.Equals(styleName, StringComparison.OrdinalIgnoreCase))
            {
                return entry.Value;
            }
        }

        return ObjectId.Null;
    }

    private static void UnmergeAllCells(Table table)
    {
        for (var row = 0; row < table.Rows.Count; row++)
        {
            for (var column = 0; column < table.Columns.Count; column++)
            {
                var range = table.Cells[row, column].GetMergeRange();
                if (range != null && range.TopRow == row && range.LeftColumn == column)
                {
                    table.UnmergeCells(range);
                }
            }
        }
    }

    private static void RemoveAllCellBorders(Table table)
    {
        for (var row = 0; row < table.Rows.Count; row++)
        {
            for (var column = 0; column < table.Columns.Count; column++)
            {
                var cell = table.Cells[row, column];
                cell.Borders.Top.IsVisible = false;
                cell.Borders.Bottom.IsVisible = false;
                cell.Borders.Left.IsVisible = false;
                cell.Borders.Right.IsVisible = false;
            }
        }
    }

    private static void AddBordersForDataCells(Table table)
    {
        var color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 7);
        for (var row = 0; row < table.Rows.Count; row++)
        {
            for (var column = 0; column < table.Columns.Count; column++)
            {
                var cell = table.Cells[row, column];
                var value = cell.TextString;
                var hasData = !string.IsNullOrWhiteSpace(value);
                if (hasData)
                {
                    SetCellBorders(cell, true, color);
                }
                else
                {
                    var top = row > 0 && !string.IsNullOrWhiteSpace(table.Cells[row - 1, column].TextString);
                    var bottom = row < table.Rows.Count - 1 && !string.IsNullOrWhiteSpace(table.Cells[row + 1, column].TextString);
                    var left = column > 0 && !string.IsNullOrWhiteSpace(table.Cells[row, column - 1].TextString);
                    var right = column < table.Columns.Count - 1 && !string.IsNullOrWhiteSpace(table.Cells[row, column + 1].TextString);

                    cell.Borders.Top.IsVisible = top;
                    cell.Borders.Bottom.IsVisible = bottom;
                    cell.Borders.Left.IsVisible = left;
                    cell.Borders.Right.IsVisible = right;
                }
            }
        }
    }

    private static void SetCellBorders(Cell cell, bool isVisible, Autodesk.AutoCAD.Colors.Color color)
    {
        cell.Borders.Top.IsVisible = isVisible;
        cell.Borders.Bottom.IsVisible = isVisible;
        cell.Borders.Left.IsVisible = isVisible;
        cell.Borders.Right.IsVisible = isVisible;
        if (isVisible)
        {
            cell.Borders.Top.LineWeight = LineWeight.LineWeight025;
            cell.Borders.Bottom.LineWeight = LineWeight.LineWeight025;
            cell.Borders.Left.LineWeight = LineWeight.LineWeight025;
            cell.Borders.Right.LineWeight = LineWeight.LineWeight025;
            cell.Borders.Top.Color = color;
            cell.Borders.Bottom.Color = color;
            cell.Borders.Left.Color = color;
            cell.Borders.Right.Color = color;
        }
    }

    private static IEnumerable<Entity> GetEntitiesOnLayer(Database database, string layer, params Type[] types)
    {
        var results = new List<Entity>();
        using (var transaction = database.TransactionManager.StartTransaction())
        {
            var blockTable = (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
            var modelSpace = (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            foreach (ObjectId id in modelSpace)
            {
                if (transaction.GetObject(id, OpenMode.ForRead) is Entity entity)
                {
                    if (!entity.Layer.Equals(layer, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (types != null && types.Length > 0)
                    {
                        var rx = entity.GetRXClass();
                        var match = types.Any(type => rx.IsDerivedFrom(RXObject.GetClass(type)));
                        if (!match)
                        {
                            continue;
                        }
                    }

                    results.Add((Entity)entity.Clone());
                }
            }

            transaction.Commit();
        }

        return results;
    }

    private static bool LayerExists(Database database, string name)
    {
        using (var transaction = database.TransactionManager.StartTransaction())
        {
            var layerTable = (LayerTable)transaction.GetObject(database.LayerTableId, OpenMode.ForRead);
            var exists = layerTable.Has(name);
            transaction.Commit();
            return exists;
        }
    }

    private static void ShowAlert(string message)
    {
        AutoCADApplication.ShowAlertDialog(message);
    }

}
