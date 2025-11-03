using System;
using System.Collections.Generic;
using System.Windows;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Compass.Infrastructure.Logging;

namespace Compass.Services;

public class DrillAttributeSyncService
{
    private readonly AutoCADBlockService _blockService;
    private readonly ILog _log;

    public DrillAttributeSyncService(ILog log, AutoCADBlockService blockService)
    {
        _log = log ?? throw new ArgumentNullException(nameof(log));
        _blockService = blockService ?? throw new ArgumentNullException(nameof(blockService));
    }

    public IReadOnlyList<string>? GetDrillNamesFromSelection(int drillCount)
    {
        var document = Application.DocumentManager.MdiActiveDocument;
        if (document == null)
        {
            MessageBox.Show("No active AutoCAD document is available.", "Update", MessageBoxButton.OK, MessageBoxImage.Information);
            return null;
        }

        if (drillCount <= 0)
        {
            return Array.Empty<string>();
        }

        var editor = document.Editor;
        var selection = editor.GetSelection();
        if (selection.Status != PromptStatus.OK)
        {
            MessageBox.Show("No objects selected.", "Update", MessageBoxButton.OK, MessageBoxImage.Information);
            return null;
        }

        try
        {
            var values = new string[drillCount];
            var updated = new bool[drillCount];

            using (document.LockDocument())
            using (var transaction = document.Database.TransactionManager.StartTransaction())
            {
                foreach (SelectedObject selected in selection.Value)
                {
                    if (selected == null)
                    {
                        continue;
                    }

                    if (transaction.GetObject(selected.ObjectId, OpenMode.ForRead) is not BlockReference block)
                    {
                        continue;
                    }

                    for (var index = 0; index < drillCount; index++)
                    {
                        var tag = $"DRILL_{index + 1}";
                        var value = _blockService.GetAttributeValue(block, tag, transaction);
                        if (!string.IsNullOrEmpty(value))
                        {
                            values[index] = value;
                            updated[index] = true;
                        }
                    }
                }

                transaction.Commit();
            }

            for (var i = 0; i < drillCount; i++)
            {
                if (!updated[i])
                {
                    values[i] = $"DRILL_{i + 1}";
                }
            }

            return values;
        }
        catch (System.Exception ex)
        {
            _log.Error("Failed to update drill names from block attributes.", ex);
            MessageBox.Show($"Error updating from block attributes: {ex.Message}", "Update", MessageBoxButton.OK, MessageBoxImage.Error);
            return null;
        }
    }
}
