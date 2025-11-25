using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.AutoCAD.ApplicationServices;
using Compass.Models;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Compass.Services;

public class JsonSettingsService
{
    private readonly AppSettings _appSettings;
    private static readonly JsonSerializerOptions SerializerOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    public JsonSettingsService(AppSettings appSettings)
    {
        _appSettings = appSettings ?? throw new ArgumentNullException(nameof(appSettings));
    }

    public void Save(DrillGridState state)
    {
        if (state == null)
        {
            throw new ArgumentNullException(nameof(state));
        }

        var path = GetPath();
        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var json = SerializeIndented(state);
        File.WriteAllText(path, json);
    }

    public DrillGridState Load(int drillCount)
    {
        var path = GetPath();
        if (!File.Exists(path))
        {
            return DrillGridState.CreateDefault(drillCount);
        }

        var json = File.ReadAllText(path);
        var state = JsonSerializer.Deserialize<DrillGridState>(json, SerializerOptions) ?? new DrillGridState();
        EnsureCount(state.DrillNames, drillCount);
        return state;
    }

    private string GetPath()
    {
        var document = Application.DocumentManager.MdiActiveDocument;
        if (document != null && !string.IsNullOrEmpty(document.Name))
        {
            var directory = Path.GetDirectoryName(document.Name);
            var drawingName = Path.GetFileNameWithoutExtension(document.Name) ?? string.Empty;
            drawingName = drawingName.TrimEnd('-');

            var parts = drawingName.Split('-');
            var prefix = parts.Length >= 2 ? $"{parts[0]}-{parts[1]}" : drawingName;

            if (!string.IsNullOrEmpty(directory))
            {
                return Path.Combine(directory, $"{prefix}.json");
            }

            if (!string.IsNullOrWhiteSpace(prefix))
            {
                return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{prefix}.json");
            }
        }

        var fallback = _appSettings.JsonConfigName;
        if (string.IsNullOrWhiteSpace(fallback))
        {
            fallback = "drillProps.json";
        }

        var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
        return Path.Combine(baseDirectory, fallback);
    }

    private static void EnsureCount(IList<string> names, int desiredCount)
    {
        if (names == null)
        {
            return;
        }

        while (names.Count < desiredCount)
        {
            names.Add(string.Empty);
        }

        if (names.Count > desiredCount)
        {
            for (var i = names.Count - 1; i >= desiredCount; i--)
            {
                names.RemoveAt(i);
            }
        }
    }

    private static string SerializeIndented(object value)
    {
        return JsonSerializer.Serialize(value, SerializerOptions);
    }
}
