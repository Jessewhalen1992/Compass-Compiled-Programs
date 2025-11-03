using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;

namespace Compass.Serialization;

public static class DrillPropsJsonStore
{
    private sealed class DrillPropsJsonModel
    {
        public List<string>? Drills { get; set; }

        public Dictionary<string, string?>? Slots { get; set; }
    }

    public static IReadOnlyList<string> Load(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("A JSON file path is required.", nameof(path));
        }

        if (!File.Exists(path))
        {
            throw new FileNotFoundException("The drill props JSON file was not found.", path);
        }

        var json = File.ReadAllText(path);
        if (string.IsNullOrWhiteSpace(json))
        {
            return Array.Empty<string>();
        }

        var serializer = new JavaScriptSerializer();
        var model = serializer.Deserialize<DrillPropsJsonModel>(json);
        if (model == null)
        {
            return Array.Empty<string>();
        }

        if (model.Drills is { Count: > 0 } list)
        {
            return list.Select(name => name ?? string.Empty).ToList();
        }

        if (model.Slots is { Count: > 0 } slots)
        {
            return slots
                .OrderBy(pair => ParseSlotKey(pair.Key))
                .Select(pair => pair.Value ?? string.Empty)
                .ToList();
        }

        return Array.Empty<string>();
    }

    public static void Save(string path, IEnumerable<string> drills)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("A JSON file path is required.", nameof(path));
        }

        if (drills == null)
        {
            throw new ArgumentNullException(nameof(drills));
        }

        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var payload = new DrillPropsJsonModel
        {
            Drills = drills.Select(name => name ?? string.Empty).ToList()
        };

        var serializer = new JavaScriptSerializer();
        var json = serializer.Serialize(payload);
        File.WriteAllText(path, json);
    }

    private static int ParseSlotKey(string? key)
    {
        if (string.IsNullOrWhiteSpace(key))
        {
            return int.MaxValue;
        }

        if (int.TryParse(key, out var index))
        {
            return index;
        }

        var digits = new string(key.Where(char.IsDigit).ToArray());
        return int.TryParse(digits, out index) ? index : int.MaxValue;
    }
}
