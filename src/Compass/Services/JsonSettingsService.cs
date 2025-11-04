using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Compass.Models;
using System.Web.Script.Serialization;

namespace Compass.Services;

public class JsonSettingsService
{
    private readonly AppSettings _appSettings;
    private static readonly JavaScriptSerializer Serializer = new();

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
        var state = Serializer.Deserialize<DrillGridState>(json) ?? new DrillGridState();
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
            return Path.Combine(directory ?? string.Empty, $"{prefix}.json");
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
        var json = Serializer.Serialize(value);
        return FormatJson(json);
    }

    private static string FormatJson(string json)
    {
        if (string.IsNullOrEmpty(json))
        {
            return json;
        }

        var builder = new StringBuilder();
        var indent = 0;
        var inQuotes = false;
        var escape = false;

        foreach (var ch in json)
        {
            if (escape)
            {
                escape = false;
                builder.Append(ch);
                continue;
            }

            switch (ch)
            {
                case '\\':
                    escape = true;
                    builder.Append(ch);
                    break;
                case '"':
                    inQuotes = !inQuotes;
                    builder.Append(ch);
                    break;
                case '{':
                case '[':
                    builder.Append(ch);
                    if (!inQuotes)
                    {
                        builder.AppendLine();
                        indent++;
                        AppendIndent(builder, indent);
                    }
                    break;
                case '}':
                case ']':
                    if (!inQuotes)
                    {
                        TrimTrailingWhitespace(builder);
                        if (!EndsWithNewLine(builder))
                        {
                            builder.AppendLine();
                        }
                        indent = Math.Max(indent - 1, 0);
                        AppendIndent(builder, indent);
                    }
                    builder.Append(ch);
                    break;
                case ',':
                    builder.Append(ch);
                    if (!inQuotes)
                    {
                        builder.AppendLine();
                        AppendIndent(builder, indent);
                    }
                    break;
                case ':':
                    builder.Append(ch);
                    if (!inQuotes)
                    {
                        builder.Append(' ');
                    }
                    break;
                default:
                    if (!inQuotes && char.IsWhiteSpace(ch))
                    {
                        break;
                    }

                    builder.Append(ch);
                    break;
            }
        }

        return builder.ToString();
    }

    private static void AppendIndent(StringBuilder builder, int indent)
    {
        if (indent > 0)
        {
            builder.Append(' ', indent * 4);
        }
    }

    private static bool EndsWithNewLine(StringBuilder builder)
    {
        if (builder.Length == 0)
        {
            return false;
        }

        var last = builder[builder.Length - 1];
        return last == '\n' || last == '\r';
    }

    private static void TrimTrailingWhitespace(StringBuilder builder)
    {
        while (builder.Length > 0)
        {
            var last = builder[builder.Length - 1];
            if (last == '\n' || last == '\r' || !char.IsWhiteSpace(last))
            {
                break;
            }

            builder.Length--;
        }
    }
}
