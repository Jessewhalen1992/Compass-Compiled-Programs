namespace Compass.Models;

public class AppSettings
{
    public string BaseBlockFolder { get; set; } = string.Empty;
    public string[] DefaultLayerNames { get; set; } = System.Array.Empty<string>();
    public string JsonConfigName { get; set; } = "drillProps.json";
}
