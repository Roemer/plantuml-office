using System.Diagnostics;

var folderPath = Path.GetFullPath(@"..\OfficeSymbols_2014_SVG_Optimized");
var inkScapePath = @"..\Inkscape-0.92.3-x64\inkscape\inkscape.exe";

foreach (var imagePath in Directory.GetFiles(folderPath, "*.svg", SearchOption.AllDirectories))
{
    var entityName = Path.GetFileNameWithoutExtension(imagePath);
    var targetFilePath = Path.Combine(Directory.GetParent(imagePath).FullName, entityName + ".png");
    var process = Process.Start(new ProcessStartInfo
    {
        //WorkingDirectory = Path.GetDirectoryName(inkScapePath),
        FileName = inkScapePath,
        Arguments = $"\"{imagePath}\" --export-png=\"{targetFilePath}\"",
        //RedirectStandardOutput = true,
        //UseShellExecute = false,
        //CreateNoWindow = true
    });
    if (!process.WaitForExit(5000))
    {
        Console.WriteLine("Killing");
        process.Kill();
    }
    Console.WriteLine($"Converted {imagePath}");
}
