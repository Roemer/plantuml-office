#r "System.Drawing"

using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;

var rootFolder = Path.GetFullPath(@"..\");
var sourceFolder = Path.Combine(rootFolder, "OfficeSymbols_2014_SVG_Optimized");
var targetMaxSize = 48;
var plantUmlPath = @"plantuml.jar";
var inkScapePath = @"..\Inkscape-0.92.3-x64\inkscape\inkscape.exe";

Main();

public void Main()
{
    foreach (var svgPath in Directory.GetFiles(sourceFolder, "*.svg", SearchOption.AllDirectories))
    {
        Console.WriteLine($"Processing {svgPath}");
        var filePath = FixName(svgPath);
        var pngPath = ConvertToPng(filePath, true);
        var pumlPath = ConvertToPuml(pngPath);
        pngPath = ConvertToPng(filePath, false);
    }
    GenerateMarkdownTable();
    Console.WriteLine("Finished");
}

public string FixName(string svgPath)
{
    var entityName = Path.GetFileNameWithoutExtension(svgPath);
    var entityNameFixed = entityName.Replace(",", " ").Replace(";", " ").Replace("-", " ").Replace(" ", "_").ToLower();
    entityNameFixed = Regex.Replace(entityNameFixed, "_+", "_");
    if (entityName.Equals(entityNameFixed))
    {
        return svgPath;
    }
    var tmpName = svgPath + ".tmp";
    var newPath = Path.Combine(Directory.GetParent(svgPath).FullName, entityNameFixed + Path.GetExtension(svgPath));
    File.Move(svgPath, tmpName);
    File.Move(tmpName, newPath);
    return newPath;
}

public string ConvertToPng(string svgPath, bool withBackground)
{
    // Convert the image
    var backgroundOpacity = withBackground ? "1.0" : "0.0";
    var entityName = Path.GetFileNameWithoutExtension(svgPath);
    var targetFilePath = Path.Combine(Directory.GetParent(svgPath).FullName, entityName + ".png");
    var process = Process.Start(new ProcessStartInfo
    {
        FileName = inkScapePath,
        Arguments = $"-z --file=\"{svgPath}\" --export-png=\"{targetFilePath}\" --export-background=#FFFFFF --export-background-opacity={backgroundOpacity} --export-area-drawing",
        RedirectStandardOutput = true,
        UseShellExecute = false,
        CreateNoWindow = true
    });
    if (!process.WaitForExit(5000))
    {
        Console.WriteLine("Killing");
        process.Kill();
    }
    // Scale the image
    Image newImage = null;
    using (var image = Image.FromFile(targetFilePath))
    {
        newImage = ScaleImage(image, targetMaxSize, targetMaxSize);
    }
    newImage.Save(targetFilePath, ImageFormat.Png);
    newImage.Dispose();
    return targetFilePath;
}

public string ConvertToPuml(string pngPath)
{
    var format = "16"; // 16z for compressed

    var entityName = Path.GetFileNameWithoutExtension(pngPath);
    var entityNameUpper = entityName.ToUpper();
    var pumlPath = Path.Combine(Directory.GetParent(pngPath).FullName, entityName + ".puml");
    var process = Process.Start(new ProcessStartInfo
    {
        WorkingDirectory = rootFolder,
        FileName = "java",
        Arguments = $"-jar {plantUmlPath} -encodesprite {format} \"{pngPath}\"",
        RedirectStandardOutput = true,
        UseShellExecute = false,
        CreateNoWindow = true
    });
    process.WaitForExit();
    var sbImage = new StringBuilder();
    sbImage.Append(process.StandardOutput.ReadToEnd());
    sbImage.AppendLine($"!define OFF_{entityNameUpper}(_alias) ENTITY(rectangle,black,{entityName},_alias,OFF {entityNameUpper})");
    sbImage.AppendLine($"!define OFF_{entityNameUpper}(_alias,_label) ENTITY(rectangle,black,{entityName},_label,_alias,OFF {entityNameUpper})");
    sbImage.AppendLine($"!define OFF_{entityNameUpper}(_alias,_label,_shape) ENTITY(_shape,black,{entityName},_label,_alias,OFF {entityNameUpper})");
    sbImage.AppendLine($"!define OFF_{entityNameUpper}(_alias,_label,_shape,_color) ENTITY(_shape,_color,{entityName},_label,_alias,OFF {entityNameUpper})");

    File.WriteAllText(pumlPath, sbImage.ToString());
    return pumlPath;
}

public void DuplicateAsMonchrome()
{
    foreach (var imagePath in Directory.GetFiles(sourceFolder, "*.png", SearchOption.AllDirectories))
    {
        Image grayImage = null;
        using (var image = Image.FromFile(imagePath))
        {
            grayImage = MakeGrayscale(image);
        }
        var grayImagePath = Path.Combine(Path.GetDirectoryName(imagePath), String.Concat(Path.GetFileNameWithoutExtension(imagePath), "(m)", Path.GetExtension(imagePath)));
        grayImage.Save(grayImagePath, ImageFormat.Png);
        grayImage.Dispose();
    }
}

public void GenerateMarkdownTable()
{
    // Create a markdown table with all entries
    var sbTable = new StringBuilder();
    sbTable.AppendLine("Category | Macro | Image | Url");
    sbTable.AppendLine("--- | --- | --- | ---");
    foreach (var filePath in Directory.GetFiles(sourceFolder, "*.puml", SearchOption.AllDirectories))
    {
        var entityName = Path.GetFileNameWithoutExtension(filePath);
        var category = Directory.GetParent(filePath).Name;
        sbTable.AppendLine($"{category} | OFF_{entityName.ToUpper()} | ![{entityName}](/office2014/{category}/{entityName}.png?raw=true) | {category}/{entityName}.puml");
    }
    File.WriteAllText("../table.md", sbTable.ToString());
}

private static Image ScaleImage(Image image, int maxWidth, int maxHeight)
{
    var ratioX = (double)maxWidth / image.Width;
    var ratioY = (double)maxHeight / image.Height;
    var ratio = Math.Min(ratioX, ratioY);

    var newWidth = (int)(image.Width * ratio);
    var newHeight = (int)(image.Height * ratio);

    var newImage = new Bitmap(newWidth, newHeight);

    using (var graphics = Graphics.FromImage(newImage))
        graphics.DrawImage(image, 0, 0, newWidth, newHeight);

    return newImage;
}

private static Image MakeGrayscale(Image original)
{
    //create a blank bitmap the same size as original
    var newBitmap = new Bitmap(original.Width, original.Height);

    //get a graphics object from the new image
    var g = Graphics.FromImage(newBitmap);

    //create the grayscale ColorMatrix
    ColorMatrix colorMatrix = new ColorMatrix(
       new float[][]
       {
         new float[] {.3f, .3f, .3f, 0, 0},
         new float[] {.59f, .59f, .59f, 0, 0},
         new float[] {.11f, .11f, .11f, 0, 0},
         new float[] {0, 0, 0, 1, 0},
         new float[] {0, 0, 0, 0, 1}
       });

    //create some image attributes
    ImageAttributes attributes = new ImageAttributes();

    //set the color matrix attribute
    attributes.SetColorMatrix(colorMatrix);

    //draw the original image on the new image
    //using the grayscale color matrix
    g.DrawImage(original, new Rectangle(0, 0, original.Width, original.Height),
       0, 0, original.Width, original.Height, GraphicsUnit.Pixel, attributes);

    //dispose the Graphics object
    g.Dispose();
    return newBitmap;
}