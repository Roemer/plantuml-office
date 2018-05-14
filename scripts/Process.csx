#r "System.Drawing"

using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;

var rootFolder = Path.GetFullPath(@"..\");
var sourceFolder = Path.Combine(rootFolder, "OfficeSymbols_2014_PNG");
var targetMaxSize = 48;
var plantUmlPath = @"plantuml.jar";

Main();

public void Main()
{
    FixNames();
    ScaleImages();
    ConvertToPumls();
    DuplicateAsMonchrome();
    GenerateMarkdownTable();
    Console.WriteLine("All work done");
}

public void FixNames()
{
    foreach (var imagePath in Directory.GetFiles(sourceFolder, "*.*", SearchOption.AllDirectories))
    {
        var entityName = Path.GetFileNameWithoutExtension(imagePath);
        var entityNameFixed = entityName.Replace(",", " ").Replace(";", " ").Replace("-", " ").Replace(" ", "_").ToLower();
        entityNameFixed = Regex.Replace(entityNameFixed, "_+", "_");
        var tmpName = imagePath + ".tmp";
        var newPath = Path.Combine(Directory.GetParent(imagePath).FullName, entityNameFixed + Path.GetExtension(imagePath));
        File.Move(imagePath, tmpName);
        File.Move(tmpName, newPath);
        Console.WriteLine("Fixed name: " + entityName + " to " + entityNameFixed);
    }
}

public void ScaleImages()
{
    foreach (var imagePath in Directory.GetFiles(sourceFolder, "*.png", SearchOption.AllDirectories))
    {
        Image newImage = null;
        using (var image = Image.FromFile(imagePath))
        {
            newImage = ScaleImage(image, targetMaxSize, targetMaxSize);
        }
        newImage.Save(imagePath, ImageFormat.Png);
        newImage.Dispose();
    }
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

public void ConvertToPumls()
{
    var count = 0;
    var format = "16"; // 16z for compressed
    foreach (var imagePath in Directory.GetFiles(sourceFolder, "*.png", SearchOption.AllDirectories))
    {
        var entityName = Path.GetFileNameWithoutExtension(imagePath);
        var entityNameUpper = entityName.ToUpper();
        var pumlFileName = Path.Combine(Directory.GetParent(imagePath).FullName, entityName + ".puml");
        var process = Process.Start(new ProcessStartInfo
        {
            WorkingDirectory = rootFolder,
            FileName = "java",
            Arguments = $"-jar {plantUmlPath} -encodesprite {format} \"{imagePath}\"",
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
        sbImage.Append($"skinparam folderBackgroundColor<<OFF {entityNameUpper}>> White");

        File.WriteAllText(pumlFileName, sbImage.ToString());

        if (++count % 20 == 0)
        {
            Console.WriteLine($"Processed {count} image(s)");
        }
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