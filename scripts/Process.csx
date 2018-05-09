#r "System.Drawing"

using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;

var rootFolder = Path.GetFullPath(@"..\");
var sourceFolder = Path.Combine(rootFolder, "office2014");
var targetMaxSize = 48;
var plantUmlPath = @"plantuml.jar";

Main();

public void Main()
{
    FilterEntities();
    PreProcessImages();
    FixNames();
    GenerateMarkdownTable();
    ConvertToPumls();
    Console.WriteLine("All work done");
}

public void FilterEntities()
{
    foreach (var imagePath in Directory.GetFiles(sourceFolder, "*.*", SearchOption.AllDirectories))
    {
        var entityName = Path.GetFileNameWithoutExtension(imagePath);
        // Delete certain entries
        if (entityName.EndsWith(" - small") || entityName.EndsWith(" - blue") || entityName.EndsWith(" - green") || entityName.EndsWith(" - orange") || entityName.EndsWith(" - red"))
        {
            Console.WriteLine("Deleting " + imagePath);
            File.Delete(imagePath);
            continue;
        }
    }
}

public void PreProcessImages()
{
    foreach (var imagePath in Directory.GetFiles(sourceFolder, "*.*", SearchOption.AllDirectories))
    {
        // Scale the images and make them monchrome
        Image grayImage = null;
        using (var image = Image.FromFile(imagePath))
        {
            using (var newImage = ScaleImage(image, targetMaxSize, targetMaxSize))
            {
                grayImage = MakeGrayscale(newImage);
            }
        }
        grayImage.Save(imagePath, ImageFormat.Png);
    }
}

public void FixNames()
{
    foreach (var imagePath in Directory.GetFiles(sourceFolder, "*.*", SearchOption.AllDirectories))
    {
        var entityName = Path.GetFileNameWithoutExtension(imagePath);
        // Fix the entity name
        var entityNameFixed = entityName.Replace(",", " ").Replace(";", " ").Replace("-", " ").Replace(" ", "_").ToLower();
        entityNameFixed = Regex.Replace(entityNameFixed, "_+", "_");
        var tmpName = imagePath + ".tmp";
        var newPath = Path.Combine(Directory.GetParent(imagePath).FullName, entityNameFixed + Path.GetExtension(imagePath));
        File.Move(imagePath, tmpName);
        File.Move(tmpName, newPath);
        Console.WriteLine("Fixed name: " + entityName + " to " + entityNameFixed);
    }
}

public void GenerateMarkdownTable()
{
    // Create a markdown table with all entries
    var sbTable = new StringBuilder();
    sbTable.AppendLine("Category | Macro | Image | Name");
    sbTable.AppendLine("--- | --- | --- | ---");
    foreach (var imagePath in Directory.GetFiles(sourceFolder, "*.png*", SearchOption.AllDirectories))
    {
        var entityName = Path.GetFileNameWithoutExtension(imagePath);
        var category = Directory.GetParent(imagePath).Name;
        var fileName = Path.GetFileName(imagePath);
        sbTable.AppendLine($"{category} | OFF_{entityName.ToUpper()} | ![{entityName}](/office2014/{category}/{fileName}?raw=true) | {entityName}");
    }
    File.WriteAllText("../table.md", sbTable.ToString());
}

public void ConvertToPumls()
{
    var count = 0;
    var format = "16"; // 16z for compressed
    foreach (var imagePath in Directory.GetFiles(sourceFolder, "*.*", SearchOption.AllDirectories))
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