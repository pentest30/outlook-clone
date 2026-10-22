// Draws the app icon (blue rounded square, white envelope) and writes:
//   app.png  256 px, for the window, tray and the installer
//   app.ico  16, 24, 32, 48, 64, 128, 256 px, PNG-compressed entries (Windows Vista+ ICO format)
//
//   dotnet run tools/IconGen/IconGen.cs -- src/Outlook.App/Assets
#:package SkiaSharp@3.119.1
#:property PublishAot=false
#:property ManagePackageVersionsCentrally=false

using SkiaSharp;

var outDir = args.Length > 0 ? args[0] : ".";
Directory.CreateDirectory(outDir);
int[] sizes = [16, 24, 32, 48, 64, 128, 256];
var pngs = sizes.ToDictionary(s => s, Draw);
File.WriteAllBytes(Path.Combine(outDir, "app.png"), pngs[256]);
File.WriteAllBytes(Path.Combine(outDir, "app.ico"), Ico(sizes.Select(s => (s, pngs[s])).ToList()));
Console.WriteLine($"wrote app.png and app.ico ({sizes.Length} sizes) to {Path.GetFullPath(outDir)}");

static byte[] Draw(int size)
{
    using var surface = SKSurface.Create(new SKImageInfo(size, size, SKColorType.Rgba8888, SKAlphaType.Premul));
    var c = surface.Canvas;
    c.Clear(SKColors.Transparent);
    float s = size;

    // Background: Outlook-ish blue rounded square with a subtle vertical gradient.
    using (var bg = new SKPaint { IsAntialias = true })
    {
        bg.Shader = SKShader.CreateLinearGradient(new SKPoint(0, 0), new SKPoint(0, s),
            [new SKColor(0x1A, 0x7F, 0xD4), new SKColor(0x0F, 0x5C, 0xA8)], SKShaderTileMode.Clamp);
        c.DrawRoundRect(new SKRect(0, 0, s, s), s * 0.2f, s * 0.2f, bg);
    }

    // Envelope body.
    var left = s * 0.18f; var right = s * 0.82f; var top = s * 0.30f; var bottom = s * 0.72f;
    using var white = new SKPaint { IsAntialias = true, Color = SKColors.White };
    c.DrawRoundRect(new SKRect(left, top, right, bottom), s * 0.05f, s * 0.05f, white);

    // Flap: a blue V across the top of the envelope (thicker at small sizes so it stays visible).
    using var flap = new SKPaint
    {
        IsAntialias = true, Color = new SKColor(0x0F, 0x5C, 0xA8), Style = SKPaintStyle.Stroke,
        StrokeWidth = Math.Max(1.5f, s * 0.055f), StrokeJoin = SKStrokeJoin.Round, StrokeCap = SKStrokeCap.Round,
    };
    using var path = new SKPath();
    path.MoveTo(left + s * 0.03f, top + s * 0.04f);
    path.LineTo(s * 0.5f, top + (bottom - top) * 0.55f);
    path.LineTo(right - s * 0.03f, top + s * 0.04f);
    c.DrawPath(path, flap);

    using var image = surface.Snapshot();
    using var data = image.Encode(SKEncodedImageFormat.Png, 100);
    return data.ToArray();
}

// ICONDIR + ICONDIRENTRY per image, followed by the PNG blobs.
static byte[] Ico(List<(int Size, byte[] Png)> images)
{
    using var ms = new MemoryStream();
    using var w = new BinaryWriter(ms);
    w.Write((ushort)0); w.Write((ushort)1); w.Write((ushort)images.Count);
    var offset = 6 + 16 * images.Count;
    foreach (var (size, png) in images)
    {
        w.Write((byte)(size >= 256 ? 0 : size)); w.Write((byte)(size >= 256 ? 0 : size));
        w.Write((byte)0); w.Write((byte)0);
        w.Write((ushort)1); w.Write((ushort)32);
        w.Write(png.Length); w.Write(offset);
        offset += png.Length;
    }
    foreach (var (_, png) in images) w.Write(png);
    return ms.ToArray();
}
