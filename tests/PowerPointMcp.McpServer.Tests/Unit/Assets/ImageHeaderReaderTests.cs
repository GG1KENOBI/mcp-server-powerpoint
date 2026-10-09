// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Buffers.Binary;
using System.Text;
using Sbroenne.PowerPointMcp.Core.Assets;

namespace Sbroenne.PowerPointMcp.McpServer.Tests.Unit.Assets;

/// <summary>Image header parsing for every supported format, built from synthetic headers.</summary>
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Assets")]
public sealed class ImageHeaderReaderTests
{
    internal static byte[] Png(int width, int height)
    {
        var data = new byte[33];
        new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, (byte)'I', (byte)'H', (byte)'D', (byte)'R' }.CopyTo(data, 0);
        BinaryPrimitives.WriteInt32BigEndian(data.AsSpan(16), width);
        BinaryPrimitives.WriteInt32BigEndian(data.AsSpan(20), height);
        return data;
    }

    internal static byte[] Jpeg(int width, int height, int? orientation = null)
    {
        var bytes = new List<byte> { 0xFF, 0xD8 };
        if (orientation is { } value)
        {
            // APP1 Exif with one IFD entry (orientation), little endian.
            var exif = new List<byte>();
            exif.AddRange("Exif\0\0"u8.ToArray());
            exif.AddRange(new byte[] { (byte)'I', (byte)'I', 42, 0, 8, 0, 0, 0, 1, 0, 0x12, 0x01, 3, 0, 1, 0, 0, 0, (byte)value, 0, 0, 0, 0, 0, 0, 0 });
            int length = exif.Count + 2;
            bytes.AddRange(new byte[] { 0xFF, 0xE1, (byte)(length >> 8), (byte)length });
            bytes.AddRange(exif);
        }
        bytes.AddRange(new byte[] { 0xFF, 0xDB, 0, 4, 0, 0 }); // DQT stub
        bytes.AddRange(new byte[] { 0xFF, 0xC0, 0, 11, 8, (byte)(height >> 8), (byte)height, (byte)(width >> 8), (byte)width, 3, 0, 0, 0 });
        bytes.AddRange(new byte[] { 0xFF, 0xDA });
        return bytes.ToArray();
    }

    [Fact]
    public void Png_ReadsIhdr() => Assert.Equal(("png", 1920, 1080), Info(Png(1920, 1080)));

    [Fact]
    public void Jpeg_ReadsSofAfterOtherSegments() => Assert.Equal(("jpeg", 640, 480), Info(Jpeg(640, 480)));

    [Fact]
    public void Jpeg_ExifRotationSwapsDisplaySize()
    {
        var info = ImageHeaderReader.Read(Jpeg(4000, 3000, orientation: 6))!;
        Assert.Equal((6, 3000, 4000), (info.ExifOrientation!.Value, info.DisplayWidth, info.DisplayHeight));
        Assert.Equal(0.75, info.Aspect, 3);
    }

    [Fact]
    public void Gif()
    {
        var data = "GIF89a"u8.ToArray().Concat(new byte[] { 0x40, 0x01, 0xF0, 0x00 }).ToArray();
        Assert.Equal(("gif", 320, 240), Info(data));
    }

    [Fact]
    public void Bmp()
    {
        var data = new byte[30];
        data[0] = (byte)'B';
        data[1] = (byte)'M';
        BinaryPrimitives.WriteInt32LittleEndian(data.AsSpan(18), 800);
        BinaryPrimitives.WriteInt32LittleEndian(data.AsSpan(22), -600); // top-down bitmap
        Assert.Equal(("bmp", 800, 600), Info(data));
    }

    [Fact]
    public void WebP_Vp8X()
    {
        var data = new byte[30];
        "RIFF"u8.CopyTo(data);
        "WEBPVP8X"u8.CopyTo(data.AsSpan(8));
        data[24] = 0xFF; data[25] = 0x03; // width-1 = 1023
        data[27] = 0xFF; data[28] = 0x01; // height-1 = 511
        Assert.Equal(("webp", 1024, 512), Info(data));
    }

    [Fact]
    public void Tiff_LittleEndian()
    {
        var data = new byte[8 + 2 + (2 * 12)];
        new byte[] { (byte)'I', (byte)'I', 42, 0, 8, 0, 0, 0 }.CopyTo(data, 0);
        BinaryPrimitives.WriteUInt16LittleEndian(data.AsSpan(8), 2);
        WriteEntry(10, 256, 1200);
        WriteEntry(22, 257, 900);
        Assert.Equal(("tiff", 1200, 900), Info(data));

        void WriteEntry(int offset, ushort tag, ushort value)
        {
            BinaryPrimitives.WriteUInt16LittleEndian(data.AsSpan(offset), tag);
            BinaryPrimitives.WriteUInt16LittleEndian(data.AsSpan(offset + 2), 3);
            BinaryPrimitives.WriteUInt32LittleEndian(data.AsSpan(offset + 4), 1);
            BinaryPrimitives.WriteUInt16LittleEndian(data.AsSpan(offset + 8), value);
        }
    }

    [Theory]
    [InlineData("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"200\" height=\"100\">", 200, 100)]
    [InlineData("<?xml version=\"1.0\"?><svg viewBox=\"0 0 64 32\" xmlns=\"http://www.w3.org/2000/svg\">", 64, 32)]
    [InlineData("<svg width=\"2in\" viewBox=\"0 0 100 50\">", 192, 96)]
    [InlineData("<svg width='10cm' height='5cm'>", 378, 189)]
    public void Svg(string header, int width, int height)
    {
        var info = ImageHeaderReader.Read(Encoding.UTF8.GetBytes(header + "</svg>"))!;
        Assert.Equal(("svg", width, height, true), (info.Format, info.Width, info.Height, info.IsVector));
    }

    [Fact]
    public void Emf()
    {
        var data = new byte[88];
        BinaryPrimitives.WriteUInt32LittleEndian(data, 1);
        BinaryPrimitives.WriteInt32LittleEndian(data.AsSpan(16), 299);
        BinaryPrimitives.WriteInt32LittleEndian(data.AsSpan(20), 149);
        " EMF"u8.CopyTo(data.AsSpan(40));
        Assert.Equal(("emf", 300, 150), Info(data));
    }

    [Theory]
    [InlineData(new byte[0])]
    [InlineData(new byte[] { 1, 2, 3, 4, 5 })]
    [InlineData(new byte[] { 0xFF, 0xD8, 0xFF })]
    public void UnknownOrTruncated_ReturnsNull(byte[] data)
    {
        Assert.Null(ImageHeaderReader.Read(data));
    }

    [Fact]
    public void ReadFile_HandlesMissingFilesAndUnicodePaths()
    {
        Assert.Null(ImageHeaderReader.ReadFile(Path.Combine(Path.GetTempPath(), "does-not-exist.png")));
        var directory = Directory.CreateTempSubdirectory("pptmcp-изображения-");
        try
        {
            var path = Path.Combine(directory.FullName, "график.png");
            File.WriteAllBytes(path, Png(10, 20));
            Assert.Equal(20, ImageHeaderReader.ReadFile(path)!.Height);
        }
        finally
        {
            directory.Delete(recursive: true);
        }
    }

    private static (string, int, int) Info(byte[] data)
    {
        var info = ImageHeaderReader.Read(data);
        Assert.NotNull(info);
        return (info.Format, info.Width, info.Height);
    }
}
