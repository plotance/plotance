// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Drawing;
using System.Globalization;
using System.Net;
using System.Text.RegularExpressions;
using System.Xml;

namespace Plotance.Image;

/// <summary>
/// Provides utility methods for extracting size information from various image
/// formats.
/// </summary>
public static class Images
{
    /// <summary>
    /// Extracts the width and height of a PNG image from its header.
    /// </summary>
    /// <param name="stream">
    /// The stream containing PNG image data. The stream position must be at the
    /// beginning of the PNG data.
    /// </param>
    /// <returns>
    /// A <see cref="Size"/> structure containing the width and height of the
    /// image.
    /// </returns>
    /// <exception cref="InvalidDataException">
    /// Thrown when the stream does not contain valid PNG data.
    /// </exception>
    public static Size GetPngImageSize(Stream stream)
    {
        // http://www.libpng.org/pub/png/spec/1.2/PNG-Structure.html

        // 8 bytes: header 89 50 4e 47 0d 0a 1a 0a
        // 4 bytes: length
        // 4 bytes: chunk type "IHDR"
        // 4 bytes: width
        // 4 bytes: height

        byte[] buffer = new byte[24];
        byte[] signature = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
        byte[] chunkType = [0x49, 0x48, 0x44, 0x52];

        try
        {
            stream.ReadExactly(buffer, 0, 24);

            if (!buffer.AsSpan(0, 8).SequenceEqual(signature))
            {
                throw new InvalidDataException("Invalid PNG format");
            }

            if (!buffer.AsSpan(12, 4).SequenceEqual(chunkType))
            {
                throw new InvalidDataException("Invalid PNG format");
            }

            return new Size(
                IPAddress.NetworkToHostOrder(
                    BitConverter.ToInt32(buffer, 16)
                ),
                IPAddress.NetworkToHostOrder(
                    BitConverter.ToInt32(buffer, 20)
                )
            );
        }
        catch (IOException e)
        {
            throw new InvalidDataException("Invalid PNG format", e);
        }
    }

    /// <summary>
    /// Extracts the width and height of a JPEG image by parsing its segments.
    /// </summary>
    /// <param name="stream">
    /// The stream containing JPEG image data. The stream position must be at
    /// the beginning of the JPEG data.
    /// </param>
    /// <returns>
    /// A <see cref="Size"/> structure containing the width and height of the
    /// image.
    /// </returns>
    /// <exception cref="InvalidDataException">
    /// Thrown when the stream does not contain valid JPEG data or uses an
    /// unsupported JPEG format.
    /// </exception>
    public static Size GetJpegImageSize(Stream stream)
    {
        // https://www.w3.org/Graphics/JPEG/itu-t81.pdf

        var buffer = new byte[65537]; // marker + 64 KiB - 1 data

        while (true)
        {
            var size = ReadJpegSegment(stream, buffer);

            if (0xc0 <= buffer[1]
                && buffer[1] < 0xd0
                && buffer[1] != 0xc4
                && buffer[1] != 0xc8
                && buffer[1] != 0xcc
                )
            {
                if (size < 9)
                {
                    throw new InvalidDataException("Invalid PNG file");
                }

                // 2 bytes: marker
                // 2 bytes: frame length (including the length itself)
                // 1 byte: sample precision
                // 2 bytes: number of lines
                // 2 bytes: number of samples per line
                // ...

                var width = IPAddress.NetworkToHostOrder(
                    BitConverter.ToInt16(buffer, 7)
                );
                var height = IPAddress.NetworkToHostOrder(
                    BitConverter.ToInt16(buffer, 5)
                );

                if (height == 0)
                {
                    // When height == 0, it is defined by the DNL header after
                    // a scan.  We don't support it for now.
                    throw new InvalidDataException("Unsupported JPEG format");
                }

                return new Size(width, height);
            }
        }
    }

    /// <summary>
    /// Reads a JPEG segment from the stream into the provided buffer.
    /// </summary>
    /// <param name="stream">
    /// The stream containing JPEG image data, positioned at the start of a
    /// segment.
    /// </param>
    /// <param name="buffer">
    /// The buffer to read the segment data into. Must be large enough to hold
    /// the entire segment.
    /// </param>
    /// <returns>The total number of bytes read for the segment.</returns>
    /// <exception cref="InvalidDataException">
    /// Thrown when the segment is invalid or unsupported.
    /// </exception>
    private static int ReadJpegSegment(Stream stream, byte[] buffer)
    {
        try
        {
            if (stream.ReadByte() != 0xff)
            {
                throw new InvalidDataException("Invalid JPEG format");
            }

            buffer[0] = 0xff;

            var marker = stream.ReadByte();

            // C0-C3: start of frame 0-3
            // C4   : define huffman tables
            // C5-C7: start of frame 5-7
            // C8   : define arithmetic coding conditionings
            // C9-CB: start of frame 9-12
            // CC   : JPEG extension
            // CD-CF: start of frame 13-15
            // D0-D7: restart markers (standalone)
            // D8   : start of image (standalone)
            // D9   : end of image (standalone)
            // DA   : start of scan
            // DB   : define quantization tables
            // DC   : define number of lines
            // DD   : define restart interval
            // DE   : define hierachical progression
            // DF   : expand reference components
            // E0-EF: application segments
            // F0-FD: JPEG extensions
            // FE   : comment

            if (marker < 0xc0 || marker == 0xff)
            {
                throw new InvalidDataException("Unsupported JPEG segment");
            }

            buffer[1] = (byte)marker;

            if (0xd0 <= marker && marker < 0xda)
            {
                return 2;
            }

            stream.ReadExactly(buffer, 2, 2);

            var length = IPAddress.NetworkToHostOrder(
                BitConverter.ToInt16(buffer, 2)
            );

            if (length < 2)
            {
                throw new InvalidDataException("Invalid JPEG format");
            }

            stream.ReadExactly(buffer, 4, length - 2);

            return 2 + length;
        }
        catch (IOException e)
        {
            throw new InvalidDataException("Invalid JPEG format", e);
        }
    }

    /// <summary>
    /// Extracts the width and height of an SVG image by parsing its XML
    /// attributes.
    /// </summary>
    /// <param name="stream">
    /// The stream containing SVG image data. The stream position must be at the
    /// beginning of the SVG data.
    /// </param>
    /// <returns>
    /// A <see cref="Size"/> structure containing the width and height of the
    /// image.
    /// </returns>
    /// <exception cref="InvalidDataException">
    /// Thrown when the SVG lacks required width/height information or uses an
    /// unsupported format.
    /// </exception>
    public static Size GetSvgImageSize(Stream stream)
    {
        // https://www.w3.org/TR/SVG11/coords.html#ViewportSpace
        try
        {
            using var reader = XmlReader.Create(stream);

            string? widthAttr = null;
            string? heightAttr = null;
            string? viewBoxAttr = null;

            // Find the root svg element and extract attributes
            while (reader.Read())
            {
                if (
                    reader.NodeType == XmlNodeType.Element
                        && reader.Name == "svg"
                )
                {
                    widthAttr = reader.GetAttribute("width");
                    heightAttr = reader.GetAttribute("height");
                    viewBoxAttr = reader.GetAttribute("viewBox");
                    break;
                }
            }

            int width, height;

            if (!string.IsNullOrEmpty(widthAttr))
            {
                width = ParseLength(widthAttr);
            }
            else if (!string.IsNullOrEmpty(viewBoxAttr))
            {
                var viewBox = ParseViewBox(viewBoxAttr);

                width = (int)Single.Round(viewBox.Width);
            }
            else
            {
                throw new InvalidDataException(
                    "Missing width or viewBox attribute"
                );
            }

            if (!string.IsNullOrEmpty(heightAttr))
            {
                height = ParseLength(heightAttr);
            }
            else if (!string.IsNullOrEmpty(viewBoxAttr))
            {
                var viewBox = ParseViewBox(viewBoxAttr);

                height = (int)Single.Round(viewBox.Height);
            }
            else
            {
                throw new InvalidDataException(
                    "Missing height or viewBox attribute"
                );
            }

            return new Size(width, height);
        }
        catch (IOException e)
        {
            throw new InvalidDataException("Invalid SVG format", e);
        }
        catch (XmlException e)
        {
            throw new InvalidDataException("Invalid SVG format", e);
        }
    }

    /// <summary>
    /// Parses an SVG viewBox attribute into a rectangle.
    /// </summary>
    /// <param name="viewBox">
    /// The viewBox attribute string in the format "min-x min-y width height".
    /// </param>
    /// <returns>
    /// A <see cref="RectangleF"/> representing the viewBox dimensions.
    /// </returns>
    /// <exception cref="InvalidDataException">
    /// Thrown when the viewBox string does not contain valid values.
    /// </exception>
    private static RectangleF ParseViewBox(string viewBox)
    {
        var numberPattern
            = @"[+-]?(?:[0-9]*\.[0-9]+|[0-9]+)(?:[Ee][+-]?[0-9]+)?";
        // https://www.w3.org/TR/SVG11/types.html#BasicDataTypes
        // comma-wsp  ::= (wsp+ ","? wsp*) | ("," wsp*)
        var separatorPattern
            = @"(?:(?>[ \t\f\r\n]+),?[ \t\f\r\n]*|,[ \t\f\r\n]*)";
        var pattern = $"""
                        ^[ \t\f\r\n]*
                        ({numberPattern})
                        {separatorPattern}
                        ({numberPattern})
                        {separatorPattern}
                        ({numberPattern})
                        {separatorPattern}
                        ({numberPattern})
                        [ \t\f\r\n]*$
                      """;
        var match = Regex.Match(
            viewBox,
            pattern,
            RegexOptions.CultureInvariant | RegexOptions.IgnorePatternWhitespace
        );

        if (!match.Success)
        {
            throw new InvalidDataException(
                $"Invalid viewBox format: {viewBox}"
            );
        }

        try
        {
            float ParseFloat(string text)
                => float.Parse(text, CultureInfo.InvariantCulture);

            return new RectangleF(
                ParseFloat(match.Groups[1].Value),
                ParseFloat(match.Groups[2].Value),
                ParseFloat(match.Groups[3].Value),
                ParseFloat(match.Groups[4].Value)
            );
        }
        catch (FormatException)
        {
            throw new InvalidDataException($"Invalid viewBox value: {viewBox}");
        }
    }

    /// <summary>
    /// Parses a length string from an SVG attribute into an integer pixel
    /// value.
    /// </summary>
    /// <param name="length">
    /// The length string to parse, which may include units (px, cm, mm, in, pt,
    /// pc, em, ex, %).
    /// </param>
    /// <returns>The length in pixels.</returns>
    /// <exception cref="InvalidDataException">
    /// Thrown when the length string cannot be parsed.
    /// </exception>
    /// <exception cref="FormatException">
    /// Thrown when the length string is invalid.
    /// </exception>
    private static int ParseLength(string length)
    {
        length = length.Trim();

        if (string.IsNullOrEmpty(length))
        {
            throw new FormatException($"Invalid length format: {length}");
        }

        var units = "mm|cm|in|pt|pc|px";
        var pattern = @$"^([0-9]*\.[0-9]+|[0-9]+)[ \t]*({units})?$";
        var match = Regex.Match(length, pattern);

        if (!match.Success)
        {
            throw new FormatException($"Invalid length format: {length}");
        }

        var value = decimal.Parse(
            match.Groups[1].Value,
            CultureInfo.InvariantCulture
        );
        var unit = match.Groups[2].Value;

        // Convert to EMU (English Metric Unit)
        return unit switch
        {
            "px" or "" => (int)decimal.Round(value),
            "in" => (int)decimal.Round(96 * value), // 96 px
            "pt" => (int)decimal.Round(4 * value / 3), // 1/72 in
            "pc" => (int)decimal.Round(16 * value), // 12 pt
            "mm" => (int)decimal.Round(96 * value / 25.4m), // 1/25.4 in
            "cm" => (int)decimal.Round(96 * value / 2.54m), // 1/2.54 in
            _ => throw new FormatException($"Unknown unit: {unit}")
        };
    }
}
