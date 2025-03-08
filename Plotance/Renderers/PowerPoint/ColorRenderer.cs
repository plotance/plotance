// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Numerics;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using Plotance.Models;

using D = DocumentFormat.OpenXml.Drawing;

namespace Plotance.Renderers.PowerPoint;

/// <summary>
/// Provides methods for converting color model values to OpenXml color
/// elements.
/// </summary>
public static class ColorRenderer
{
    /// <summary>
    /// Mapping of scheme color names to their corresponding OpenXml
    /// SchemeColorValues.
    /// </summary>
    private static
        IReadOnlyDictionary<string, D.SchemeColorValues>
        s_schemeColors
        = new Dictionary<string, D.SchemeColorValues>()
        {
            { "dark1", D.SchemeColorValues.Dark1 },
            { "light1", D.SchemeColorValues.Light1 },
            { "accent1", D.SchemeColorValues.Accent1 },
            { "accent2", D.SchemeColorValues.Accent2 },
            { "accent3", D.SchemeColorValues.Accent3 },
            { "accent4", D.SchemeColorValues.Accent4 },
            { "accent5", D.SchemeColorValues.Accent5 },
            { "accent6", D.SchemeColorValues.Accent6 },
            { "hyperlink", D.SchemeColorValues.Hyperlink },
            { "followed_hyperlink", D.SchemeColorValues.FollowedHyperlink }
        };

    /// <summary>Parses a color string to an OpenXml color element.</summary>
    /// <param name="color">
    /// The color string to parse. Can be either a theme color name or a hex
    /// color code.
    /// </param>
    /// <returns>
    /// An OpenXml composite element (e.g. SchemeColor or RgbColorModelHex)
    /// representing the color, or null if the input is null.
    /// </returns>
    /// <remarks>
    /// Theme colors are defined in s_schemeColors. Hex color codes can be
    /// either 3 or 6 digits, prefixed with #.
    /// </remarks>
    /// <exception cref="PlotanceException">
    /// Thrown when the color string is invalid.
    /// </exception>
    [return: NotNullIfNotNull(nameof(color))]
    public static OpenXmlCompositeElement? ParseColor(Color? color)
    {
        if (color == null)
        {
            return null;
        }
        else if (
            s_schemeColors.TryGetValue(
                color.Text.ToLowerInvariant(),
                out var schemeColor
            )
        )
        {
            return new D.SchemeColor() { Val = schemeColor };
        }
        else
        {
            var match = Regex.Match(
                color.Text,
                @"^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$"
            );

            if (match.Success)
            {
                var hex = match.Groups[1].Value;

                if (hex.Length == 3)
                {
                    hex = new string([
                        hex[0],
                        hex[0],
                        hex[1],
                        hex[1],
                        hex[2],
                        hex[2]
                    ]);
                }

                return new D.RgbColorModelHex() { Val = hex };
            }
            else
            {
                throw new PlotanceException(
                    color.Path,
                    color.Line,
                    $"Invalid color format: {color.Text}"
                );
            }
        }
    }

    /// <summary>
    /// Parses a list of color strings to a list of OpenXml color elements.
    /// </summary>
    /// <param name="colors">
    /// The list of color strings to parse. Each color can be either a theme
    /// color name or a hex color code.
    /// </param>
    /// <returns>
    /// A list of OpenXml composite elements (e.g. SchemeColor or
    /// RgbColorModelHex) representing the colors, or null if the input is null.
    /// </returns>
    [return: NotNullIfNotNull(nameof(colors))]
    public static IReadOnlyList<OpenXmlCompositeElement>? ParseColors(
        IReadOnlyList<Color>? colors
    ) => colors?.Select(color => ParseColor(color))?.ToList();

    /// <summary>
    /// Parses a list of color strings with source location to a list of OpenXml
    /// color elements.
    /// </summary>
    /// <param name="colors">
    /// The list of color strings with source location to parse. Each color can
    /// be either a theme color name or a hex color code.
    /// </param>
    /// <returns>
    /// A list of OpenXml composite elements (e.g. SchemeColor or
    /// RgbColorModelHex) representing the colors, or null if the input is null.
    /// </returns>
    [return: NotNullIfNotNull(nameof(colors))]
    public static IReadOnlyList<OpenXmlCompositeElement>? ParseColors(
        ValueWithLocation<IReadOnlyList<Color>>? colors
    ) => ParseColors(colors?.Value);
}
