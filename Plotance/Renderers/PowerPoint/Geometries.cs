// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using DocumentFormat.OpenXml.Presentation;
using Plotance.Models;

using D = DocumentFormat.OpenXml.Drawing;

namespace Plotance.Renderers.PowerPoint;

/// <summary>Provides methods for geometry-related operations.</summary>
public static class Geometries
{
    /// <summary>Gets the offset property from a shape.</summary>
    /// <param name="shape">The shape to get the offset from.</param>
    /// <returns>The offset of the shape or null if not available.</returns>
    public static D.Offset? GetOffset(Shape? shape)
        => shape
            ?.ShapeProperties
            ?.Transform2D
            ?.Offset;

    /// <summary>Gets the extents property from a shape.</summary>
    /// <param name="shape">The shape to get the extents from.</param>
    /// <returns>The extents of the shape or null if not available.</returns>
    public static D.Extents? GetExtents(Shape? shape)
        => shape
            ?.ShapeProperties
            ?.Transform2D
            ?.Extents;

    /// <summary>Parses a horizontal alignment value from a string.</summary>
    /// <param name="align">
    /// The alignment string with location information.
    /// </param>
    /// <returns>
    /// The corresponding TextAlignmentTypeValues enum value or null if input is
    /// null.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the alignment value is not recognized.
    /// </exception>
    public static D.TextAlignmentTypeValues? ParseHorizontalAlign(
        ValueWithLocation<string>? align
    ) => align?.Value switch
    {
        "left" => D.TextAlignmentTypeValues.Left,
        "center" => D.TextAlignmentTypeValues.Center,
        "right" => D.TextAlignmentTypeValues.Right,
        "justified" => D.TextAlignmentTypeValues.Justified,
        "distributed" => D.TextAlignmentTypeValues.Distributed,
        null => null,
        _ => throw new PlotanceException(
            align?.Path,
            align?.Line,
            $"Unknown horizontal alignment: {align?.Value}"
        )
    };

    /// <summary>Parses a vertical alignment value from a string.</summary>
    /// <param name="align">
    /// The alignment string with location information.
    /// </param>
    /// <returns>
    /// The corresponding TextAnchoringTypeValues enum value or null if input is
    /// null.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the alignment value is not recognized.
    /// </exception>
    public static D.TextAnchoringTypeValues? ParseVerticalAlign(
        ValueWithLocation<string>? align
    ) => align?.Value switch
    {
        "top" => D.TextAnchoringTypeValues.Top,
        "center" => D.TextAnchoringTypeValues.Center,
        "bottom" => D.TextAnchoringTypeValues.Bottom,
        null => null,
        _ => throw new PlotanceException(
            align?.Path,
            align?.Line,
            $"Unknown vertical alignment: {align?.Value}"
        )
    };
}

/// <summary>Represents a rectangular region with position and size.</summary>
/// <param name="X">The X coordinate of the top-left corner.</param>
/// <param name="Y">The Y coordinate of the top-left corner.</param>
/// <param name="Width">The width of the rectangle.</param>
/// <param name="Height">The height of the rectangle.</param>
public record Rectangle(long X, long Y, long Width, long Height);
