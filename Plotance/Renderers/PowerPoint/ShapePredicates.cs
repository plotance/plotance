// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using DocumentFormat.OpenXml.Presentation;

namespace Plotance.Renderers.PowerPoint;

/// <summary>
/// Provides predicate functions for matching shapes in PowerPoint
/// presentations.
/// </summary>
public static class ShapePredicates
{
    /// <summary>
    /// Creates a predicate that matches shapes with the same placeholder index
    /// or placeholder type as the given shape.
    /// </summary>
    /// <param name="shape">
    /// The shape to extract placeholder information from.
    /// </param>
    /// <returns>
    /// A predicate function that returns true for shapes with matching
    /// placeholder index or placeholder type.
    /// </returns>
    public static Func<Shape, bool> HaveMatchingPlaceholder(Shape shape)
    {
        var placeholderShape = SlideLayouts.ExtractPlaceholderShape(shape);
        var index = placeholderShape?.Index?.Value;
        var placeholderType = placeholderShape?.Type?.Value;

        if (index is uint i)
        {
            return HavePlaceholderIndex(i);
        }

        if (placeholderType is PlaceholderValues t)
        {
            return IsPlaceholderOfType(t);
        }

        return False;
    }

    /// <summary>
    /// Creates a predicate that matches shapes with the specified placeholder
    /// index.
    /// </summary>
    /// <param name="index">The placeholder index to match.</param>
    /// <returns>
    /// A predicate function that returns true for shapes with the specified
    /// placeholder index.
    /// </returns>
    public static Func<Shape, bool> HavePlaceholderIndex(uint index)
        => shape => shape
        .NonVisualShapeProperties
        ?.GetFirstChild<ApplicationNonVisualDrawingProperties>()
        ?.PlaceholderShape
        ?.Index
        ?.Value
        == index;

    /// <summary>
    /// Creates a predicate that matches shapes with the specified placeholder
    /// type.
    /// </summary>
    /// <param name="placeholderType">The placeholder type to match.</param>
    /// <returns>
    /// A predicate function that returns true for shapes with the specified
    /// placeholder type.
    /// </returns>
    public static Func<Shape, bool> IsPlaceholderOfType(
        PlaceholderValues placeholderType
    ) => shape => shape
        .NonVisualShapeProperties
        ?.GetFirstChild<ApplicationNonVisualDrawingProperties>()
        ?.PlaceholderShape
        ?.Type
        ?.Value
        == placeholderType;

    /// <summary>
    /// Creates a predicate that matches shapes that match any of the specified
    /// predicates.
    /// </summary>
    /// <param name="predicates">
    /// The predicates to combine with OR logic.</param>
    /// <returns>
    /// A predicate function that returns true if any of the specified
    /// predicates return true.
    /// </returns>
    public static Func<Shape, bool> Or(params Func<Shape, bool>[] predicates)
        => shape => predicates.Any(predicate => predicate(shape));

    /// <summary>
    /// A predicate that always returns true, matching any shape.
    /// </summary>
    public static Func<Shape, bool> True => (_ => true);

    /// <summary>
    /// A predicate that always returns false, matching no shapes.
    /// </summary>
    public static Func<Shape, bool> False => (_ => false);
}
