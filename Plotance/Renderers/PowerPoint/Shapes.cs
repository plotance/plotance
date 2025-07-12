// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Plotance.Models;

using D = DocumentFormat.OpenXml.Drawing;

namespace Plotance.Renderers.PowerPoint;

/// <summary>Provides methods for manipulating shapes.</summary>
public static class Shapes
{
    /// <summary>Creates a title shape with default properties.</summary>
    /// <returns>A new Shape instance configured as a title shape.</returns>
    public static Shape CreateTitleShape()
    {
        return new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties()
                {
                    Id = (UInt32Value)2U,
                    Name = "Title"
                },
                new NonVisualShapeDrawingProperties(
                    new D.ShapeLocks() { NoGrouping = true }
                ),
                new ApplicationNonVisualDrawingProperties(
                    new PlaceholderShape() { Type = PlaceholderValues.Title }
                )
            ),
            new ShapeProperties(
                new D.Transform2D(
                    new D.Offset()
                    {
                        X = (Int64Value)(8L * 36000L),
                        Y = (Int64Value)(8L * 36000L)
                    },
                    new D.Extents()
                    {
                        Cx = (Int64Value)(9144000L - 16L * 36000L),
                        Cy = (Int64Value)(23L * 36000L)
                    }
                ),
                new D.PresetGeometry(new D.AdjustValueList())
                {
                    Preset = D.ShapeTypeValues.Rectangle
                }
            ),
            new TextBody(
                new D.BodyProperties(
                    new D.NoAutoFit()
                ),
                new D.Paragraph(
                    new D.Run(
                        new D.Text("")
                    )
                )
            )
        );
    }

    /// <summary>
    /// Creates a new shape based on a template shape. Text body is cleared.
    /// </summary>
    /// <param name="templateShape">
    /// The template shape to base the new shape on.
    /// </param>
    /// <returns>
    /// A new Shape instance with properties from the template.
    /// </returns>
    public static Shape CreateShapeFromTemplate(Shape templateShape)
    {
        var shape = (Shape)templateShape.CloneNode(true);

        CleanPlaceholderShape(shape);

        return shape;
    }

    /// <summary>
    /// Cleans a placeholder shape by removing all paragraphs from its text body
    /// and adding a new one.
    /// </summary>
    /// <param name="element">The element to clean.</param>
    public static void CleanPlaceholderShape(OpenXmlElement element)
    {
        var textBody = element.GetFirstChild<TextBody>();

        textBody?.RemoveAllChildren<D.Paragraph>();
        textBody?.AddChild(new D.Paragraph());
    }

    /// <summary>Moves a shape to the specified position.</summary>
    /// <param name="shape">The shape to move.</param>
    /// <param name="x">The X coordinate to move to.</param>
    /// <param name="y">The Y coordinate to move to.</param>
    /// <param name="width">The new width of the shape.</param>
    /// <param name="height">The new height of the shape.</param>
    public static void MoveShape(
        Shape shape,
        long x,
        long y,
        long width,
        long height
    )
    {
        var shapeProperties = shape
            .ShapeProperties
            ??= new ShapeProperties(
                new D.PresetGeometry(new D.AdjustValueList())
                {
                    Preset = D.ShapeTypeValues.Rectangle
                }
            );

        shapeProperties.RemoveAllChildren<D.Transform2D>();

        shapeProperties.AddChild(
            new D.Transform2D(
                new D.Offset()
                {
                    X = (Int64Value)x,
                    Y = (Int64Value)y
                },
                new D.Extents()
                {
                    Cx = (Int64Value)width,
                    Cy = (Int64Value)height
                }
            )
        );
    }

    /// <summary>Replaces the text content of a shape.</summary>
    /// <param name="shape">The shape to modify.</param>
    /// <param name="paragraphs">The new text content.</param>
    public static void ReplaceTextBodyContent(
        Shape shape,
        IEnumerable<D.Paragraph> paragraphs
    )
    {
        var textBody = shape.GetFirstChild<TextBody>();

        if (textBody == null)
        {
            textBody = new TextBody(
                new D.BodyProperties(
                    new D.NoAutoFit()
                )
            );

            shape.AddChild(textBody);
        }

        textBody.RemoveAllChildren<D.Paragraph>();
        textBody.Append(paragraphs);
    }

    /// <summary>Assigns fresh IDs to shapes in the document.</summary>
    /// <param name="root">The root OpenXml element to fix shape IDs in.</param>
    public static void FixShapeIds(OpenXmlElement root)
    {
        var descendants = root.Descendants<NonVisualDrawingProperties>();
        uint nextId = 1U;

        foreach (var nonVisualDrawingProperties in descendants)
        {
            nonVisualDrawingProperties.Id = nextId;

            nextId++;
        }
    }

    /// <summary>Extracts the title shape from an OpenXml part.</summary>
    /// <remarks>
    /// Title shape is the first shape with a title or centered title
    /// placeholder.
    /// </remarks>
    /// <param name="part">The OpenXml part to find a title shape in.</param>
    /// <returns>
    /// The first shape with a title placeholder, or null if none exists.
    /// </returns>
    public static Shape? ExtractTitleShape(OpenXmlPart part)
    {
        bool IsTitleShape(Shape shape)
        {
            var placeHolderType = shape
                .NonVisualShapeProperties
                ?.ApplicationNonVisualDrawingProperties
                ?.PlaceholderShape
                ?.Type
                ?.Value;

            return placeHolderType == PlaceholderValues.Title
                || placeHolderType == PlaceholderValues.CenteredTitle;
        }

        return FindShape(part, IsTitleShape);
    }

    /// <summary>Extracts the body shape from an OpenXml part.</summary>
    /// <remarks>
    /// Body shape is the first shape with placeholder with index one, or
    /// placeholder type of chart, diagram, table, subtitle, or body.
    /// It is used for a template and its region defines where the content
    /// is rendered.
    /// </remarks>
    /// <param name="part">The OpenXml part to find a body shape in.</param>
    /// <returns>
    /// The first shape with a body placeholder, or null if none exists.
    /// </returns>
    public static Shape? ExtractBodyShape(OpenXmlPart part)
    {
        return FindShape(
            part,
            ShapePredicates.Or(
                ShapePredicates.HavePlaceholderIndex(1),
                ShapePredicates.IsPlaceholderOfType(PlaceholderValues.Chart),
                ShapePredicates.IsPlaceholderOfType(PlaceholderValues.Diagram),
                ShapePredicates.IsPlaceholderOfType(PlaceholderValues.Table),
                ShapePredicates.IsPlaceholderOfType(PlaceholderValues.SubTitle),
                ShapePredicates.IsPlaceholderOfType(PlaceholderValues.Body)
            )
        );
    }

    /// <summary>
    /// Finds a shape in an OpenXml part that matches the predicate.
    /// </summary>
    /// <param name="part">The OpenXml part to search in.</param>
    /// <param name="predicate">The predicate to match shapes against.</param>
    /// <returns>The first matching shape, or null if none found.</returns>
    public static Shape? FindShape(
        OpenXmlPart? part,
        Func<Shape, bool> predicate
    )
    {
        return part
            ?.RootElement
            ?.GetFirstChild<CommonSlideData>()
            ?.ShapeTree
            ?.Elements<Shape>()
            ?.FirstOrDefault(predicate);
    }

    public static Shape? FindShapeFromAncestorParts(
        SlidePart slidePart,
        Func<OpenXmlPart, Shape?> extractor
    )
    {
        var slideLayoutPart = slidePart.SlideLayoutPart;
        List<OpenXmlPart?> ancestorParts = [
            slideLayoutPart,
            slideLayoutPart?.SlideMasterPart,
            slidePart
                .OpenXmlPackage
                ?.GetPartsOfType<SlideMasterPart>()
                ?.FirstOrDefault()
        ];

        return Presentations.FindOrDefault(
            ancestorParts,
            part => part == null ? null : extractor(part)
        );
    }
}
