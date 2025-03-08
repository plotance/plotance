// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using D = DocumentFormat.OpenXml.Drawing;

namespace Plotance.Renderers.PowerPoint;

/// <summary>Provides methods for manipulating slide layouts.</summary>
public static class SlideLayouts
{
    /// <summary>Unique identifier for slide number field.</summary>
    internal static string s_slideNumberFieldGuidString
        = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";

    /// <summary>Names of slide layouts.</summary>
    private static IReadOnlyDictionary<SlideLayoutValues, string>
        s_slideLayoutNames = new Dictionary<SlideLayoutValues, string>()
        {
            {
                SlideLayoutValues.Title,
                "Title"
            },
            {
                SlideLayoutValues.Text,
                "Text"
            },
            {
                SlideLayoutValues.TwoColumnText,
                "Two Column Text"
            },
            {
                SlideLayoutValues.Table,
                "Table"
            },
            {
                SlideLayoutValues.TextAndChart,
                "Text and Chart"
            },
            {
                SlideLayoutValues.ChartAndText,
                "Chart and Text"
            },
            {
                SlideLayoutValues.Diagram,
                "Diagram"
            },
            {
                SlideLayoutValues.Chart,
                "Chart"
            },
            {
                SlideLayoutValues.TextAndClipArt,
                "Text and Clip Art"
            },
            {
                SlideLayoutValues.ClipArtAndText,
                "Clip Art and Text"
            },
            {
                SlideLayoutValues.TitleOnly,
                "Title Only"
            },
            {
                SlideLayoutValues.Blank,
                "Blank"
            },
            {
                SlideLayoutValues.TextAndObject,
                "Text and Object"
            },
            {
                SlideLayoutValues.ObjectAndText,
                "Object and Text"
            },
            {
                SlideLayoutValues.ObjectOnly,
                "Object"
            },
            {
                SlideLayoutValues.Object,
                "Title and Object"
            },
            {
                SlideLayoutValues.TextAndMedia,
                "Text and Media"
            },
            {
                SlideLayoutValues.MidiaAndText,
                "Media and Text"
            },
            {
                SlideLayoutValues.ObjectOverText,
                "Object over Text"
            },
            {
                SlideLayoutValues.TextOverObject,
                "Text over Object"
            },
            {
                SlideLayoutValues.TextAndTwoObjects,
                "Text and Two Objects"
            },
            {
                SlideLayoutValues.TwoObjectsAndText,
                "Two Objects and Text"
            },
            {
                SlideLayoutValues.TwoObjectsOverText,
                "Two Objects over Text"
            },
            {
                SlideLayoutValues.FourObjects,
                "Four Objects"
            },
            {
                SlideLayoutValues.VerticalText,
                "Vertical Text"
            },
            {
                SlideLayoutValues.ClipArtAndVerticalText,
                "Clip Art and Vertical Text"
            },
            {
                SlideLayoutValues.VerticalTitleAndText,
                "Vertical Title and Text"
            },
            {
                SlideLayoutValues.VerticalTitleAndTextOverChart,
                "Vertical Title and Text over Chart"
            },
            {
                SlideLayoutValues.TwoObjects,
                "Two Objects"
            },
            {
                SlideLayoutValues.ObjectAndTwoObjects,
                "Object and Two Objects"
            },
            {
                SlideLayoutValues.TwoObjectsAndObject,
                "Two Objects and Object"
            },
            {
                SlideLayoutValues.SectionHeader,
                "Section Header"
            },
            {
                SlideLayoutValues.TwoTextAndTwoObjects,
                "Two Text and Two Objects"
            },
            {
                SlideLayoutValues.ObjectText,
                "Title, Object, and Caption"
            },
            {
                SlideLayoutValues.PictureText,
                "Picture and Caption"
            },
            {
                SlideLayoutValues.Custom,
                "Custom"
            }
        };

    /// <summary>
    /// Checks if the slide master part has a slide layout of the specified
    /// type.
    /// </summary>
    /// <param name="slideMasterPart">The slide master part to check.</param>
    /// <param name="layoutType">The type of layout to check for.</param>
    /// <returns>
    /// True if the slide master part has a slide layout of the specified type;
    /// false otherwise.
    /// </returns>
    public static bool HasSlideLayout(
        SlideMasterPart slideMasterPart,
        SlideLayoutValues layoutType
    ) => slideMasterPart
        .SlideLayoutParts
        .Any(
            slideLayoutPart =>
                slideLayoutPart.SlideLayout?.Type?.Value == layoutType
        );

    /// <summary>
    /// Extracts the slide layout part from a slide master part based on the
    /// layout type.
    /// </summary>
    /// <param name="slideMasterPart">
    /// The slide master part to extract the layout from.
    /// </param>
    /// <param name="layoutType">The type of layout to extract.</param>
    /// <returns>The slide layout part matching the layout type.</returns>
    public static SlideLayoutPart ExtractSlideLayoutPart(
        SlideMasterPart slideMasterPart,
        SlideLayoutValues layoutType
    )
    {
        var slideLayoutParts = slideMasterPart.SlideLayoutParts;
        var slideLayoutPart = slideLayoutParts
            .FirstOrDefault(
                slideLayoutPart =>
                    slideLayoutPart.SlideLayout?.Type?.Value == layoutType
            )
            ?? slideMasterPart.AddNewPart<SlideLayoutPart>();

        if (slideLayoutPart.SlideMasterPart == null)
        {
            slideLayoutPart.AddPart(slideMasterPart);
        }

        ExtractSlideLayout(slideLayoutPart, layoutType);

        var slideMaster = SlideMasters
            .ExtractSlideMaster(slideMasterPart);
        var slideLayoutIdList = (
            slideMaster.SlideLayoutIdList ??= new SlideLayoutIdList()
        );
        var slideLayoutRelationshipId = slideMasterPart
            .GetIdOfPart(slideLayoutPart);

        if (
            slideLayoutIdList.Elements<SlideLayoutId>().FirstOrDefault(
                id => id.RelationshipId == slideLayoutRelationshipId
            ) == null
        )
        {
            var nextId = (
                slideLayoutIdList
                    .Elements<SlideLayoutId>()
                    .Max(id => id.Id)
                    ?? 2147483649
            ) + 1;

            slideLayoutIdList.AppendChild(
                new SlideLayoutId()
                {
                    Id = nextId,
                    RelationshipId = slideLayoutRelationshipId
                }
            );
        }

        return slideLayoutPart;
    }

    /// <summary>
    /// Extracts or adds the slide layout from a slide layout part based on the
    /// layout type.
    /// </summary>
    /// <param name="slideLayoutPart">
    /// The slide layout part to extract the layout from.
    /// </param>
    /// <param name="layoutType">The type of layout to extract.</param>
    /// <returns>The slide layout matching the layout type.</returns>
    public static SlideLayout ExtractSlideLayout(
        SlideLayoutPart slideLayoutPart,
        SlideLayoutValues layoutType
    ) => slideLayoutPart.SlideLayout ??= CreateSlideLayout(layoutType);

    /// <summary>
    /// Creates a new slide layout with the specified layout type.
    /// </summary>
    /// <param name="layoutType">The type of layout to create.</param>
    /// <returns>A new SlideLayout instance.</returns>
    public static SlideLayout CreateSlideLayout(SlideLayoutValues layoutType)
    {
        var commonSlideData = layoutType switch
        {
            _ when layoutType == SlideLayoutValues.Title
                => CreateTitleCommonSlideData(),
            _ when layoutType == SlideLayoutValues.SectionHeader
                => CreateTitleCommonSlideData(),
            _ => CreateMainCommonSlideData("‹#›")
        };

        commonSlideData.Name = s_slideLayoutNames[layoutType];

        return new SlideLayout(commonSlideData)
        {
            Type = layoutType
        };
    }

    /// <summary>Creates a CommonSlideData for the main slide layout.</summary>
    /// <param name="slideNumberText">Text for slide number.</param>
    /// <returns>A new CommonSlideData instance.</returns>
    public static CommonSlideData CreateMainCommonSlideData(
        string slideNumberText
    ) => new CommonSlideData(
        new Background(
            new BackgroundProperties(
                new D.SolidFill(
                    new D.SchemeColor()
                    {
                        Val = D.SchemeColorValues.Background1
                    }
                )
            )
        ),
        new ShapeTree(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties()
                {
                    Id = (UInt32Value)1U,
                    Name = ""
                },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()
            ),
            new GroupShapeProperties(
                new D.TransformGroup()
            ),
            Shapes.CreateTitleShape(),
            new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties()
                    {
                        Id = (UInt32Value)3U,
                        Name = "Content"
                    },
                    new NonVisualShapeDrawingProperties(
                        new D.ShapeLocks() { NoGrouping = true }
                    ),
                    new ApplicationNonVisualDrawingProperties(
                        new PlaceholderShape()
                        {
                            Index = 1,
                            Type = PlaceholderValues.Body
                        }
                    )
                ),
                new ShapeProperties(
                    new D.Transform2D(
                        new D.Offset()
                        {
                            X = (Int64Value)(8L * 36000L),
                            Y = (Int64Value)(32L * 36000L)
                        },
                        new D.Extents()
                        {
                            Cx = (Int64Value)(9144000L - 16L * 36000L),
                            Cy = (Int64Value)(6858000L - 49L * 36000L)
                        }
                    )
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
            ),
            new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties()
                    {
                        Id = (UInt32Value)4U,
                        Name = "Page Numbering"
                    },
                    new NonVisualShapeDrawingProperties(
                        new D.ShapeLocks() { NoGrouping = true }
                    ),
                    new ApplicationNonVisualDrawingProperties(
                        new PlaceholderShape()
                        {
                            Type = PlaceholderValues.SlideNumber,
                            Index = 12
                        }
                    )
                ),
                new ShapeProperties(
                    new D.Transform2D(
                        new D.Offset()
                        {
                            X = (Int64Value)(9144000 - 16L * 36000L),
                            Y = (Int64Value)(6858000 - 16L * 36000L)
                        },
                        new D.Extents()
                        {
                            Cx = (Int64Value)(8L * 36000L),
                            Cy = (Int64Value)(8L * 36000L)
                        }
                    )
                ),
                new TextBody(
                    new D.BodyProperties()
                    {
                        Wrap = D.TextWrappingValues.None
                    },
                    new D.ListStyle(
                        new D.Level1ParagraphProperties(
                            new D.NoBullet()
                        )
                        {
                            Alignment = D.TextAlignmentTypeValues.Right
                        }
                    ),
                    new D.Paragraph(
                        new D.Field(
                            new D.Text(slideNumberText)
                        )
                        {
                            Id = s_slideNumberFieldGuidString,
                            Type = "slidenum"
                        }
                    )
                )
            )
        )
    );

    /// <summary>Creates a CommonSlideData for the title slide layout.</summary>
    /// <returns>A new CommonSlideData instance.</returns>
    public static CommonSlideData CreateTitleCommonSlideData()
        => new CommonSlideData(
            new Background(
                new BackgroundProperties(
                    new D.SolidFill(
                        new D.SchemeColor()
                        {
                            Val = D.SchemeColorValues.Background1
                        }
                    )
                )
            ),
            new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties()
                    {
                        Id = (UInt32Value)1U,
                        Name = ""
                    },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()
                ),
                new GroupShapeProperties(),
                new Shape(
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
                            new PlaceholderShape()
                            {
                                Type = PlaceholderValues.CenteredTitle
                            }
                        )
                    ),
                    new ShapeProperties(
                        new D.Transform2D(
                            new D.Offset()
                            {
                                X = (Int64Value)(8L * 36000L),
                                Y = (Int64Value)((6858000L - 23L * 36000L) / 2)
                            },
                            new D.Extents()
                            {
                                Cx = (Int64Value)(9144000L - 16L * 36000L),
                                Cy = (Int64Value)(23L * 36000L)
                            }
                        )
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
                ),
                new Shape(
                    new NonVisualShapeProperties(
                        new NonVisualDrawingProperties()
                        {
                            Id = (UInt32Value)3U,
                            Name = "Subtitle"
                        },
                        new NonVisualShapeDrawingProperties(
                            new D.ShapeLocks() { NoGrouping = true }
                        ),
                        new ApplicationNonVisualDrawingProperties(
                            new PlaceholderShape()
                            {
                                Index = 1,
                                Type = PlaceholderValues.SubTitle
                            }
                        )
                    ),
                    new ShapeProperties(
                        new D.Transform2D(
                            new D.Offset()
                            {
                                X = (Int64Value)(8L * 36000L),
                                Y = (Int64Value)(
                                    (6858000L - 23L * 36000L) / 2 + 24L * 36000L
                                )
                            },
                            new D.Extents()
                            {
                                Cx = (Int64Value)(9144000L - 16L * 36000L),
                                Cy = (Int64Value)(23L * 36000L)
                            }
                        )
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
                )
            )
        );

    /// <summary>Extracts the placeholder shape from an element.</summary>
    /// <param name="element">
    /// The element to extract the placeholder shape from.
    /// </param>
    /// <returns>The placeholder shape, or null if not found.</returns>
    public static PlaceholderShape? ExtractPlaceholderShape(
        OpenXmlElement element
    )
    {
        var applicationNonVisualDrawingProperties = element switch
        {
            Shape shape => shape
                .NonVisualShapeProperties
                ?.ApplicationNonVisualDrawingProperties,

            ConnectionShape connectionShape => connectionShape
                .NonVisualConnectionShapeProperties
                ?.ApplicationNonVisualDrawingProperties,

            GraphicFrame graphicFrame => graphicFrame
                .NonVisualGraphicFrameProperties
                ?.ApplicationNonVisualDrawingProperties,

            GroupShape groupShape => groupShape
                .NonVisualGroupShapeProperties
                ?.ApplicationNonVisualDrawingProperties,

            Picture picture => picture
                .NonVisualPictureProperties
                ?.ApplicationNonVisualDrawingProperties,

            _ => null
        };

        return applicationNonVisualDrawingProperties?.PlaceholderShape;
    }

    /// <summary>
    /// Extracts the region of a placeholder shape from a slide layout or slide
    /// master corresponding to the given placeholder shape.
    /// </summary>
    /// <param name="slidePart">
    /// The slide part to extract the region from.
    /// </param>
    /// <param name="bodyShape">
    /// The body shape to use as a template, or null to find one.
    /// </param>
    /// <returns>
    /// A Rectangle representing the region of the placeholder shape or a
    /// default value.
    /// </returns>
    public static Rectangle ExtractBodyRegion(
        SlidePart slidePart,
        Shape? bodyShape
    )
    {
        // TODO use slide size
        var defaultBodyRegion = new Rectangle(
            X: 8L * 36000L,
            Y: 24L * 36000L,
            Width: 9144000L - 16L * 36000L,
            Height: 6858000L - 41L * 36000L
        );

        if (bodyShape == null)
        {
            return defaultBodyRegion;
        }

        var placeholderMatches = ShapePredicates
            .HaveMatchingPlaceholder(bodyShape);
        var slideLayoutPart = slidePart.SlideLayoutPart;
        IEnumerable<OpenXmlPart?> parts = [
            slideLayoutPart,
            slideLayoutPart?.SlideMasterPart,
            slidePart
                .OpenXmlPackage
                ?.GetPartsOfType<SlideMasterPart>()
                ?.FirstOrDefault()
        ];
        var shapes = parts
            .Select(part => Shapes.FindShape(part, placeholderMatches))
            .Prepend(bodyShape);
        var maybeRegion = (
            Presentations.FindOrDefault(
                shapes,
                shape => Geometries.GetOffset(shape)?.X
            ),
            Presentations.FindOrDefault(
                shapes,
                shape => Geometries.GetOffset(shape)?.Y
            ),
            Presentations.FindOrDefault(
                shapes,
                shape => Geometries.GetExtents(shape)?.Cx
            ),
            Presentations.FindOrDefault(
                shapes,
                shape => Geometries.GetExtents(shape)?.Cy
            )
        );

        if (
            maybeRegion is (
                Int64Value x,
                Int64Value y,
                Int64Value cx,
                Int64Value cy
            )
        )
        {
            return new Rectangle(x, y, cx, cy);
        }
        else
        {
            return defaultBodyRegion;
        }
    }
}
