// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using Markdig.Syntax;
using Plotance.Models;

using D = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace Plotance.Renderers.PowerPoint;

/// <summary>
/// Provides methods for converting Markdown blocks to PowerPoint presentation.
/// </summary>
public static class PowerPointRenderer
{
    /// <summary>Whether to suppress messages.</summary>
    public static bool Quiet { get; set; }

    /// <summary>Convert Markdown blocks to PowerPoint presentation.</summary>
    /// <param name="template">
    /// The PowerPoint template document to use as a base.
    /// </param>
    /// <param name="variables">
    /// Dictionary of variables to expand in text content.
    /// </param>
    /// <param name="blocks">The Markdown blocks to convert to slides.</param>
    /// <returns>The converted PowerPoint presentation document.</returns>
    /// <exception cref="InvalidOperationException">
    /// Thrown if the slide layout does not have common slide data.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// Thrown when the block has an invalid structure, the configuration is
    /// invalid, image URL is invalid, image format is not supported, or the
    /// image data is invalid..
    /// </exception>
    public static PresentationDocument Render(
        PresentationDocument template,
        IReadOnlyDictionary<string, string> variables,
        IEnumerable<Block> blocks
    )
    {
        var variablesWithDefault = new Dictionary<string, string>(variables);

        variablesWithDefault["$"] = "$";

        var sections = ImplicitSection.Create(blocks, variablesWithDefault);
        var document = template.Clone();

        Slides.RemoveAllSlides(Presentations.ExtractPresentationPart(document));

        foreach (var section in sections)
        {
            RenderSection(document, section);
        }

        AddSectionList(
            Presentations.ExtractPresentation(
                Presentations.ExtractPresentationPart(document)
            ),
            sections
        );

        document.Save();

        if (!Quiet)
        {
            Validate(document);
        }

        return document;
    }

    /// <summary>Renders a section to a slide.</summary>
    /// <param name="document">The PowerPoint document to render to.</param>
    /// <param name="section">The section to render.</param>
    /// <exception cref="InvalidOperationException">
    /// Thrown if the slide layout does not have common slide data.
    /// </exception>
    /// <exception cref="PlotanceException">
    /// Thrown when the block has an invalid structure, the configuration is
    /// invalid, image URL is invalid, image format is not supported, or the
    /// image data is invalid.
    /// </exception>
    private static void RenderSection(
        PresentationDocument document,
        ImplicitSection section
    )
    {
        var presentationPart = Presentations.ExtractPresentationPart(document);
        var slideMasterPart = SlideMasters
            .ExtractSlideMasterPart(presentationPart);
        var defaultSlideLayoutType = SlideLayouts.HasSlideLayout(
            slideMasterPart,
            SlideLayoutValues.Text
        )
            ? SlideLayoutValues.Text
            : SlideLayoutValues.Object;
        var slideLayoutType = GetSlideLayoutType(
            section,
            defaultSlideLayoutType
        );
        var slidePart = Slides.AddSlidePart(presentationPart, slideLayoutType);
        var bodyShape = Shapes.ExtractBodyShape(slidePart);

        bodyShape?.Remove();
        bodyShape ??= Shapes.FindShapeFromAncestorParts(
            slidePart,
            Shapes.ExtractBodyShape
        );

        var templateShape = bodyShape ?? CreateTemplateShape();
        var slideParagraphStyles = SlideParagraphStyles
            .ExtractParagraphStyles(slidePart);

        RenderHeading(slidePart, slideParagraphStyles, section);

        var isFirst = true;
        var bodyRegion = SlideLayouts.ExtractBodyRegion(slidePart, bodyShape);
        var y = bodyRegion.Y;
        var yPositions = ComputeYPositions(bodyRegion.Height, section);

        foreach (var (row, (yOffset, height)) in section.Rows.Zip(yPositions))
        {
            RenderRow(
                slidePart,
                bodyRegion,
                slideParagraphStyles,
                templateShape,
                isFirst,
                y + yOffset,
                height,
                row
            );

            isFirst = false;
        }

        if (slidePart.Slide is Slide slide)
        {
            Shapes.FixShapeIds(slide);
        }
    }

    /// <summary>Computes the Y position and height of the rows.</summary>
    /// <param name="totalHeight">The total height of the section.</param>
    /// <param name="section">
    /// The section to compute the Y positions for.
    /// </param>
    /// <returns>The Y position and height of the rows.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the rows do not fit into the space.
    /// </exception>
    private static IReadOnlyList<(long Start, long Length)> ComputeYPositions(
        long totalHeight,
        ImplicitSection section
    )
    {
        try
        {
            return ILengthWeight
                .Divide(
                    totalHeight,
                    section
                        .Rows
                        .Select(
                            row => (
                                Weight: row.Weight,
                                Gap: row.GapBefore?.ToEmu() ?? 0
                            )
                        )
                        .ToList()
                );
        }
        catch (ArgumentException e)
        {
            throw new PlotanceException(
                section.Path,
                section.Line,
                "Rows does not fit into the space.",
                e
            );
        }
    }

    /// <summary>
    /// Gets the slide layout type for the given section.  Sections are
    /// classified as title, section header, or text.
    ///
    /// If the heading level is greater than or equal to the slide level,
    /// it is a text slide. Otherwise, if the heading level is 1, it is a
    /// title slide. Otherwise, it is a section header slide.
    /// </summary>
    /// <param name="section">
    /// The section to get the slide layout type for.
    /// </param>
    /// <returns>The slide layout type.</returns>
    private static SlideLayoutValues GetSlideLayoutType(
        ImplicitSection section,
        SlideLayoutValues defaultSlideLayoutType
    )
    {
        var headingLevel = section.HeadingBlock?.Level;
        var slideLevel = section.SlideLevel?.Value ?? 2;

        if (slideLevel <= headingLevel)
        {
            return defaultSlideLayoutType;
        }
        else if (headingLevel == 1)
        {
            return SlideLayoutValues.Title;
        }
        else
        {
            return SlideLayoutValues.SectionHeader;
        }
    }

    /// <summary>Creates a new template shape.</summary>
    /// <returns>A new Shape instance configured as a template shape.</returns>
    private static Shape CreateTemplateShape()
        => new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties()
                {
                    Id = (UInt32Value)1U,
                    Name = ""
                },
                new NonVisualShapeDrawingProperties(
                    new D.ShapeLocks() { NoGrouping = true }
                ),
                new ApplicationNonVisualDrawingProperties(
                    new PlaceholderShape()
                    {
                        Type = PlaceholderValues.Object,
                        Index = 1
                    }
                )
            ),
            new ShapeProperties(),
            new TextBody(
                new D.BodyProperties(
                    new D.NoAutoFit()
                ),
                new D.Paragraph()
            )
        );

    /// <summary>Render the heading block of a section to the slide.</summary>
    /// <param name="slidePart">The slide part to render to.</param>
    /// <param name="slideParagraphStyles">The slide paragraph styles.</param>
    /// <param name="section">The section to render.</param>
    /// <exception cref="PlotanceException">
    /// Thrown when the block has an invalid structure or the configuration is
    /// invalid.
    /// </exception>
    private static void RenderHeading(
        SlidePart slidePart,
        SlideParagraphStyles slideParagraphStyles,
        ImplicitSection section
    )
    {
        if (section.HeadingBlock is HeadingBlock headingBlock)
        {
            var titleShape = Shapes.ExtractTitleShape(slidePart);

            if (titleShape == null)
            {
                var sourceShape = Shapes.FindShapeFromAncestorParts(
                    slidePart,
                    Shapes.ExtractTitleShape
                ) ?? Shapes.CreateTitleShape();

                titleShape = Shapes.CreateShapeFromTemplate(sourceShape);

                slidePart
                    .Slide
                    ?.CommonSlideData
                    ?.ShapeTree
                    ?.AppendChild(titleShape);
            }

            var titleFontScale = section.TitleFontScale?.Value ?? 1m;

            if (titleFontScale != 1m || section.Language != null)
            {
                var textBody = (
                    titleShape.TextBody ??= new TextBody(
                        new D.BodyProperties(),
                        new D.Paragraph()
                    )
                );

                textBody.ListStyle = slideParagraphStyles
                    .Title
                    .Scaled(titleFontScale)
                    .WithLanguage(section.Language)
                    .ToListStyle();
            }

            Shapes.ReplaceTextBodyContent(
                titleShape,
                Paragraphs.CreateParagraphs(
                    slidePart,
                    section.Variables,
                    ExtractPath(headingBlock),
                    headingBlock
                )
            );
        }
    }

    /// <summary>Renders a row to a slide.</summary>
    /// <param name="slidePart">The slide part to render to.</param>
    /// <param name="bodyRegion">The body region to render to.</param>
    /// <param name="slideParagraphStyles">The slide paragraph styles.</param>
    /// <param name="templateShape">The template shape.</param>
    /// <param name="isFirstRow">Whether this is the first row.</param>
    /// <param name="y">The Y coordinate of the row.</param>
    /// <param name="height">The height of the row.</param>
    /// <param name="row">The row to render.</param>
    /// <exception cref="PlotanceException">
    /// Thrown when the block has an invalid structure, the configuration is
    /// invalid, image URL is invalid, image format is not supported, or the
    /// image data is invalid.
    /// </exception>
    private static void RenderRow(
        SlidePart slidePart,
        Rectangle bodyRegion,
        SlideParagraphStyles slideParagraphStyles,
        Shape templateShape,
        bool isFirstRow,
        long y,
        long height,
        ImplicitSectionRow row
    )
    {
        var isFirstColumn = true;
        var x = bodyRegion.X;
        var xPositions = ComputeXPositions(bodyRegion.Width, row);

        foreach (
            var (column, (xOffset, width)) in row.Columns.Zip(xPositions)
        )
        {
            RenderColumn(
                slidePart,
                slideParagraphStyles,
                templateShape,
                isFirstRow && isFirstColumn,
                x + xOffset,
                y,
                width,
                height,
                column
            );

            isFirstColumn = false;
        }
    }

    /// <summary>Computes the X position and width of the columns.</summary>
    /// <param name="totalWidth">The total width of the row.</param>
    /// <param name="row">The row to compute the X positions for.</param>
    /// <returns>The X position and width of the columns.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the columns do not fit into the space.
    /// </exception>
    private static IReadOnlyList<(long Start, long Length)> ComputeXPositions(
        long totalWidth,
        ImplicitSectionRow row
    )
    {
        try
        {
            return ILengthWeight
                .Divide(
                    totalWidth,
                    row
                        .Columns
                        .Select(
                            column => (
                                Weight: column.Weight,
                                Gap: column.GapBefore?.ToEmu() ?? 0
                            )
                        )
                        .ToList()
                );
        }
        catch (ArgumentException e)
        {
            throw new PlotanceException(
                row.Path,
                row.Line,
                "Columns does not fit into the space.",
                e
            );
        }
    }

    /// <summary>Renders a column to a slide.</summary>
    /// <param name="slidePart">The slide part to render to.</param>
    /// <param name="slideParagraphStyles">The slide paragraph styles.</param>
    /// <param name="templateShape">The template shape.</param>
    /// <param name="isFirst">
    /// Whether this is the first block of the slide.
    /// </param>
    /// <param name="x">The X coordinate of the block.</param>
    /// <param name="y">The Y coordinate of the block.</param>
    /// <param name="width">The width of the block.</param>
    /// <param name="height">The height of the block.</param>
    /// <param name="column">The column to render.</param>
    /// <exception cref="PlotanceException">
    /// Thrown when the block has an invalid structure, the configuration is
    /// invalid, image URL is invalid, image format is not supported, or the
    /// image data is invalid.
    /// </exception>
    private static void RenderColumn(
        SlidePart slidePart,
        SlideParagraphStyles slideParagraphStyles,
        Shape templateShape,
        bool isFirst,
        long x,
        long y,
        long width,
        long height,
        ImplicitSectionColumn column
    )
    {
        var block = column.Block;

        if (
            block.GetData("query_results")
                is IReadOnlyList<QueryResultSet> queryResults
        )
        {
            if (column.ChartOptions?.Format?.Value == "table")
            {
                var bodyFontScale = column.BodyFontScale?.Value ?? 1m;

                TableRenderer.Render(
                    slidePart,
                    slideParagraphStyles.Body.Scaled(bodyFontScale),
                    isFirst,
                    x,
                    y,
                    width,
                    height,
                    column,
                    queryResults
                );
            }
            else
            {
                ChartRenderer.Render(
                    slidePart,
                    isFirst,
                    x,
                    y,
                    width,
                    height,
                    column,
                    queryResults
                );
            }
        }
        else if (
            SlideImages.ExtractImageLink(column.Variables, block)
                is ImageLink imageLink
        )
        {
            SlideImages.EmbedImage(
                slidePart,
                imageLink,
                x,
                y,
                width,
                height,
                isFirst,
                Geometries.ParseHorizontalAlign(column.BodyHorizontalAlign)
                    ?? D.TextAlignmentTypeValues.Left,
                Geometries.ParseVerticalAlign(column.BodyVerticalAlign)
                    ?? D.TextAnchoringTypeValues.Top
            );
        }
        else
        {
            RenderTextBlock(
                slidePart,
                slideParagraphStyles,
                templateShape,
                isFirst,
                x,
                y,
                width,
                height,
                column
            );
        }
    }

    /// <summary>Render a text block to a slide.</summary>
    /// <param name="slidePart">The slide part to render to.</param>
    /// <param name="slideParagraphStyles">The slide paragraph styles.</param>
    /// <param name="templateShape">The template shape.</param>
    /// <param name="isFirst">
    /// Whether this is the first block of the slide.
    /// </param>
    /// <param name="x">The X coordinate of the block.</param>
    /// <param name="y">The Y coordinate of the block.</param>
    /// <param name="width">The new width of the shape.</param>
    /// <param name="height">The new height of the shape.</param>
    /// <param name="column">The column to render.</param>
    /// <exception cref="PlotanceException">
    /// Thrown when the block has an invalid structure or the configuration is
    /// invalid.
    /// </exception>
    private static void RenderTextBlock(
        SlidePart slidePart,
        SlideParagraphStyles slideParagraphStyles,
        Shape templateShape,
        bool isFirst,
        long x,
        long y,
        long width,
        long height,
        ImplicitSectionColumn column
    )
    {
        var bodyFontScale = column.BodyFontScale?.Value ?? 1m;
        var shape = Shapes.CreateShapeFromTemplate(templateShape);

        if (
            bodyFontScale != 1m
                || column.Language != null
                || column.BodyHorizontalAlign != null
                || !isFirst
        )
        {
            var textBody = (
                shape.TextBody ??= new TextBody(
                    new D.BodyProperties(),
                    new D.Paragraph()
                )
            );

            var listStyle = slideParagraphStyles
                .Body
                .Scaled(bodyFontScale)
                .WithLanguage(column.Language);

            if (
                Geometries.ParseHorizontalAlign(column.BodyHorizontalAlign)
                    is D.TextAlignmentTypeValues horizontalAlign
            )
            {
                listStyle = listStyle.WithAlignment(horizontalAlign);
            }

            textBody.ListStyle = listStyle.ToListStyle();
        }

        Shapes.MoveShape(shape, x, y, width, height);
        Shapes.ReplaceTextBodyContent(
            shape,
            Paragraphs.CreateParagraphs(
                slidePart,
                column.Variables,
                ExtractPath(column.Block),
                column.Block
            )
        );

        if (
            Geometries.ParseVerticalAlign(column.BodyVerticalAlign)
                is D.TextAnchoringTypeValues verticalAlign
        )
        {
            var textBody = shape.TextBody ??= new TextBody(
                new D.BodyProperties(),
                new D.Paragraph()
            );
            var bodyProperties = textBody.BodyProperties
                ??= new D.BodyProperties();

            bodyProperties.Anchor = verticalAlign;
        }

        if (!isFirst)
        {
            var placeholderShape = shape
                .NonVisualShapeProperties
                ?.ApplicationNonVisualDrawingProperties
                ?.PlaceholderShape;

            if (placeholderShape != null)
            {
                placeholderShape.Index = null;
            }
        }

        slidePart
            .Slide
            ?.CommonSlideData
            ?.ShapeTree
            ?.AppendChild(shape);
    }

    /// <summary>Extracts the Markdown file path from a block.</summary>
    /// <param name="block">The block to extract the path from.</param>
    /// <returns>The Markdown file path.</returns>
    /// <exception cref="ArgumentException">
    /// Thrown when the path is not set.
    /// </exception>
    public static string ExtractPath(Block block)
        => block.GetData("path") as string ?? throw new ArgumentException(
            $"path is not set."
        );

    /// <summary>
    /// Adds section list to the presentation when there are two or more
    /// presentation sections.
    /// </summary>
    /// <param name="presentation">The presentation.</param>
    /// <param name="implicitSections">
    /// Implicit sections in rendering order.
    /// </param>
    private static void AddSectionList(
        Presentation presentation,
        IReadOnlyList<ImplicitSection> implicitSections
    )
    {
        var slideIds = presentation.SlideIdList?.Elements<SlideId>()?.ToList()
            ?? [];

        var presentationSections
            = new List<(string Name, List<uint> SlideIds)>();

        foreach (var (section, slideId) in implicitSections.Zip(slideIds))
        {
            var headingLevel = section.HeadingBlock?.Level ?? Int32.MaxValue;
            var slideLevel = section.SlideLevel?.Value ?? 2;
            var startsNew = headingLevel < slideLevel;

            if (startsNew || presentationSections.Count == 0)
            {
                var defaultName = $"Section {presentationSections.Count + 1}";
                var name = section.HeadingBlock == null
                    ? defaultName
                    : ExtractHeadingText(
                        section.Variables,
                        section.HeadingBlock
                    );

                if (string.IsNullOrWhiteSpace(name))
                {
                    name = defaultName;
                }

                presentationSections.Add((name, []));
            }

            presentationSections[^1].SlideIds.Add(slideId.Id ?? 0);
        }

        if (presentationSections.Count < 2)
        {
            return;
        }

        var extList = presentation.PresentationExtensionList
            ??= new PresentationExtensionList();

        const string SectionListUri = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";

        extList.AddChild(
            new PresentationExtension(
                new P14.SectionList(
                    presentationSections.Select(
                        section => new P14.Section(
                            new P14.SectionSlideIdList(
                                section.SlideIds.Select(
                                    slideId => new P14.SectionSlideIdListEntry()
                                    {
                                        Id = slideId
                                    }
                                )
                            )
                        )
                        {
                            Name = section.Name,
                            Id = Guid.NewGuid().ToString("B").ToUpperInvariant()
                        }
                    )
                )
            )
            {
                Uri = SectionListUri
            }
        );
    }

    /// <summary>
    /// Extracts plain text from a heading block.
    /// </summary>
    /// <param name="variables">
    /// Dictionary of variables to expand in text content.
    /// </param>
    /// <param name="heading">The heading block to extract text from.</param>
    /// <returns>
    /// The extracted text or null if the block has no inline content.
    /// </returns>
    private static string ExtractHeadingText(
        IReadOnlyDictionary<string, string> variables,
        HeadingBlock heading
    )
    {
        var inline = heading.Inline;

        return inline == null
            ? string.Empty
            : Paragraphs.ToPlainText(variables, inline);
    }

    /// <summary>
    /// Validates the presentation document. Errors are written to
    /// <see cref="Console.Error"/>.
    /// </summary>
    /// <param name="document">The presentation document to validate.</param>
    private static void Validate(PresentationDocument document)
    {
        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);

        foreach (var error in validator.Validate(document))
        {
            Console.Error.WriteLine(
                string.Join(
                    "\n  ",
                    [
                        error.Id,
                        error.Part?.Uri,
                        error.Path?.XPath,
                        error.RelatedPart,
                        error.RelatedNode,
                        error.Description
                    ]
                )
            );
            Console.Error.WriteLine();
        }
    }
}
