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
    /// image data is invalid.
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

        var layoutDirection = section.LayoutDirection?.Value
            ?? LayoutDirection.Row;
        var isFirst = true;
        var bodyRegion = SlideLayouts.ExtractBodyRegion(slidePart, bodyShape);
        var origin = layoutDirection.Choose(
            row: bodyRegion.Y,
            column: bodyRegion.X
        );
        var totalLength = layoutDirection.Choose(
            row: bodyRegion.Height,
            column: bodyRegion.Width
        );
        var offsets = ComputeBlockGroupOffsets(totalLength, section);

        foreach (
            var (blockGroup, (offset, length))
                in section.BlockGroups.Zip(offsets)
        )
        {
            RenderBlockGroup(
                slidePart,
                bodyRegion,
                slideParagraphStyles,
                templateShape,
                layoutDirection,
                isFirst,
                origin + offset,
                length,
                blockGroup
            );

            isFirst = false;
        }

        if (slidePart.Slide is Slide slide)
        {
            Shapes.FixShapeIds(slide);
        }
    }

    /// <summary>
    /// Computes the offsets and lengths of the block groups.
    ///
    /// If LayoutDirection is "row", computes the Y positions and heights of
    /// the rows. If LayoutDirection is "column", computes the X positions and
    /// widths of the columns.
    /// </summary>
    /// <param name="totalLength">The total length of the section.</param>
    /// <param name="section">
    /// The section to compute the offsets and lengths.
    /// </param>
    /// <returns>The offsets and lengths of the block groups.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the block group do not fit into the space.
    /// </exception>
    private static IReadOnlyList<(long Start, long Length)>
        ComputeBlockGroupOffsets(
            long totalLength,
            ImplicitSection section
        )
    {
        try
        {
            return ILengthWeight
                .Divide(
                    totalLength,
                    section
                        .BlockGroups
                        .Select(
                            blockGroup => (
                                Weight: blockGroup.Weight,
                                Gap: blockGroup.GapBefore?.ToEmu() ?? 0
                            )
                        )
                        .ToList()
                );
        }
        catch (ArgumentException e)
        {
            var layoutDirection = section.LayoutDirection?.Value
                ?? LayoutDirection.Row;

            throw new PlotanceException(
                section.Path,
                section.Line,
                layoutDirection.Choose(
                    row: "Rows does not fit into the space.",
                    column: "Columns does not fit into the space."
                ),
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

        if (headingLevel == null || slideLevel <= headingLevel)
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
            new ShapeProperties(
                new D.PresetGeometry(new D.AdjustValueList())
                {
                    Preset = D.ShapeTypeValues.Rectangle
                }
            ),
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

    /// <summary>
    /// Renders a block group (row if layout direction is "row" or column if
    /// layout direction is "column") to a slide.
    /// </summary>
    /// <param name="slidePart">The slide part to render to.</param>
    /// <param name="bodyRegion">The body region to render to.</param>
    /// <param name="slideParagraphStyles">The slide paragraph styles.</param>
    /// <param name="templateShape">The template shape.</param>
    /// <param name="layoutDirection">
    /// The layout direction of the block group.
    /// </param>
    /// <param name="isFirstGroup">
    /// Whether this is the first block group.
    /// </param>
    /// <param name="groupOffset">
    /// The offset (Y or X positions) of the block group.
    /// </param>
    /// <param name="groupLength">
    /// The length (height or width) of the block group.
    /// </param>
    /// <param name="blockGroup">The block group to render.</param>
    /// <exception cref="PlotanceException">
    /// Thrown when the block has an invalid structure, the configuration is
    /// invalid, image URL is invalid, image format is not supported, or the
    /// image data is invalid.
    /// </exception>
    private static void RenderBlockGroup(
        SlidePart slidePart,
        Rectangle bodyRegion,
        SlideParagraphStyles slideParagraphStyles,
        Shape templateShape,
        LayoutDirection layoutDirection,
        bool isFirstGroup,
        long groupOffset,
        long groupLength,
        BlockGroup blockGroup
    )
    {
        var isFirstBlock = true;
        var origin = layoutDirection.Choose(
            row: bodyRegion.X,
            column: bodyRegion.Y
        );
        var totalLength = layoutDirection.Choose(
            row: bodyRegion.Width,
            column: bodyRegion.Height
        );
        var offsets = ComputeBlockOffsets(
            layoutDirection,
            totalLength,
            blockGroup
        );

        foreach (
            var (block, (offset, length)) in blockGroup.Blocks.Zip(offsets)
        )
        {
            var x = layoutDirection.Choose(
                row: origin + offset,
                column: groupOffset
            );
            var y = layoutDirection.Choose(
                row: groupOffset,
                column: origin + offset
            );
            var width = layoutDirection.Choose(
                row: length,
                column: groupLength
            );
            var height = layoutDirection.Choose(
                row: groupLength,
                column: length
            );

            RenderBlock(
                slidePart,
                slideParagraphStyles,
                templateShape,
                isFirstGroup && isFirstBlock,
                x,
                y,
                width,
                height,
                block
            );

            isFirstBlock = false;
        }
    }

    /// <summary>
    /// Computes the offsets and lengths of the blocks.
    ///
    /// If LayoutDirection is "row", computes the X positions and widths of
    /// the rows. If LayoutDirection is "column", computes the Y positions and
    /// heights of the columns.
    /// </summary>
    /// <param name="totalLength">The total length of the block group.</param>
    /// <param name="blockGroup">
    /// The block group to compute the offsets and lengths.
    /// </param>
    /// <returns>The offsets and length of the blocks.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the blocks do not fit into the space.
    /// </exception>
    private static IReadOnlyList<(long Start, long Length)> ComputeBlockOffsets(
        LayoutDirection layoutDirection,
        long totalLength,
        BlockGroup blockGroup
    )
    {
        try
        {
            return ILengthWeight
                .Divide(
                    totalLength,
                    blockGroup
                        .Blocks
                        .Select(
                            block => (
                                Weight: block.Weight,
                                Gap: block.GapBefore?.ToEmu() ?? 0
                            )
                        )
                        .ToList()
                );
        }
        catch (ArgumentException e)
        {
            throw new PlotanceException(
                blockGroup.Path,
                blockGroup.Line,
                layoutDirection.Choose(
                    row: "Columns does not fit into the space.",
                    column: "Rows does not fit into the space."
                ),
                e
            );
        }
    }

    /// <summary>Renders a block to a slide.</summary>
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
    /// <param name="blockContainer">The block container to render.</param>
    /// <exception cref="PlotanceException">
    /// Thrown when the block has an invalid structure, the configuration is
    /// invalid, image URL is invalid, image format is not supported, or the
    /// image data is invalid.
    /// </exception>
    private static void RenderBlock(
        SlidePart slidePart,
        SlideParagraphStyles slideParagraphStyles,
        Shape templateShape,
        bool isFirst,
        long x,
        long y,
        long width,
        long height,
        BlockContainer blockContainer
    )
    {
        var block = blockContainer.Block;

        if (
            block.GetData("query_results")
                is IReadOnlyList<QueryResultSet> queryResults
        )
        {
            if (blockContainer.ChartOptions?.Format?.Value == "table")
            {
                var bodyFontScale = blockContainer.BodyFontScale?.Value ?? 1m;

                TableRenderer.Render(
                    slidePart,
                    slideParagraphStyles.Body.Scaled(bodyFontScale),
                    isFirst,
                    x,
                    y,
                    width,
                    height,
                    blockContainer,
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
                    blockContainer,
                    queryResults
                );
            }
        }
        else if (
            SlideImages.ExtractImageLink(blockContainer.Variables, block)
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
                Geometries
                    .ParseHorizontalAlign(blockContainer.BodyHorizontalAlign)
                    ?? D.TextAlignmentTypeValues.Left,
                Geometries
                    .ParseVerticalAlign(blockContainer.BodyVerticalAlign)
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
                blockContainer
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
    /// <param name="block">The block to render.</param>
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
        BlockContainer blockContainer
    )
    {
        var bodyFontScale = blockContainer.BodyFontScale?.Value ?? 1m;
        var shape = Shapes.CreateShapeFromTemplate(templateShape);

        if (
            bodyFontScale != 1m
                || blockContainer.Language != null
                || blockContainer.BodyHorizontalAlign != null
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
                .WithLanguage(blockContainer.Language);

            if (
                Geometries.ParseHorizontalAlign(blockContainer.BodyHorizontalAlign)
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
                blockContainer.Variables,
                ExtractPath(blockContainer.Block),
                blockContainer.Block
            )
        );

        if (
            Geometries.ParseVerticalAlign(blockContainer.BodyVerticalAlign)
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
            SlideLayouts.ExtractPlaceholderShape(shape)?.Remove();

            var shapeProperties = (
                shape.ShapeProperties ??= new ShapeProperties()
            );

            shapeProperties.AddChild(
                new D.PresetGeometry()
                {
                    Preset = D.ShapeTypeValues.Rectangle
                }
            );
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
