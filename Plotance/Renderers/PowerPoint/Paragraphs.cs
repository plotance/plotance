// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Markdig.Helpers;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Plotance.Models;

using D = DocumentFormat.OpenXml.Drawing;

namespace Plotance.Renderers.PowerPoint;

/// <summary>
/// Provides methods for converting Markdown blocks to OpenXml paragraphs for
/// PowerPoint slides.
/// </summary>
public static class Paragraphs
{
    /// <summary>
    /// Creates a list of font elements used for rendering code blocks in
    /// slides.
    /// </summary>
    private static List<OpenXmlElement> CodeFonts => [
        new D.LatinFont() { Typeface = "Courier New" },
        new D.EastAsianFont() { Typeface = "Courier New" },
        new D.ComplexScriptFont() { Typeface = "Courier New" }
    ];

    /// <summary>
    /// Creates OpenXml paragraph elements from a Markdown block.
    /// </summary>
    /// <param name="slidePart">The slide part to add the paragraphs to.</param>
    /// <param name="variables">
    /// Dictionary of variables to expand in text content.
    /// </param>
    /// <param name="path">
    /// The path to the file containing the block, for error reporting.
    /// </param>
    /// <param name="block">The Markdown block to convert to paragraphs.</param>
    /// <returns>A collection of OpenXml paragraph elements.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the block has an invalid structure.
    /// </exception>
    public static IEnumerable<D.Paragraph> CreateParagraphs(
        SlidePart slidePart,
        IReadOnlyDictionary<string, string> variables,
        string path,
        Block block
    )
    {
        D.ParagraphProperties CreateLeafParagraphProperties()
        {
            var paragraphProperties = new D.ParagraphProperties(
                new D.NoBullet()
            )
            {
                Indent = 0,
                LeftMargin = 0
            };

            if (block is CodeBlock)
            {
                paragraphProperties.AddChild(
                    new D.DefaultRunProperties(CodeFonts)
                );
            }

            return paragraphProperties;
        }

        switch (block)
        {
            case LeafBlock { Inline: ContainerInline inline }:
                return [
                    new D.Paragraph([
                        CreateLeafParagraphProperties(),
                        .. inline.SelectMany(
                            inline => CreateRuns(
                                slidePart,
                                variables,
                                path,
                                inline
                            )
                        )
                    ])
                ];

            case LeafBlock { Lines: StringLineGroup lineGroup }:
                return lineGroup.Lines.Select(
                    line => new D.Paragraph(
                        CreateLeafParagraphProperties(),
                        new D.Run(
                            new D.Text(
                                Variables.ExpandVariables(
                                    line.ToString(),
                                    variables
                                )
                            )
                        )
                    )
                );

            case ListBlock listBlock:
                return CreateListParagraphs(
                    slidePart,
                    variables,
                    path,
                    listBlock
                );

            case ContainerBlock containerBlock:
                return containerBlock.SelectMany(
                    child => CreateParagraphs(slidePart, variables, path, child)
                );

            default:
                return [];
        }
    }

    /// <summary>
    /// Creates OpenXml paragraph elements from a Markdown list block.
    /// </summary>
    /// <remarks>
    /// The list items can have at most one paragraph and at most one nested
    /// list, in order.
    /// </remarks>
    /// <param name="slidePart">The slide part to add the paragraphs to.</param>
    /// <param name="variables">
    /// Dictionary of variables to expand in text content.
    /// </param>
    /// <param name="path">
    /// The path to the file containing the block, for error reporting.
    /// </param>
    /// <param name="listBlock">
    /// The Markdown list block to convert to paragraphs.
    /// </param>
    /// <param name="level">
    /// The nesting level of the list, used for indentation.
    /// </param>
    /// <returns>
    /// A collection of OpenXml paragraph elements representing the list.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the list item has an invalid structure.
    /// </exception>
    private static IEnumerable<D.Paragraph> CreateListParagraphs(
        SlidePart slidePart,
        IReadOnlyDictionary<string, string> variables,
        string path,
        ListBlock listBlock,
        int level = 0
    )
    {
        var listItems = listBlock.OfType<ListItemBlock>();
        var startAt = listBlock.IsOrdered
            ? listItems.FirstOrDefault()?.Order
            : null;

        return listItems
            .SelectMany(
                listItemBlock => CreateListItemParagraphs(
                    slidePart,
                    variables,
                    path,
                    listBlock,
                    listItemBlock,
                    startAt,
                    level
                )
            );
    }

    /// <summary>
    /// Creates OpenXml paragraph elements from a Markdown list item block.
    /// </summary>
    /// <param name="slidePart">The slide part to add the paragraphs to.</param>
    /// <param name="variables">
    /// Dictionary of variables to expand in text content.
    /// </param>
    /// <param name="path">
    /// The path to the file containing the block, for error reporting.
    /// </param>
    /// <param name="listBlock">
    /// The parent list block containing this list item.
    /// </param>
    /// <param name="listItemBlock">
    /// The Markdown list item block to convert to paragraphs.
    /// </param>
    /// <param name="startAt">
    /// For ordered lists, the starting number of the list.
    /// </param>
    /// <param name="level">
    /// The nesting level of the list, used for indentation.
    /// </param>
    /// <returns>
    /// A collection of OpenXml paragraph elements representing the list item.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the list item has an invalid structure.
    /// </exception>
    private static IEnumerable<D.Paragraph> CreateListItemParagraphs(
        SlidePart slidePart,
        IReadOnlyDictionary<string, string> variables,
        string path,
        ListBlock listBlock,
        ListItemBlock listItemBlock,
        int? startAt,
        int level
    )
    {
        bool IsVisibleInline(Inline inline) => inline switch
        {
            LiteralInline literal
                => !String.IsNullOrWhiteSpace(literal.ToString()),

            HtmlInline html
                => !html.Tag.StartsWith("<!--")
                && !html.Tag.StartsWith("<?"),

            LineBreakInline lineBreak
                => false,

            _ => true
        };

        bool IsVisibleChild(Block block) => block switch
        {
            ParagraphBlock paragraph
                => paragraph
                .Inline
                ?.Any(IsVisibleInline)
                ?? false,

            HtmlBlock html
                => html.Type != HtmlBlockType.Comment
                && html.Type != HtmlBlockType.ProcessingInstruction,

            LinkReferenceDefinition
                => false,

            _ => true
        };

        var children = listItemBlock.Where(IsVisibleChild).ToList();

        if (listItemBlock.Count == 0)
        {
            return [new D.Paragraph()];
        }

        var (paragraphBlock, nestedListBlock) = children switch
        {
            [ParagraphBlock p] => (p, null),
            [ListBlock l] => (null, l),
            [ParagraphBlock p, ListBlock l] => (p, l),
            _ => throw new PlotanceException(
                path,
                listItemBlock.Line,
                "List item can contain at most one paragraph and at most one"
                    + " list in order."
            )
        };
        var paragraphs = (
            paragraphBlock == null
                ? []
                : CreateParagraphs(
                    slidePart,
                    variables,
                    path,
                    paragraphBlock
                )
        )
            .Select(
                paragraph =>
                {
                    var paragraphProperties = paragraph.ParagraphProperties
                        ??= new D.ParagraphProperties();

                    paragraphProperties.RemoveAllChildren<D.NoBullet>();
                    paragraphProperties.Level = level;
                    paragraphProperties.Indent = null;
                    paragraphProperties.LeftMargin = null;

                    if (listBlock.IsOrdered)
                    {
                        var bulletType = D.TextAutoNumberSchemeValues
                            .ArabicPeriod;

                        paragraphProperties.AddChild(
                            new D.AutoNumberedBullet()
                            {
                                Type = bulletType,
                                StartAt = startAt
                            }
                        );
                    }

                    return paragraph;
                }
            );
        var nestedListParagraphs = nestedListBlock == null
            ? []
            : CreateListParagraphs(
                slidePart,
                variables,
                path,
                nestedListBlock,
                level + 1
            );

        return paragraphs.Concat(nestedListParagraphs);
    }

    /// <summary>
    /// Creates OpenXml run elements from a Markdown inline element.
    /// </summary>
    /// <param name="slidePart">The slide part to add the runs to.</param>
    /// <param name="variables">
    /// Dictionary of variables to expand in text content.
    /// </param>
    /// <param name="path">
    /// The path to the file containing the inline, for error reporting.
    /// </param>
    /// <param name="inline">
    /// The Markdown inline element to convert to runs.
    /// </param>
    /// <returns>
    /// A collection of OpenXml elements representing the inline content.
    /// </returns>
    private static IEnumerable<OpenXmlElement> CreateRuns(
        SlidePart slidePart,
        IReadOnlyDictionary<string, string> variables,
        string path,
        Inline inline
    ) => CreateRuns(slidePart, variables, path, inline, new());

    /// <summary>
    /// Creates OpenXml run elements from a Markdown inline element with the
    /// specified run properties.
    /// </summary>
    /// <param name="slidePart">The slide part to add the runs to.</param>
    /// <param name="variables">
    /// Dictionary of variables to expand in text content.
    /// </param>
    /// <param name="path">
    /// The path to the file containing the inline, for error reporting.
    /// </param>
    /// <param name="inline">
    /// The Markdown inline element to convert to runs.
    /// </param>
    /// <param name="runPropertiesChildren">
    /// The run properties to apply to the created runs.
    /// </param>
    /// <returns>
    /// A collection of OpenXml elements representing the inline content with
    /// styling.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown when an inline image is encountered, which is not supported.
    /// </exception>
    private static IEnumerable<OpenXmlElement> CreateRuns(
        SlidePart slidePart,
        IReadOnlyDictionary<string, string> variables,
        string path,
        Inline inline,
        RunPropertiesChildren runPropertiesChildren
    )
    {
        D.Run CreateLeafRun(
            string text,
            RunPropertiesChildren runPropertiesChildren
        )
        {
            var run = new D.Run(
                new D.Text(Variables.ExpandVariables(text, variables))
            );
            var runProperties = runPropertiesChildren.CreateRunProperties();

            if (runProperties != null)
            {
                run.AddChild(runProperties);
            }

            return run;
        }

        R WithHyperlinkOnClick<R>(
            string url,
            string? title,
            Func<RunPropertiesChildren, R> body
        ) => runPropertiesChildren.WithValue(
            CreateHyperlinkOnClick(slidePart, url, title),
            body
        );

        return inline switch
        {
            AutolinkInline autolink => [
                WithHyperlinkOnClick(
                    autolink.Url,
                    null,
                    c => CreateLeafRun("<" + autolink.Url + ">", c)
                )
            ],

            CodeInline code => [
                runPropertiesChildren.WithValues(
                    CodeFonts,
                    c => CreateLeafRun(code.Content, c)
                )
            ],

            EmphasisInline emphasis => runPropertiesChildren.WithValue(
                emphasis.DelimiterCount == 2
                    ? runProperties => runProperties.Bold = true
                    : runProperties => runProperties.Italic = true,
                runPropertiesChildren => emphasis.SelectMany(
                    child => CreateRuns(
                        slidePart,
                        variables,
                        path,
                        child,
                        runPropertiesChildren
                    )
                ).ToList()
            ),

            HtmlEntityInline htmlEntity => [
                CreateLeafRun(
                    htmlEntity.Transcoded.ToString(),
                    runPropertiesChildren
                )
            ],

            HtmlInline html => [
                runPropertiesChildren.WithValues(
                    CodeFonts,
                    c => CreateLeafRun(html.Tag, c)
                )
            ],

            LineBreakInline lineBreak => [
                lineBreak.IsHard
                    ? new D.Break()
                    : CreateLeafRun(" ", runPropertiesChildren)
            ],

            LinkInline image when image.IsImage
                => throw new PlotanceException(
                    path,
                    image.Line,
                    "Inline image is not supported."
                ),

            LinkInline link => WithHyperlinkOnClick(
                link.Url ?? "",
                link.Title,
                runPropertiesChildren => link.SelectMany(
                    child => CreateRuns(
                        slidePart,
                        variables,
                        path,
                        child,
                        runPropertiesChildren
                    )
                ).ToList()
            ),

            LiteralInline literal => [
                CreateLeafRun(literal.ToString(), runPropertiesChildren)
            ],

            ContainerInline container => container.SelectMany(
                child => CreateRuns(
                    slidePart,
                    variables,
                    path,
                    child,
                    runPropertiesChildren
                )
            ).ToList(),

            _ => [
                CreateLeafRun(inline.ToString() ?? "", runPropertiesChildren)
            ],
        };
    }

    /// <summary>
    /// Creates a hyperlink for a PowerPoint slide that activates when clicked.
    /// </summary>
    /// <param name="slidePart">
    /// The slide part to add the hyperlink relationship to.
    /// </param>
    /// <param name="url">The URL that the hyperlink points to.</param>
    /// <param name="tooltip">
    /// The optional tooltip text to display when hovering over the hyperlink.
    /// </param>
    /// <returns>A HyperlinkOnClick object.</returns>
    public static D.HyperlinkOnClick CreateHyperlinkOnClick(
        SlidePart slidePart,
        string url,
        string? tooltip
    )
    {
        var hyperlinkRelationship = slidePart
            .AddHyperlinkRelationship(new Uri(url), true);

        var hyperlinkOnClick = new D.HyperlinkOnClick()
        {
            Id = hyperlinkRelationship.Id,
        };

        if (!string.IsNullOrEmpty(tooltip))
        {
            hyperlinkOnClick.Tooltip = tooltip;
        }

        return hyperlinkOnClick;
    }

    /// <summary>Converts a Markdown inline element to plain text.</summary>
    /// <param name="variables">
    /// Dictionary of variables to expand in text content.
    /// </param>
    /// <param name="inline">The Markdown inline element to convert.</param>
    /// <returns>The plain text representation of the inline element.</returns>
    public static string ToPlainText(
        IReadOnlyDictionary<string, string> variables,
        Inline inline
    )
    {
        var builder = new StringBuilder();

        void AppendText(string text)
        {
            builder.Append(
                Variables.ExpandVariables(text, variables) ?? string.Empty
            );
        }

        void AppendInline(Inline inline)
        {
            switch (inline)
            {
                case LiteralInline literal:
                    AppendText(literal.Content.ToString());
                    break;

                case CodeInline code:
                    AppendText(code.Content);
                    break;

                case LineBreakInline:
                    AppendText(" ");
                    break;

                case AutolinkInline autolink:
                    AppendText("<" + autolink.Url + ">");
                    break;

                case HtmlEntityInline entity:
                    AppendText(entity.Transcoded.ToString());
                    break;

                case HtmlInline html:
                    AppendText(html.Tag);
                    break;

                case ContainerInline container:
                    foreach (var child in container)
                    {
                        AppendInline(child);
                    }
                    break;

                default:
                    AppendText(inline.ToString() ?? string.Empty);
                    break;
            }
        }

        AppendInline(inline);

        return builder.ToString();
    }
}
