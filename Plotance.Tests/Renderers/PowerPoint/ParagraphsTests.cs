// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using Markdig;
using Markdig.Helpers;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Plotance.Renderers.PowerPoint;

namespace Plotance.Tests.Renderers.PowerPoint;

public static class MarkdownHelper
{
    public static Inline ParseInline(string markdown)
    {
        var pipeline = new MarkdownPipelineBuilder().Build();
        var document = Markdown.Parse(markdown, pipeline);
        var paragraph = (ParagraphBlock)document.First();

        return paragraph.Inline ?? throw new Exception("Inline is null");
    }
}

public class ParagraphsTests
{
    [Theory]
    [InlineData("text", "text")]
    [InlineData("**bold** *italic*", "bold italic")]
    [InlineData("**_bold and italic_**", "bold and italic")]
    [InlineData("`code`", "code")]
    [InlineData("&amp;", "&")]
    [InlineData("aaa<br>bbb", "aaa<br>bbb")]
    [InlineData("aaa\\\nbbb", "aaa bbb")]
    [InlineData("[link](https://example.org)", "link")]
    [InlineData("<https://example.org>", "<https://example.org>")]
    public void ToPlainText_WithInline_ReturnsExpected(
        string markdown,
        string expected
    )
    {
        // Arrange
        var inline = MarkdownHelper.ParseInline(markdown);

        // Act
        var text = Paragraphs.ToPlainText(
            new Dictionary<string, string>(),
            inline
        );

        // Assert
        Assert.Equal(expected, text);
    }

    [Fact]
    public void ToPlainText_WithVariable_ExpandsVariables()
    {
        // Arrange
        var inline = MarkdownHelper.ParseInline(
            "${var} `${var}` **${var}** [${var}](https://example.org)"
        );

        // Act
        var text = Paragraphs.ToPlainText(
            new Dictionary<string, string>
            {
                { "var", "value" }
            },
            inline
        );

        // Assert
        Assert.Equal("value value value value", text);
    }
}

