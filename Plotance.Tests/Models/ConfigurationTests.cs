// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using Markdig.Helpers;
using Markdig.Syntax;
using Plotance.Models;

namespace Plotance.Tests.Models;

public class ConfigurationTests
{
    [Fact]
    public void TryCreate_WithValidFencedCodeBlock_ReturnsConfig()
    {
        // Arrange
        var codeBlock = new FencedCodeBlock(null!)
        {
            Info = "plotance",
            Lines = new StringLineGroup("""
                  slide_level: 3
                  rows: 1:2.1:3
                  columns: 4:5cm:6
                  body_font_scale: 1.5
                  title_font_scale: 2.0
                  language: en-GB
                """)
        };

        // Act
        var config = Configuration.TryCreate("test.md", codeBlock);

        // Assert
        Assert.NotNull(config);
        Assert.Equal(3, config.SlideLevel?.Value);
        Assert.Equal("test.md", config.SlideLevel?.Path);
        Assert.Equal(2, config.SlideLevel?.Line); // +1 for the code fence
        Assert.Equal(
            new List<ILengthWeight> {
                new RelativeLengthWeight("test.md", 3, 1m),
                new RelativeLengthWeight("test.md", 3, 2.1m),
                new RelativeLengthWeight("test.md", 3, 3m)
            },
            config.Rows?.Value
        );
        Assert.Equal(
            new List<ILengthWeight> {
                new RelativeLengthWeight("test.md", 4, 4m),
                new AbsoluteLengthWeight("test.md", 4, "5cm"),
                new RelativeLengthWeight("test.md", 4, 6m)
            },
            config.Columns?.Value
        );
        Assert.Equal(1.5m, config.BodyFontScale?.Value);
        Assert.Equal(2.0m, config.TitleFontScale?.Value);
        Assert.Equal("en-GB", config.Language?.Value);
        Assert.Equal("test.md", config.Language?.Path);
        Assert.Equal(7, config.Language?.Line);
    }

    [Fact]
    public void TryCreate_WithInvalidFencedCodeBlock_ReturnsNull()
    {
        // Arrange
        var codeBlock = new FencedCodeBlock(null!)
        {
            Info = "not-plotance",
            Lines = new StringLineGroup("slide_level: 3")
        };

        // Act
        var config = Configuration.TryCreate("test.md", codeBlock);

        // Assert
        Assert.Null(config);
    }

    [Fact]
    public void TryCreate_WithSingleLineProcessingInstruction_ReturnsConfig()
    {
        // Arrange
        var htmlBlock = new HtmlBlock(null)
        {
            Type = HtmlBlockType.ProcessingInstruction,
            Lines = new StringLineGroup(
                "<?plotance rows: 1:2:3, columns: 4:5:6 ?>"
            )
        };

        // Act
        var config = Configuration.TryCreate("test.md", htmlBlock);

        // Assert
        Assert.NotNull(config);
        Assert.Equal(
            new List<ILengthWeight> {
                new RelativeLengthWeight("test.md", 1, 1m),
                new RelativeLengthWeight("test.md", 1, 2m),
                new RelativeLengthWeight("test.md", 1, 3m)
            },
            config.Rows?.Value
        );
        Assert.Equal(
            new List<ILengthWeight> {
                new RelativeLengthWeight("test.md", 1, 4m),
                new RelativeLengthWeight("test.md", 1, 5m),
                new RelativeLengthWeight("test.md", 1, 6m)
            },
            config.Columns?.Value
        );
    }

    [Theory]
    [InlineData(
        """
        <?plotance
          slide_level: 3
          rows: 1:2:3
          columns: 4:5:6
        ?>
        """
    )]
    [InlineData(
        """
        <?plotance slide_level: 3
                          rows: 1:2:3
                          columns: 4:5:6?>
        """
    )]
    public void TryCreate_WithMultiLineProcessingInstruction_ReturnsConfig(
        string text
    )
    {
        // Arrange
        var htmlBlock = new HtmlBlock(null)
        {
            Type = HtmlBlockType.ProcessingInstruction,
            Lines = new StringLineGroup(text)
        };

        // Act
        var config = Configuration.TryCreate("test.md", htmlBlock);

        decimal? GetRelativeWeight(ILengthWeight? lengthWeight)
            => lengthWeight is RelativeLengthWeight w
            ? w.Weight
            : null;

        // Assert
        Assert.NotNull(config);
        Assert.Equal(3, config.SlideLevel?.Value);
        Assert.Equal(1m, GetRelativeWeight(config.Rows?.Value?[0]));
        Assert.Equal(2m, GetRelativeWeight(config.Rows?.Value?[1]));
        Assert.Equal(3m, GetRelativeWeight(config.Rows?.Value?[2]));
        Assert.Equal(4m, GetRelativeWeight(config.Columns?.Value?[0]));
        Assert.Equal(5m, GetRelativeWeight(config.Columns?.Value?[1]));
        Assert.Equal(6m, GetRelativeWeight(config.Columns?.Value?[2]));
    }

    [Fact]
    public void TryCreate_WithInvalidProcessingInstruction_ReturnsNull()
    {
        // Arrange
        var htmlBlock = new HtmlBlock(null)
        {
            Type = HtmlBlockType.ProcessingInstruction,
            Lines = new StringLineGroup("<?not-plotance rows: 1:2:3 ?>")
        };

        // Act
        var config = Configuration.TryCreate("test.md", htmlBlock);

        // Assert
        Assert.Null(config);
    }

    private Configuration ParseConfig(string? path, string text)
    {
        var codeBlock = new FencedCodeBlock(null!)
        {
            Info = "plotance",
            Lines = new StringLineGroup(text)
        };

        return Configuration.TryCreate(path, codeBlock)
            ?? throw new Exception("Failed to parse config");
    }

    [Fact]
    public void Update_WithNewValues_UpdatesExistingConfig()
    {
        // Arrange
        var config = ParseConfig("test1.md", """
          data_source: original_source
          slide_level: 2
          body_font_scale: 1.0
          title_font_scale: 1.0
          language: en-US
          db_config:
            key1: value1
        """);

        var newConfig = ParseConfig("test2.md", """

          data_source: new_source
          slide_level: 3
          body_font_scale: 1.5
          title_font_scale: 2.0
          language: en-GB
          db_config:
            key2: value2
        """);

        // Act
        config.Update(newConfig);

        // Assert
        Assert.Equal("new_source", config.DataSource?.Value);
        Assert.Equal("test2.md", config.DataSource?.Path);
        Assert.Equal(3, config.DataSource?.Line);
        Assert.Equal(3, config.SlideLevel?.Value);
        Assert.Equal(1.5m, config.BodyFontScale?.Value);
        Assert.Equal(2.0m, config.TitleFontScale?.Value);
        Assert.Equal("en-GB", config.Language?.Value);
        Assert.Equal(2, config.DbConfig?.KeyValues.Count);
        Assert.Equal("value1", config.DbConfig?["key1"]?.Value);
        Assert.Equal("test1.md", config.DbConfig?["key1"]?.Path);
        Assert.Equal(8, config.DbConfig?["key1"]?.Line);
        Assert.Equal("value2", config.DbConfig?["key2"]?.Value);
        Assert.Equal("test2.md", config.DbConfig?["key2"]?.Path);
        Assert.Equal(9, config.DbConfig?["key2"]?.Line);
    }

    [Fact]
    public void Update_WithNullValues_KeepsOriginalValues()
    {
        // Arrange
        var config = ParseConfig("test.md", """
          data_source: original_source
          slide_level: 2
          body_font_scale: 1.0
          title_font_scale: 1.0
          language: en-US
        """);

        var newConfig = new Configuration(); // All properties are null

        // Act
        config.Update(newConfig);

        // Assert
        Assert.Equal("original_source", config.DataSource?.Value);
        Assert.Equal("test.md", config.DataSource?.Path);
        Assert.Equal(2, config.DataSource?.Line);
        Assert.Equal(2, config.SlideLevel?.Value);
        Assert.Equal(1.0m, config.BodyFontScale?.Value);
        Assert.Equal(1.0m, config.TitleFontScale?.Value);
        Assert.Equal("en-US", config.Language?.Value);
    }

    [Fact]
    public void Update_WithOverlappingConfigKeys_UpdatesExistingKeys()
    {
        // Arrange
        var config = ParseConfig("test1.md", """
          db_config:
            key1: value1
            key2: original_value2
        """);

        var newConfig = ParseConfig("test2.md", """


          db_config:
            key2: new_value2
            key3: value3
        """);

        // Act
        config.Update(newConfig);

        // Assert
        Assert.Equal(3, config.DbConfig?.KeyValues?.Count);
        Assert.Equal("test1.md", config.DbConfig?.Path);
        // +1 for the code fence
        Assert.Equal(3, config.DbConfig?.Line);
        Assert.Equal("value1", config.DbConfig?["key1"]?.Value);
        Assert.Equal("test1.md", config.DbConfig?["key1"]?.Path);
        // Updated value
        Assert.Equal("new_value2", config.DbConfig?["key2"]?.Value);
        Assert.Equal("test2.md", config.DbConfig?["key2"]?.Path);
        Assert.Equal(5, config.DbConfig?["key2"]?.Line);
        // New key
        Assert.Equal("value3", config.DbConfig?["key3"]?.Value);
        Assert.Equal("test2.md", config.DbConfig?["key3"]?.Path);
        Assert.Equal(6, config.DbConfig?["key3"]?.Line);
    }

    [Fact]
    public void TryCreate_WithSeriesColors_ParsesStringList()
    {
        // Arrange
        var config = ParseConfig("test.md", """
          series_colors:
            - dark1
            - light1
            - accent1
        """);

        // Assert
        Assert.NotNull(config);
        Assert.NotNull(config.ChartOptions.SeriesColors?.Value);
        Assert.Equal(3, config.ChartOptions.SeriesColors?.Value.Count);
        Assert.Equal("test.md", config.ChartOptions.SeriesColors?.Path);
        // +1 for the code fence
        Assert.Equal(3, config.ChartOptions.SeriesColors?.Line);
        Assert.Equal(
            "dark1",
            config.ChartOptions.SeriesColors?.Value[0]?.Text
        );
        Assert.Equal(
            "test.md",
            config.ChartOptions.SeriesColors?.Value[0]?.Path
        );
        Assert.Equal(3, config.ChartOptions.SeriesColors?.Value[0]?.Line);
        Assert.Equal(
            "light1",
            config.ChartOptions.SeriesColors?.Value[1]?.Text
        );
        Assert.Equal(4, config.ChartOptions.SeriesColors?.Value[1]?.Line);
        Assert.Equal(
            "accent1",
            config.ChartOptions.SeriesColors?.Value[2]?.Text
        );
        Assert.Equal(5, config.ChartOptions.SeriesColors?.Value[2]?.Line);
    }

    [Fact]
    public void TryCreate_WithSeriesColorsAsString_ParsesStringList()
    {
        // Arrange
        var config = ParseConfig(
            "test.md",
            "series_colors: dark1, light1, accent1"
        );

        // Assert
        Assert.NotNull(config);
        Assert.NotNull(config.ChartOptions.SeriesColors?.Value);
        Assert.Equal(3, config.ChartOptions.SeriesColors?.Value.Count);
        Assert.Equal(
            "dark1",
            config.ChartOptions.SeriesColors?.Value[0]?.Text
        );
        Assert.Equal(
            "light1",
            config.ChartOptions.SeriesColors?.Value[1]?.Text
        );
        Assert.Equal(
            "accent1",
            config.ChartOptions.SeriesColors?.Value[2]?.Text
        );
    }

    [Fact]
    public void Update_WithSeriesColors_UpdatesExistingConfig()
    {
        // Arrange
        var config = ParseConfig(
            "test1.md",
            "series_colors: [dark1, light1]"
        );

        var newConfig = ParseConfig(
            "test2.md",
            "series_colors: [dark2, light2]"
        );

        // Act
        config.Update(newConfig);

        // Assert
        Assert.NotNull(config.ChartOptions.SeriesColors?.Value);
        Assert.Equal(2, config.ChartOptions.SeriesColors?.Value.Count);
        Assert.Equal(
            "dark2",
            config.ChartOptions.SeriesColors?.Value[0]?.Text
        );
        Assert.Equal(
            "test2.md",
            config.ChartOptions.SeriesColors?.Value[0]?.Path
        );
        Assert.Equal(
            "light2",
            config.ChartOptions.SeriesColors?.Value[1]?.Text
        );
        Assert.Equal(
            "test2.md",
            config.ChartOptions.SeriesColors?.Value[1]?.Path
        );
    }

    [Fact]
    public void Update_WithParameters_UpdatesExistingConfig()
    {
        // Arrange
        var config = ParseConfig("test1.md", """
          parameters:
            - name: name1
            - name: name2
        """);

        var newConfig = ParseConfig("test2.md", """
          parameters:
            - name: name3
            - name: name4
        """);

        // Act
        config.Update(newConfig);

        // Assert
        Assert.NotNull(config.Parameters?.Value);
        Assert.Equal(4, config.Parameters?.Value.Count);
        Assert.Equal("name1", config.Parameters?.Value[0]?.Name?.Value);
        Assert.Equal("test1.md", config.Parameters?.Value[0]?.Name?.Path);
        Assert.Equal(3, config.Parameters?.Value[0]?.Name?.Line);
        Assert.Equal("name2", config.Parameters?.Value[1]?.Name?.Value);
        Assert.Equal("test1.md", config.Parameters?.Value[1]?.Name?.Path);
        Assert.Equal(4, config.Parameters?.Value[1]?.Name?.Line);
        Assert.Equal("name3", config.Parameters?.Value[2]?.Name?.Value);
        Assert.Equal("test2.md", config.Parameters?.Value[2]?.Name?.Path);
        Assert.Equal(3, config.Parameters?.Value[2]?.Name?.Line);
        Assert.Equal("name4", config.Parameters?.Value[3]?.Name?.Value);
        Assert.Equal("test2.md", config.Parameters?.Value[3]?.Name?.Path);
        Assert.Equal(4, config.Parameters?.Value[3]?.Name?.Line);
    }
}
