// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using Markdig.Syntax;
using YamlDotNet.Core;
using YamlDotNet.RepresentationModel;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace Plotance.Models;

/// <summary>
/// Represents the configuration for the plotter, containing all the settings
/// needed to generate charts and other visualization elements.
/// </summary>
public class Configuration
{
    /// <summary>
    /// The set of info strings that are recognized as plotter configuration
    /// code blocks.
    /// </summary>
    public static readonly ISet<string> ConfigBlockNames = new HashSet<string>
    {
        "plotance"
    };

    /// <summary>
    /// The set of configuration keys that are treated as comma-separated string
    /// lists when they appear as scalar values.
    /// </summary>
    private static readonly ISet<string> StringListKeys = new HashSet<string>
    {
        "series_colors",
        "data_label_contents"
    };

    /// <summary>
    /// The set of configuration keys that are treated as appending lists when
    /// merging configurations.
    /// </summary>
    private static readonly ISet<string> AppendingListKeys = new HashSet<string>
    {
        "parameters"
    };

    /// <summary>The root configuration node containing all settings.</summary>
    public MappingConfigNode Root { get; private set; }

    /// <summary>
    /// Initializes a new instance of the <see cref="Configuration"/> class with
    /// the specified root configuration node.
    /// </summary>
    /// <param name="root">The root configuration node.</param>
    public Configuration(MappingConfigNode root)
    {
        Root = root;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="Configuration"/> class with
    /// an empty configuration.
    /// </summary>
    public Configuration() : this(MappingConfigNode.Empty)
    {
    }

    // Path options

    /// <summary>
    /// The output path for the generated content. Relative to the current
    /// file.
    /// </summary>
    public ValueWithLocation<string>? Output => GetString("output");

    /// <summary>
    /// The template path to use for generation. Relative to the current
    /// file.
    /// </summary>
    public ValueWithLocation<string>? Template => GetString("template");

    // Database options

    /// <summary>The data source identifier for database connections.</summary>
    public ValueWithLocation<string>? DataSource => GetString("data_source");

    /// <summary>The database configuration properties.</summary>
    public DictionaryWithLocation<string, string>? DbConfig
        => GetStringDictionary("db_config");

    // Include options

    /// <summary>
    /// The path to a YAML file to include in the configuration, or a Markdown
    /// file to be inserted before the current block.
    /// </summary>
    public ValueWithLocation<string>? Include => GetString("include");

    // Slide options

    /// <summary>The heading level to use for slide breaks.</summary>
    public ValueWithLocation<int>? SlideLevel => GetInt("slide_level");

    /// <summary>The slide row height weights.</summary>
    public ValueWithLocation<IReadOnlyList<ILengthWeight>>? Rows
        => GetWeights("rows");

    /// <summary>The slide column width weights.</summary>
    public ValueWithLocation<IReadOnlyList<ILengthWeight>>? Columns
        => GetWeights("columns");

    /// <summary>The gap between slide rows.</summary>
    public ILength? RowGap => GetLength("row_gap");

    /// <summary>The gap between slide columns.</summary>
    public ILength? ColumnGap => GetLength("column_gap");

    /// <summary>The scale factor for body text font size.</summary>
    public ValueWithLocation<decimal>? BodyFontScale
        => GetDecimal("body_font_scale");

    /// <summary>The scale factor for title text font size.</summary>
    public ValueWithLocation<decimal>? TitleFontScale
        => GetDecimal("title_font_scale");

    /// <summary>The language code for content.</summary>
    public ValueWithLocation<string>? Language => GetString("language");

    /// <summary>
    /// The horizontal alignment for body content. Valid values are "left",
    /// "center", "right", "justified", "distributed".
    /// </summary>
    public ValueWithLocation<string>? BodyHorizontalAlign
        => GetString("body_horizontal_align");

    /// <summary>
    /// The vertical alignment for body content. Valid values are "top",
    /// "center", "bottom".
    /// </summary>
    public ValueWithLocation<string>? BodyVerticalAlign
        => GetString("body_vertical_align");

    // Query options

    /// <summary>
    /// The path to the file containing SQL queries. Relative to the current
    /// file.
    /// </summary>
    public ValueWithLocation<string>? QueryFile => GetString("query_file");

    /// <summary>The SQL query text.</summary>
    public ValueWithLocation<string>? Query => GetString("query");

    // Document parameters

    /// <summary>The list of parameter declarations for the document.</summary>
    public ValueWithLocation<IReadOnlyList<ParameterDeclaration>>? Parameters
        => GetParameterList("parameters");

    // Chart options

    /// <summary>The chart configuration options.</summary>
    public ChartOptions ChartOptions => new ChartOptions(this);

    /// <summary>
    /// Attempts to create a plotter configuration from a Markdown block.
    /// </summary>
    /// <param name="path">
    /// The path to the file containing the block, for error reporting.
    /// </param>
    /// <param name="block">The Markdown block to parse.</param>
    /// <returns>
    /// A Configuration if the block contains valid configuration, or null if
    /// the block is not a recognized configuration block.
    /// </returns>
    public static Configuration? TryCreate(
        string? path,
        Block block
    ) => TryCreate(path, block, new Dictionary<string, string>());

    /// <summary>
    /// Attempts to create a plotter configuration from a Markdown block with
    /// variable expansion.
    /// </summary>
    /// <param name="path">
    /// The path to the file containing the block, for error reporting.
    /// </param>
    /// <param name="block">The Markdown block to parse.</param>
    /// <param name="variables">
    /// Dictionary of variables to expand in the configuration.
    /// </param>
    /// <returns>
    /// A Configuration if the block contains valid configuration, or null if
    /// the block is not a recognized configuration block.
    /// </returns>
    /// <seealso cref="Variables.ExpandVariables"/>
    public static Configuration? TryCreate(
        string? path,
        Block block,
        IReadOnlyDictionary<string, string> variables
    ) => block switch
    {
        FencedCodeBlock fencedCodeBlock
            => TryCreate(path, fencedCodeBlock, variables),

        HtmlBlock htmlBlock
                when htmlBlock.Type == HtmlBlockType.ProcessingInstruction
            => TryCreate(path, htmlBlock, variables),

        _ => null
    };

    /// <summary>
    /// Attempts to create a plotter configuration from a Markdown fenced code
    /// block.
    /// </summary>
    /// <param name="path">
    /// The path to the file containing the block, for error reporting.
    /// </param>
    /// <param name="fencedCodeBlock">The fenced code block to parse.</param>
    /// <param name="variables">
    /// Dictionary of variables to expand in the configuration.
    /// </param>
    /// <returns>
    /// A Configuration if the block contains valid configuration, or null if
    /// the block is not a recognized configuration block.
    /// </returns>
    private static Configuration? TryCreate(
        string? path,
        FencedCodeBlock codeBlock,
        IReadOnlyDictionary<string, string> variables
    )
    {
        var info = codeBlock.Info?.Trim().ToLowerInvariant() ?? "";

        if (ConfigBlockNames.Contains(info))
        {
            // Note codeBlock.Line is zero-based.
            // +1 for the opening code fence.
            return Create(
                path,
                codeBlock.Lines.ToString(),
                codeBlock.Line + 1,
                variables
            );
        }
        else
        {
            return null;
        }
    }

    /// <summary>
    /// Attempts to create a plotter configuration from a Markdown HTML block.
    /// </summary>
    /// <param name="path">
    /// The path to the file containing the block, for error reporting.
    /// </param>
    /// <param name="htmlBlock">The HTML block to parse.</param>
    /// <param name="variables">
    /// Dictionary of variables to expand in the configuration.
    /// </param>
    /// <returns>
    /// A Configuration if the block contains valid configuration, or null if
    /// the block is not a recognized configuration block.
    /// </returns>
    private static Configuration? TryCreate(
        string? path,
        HtmlBlock htmlBlock,
        IReadOnlyDictionary<string, string> variables
    )
    {
        var content = htmlBlock.Lines.ToString();
        var match = Regex.Match(
            content,
            @"<\?plotance(?: +|(?=\n))(.+)\s*\?>$",
            RegexOptions.Singleline
        );

        if (!match.Success)
        {
            return null;
        }

        string yamlContent = match.Groups[1].Value;

        if (yamlContent.Contains('\n'))
        {
            // Trim minimum indents from second or following lines.
            // (excluding empty lines)
            var lines = yamlContent.Split('\n');
            var restLines = lines.Skip(1);
            var nonEmptyLines = restLines
                .Where(line => !string.IsNullOrWhiteSpace(line));

            if (nonEmptyLines.Any())
            {
                int minIndent = nonEmptyLines.Min(
                    line => Regex.Match(line, @"^ *").Length
                );

                yamlContent = lines[0] + "\n" + string.Join(
                    "\n",
                    restLines.Select(
                        line => string.IsNullOrWhiteSpace(line)
                            ? line
                            : line.Substring(minIndent)
                    )
                );
            }
        }
        else
        {
            // Wrap with { } if not.
            if (!yamlContent.StartsWith('{'))
            {
                yamlContent = "{" + yamlContent + "}";
            }
        }

        // Note htmlBlock.Line is zero-based.
        return Create(path, yamlContent, htmlBlock.Line, variables);
    }

    /// <summary>
    /// Creates a plotter configuration from YAML content with variable
    /// expansion.
    /// </summary>
    /// <param name="path">
    /// The path to the file containing the YAML, for error reporting.
    /// </param>
    /// <param name="source">The YAML content to parse.</param>
    /// <param name="lineOffset">
    /// The line number offset to add to all line numbers.
    /// </param>
    /// <param name="variables">
    /// Dictionary of variables to expand in the configuration.
    /// </param>
    /// <returns>A new Configuration instance.</returns>
    /// <seealso cref="Variables.ExpandVariables"/>
    public static Configuration Create(
        string? path,
        string source,
        long lineOffset,
        IReadOnlyDictionary<string, string> variables
    )
    {
        PlotanceException NewPlotanceException(Exception e, long line)
            => new PlotanceException(
                path,
                line + lineOffset,
                "Invalid YAML format.",
                e
            );

        try
        {
            var reader = new StringReader(source);
            var stream = new YamlStream();

            stream.Load(reader);

            var configs = stream
                .Documents
                .Select(
                    document => ConfigNode.Create(path, document, lineOffset)
                )
                .Select(node => node.ExpandVariables(variables))
                .Select(FixSequenceNodes)
                .Select(node => new Configuration(node));
            var config = new Configuration();

            foreach (var newConfig in configs)
            {
                config.Update(newConfig);
            }

            return config;
        }
        catch (YamlException e)
        {
            throw NewPlotanceException(e, e.Start.Line);
        }
        catch (Exception e)
            when (
                e is ArgumentException
                    or IOException
                    or InvalidOperationException
            )
        {
            throw NewPlotanceException(e, 1);
        }
    }

    /// <summary>
    /// Parse comma-separated list values of a mapping node into sequence nodes
    /// for StringListKeys.
    /// </summary>
    /// <param name="mappingNode">The mapping configuration node to fix.</param>
    /// <returns>The fixed mapping configuration node.</returns>
    private static MappingConfigNode FixSequenceNodes(
        MappingConfigNode mappingNode
    )
    {
        var newKeyValues = new Dictionary<string, ConfigNode>(
            mappingNode.KeyValues
        );

        foreach (var key in mappingNode.KeyValues.Keys)
        {
            if (
                StringListKeys.Contains(key)
                    && newKeyValues[key] is ScalarConfigNode scalarNode
            )
            {
                newKeyValues[key] = new SequenceConfigNode(
                    scalarNode.Path,
                    scalarNode.Line,
                    ParseCommaSeparatedList(scalarNode.Value)
                        .Select(
                            text => new ScalarConfigNode(
                                scalarNode.Path,
                                scalarNode.Line,
                                text
                            )
                        )
                        .ToList(),
                    AppendingListKeys.Contains(key)
                );
            }
            else if (
                AppendingListKeys.Contains(key)
                    && newKeyValues[key] is SequenceConfigNode sequenceNode
            )
            {
                newKeyValues[key] = sequenceNode with { Appending = true };
            }
        }

        return mappingNode with { KeyValues = newKeyValues };
    }

    /// <summary>
    /// Parses a comma-separated list into an enumerable of strings.
    /// </summary>
    /// <param name="list">The comma-separated list to parse.</param>
    /// <returns>An enumerable of strings.</returns>
    private static IEnumerable<string> ParseCommaSeparatedList(string list)
    {
        list = list.Trim();

        if (string.IsNullOrEmpty(list))
        {
            yield break;
        }

        var items = Regex.Split(list, @"[ \t]+(?:,[ \t]*)?|,[ \t]*");

        foreach (var item in items)
        {
            yield return item;
        }
    }

    /// <summary>
    /// Updates this configuration by merging it with another configuration.
    /// </summary>
    /// <param name="newConfig">
    /// The configuration to merge with this one.
    /// </param>
    public void Update(Configuration newConfig)
    {
        Root = Root.Merge(newConfig.Root);
    }

    /// <summary>Resets the font scale values in the configuration.</summary>
    public void ResetFontScales()
    {
        Root = Root.WithoutKeys("title_font_scale", "body_font_scale");
    }

    /// <summary>Returns a string value from the configuration.</summary>
    /// <param name="key">The configuration key.</param>
    /// <returns>
    /// The string value with its source location, or null if not found.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the value is not a scalar.
    /// </exception>
    public ValueWithLocation<string>? GetString(string key)
        => Root.KeyValues.GetValueOrDefault(key) switch
        {
            null => null,
            NullConfigNode => null,
            ScalarConfigNode scalarNode
                => new ValueWithLocation<string>(
                    scalarNode.Path,
                    scalarNode.Line,
                    scalarNode.Value
                ),
            var value => throw new PlotanceException(
                value.Path,
                value.Line,
                "Scalar is expected."
            )
        };

    /// <summary>Returns an integer value from the configuration.</summary>
    /// <param name="key">The configuration key.</param>
    /// <returns>
    /// The integer value with its source location, or null if not found.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the value cannot be parsed as an integer.
    /// </exception>
    public ValueWithLocation<int>? GetInt(string key)
        => GetString(key)?.Parse("integer", int.Parse);

    /// <summary>Returns a boolean value from the configuration.</summary>
    /// <param name="key">The configuration key.</param>
    /// <returns>
    /// The boolean value with its source location, or null if not found.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the value cannot be parsed as a boolean.
    /// </exception>
    public ValueWithLocation<bool>? GetBool(string key)
        => GetString(key)?.Parse("boolean", bool.Parse);

    /// <summary>Returns a decimal value from the configuration.</summary>
    /// <param name="key">The configuration key.</param>
    /// <returns>
    /// The decimal value with its source location, or null if not found.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the value cannot be parsed as a decimal.
    /// </exception>
    public ValueWithLocation<decimal>? GetDecimal(string key)
        => GetString(key)?.Parse("decimal", decimal.Parse);

    /// <summary>Returns a color value from the configuration.</summary>
    /// <param name="key">The configuration key.</param>
    /// <returns>
    /// The color value with its source location, or null if not found.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the value is not a scalar.
    /// </exception>
    public Color? GetColor(string key)
        => GetString(key) is ValueWithLocation<string> raw
        ? new Color(raw.Path, raw.Line, raw.Value)
        : null;

    /// <summary>Returns a length value from the configuration.</summary>
    /// <param name="key">The configuration key.</param>
    /// <returns>
    /// The length value with its source location, or null if not found.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the value is not a scalar.
    /// </exception>
    public ILength? GetLength(string key)
        => GetString(key) is ValueWithLocation<string> raw
        ? new TextLength(raw.Path, raw.Line, raw.Value)
        : null;

    /// <summary>Returns an axis unit value from the configuration.</summary>
    /// <param name="key">The configuration key.</param>
    /// <returns>
    /// The axis unit value with its source location, or null if not found.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the value is not a scalar.
    /// </exception>
    public IAxisUnitValue? GetAxisUnitValue(string key)
        => GetString(key) is ValueWithLocation<string> raw
        ? new TextAxisUnitValue(raw.Path, raw.Line, raw.Value)
        : null;

    /// <summary>Returns an axis range value from the configuration.</summary>
    /// <param name="key">The configuration key.</param>
    /// <returns>
    /// The axis range value with its source location, or null if not found.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the value is not a scalar.
    /// </exception>
    public IAxisRangeValue? GetAxisRangeValue(string key)
        => GetString(key) is ValueWithLocation<string> raw
        ? new TextAxisRangeValue(raw.Path, raw.Line, raw.Value)
        : null;

    /// <summary>Returns a list of strings from the configuration.</summary>
    /// <param name="key">The configuration key.</param>
    /// <returns>
    /// A list of string values with their source locations, or null if not
    /// found.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown if any of the values in the sequence are not strings.
    /// </exception>
    public ValueWithLocation<
        IReadOnlyList<ValueWithLocation<string>>
    >? GetStringList(
        string key
    ) => Root.KeyValues.GetValueOrDefault(key) switch
    {
        null => null,
        NullConfigNode => null,
        SequenceConfigNode sequenceNode
            => new ValueWithLocation<IReadOnlyList<ValueWithLocation<string>>>(
                sequenceNode.Path,
                sequenceNode.Line,
                sequenceNode
                    .Values
                    .Select(
                        value => value switch
                        {
                            ScalarConfigNode scalarNode
                                => new ValueWithLocation<string>(
                                    scalarNode.Path,
                                    scalarNode.Line,
                                    scalarNode.Value
                                ),
                            _ => throw new PlotanceException(
                                value.Path,
                                value.Line,
                                "String is expected."
                            )
                        }
                    )
                    .ToList()
            ),
        var value => throw new PlotanceException(
            value.Path,
            value.Line,
            "String list is expected."
        )
    };

    /// <summary>Returns a list of colors from the configuration.</summary>
    /// <param name="key">The configuration key.</param>
    /// <returns>
    /// A list of color values with their source locations, or null if not
    /// found.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown if any of the values in the sequence cannot be parsed as colors.
    /// </exception>
    public ValueWithLocation<IReadOnlyList<Color>>? GetColorList(
        string key
    ) => Root.KeyValues.GetValueOrDefault(key) switch
    {
        null => null,
        NullConfigNode => null,
        SequenceConfigNode sequenceNode
            => sequenceNode.Values.Any()
            ? new ValueWithLocation<IReadOnlyList<Color>>(
                sequenceNode.Path,
                sequenceNode.Line,
                sequenceNode
                    .Values
                        .Select(
                        value => value switch
                            {
                                ScalarConfigNode scalarNode => new Color(
                                    scalarNode.Path,
                                    scalarNode.Line,
                                    scalarNode.Value
                                ),
                                _ => throw new PlotanceException(
                                    value.Path,
                                    value.Line,
                                    "Color is expected."
                                )
                            }
                    )
                    .ToList()
            )
            : null,
        var value => throw new PlotanceException(
            value.Path,
            value.Line,
            "Color list is expected."
        )
    };

    /// <summary>
    /// Returns a list of parameter declarations from the configuration.
    /// </summary>
    /// <param name="key">The configuration key.</param>
    /// <returns>
    /// A list of parameter declarations with their source locations, or null if
    /// not found.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown if any of the values in the sequence cannot be parsed as
    /// parameter declarations.
    /// </exception>
    public ValueWithLocation<
        IReadOnlyList<ParameterDeclaration>
    >? GetParameterList(
        string key
    ) => Root.KeyValues.GetValueOrDefault(key) switch
    {
        null => null,
        NullConfigNode => null,
        SequenceConfigNode sequenceNode
            => new ValueWithLocation<IReadOnlyList<ParameterDeclaration>>(
                sequenceNode.Path,
                sequenceNode.Line,
                sequenceNode
                    .Values
                    .Select(ParameterDeclaration.Create)
                    .ToList()
            ),
        var value => throw new PlotanceException(
            value.Path,
            value.Line,
            "Parameter list is expected."
        )
    };

    /// <summary>
    /// Returns a dictionary of string key-value pairs from the configuration.
    /// </summary>
    /// <param name="key">The configuration key.</param>
    /// <returns>
    /// A dictionary of string values with their source locations, or null if
    /// not found.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown if any of the values in the mapping are not strings.
    /// </exception>
    public DictionaryWithLocation<string, string>? GetStringDictionary(
        string key
    )
    {
        ValueWithLocation<string> ToString(ConfigNode node)
            => node is ScalarConfigNode scalarNode
            ? new ValueWithLocation<string>(
                scalarNode.Path,
                scalarNode.Line,
                scalarNode.Value
            )
            : throw new PlotanceException(
                node.Path,
                node.Line,
                "String is expected."
            );

        DictionaryWithLocation<string, string> ToDictionary(
            MappingConfigNode mappingNode
        ) => new DictionaryWithLocation<string, string>(
            mappingNode.Path,
            mappingNode.Line,
            new Dictionary<string, ValueWithLocation<string>>(
                mappingNode.KeyValues.Select(
                    pair => new KeyValuePair<string, ValueWithLocation<string>>(
                        pair.Key,
                        ToString(pair.Value)
                    )
                )
            ),
            mappingNode.KeyLocations
        );

        return Root.KeyValues.GetValueOrDefault(key) switch
        {
            null => null,
            NullConfigNode => null,
            MappingConfigNode mappingNode => ToDictionary(mappingNode),
            var value => throw new PlotanceException(
                value.Path,
                value.Line,
                "String dictionary is expected."
            )
        };
    }

    /// <summary>
    /// Returns a list of length weights of rows/columns from the
    /// configuration.
    /// </summary>
    /// <param name="key">The configuration key.</param>
    /// <returns>
    /// A list of length weights with their source location, or null if not
    /// found.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown if value cannot be parsed as a list of length weights.
    /// </exception>
    public ValueWithLocation<IReadOnlyList<ILengthWeight>>? GetWeights(
        string key
    ) => GetString(key) is ValueWithLocation<string> raw
        ? new ValueWithLocation<IReadOnlyList<ILengthWeight>>(
            raw.Path,
            raw.Line,
            raw
                .Value
                .Split(':')
                .Select(
                    element => ILengthWeight.FromString(
                        new ValueWithLocation<string>(
                            raw.Path,
                            raw.Line,
                            element
                        )
                    )
                )
                .ToList()
        )
        : null;

    public Configuration Clone() => new Configuration(Root);
}

/// <summary>
/// Represents a parameter declaration for a document, including its name,
/// description, and default value.
/// </summary>
/// <param name="Name">
/// The name of the parameter with its source location.
/// </param>
/// <param name="Description">
/// The optional description of the parameter with its source location.
/// </param>
/// <param name="Default">
/// The optional default value of the parameter with its source location.
/// </param>
public record ParameterDeclaration(
    ValueWithLocation<string> Name,
    ValueWithLocation<string>? Description,
    ValueWithLocation<string>? Default
)
{
    /// <summary>
    /// Creates a parameter declaration from a configuration node.
    /// </summary>
    /// <param name="configNode">
    /// The configuration node to create the parameter declaration from.
    /// </param>
    /// <returns>A new parameter declaration.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the configuration node is not a mapping, if required fields
    /// are missing, or if values are not in the expected format.
    /// </exception>
    public static ParameterDeclaration Create(ConfigNode configNode)
        => configNode switch
        {
            MappingConfigNode mappingNode => Create(mappingNode),
            _ => throw new PlotanceException(
                configNode.Path,
                configNode.Line,
                "Mapping is expected."
            )
        };

    /// <summary>
    /// Creates a parameter declaration from a mapping configuration node.
    /// </summary>
    /// <param name="configNode">
    /// The mapping configuration node to create the parameter declaration from.
    /// </param>
    /// <returns>A new parameter declaration.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown if required fields are missing or if values are not in the
    /// expected format.
    /// </exception>
    public static ParameterDeclaration Create(MappingConfigNode configNode)
    {
        ValueWithLocation<string>? GetString(string key)
            => configNode.KeyValues.GetValueOrDefault(key) switch
            {
                null => null,
                NullConfigNode => null,
                ScalarConfigNode scalarNode => new ValueWithLocation<string>(
                    scalarNode.Path,
                    scalarNode.Line,
                    scalarNode.Value
                ),
                var value => throw new PlotanceException(
                    value.Path,
                    value.Line,
                    "String is expected."
                )
            };

        var name = GetString("name");
        var description = GetString("description");
        var defaultValue = GetString("default");

        if (name == null)
        {
            throw new PlotanceException(
                configNode.Path,
                configNode.Line,
                "Parameter name is required."
            );
        }

        return new ParameterDeclaration(name, description, defaultValue);
    }
}

/// <summary>
/// Represents the configuration options for an axis in a chart.
/// </summary>
/// <param name="Config">
/// The plotter configuration containing the axis settings.
/// </param>
/// <param name="Prefix">
/// The prefix used to identify axis-specific configuration keys.
/// </param>
public record AxisOptions(Configuration Config, string Prefix)
{
    /// <summary>The minimum value for the axis range.</summary>
    public IAxisRangeValue? Minimum
        => Config.GetAxisRangeValue(Prefix + "_range_minimum");

    /// <summary>The maximum value for the axis range.</summary>
    public IAxisRangeValue? Maximum
        => Config.GetAxisRangeValue(Prefix + "_range_maximum");

    /// <summary>The title text for the axis.</summary>
    public ValueWithLocation<string>? Title
        => Config.GetString(Prefix + "_title");

    /// <summary>The format string for axis labels.</summary>
    public ValueWithLocation<string>? LabelFormat
        => Config.GetString(Prefix + "_label_format");

    /// <summary>The rotation angle in degrees for axis labels.</summary>
    public ValueWithLocation<decimal>? LabelRotate
        => Config.GetDecimal(Prefix + "_label_rotate");

    /// <summary>The width of the axis line.</summary>
    public ILength? LineWidth
        => Config.GetLength(Prefix + "_line_width");

    /// <summary>The interval between major tick marks on the axis.</summary>
    public IAxisUnitValue? MajorUnit
        => Config.GetAxisUnitValue(Prefix + "_major_unit");

    /// <summary>The interval between minor tick marks on the axis.</summary>
    public IAxisUnitValue? MinorUnit
        => Config.GetAxisUnitValue(Prefix + "_minor_unit");

    /// <summary>The base value for logarithmic axis scaling.</summary>
    public ValueWithLocation<decimal>? LogBase
        => Config.GetDecimal(Prefix + "_log_base");

    /// <summary>
    /// Whether the axis should be displayed in reverse order.
    /// </summary>
    public ValueWithLocation<bool>? Reversed
        => Config.GetBool(Prefix + "_reversed");

    /// <summary>The width of major grid lines for the axis.</summary>
    public ILength? GridMajorWidth
        => Config.GetLength(Prefix + "_grid_major_width");

    /// <summary>The width of minor grid lines for the axis.</summary>
    public ILength? GridMinorWidth
        => Config.GetLength(Prefix + "_grid_minor_width");
}

/// <summary>Represents the configuration options for a table border.</summary>
/// <param name="Config">
/// The plotter configuration containing the border settings.
/// </param>
/// <param name="Prefix">
/// The prefix used to identify border-specific configuration keys.
/// </param>
public record TableBorderOptions(Configuration Config, string Prefix)
{
    /// <summary>The width of the table border.</summary>
    public ILength? Width => Config.GetLength(Prefix + "_border_width");

    /// <summary>The color of the table border.</summary>
    public Color? Color => Config.GetColor(Prefix + "_border_color");

    /// <summary>
    /// The style of the table border ("single", "double", "thick_thin",
    /// "thin_thick", "triple").
    /// </summary>
    public ValueWithLocation<string>? Style
        => Config.GetString(Prefix + "_border_style");
}

/// <summary>
/// Represents the configuration options for a table or a specific region of a
/// table.
/// </summary>
/// <param name="Config">
/// The plotter configuration containing the table settings.
/// </param>
/// <param name="Prefix">
/// The prefix used to identify table-specific configuration keys.
/// </param>
/// <param name="CellPrefix">
/// The suffix used to identify cell-specific configuration keys.
/// </param>
public record TableOptions(
    Configuration Config,
    string Prefix,
    string CellPrefix = "_cell"
)
{
    /// <summary>The background color of the cells.</summary>
    public Color? BackgroundColor
        => Config.GetColor(Prefix + "_background_color");

    /// <summary>The options for the left border of the region.</summary>
    public TableBorderOptions LeftBorderOptions
        => new TableBorderOptions(Config, Prefix + "_left");

    /// <summary>The options for the right border of the region.</summary>
    public TableBorderOptions RightBorderOptions
        => new TableBorderOptions(Config, Prefix + "_right");

    /// <summary>The options for the top border of the region.</summary>
    public TableBorderOptions TopBorderOptions
        => new TableBorderOptions(Config, Prefix + "_top");

    /// <summary>The options for the bottom border of the region.</summary>
    public TableBorderOptions BottomBorderOptions
        => new TableBorderOptions(Config, Prefix + "_bottom");

    /// <summary>
    /// The options for the inside horizontal borders of the region.
    /// </summary>
    public TableBorderOptions InsideHorizontalBorderOptions
        => new TableBorderOptions(Config, Prefix + "_inside_horizontal");

    /// <summary>
    /// The options for the inside vertical borders of the region.
    /// </summary>
    public TableBorderOptions InsideVerticalBorderOptions
        => new TableBorderOptions(Config, Prefix + "_inside_vertical");

    /// <summary>The font weight for the text in the region.</summary>
    public ValueWithLocation<string>? FontWeight
        => Config.GetString(Prefix + "_font_weight");

    /// <summary>The font color for the text in the region.</summary>
    public Color? FontColor
        => Config.GetColor(Prefix + "_font_color");

    /// <summary>The left margin for table cells.</summary>
    public ILength? CellLeftMargin
        => Config.GetLength(Prefix + CellPrefix + "_left_margin");

    /// <summary>The right margin for table cells.</summary>
    public ILength? CellRightMargin
        => Config.GetLength(Prefix + CellPrefix + "_right_margin");

    /// <summary>The top margin for table cells.</summary>
    public ILength? CellTopMargin
        => Config.GetLength(Prefix + CellPrefix + "_top_margin");

    /// <summary>The bottom margin for table cells.</summary>
    public ILength? CellBottomMargin
        => Config.GetLength(Prefix + CellPrefix + "_bottom_margin");

    /// <summary>
    /// The horizontal alignment for table cell content. Valid values are
    /// "left", "center", "right", "justified", "distributed".
    /// </summary>
    public ValueWithLocation<string>? CellHorizontalAlign
        => Config.GetString(Prefix + CellPrefix + "_horizontal_align");

    /// <summary>
    /// The vertical alignment for table cell content. Valid values are
    /// "top", "center", "bottom".
    /// </summary>
    public ValueWithLocation<string>? CellVerticalAlign
        => Config.GetString(Prefix + CellPrefix + "_vertical_align");
}

/// <summary>Represents the configuration options for charts.</summary>
/// <param name="Config">
/// The plotter configuration containing the chart settings.
/// </param>
public record ChartOptions(Configuration Config)
{
    /// <summary>The format of the chart.</summary>
    public ValueWithLocation<string>? Format => Config.GetString("format");

    /// <summary>
    /// The list of colors to use for data series in the chart.
    /// </summary>
    public ValueWithLocation<IReadOnlyList<Color>>? SeriesColors
        => Config.GetColorList("series_colors");

    /// <summary>
    /// Whether to group the data by the series name column (not implemented
    /// yet).
    /// </summary>
    public ValueWithLocation<bool>? GroupBySeries
        => Config.GetBool("group_by_series");

    /// <summary>
    /// The grouping style for bar charts ("clustered", "stacked",
    /// "percent_stacked").
    /// </summary>
    public ValueWithLocation<string>? BarGrouping
        => Config.GetString("bar_grouping");

    /// <summary>
    /// The direction for bar charts ("horizontal", "vertical").
    /// </summary>
    public ValueWithLocation<string>? BarDirection
        => Config.GetString("bar_direction");

    /// <summary>
    /// The gap between bar groups, as a percentage of the bar width from 0 to
    /// 500.
    /// </summary>
    public ValueWithLocation<int>? BarGap => Config.GetInt("bar_gap");

    /// <summary>
    /// The overlap between bars in a group, as a percentage of the bar width
    /// from -100 to 100.
    /// </summary>
    public ValueWithLocation<int>? BarOverlap => Config.GetInt("bar_overlap");

    /// <summary>The width of lines in the chart.</summary>
    public ILength? LineWidth => Config.GetLength("line_width");

    /// <summary>The color of lines in the chart.</summary>
    public Color? LineColor => Config.GetString("line_color")?.Value == "auto"
        ? null
        : Config.GetColor("line_color");

    /// <summary>The opacity of filled areas in the chart.</summary>
    public ValueWithLocation<decimal>? FillOpacity
        => Config.GetDecimal("fill_opacity");

    /// <summary>The title of the chart.</summary>
    public ValueWithLocation<string>? Title => Config.GetString("chart_title");

    /// <summary>The font size for the chart title.</summary>
    public ILength? TitleFontSize => Config.GetLength("chart_title_font_size");

    /// <summary>The color for the chart title.</summary>
    public Color? TitleColor => Config.GetColor("chart_title_color");

    /// <summary>The configuration options for the X axis.</summary>
    public AxisOptions? XAxisOptions => new AxisOptions(Config, "x_axis");

    /// <summary>The configuration options for the Y axis.</summary>
    public AxisOptions? YAxisOptions => new AxisOptions(Config, "y_axis");

    /// <summary>The font size for axis titles.</summary>
    public ILength? AxisTitleFontSize
        => Config.GetLength("axis_title_font_size");

    /// <summary>The color for axis titles.</summary>
    public Color? AxisTitleColor
        => Config.GetColor("axis_title_color");

    /// <summary>The font size for axis labels.</summary>
    public ILength? AxisLabelFontSize
        => Config.GetLength("axis_label_font_size");

    /// <summary>The color for axis labels.</summary>
    public Color? AxisLabelColor => Config.GetColor("axis_label_color");

    /// <summary>The color for axis lines.</summary>
    public Color? AxisLineColor => Config.GetColor("axis_line_color");

    /// <summary>The color for major grid lines.</summary>
    public Color? GridMajorColor => Config.GetColor("grid_major_color");

    /// <summary>The color for minor grid lines.</summary>
    public Color? GridMinorColor => Config.GetColor("grid_minor_color");

    /// <summary>
    /// The position of the legend ("bottom", "top_right", "left", "right",
    /// "top").
    /// </summary>
    public ValueWithLocation<string>? LegendPosition
        => Config.GetString("legend_position");

    /// <summary>The width of the legend box borders.</summary>
    public ILength? LegendLineWidth => Config.GetLength("legend_line_width");

    /// <summary>The color of the legend box borders.</summary>
    public Color? LegendLineColor => Config.GetColor("legend_line_color");

    /// <summary>The font size for legend text.</summary>
    public ILength? LegendFontSize => Config.GetLength("legend_font_size");

    /// <summary>The color for legend text.</summary>
    public Color? LegendColor => Config.GetColor("legend_color");

    /// <summary>The position of the data labels.</summary>
    /// <remarks>
    /// Valid values depend on the chart type:
    /// Bar chart: "inside_end", "outside_end", "inside_center", "inside_base",
    /// "center" (synonym for "inside_center").
    /// Scatter and bubble: "center", "left", "right", "above", "below".
    /// Area chart doesn't take a data label position.  "inside_center" and
    /// "center" are accepted but ignored.
    /// </remarks>
    public ValueWithLocation<string>? DataLabelPosition
        => Config.GetString("data_label_position");

    /// <summary>
    /// The contents of the data labels ("legend_key", "x_value", "y_value",
    /// "series_name", "percent", "bubble_size").  Can be a list of a string or
    /// a comma separated list of strings.
    /// </summary>
    public ValueWithLocation<IReadOnlyList<ValueWithLocation<string>>>?
        DataLabelContents => Config.GetStringList("data_label_contents");

    /// <summary>The format for the data labels.</summary>
    public ValueWithLocation<string>? DataLabelFormat
        => Config.GetString("data_label_format");

    /// <summary>The rotation angle for the data labels.</summary>
    public ValueWithLocation<decimal>? DataLabelRotate
        => Config.GetDecimal("data_label_rotate");

    /// <summary>The font size for the data labels.</summary>
    public ILength? DataLabelFontSize
        => Config.GetLength("data_label_font_size");

    /// <summary>The color for the data labels.</summary>
    public Color? DataLabelColor => Config.GetColor("data_label_color");

    /// <summary>The size of markers in the chart.</summary>
    public ILength? MarkerSize => Config.GetLength("marker_size");

    /// <summary>The width of the marker lines.</summary>
    public ILength? MarkerLineWidth => Config.GetLength("marker_line_width");

    /// <summary>The opacity of the marker fill.</summary>
    public ValueWithLocation<decimal>? MarkerFillOpacity
        => Config.GetDecimal("marker_fill_opacity");

    // Table options

    /// <summary>The table row height weights.</summary>
    public ValueWithLocation<IReadOnlyList<ILengthWeight>>? TableRows
        => Config.GetWeights("table_rows");

    /// <summary>The table column width weights.</summary>
    public ValueWithLocation<IReadOnlyList<ILengthWeight>>? TableColumns
        => Config.GetWeights("table_columns");

    /// <summary>The options for the entire table.</summary>
    public TableOptions WholeTableOptions => new TableOptions(Config, "table");

    /// <summary>The options for the first row of the table.</summary>
    public TableOptions FirstRowOptions
        => new TableOptions(Config, "first_row");

    /// <summary>The options for the first column of the table.</summary>
    public TableOptions FirstColumnOptions
        => new TableOptions(Config, "first_column");

    /// <summary>The options for the last row of the table.</summary>
    public TableOptions LastRowOptions => new TableOptions(Config, "last_row");

    /// <summary>The options for the last column of the table.</summary>
    public TableOptions LastColumnOptions
        => new TableOptions(Config, "last_column");

    /// <summary>
    /// The options for the odd rows (one-origin) of the table.
    /// </summary>
    public TableOptions Band1RowOptions
        => new TableOptions(Config, "band1_row");

    /// <summary>
    /// The options for the odd columns (one-origin) of the table.
    /// </summary>
    public TableOptions Band1ColumnOptions
        => new TableOptions(Config, "band1_column");

    /// <summary>
    /// The options for the even rows (one-origin) of the table.
    /// </summary>
    public TableOptions Band2RowOptions
        => new TableOptions(Config, "band2_row");

    /// <summary>
    /// The options for the even columns (one-origin) of the table.
    /// </summary>
    public TableOptions Band2ColumnOptions
        => new TableOptions(Config, "band2_column");

    /// <summary>The options for the southeast cell of the table.</summary>
    public TableOptions SoutheastCellOptions
        => new TableOptions(Config, "southeast_cell", "");

    /// <summary>The options for the southwest cell of the table.</summary>
    public TableOptions SouthwestCellOptions
        => new TableOptions(Config, "southwest_cell", "");

    /// <summary>The options for the northeast cell of the table.</summary>
    public TableOptions NortheastCellOptions
        => new TableOptions(Config, "northeast_cell", "");

    /// <summary>The options for the northwest cell of the table.</summary>
    public TableOptions NorthwestCellOptions
        => new TableOptions(Config, "northwest_cell", "");
}
