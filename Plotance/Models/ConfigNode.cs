// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using YamlDotNet.Core;
using YamlDotNet.RepresentationModel;

namespace Plotance.Models;

/// <summary>
/// Represents a node in a configuration hierarchy, which can be a scalar value,
/// a sequence, or a mapping. Node have a path and a line number for error
/// reporting. Scalar values are strings or null and parsed further by the
/// application. Mapping keys are strings.
/// </summary>
/// <param name="Path">
/// The path to the file containing the node, for error reporting.
/// </param>
/// <param name="Line">
/// The line number in the file, for error reporting.
/// </param>
public abstract record ConfigNode(string? Path, long Line)
{
    /// <summary>
    /// Merges this configuration node with another one. Scalar values are
    /// overridden by the other node, sequences are overridden or appended
    /// depending on the <see cref="SequenceConfigNode.Appending"/> option, and
    /// mappings are merged recursively by key.
    /// </summary>
    /// <param name="otherNode">The node to merge with this node.</param>
    /// <returns>The merged node.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the node type is not supported.
    /// </exception>
    public abstract ConfigNode Merge(ConfigNode otherNode);

    /// <summary>
    /// Expands variables in this node using the provided dictionary.
    /// </summary>
    /// <param name="variables">
    /// Dictionary of variable names to their values.
    /// </param>
    /// <returns>A new node with variables expanded.</returns>
    /// <seealso cref="Variables.ExpandVariables"/>
    public abstract ConfigNode ExpandVariables(
        IReadOnlyDictionary<string, string> variables
    );

    /// <summary>Creates a configuration node from a YAML document.</summary>
    /// <param name="path">
    /// The path to the file containing the document, for error reporting.
    /// </param>
    /// <param name="document">The YAML document to convert.</param>
    /// <param name="lineOffset">
    /// The line number offset to add to all line numbers.
    /// </param>
    /// <returns>A mapping configuration node.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the document root is not a mapping.
    /// </exception>
    public static MappingConfigNode Create(
        string? path,
        YamlDocument document,
        long lineOffset
    ) => Create(path, document.RootNode, lineOffset) as MappingConfigNode
        ?? throw new PlotanceException(
            path,
            lineOffset + 1,
            "Mapping expected."
        );

    /// <summary>Creates a configuration node from a YAML node.</summary>
    /// <param name="path">
    /// The path to the file containing the node, for error reporting.
    /// </param>
    /// <param name="yamlNode">The YAML node to convert.</param>
    /// <param name="lineOffset">
    /// The line number offset to add to all line numbers.
    /// </param>
    /// <returns>
    /// A configuration node of the appropriate type for the YAML node.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the YAML node type is unsupported.
    /// </exception>
    public static ConfigNode Create(
        string? path,
        YamlNode yamlNode,
        long lineOffset
    ) => yamlNode switch
    {
        YamlScalarNode scalarNode
            => ScalarConfigNode.Create(path, scalarNode, lineOffset),

        YamlSequenceNode sequenceNode
            => SequenceConfigNode.Create(path, sequenceNode, lineOffset),

        YamlMappingNode mappingNode
            => MappingConfigNode.Create(path, mappingNode, lineOffset),

        _ => throw new PlotanceException(
            path,
            yamlNode.Start.Line + lineOffset,
            "Unsupported node type."
        )
    };
}

/// <summary>
/// Represents a scalar (string) value in a configuration hierarchy.
/// </summary>
/// <param name="Path">
/// The path to the file containing the node, for error reporting.
/// </param>
/// <param name="Line">The line number in the file, for error reporting.</param>
/// <param name="Value">The string value of the node.</param>
public record ScalarConfigNode(
    string? Path,
    long Line,
    string Value
) : ConfigNode(Path, Line)
{
    /// <summary>Creates a configuration node from a YAML scalar node.</summary>
    /// <param name="path">
    /// The path to the file containing the node, for error reporting.
    /// </param>
    /// <param name="scalarNode">The YAML scalar node to convert.</param>
    /// <param name="lineOffset">
    /// The line number offset to add to all line numbers.
    /// </param>
    /// <returns>
    /// A scalar configuration node, or a null configuration node if the scalar
    /// value is null.
    /// </returns>
    public static ConfigNode Create(
        string? path,
        YamlScalarNode scalarNode,
        long lineOffset
    ) => scalarNode switch
    {
        { Value: null }
            => new NullConfigNode(path, scalarNode.Start.Line + lineOffset),
        { Value: "null", Style: ScalarStyle.Plain }
            => new NullConfigNode(path, scalarNode.Start.Line + lineOffset),
        _ => new ScalarConfigNode(
            path,
            scalarNode.Start.Line + lineOffset,
            scalarNode.Value
        )
    };

    /// <inheritdoc/>
    public override ConfigNode Merge(ConfigNode otherNode) => otherNode;

    /// <inheritdoc/>
    public override ScalarConfigNode ExpandVariables(
        IReadOnlyDictionary<string, string> variables
    ) => this with { Value = Variables.ExpandVariables(Value, variables) };
}

/// <summary>
/// Represents a null value in a configuration hierarchy.
/// </summary>
/// <param name="Path">
/// The path to the file containing the node, for error reporting.
/// </param>
/// <param name="Line">
/// The line number in the file, for error reporting.
/// </param>
public record NullConfigNode(string? Path, long Line) : ConfigNode(Path, Line)
{
    /// <inheritdoc/>
    public override ConfigNode Merge(ConfigNode otherNode) => otherNode;

    /// <inheritdoc/>
    public override NullConfigNode ExpandVariables(
        IReadOnlyDictionary<string, string> variables
    ) => this;
}

/// <summary>Represents a sequence of configuration nodes.</summary>
/// <param name="Path">
/// The path to the file containing the node, for error reporting.
/// </param>
/// <param name="Line">The line number in the file, for error reporting.</param>
/// <param name="Values">
/// The list of configuration nodes in the sequence.
/// </param>
/// <param name="Appending">
/// Whether this sequence should be appended to when merging with another
/// sequence, rather than being replaced by it.
/// </param>
public record SequenceConfigNode(
    string? Path,
    long Line,
    IReadOnlyList<ConfigNode> Values,
    bool Appending = false
) : ConfigNode(Path, Line)
{
    /// <summary>
    /// Creates a sequence configuration node from a YAML sequence node.
    /// </summary>
    /// <param name="path">
    /// The path to the file containing the node, for error reporting.
    /// </param>
    /// <param name="sequenceNode">The YAML sequence node to convert.</param>
    /// <param name="lineOffset">
    /// The line number offset to add to all line numbers.
    /// </param>
    /// <returns>A sequence configuration node.</returns>
    public static SequenceConfigNode Create(
        string? path,
        YamlSequenceNode sequenceNode,
        long lineOffset
    )
    {
        var values = new List<ConfigNode>();

        foreach (var item in sequenceNode.Children)
        {
            values.Add(ConfigNode.Create(path, item, lineOffset));
        }

        return new(path, sequenceNode.Start.Line + lineOffset, values);
    }

    /// <inheritdoc/>
    public override ConfigNode Merge(ConfigNode otherNode)
        => otherNode is SequenceConfigNode otherSequenceNode
        ? Merge(otherSequenceNode)
        : throw new PlotanceException(
            Path,
            Line,
            $"Cannot merge value at {otherNode.Path}:{otherNode.Line}"
        );

    /// <summary>Merges this sequence node with another sequence node.</summary>
    /// <param name="otherNode">
    /// The sequence node to merge with this node.
    /// </param>
    /// <returns>
    /// If this node has Appending set to true, returns a new sequence with
    /// the values from both sequences. Otherwise, returns the other node.
    /// </returns>
    public SequenceConfigNode Merge(SequenceConfigNode otherNode)
        => Appending
        ? this with { Values = [.. Values, .. otherNode.Values] }
        : otherNode;

    /// <inheritdoc/>
    public override SequenceConfigNode ExpandVariables(
        IReadOnlyDictionary<string, string> variables
    ) => this with
    {
        Values = Values
            .Select(value => value.ExpandVariables(variables))
            .ToList()
    };
}

/// <summary>
/// Represents a mapping of string keys to configuration nodes.
/// </summary>
/// <param name="Path">
/// The path to the file containing the node, for error reporting.
/// </param>
/// <param name="Line">The line number in the file, for error reporting.</param>
/// <param name="KeyLocations">
/// Dictionary mapping keys to their source locations (file path and line
/// number) for error reporting.
/// </param>
/// <param name="KeyValues">
/// Dictionary mapping keys to their corresponding configuration nodes.
/// </param>
public record MappingConfigNode(
    string? Path,
    long Line,
    IReadOnlyDictionary<string, (string?, long)> KeyLocations,
    IReadOnlyDictionary<string, ConfigNode> KeyValues
) : ConfigNode(Path, Line)
{
    /// <summary>An empty mapping configuration node.</summary>
    public static MappingConfigNode Empty => new MappingConfigNode(
        null,
        1,
        new Dictionary<string, (string?, long)>(),
        new Dictionary<string, ConfigNode>()
    );

    /// <summary>
    /// Creates a mapping configuration node from a YAML mapping node.
    /// </summary>
    /// <param name="path">
    /// The path to the file containing the node, for error reporting.
    /// </param>
    /// <param name="mappingNode">The YAML mapping node to convert.</param>
    /// <param name="lineOffset">
    /// The line number offset to add to all line numbers.
    /// </param>
    /// <returns>A mapping configuration node.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown when a non-string key is encountered.
    /// </exception>
    public static MappingConfigNode Create(
        string? path,
        YamlMappingNode mappingNode,
        long lineOffset
    )
    {
        var values = new Dictionary<string, ConfigNode>();
        var keyLocations = new Dictionary<string, (string?, long)>();

        foreach (var (key, value) in mappingNode.Children)
        {
            var keyLine = key.Start.Line + lineOffset;

            if (key is YamlScalarNode scalarKey && scalarKey.Value != null)
            {
                values[scalarKey.Value] = ConfigNode.Create(
                    path,
                    value,
                    lineOffset
                );
                keyLocations[scalarKey.Value] = (path, keyLine);
            }
            else
            {
                throw new PlotanceException(
                    path,
                    keyLine,
                    "Non-string key is not supported."
                );
            }
        }

        return new(
            path,
            mappingNode.Start.Line + lineOffset,
            keyLocations,
            values
        );
    }

    /// <inheritdoc/>
    public override ConfigNode Merge(ConfigNode otherNode)
        => otherNode is MappingConfigNode otherMappingNode
        ? Merge(otherMappingNode)
        : throw new PlotanceException(
            Path,
            Line,
            $"Cannot merge value at {otherNode.Path}:{otherNode.Line}"
        );

    /// <summary>Merges this mapping node with another mapping node.</summary>
    /// <param name="otherNode">
    /// The mapping node to merge with this node.
    /// </param>
    /// <returns>
    /// A new mapping node containing all keys from both mappings, with values
    /// from this mapping merged with corresponding values from the other
    /// mapping when keys are present in both.
    /// </returns>
    public MappingConfigNode Merge(MappingConfigNode otherNode)
    {
        var mergedKeyValues = new Dictionary<string, ConfigNode>(KeyValues);
        var mergedKeyLocations = new Dictionary<string, (string?, long)>(
            KeyLocations
        );

        foreach (var (otherKey, otherValue) in otherNode.KeyValues)
        {
            if (mergedKeyValues.ContainsKey(otherKey))
            {
                mergedKeyValues[otherKey] = mergedKeyValues[otherKey]
                    .Merge(otherValue);
            }
            else
            {
                mergedKeyValues[otherKey] = otherValue;
                mergedKeyLocations[otherKey] = otherNode.KeyLocations[otherKey];
            }
        }

        return this with
        {
            KeyLocations = mergedKeyLocations,
            KeyValues = mergedKeyValues
        };
    }

    /// <summary>
    /// Creates a new mapping node without the specified keys.
    /// </summary>
    /// <param name="keys">The keys to remove from the mapping.</param>
    /// <returns>A new mapping node without the specified keys.</returns>
    public MappingConfigNode WithoutKeys(params IEnumerable<string> keys)
    {
        var newKeyValues = new Dictionary<string, ConfigNode>(KeyValues);
        var newKeyLocations = new Dictionary<string, (string?, long)>(
            KeyLocations
        );

        foreach (var key in keys)
        {
            newKeyValues.Remove(key);
            newKeyLocations.Remove(key);
        }

        return this with
        {
            KeyLocations = newKeyLocations,
            KeyValues = newKeyValues
        };
    }

    /// <inheritdoc/>
    public override MappingConfigNode ExpandVariables(
        IReadOnlyDictionary<string, string> variables
    ) => this with
    {
        KeyValues = KeyValues.ToDictionary(
            kv => kv.Key,
            kv => kv.Value.ExpandVariables(variables)
        )
    };
}
