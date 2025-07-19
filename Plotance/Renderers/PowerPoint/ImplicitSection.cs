// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Collections.Generic;
using Markdig.Syntax;
using Plotance.Models;

namespace Plotance.Renderers.PowerPoint;

/// <summary>
/// Represents a section of a Markdown document to be rendered as a single slide
/// in a PowerPoint presentation.
/// </summary>
/// <param name="HeadingBlock">
/// The heading block that defines the section, if any.
/// </param>
/// <param name="BlockGroups">
/// The rows (if LayoutDirection is "row") or columns (if LayoutDirection is
/// "column") of content in the section.
/// </param>
/// <param name="Variables">
/// Dictionary of variables to expand in the section content.
/// </param>
/// <param name="TitleFontScale">
/// The scale factor for the font size of the title, if any.
/// </param>
/// <param name="Language">
/// The language for the section content, if any.
/// </param>
/// <param name="SlideLevel">
/// The maximum heading level that starts a new slide, if any.
/// </param>
/// <param name="LayoutDirection">
/// The layout direction of the section, if any.
/// </param>
public record ImplicitSection(
    HeadingBlock? HeadingBlock,
    IReadOnlyList<BlockGroup> BlockGroups,
    IReadOnlyDictionary<string, string> Variables,
    ValueWithLocation<decimal>? TitleFontScale = null,
    ValueWithLocation<string>? Language = null,
    ValueWithLocation<int>? SlideLevel = null,
    ValueWithLocation<LayoutDirection>? LayoutDirection = null
)
{
    /// <summary>The path to the file containing the section.</summary>
    public string? Path => BlockGroups
        .FirstOrDefault(group => group.Path != null)?.Path;

    /// <summary>The line number in the file containing the section.</summary>
    public long Line => BlockGroups
        .FirstOrDefault(group => group.Path != null)?.Line ?? 0;

    /// <summary>
    /// Creates a list of implicit sections from a list of Markdown blocks.
    /// </summary>
    /// <param name="blocks">The list of Markdown blocks to process.</param>
    /// <param name="variables">
    /// The initial dictionary of variables to expand in content.
    /// </param>
    /// <returns>
    /// A list of implicit sections, where each section represents content for
    /// a single PowerPoint slide.
    /// </returns>
    /// <remarks>
    /// Sections are created based on thematic breaks, heading blocks with level
    /// less than or equal to the slide level (default is 2), or at the
    /// beginning of the document.
    /// </remarks>
    public static IReadOnlyList<ImplicitSection> Create(
        IEnumerable<Block> blocks,
        IReadOnlyDictionary<string, string> variables
    )
    {
        var accumulatedConfig = new Configuration();
        var accumulatedVariables = new Dictionary<string, string>(variables);
        var sections = new List<ImplicitSection>();
        var currentSection = new List<Block>();
        HeadingBlock? currentHeadingBlock = null;

        void AddSection()
        {
            sections.Add(
                Create(
                    currentHeadingBlock,
                    currentSection,
                    accumulatedConfig,
                    accumulatedVariables
                )
            );
            currentSection = new List<Block>();
            currentHeadingBlock = null;
        }

        foreach (var block in blocks)
        {
            switch (block)
            {
                case ThematicBreakBlock:
                    AddSection();
                    break;

                case HeadingBlock heading when (
                    heading.Level <= (accumulatedConfig.SlideLevel?.Value ?? 2)
                ):
                    AddSection();
                    currentHeadingBlock = heading;
                    break;

                default:
                    currentSection.Add(block);
                    break;
            }
        }

        AddSection();

        if (!sections[0].BlockGroups.Any())
        {
            sections.RemoveAt(0);
        }

        return sections;
    }

    /// <summary>
    /// Creates an implicit section from a heading block and a collection of
    /// blocks. Blocks are grouped into rows (if LayoutDirection is "row") or
    /// columns (if LayoutDirection is "column").
    /// </summary>
    /// <param name="headingBlock">
    /// The heading block that defines the section, if any.
    /// </param>
    /// <param name="section">
    /// The collection of Markdown blocks in the section.
    /// </param>
    /// <param name="accumulatedConfig">
    /// The accumulated plotter configuration from previous blocks.
    /// </param>
    /// <param name="accumulatedVariables">
    /// The accumulated variables from previous blocks.
    /// </param>
    /// <returns>
    /// An implicit section representing content for a single PowerPoint slide.
    /// </returns>
    public static ImplicitSection Create(
        HeadingBlock? headingBlock,
        IEnumerable<Block> section,
        Configuration accumulatedConfig,
        IDictionary<string, string> accumulatedVariables
    )
    {
        var titleFontScale = accumulatedConfig.TitleFontScale;
        var titleLanguage = accumulatedConfig.Language;
        var slideLevel = accumulatedConfig.SlideLevel;
        List<IEnumerable<ILengthWeight>> rowWeightsRaw = [];
        List<IEnumerable<ILengthWeight>> columnWeightsRaw = [];
        List<ILengthWeight> blockGroupWeights = [];
        int blockGroupIndex = 0;
        var blockGroups = new List<BlockGroup>();
        ILength? blockGroupGap = null;
        var currentBlockGroupContents = new List<BlockContainer>();
        List<ILengthWeight> blockWeights = [];
        int blockIndex = 0;
        ValueWithLocation<LayoutDirection>? layoutDirectionRaw = null;
        LayoutDirection? layoutDirection = null;

        accumulatedConfig.ResetFontScales();

        void AddBlockGroup()
        {
            while (blockGroupWeights.Count <= blockGroupIndex)
            {
                blockGroupWeights.Add(new RelativeLengthWeight(null, 0, 1));
            }

            blockGroups.Add(
                new BlockGroup(
                    blockGroupWeights[blockGroupIndex],
                    currentBlockGroupContents,
                    blockGroupGap
                )
            );
            blockGroupGap = null;
            currentBlockGroupContents = new List<BlockContainer>();
            blockIndex = 0;
            blockGroupIndex++;
        }

        bool IsInvisibleBlock(Block block)
        {
            if (block.GetData("query_results") != null)
            {
                return accumulatedConfig.ChartOptions.Format?.Value == "none";
            }

            if (block.GetData("plotter_config") is Configuration)
            {
                return true;
            }

            var invisibleHtmlBlockTypes = new HashSet<HtmlBlockType>() {
                HtmlBlockType.DocumentType,
                HtmlBlockType.Comment,
                HtmlBlockType.ProcessingInstruction
            };

            switch (block)
            {
                case EmptyBlock:
                case LinkReferenceDefinitionGroup:
                case LinkReferenceDefinition:
                case ThematicBreakBlock:
                case HtmlBlock htmlBlock
                    when invisibleHtmlBlockTypes.Contains(htmlBlock.Type):
                    return true;

                default:
                    return false;
            }
        }

        void AccumulateVariables(
            IDictionary<string, string> variables,
            Configuration config
        )
        {
            if (config.Parameters != null)
            {
                foreach (var parameter in config.Parameters.Value)
                {
                    if (
                        parameter.Name != null
                            && !variables.ContainsKey(parameter.Name.Value)
                            && parameter.Default != null
                    )
                    {
                        variables[parameter.Name.Value]
                            = parameter.Default.Value;
                    }
                }
            }
        }

        /// <summary>
        /// Accumulates the plotter configuration from the current block. If the
        /// current block is to be rendered as a chart, the configuration is
        /// accumulated only temporarily, that is, a temporary copy of the
        /// accumulated configuration is created and merged with the given
        /// configuration.
        /// </summary>
        /// <param name="config">
        /// The plotter configuration to accumulate.</param>
        /// <returns>The accumulated plotter configuration.</returns>
        Configuration AccumulateConfig(Configuration config)
        {
            var format = (
                config.ChartOptions?.Format
                    ?? accumulatedConfig.ChartOptions?.Format
            )?.Value;

            if (
                (config.Query == null && config.QueryFile == null)
                    || format == "none"
            )
            {
                accumulatedConfig.Update(config);

                return accumulatedConfig;
            }
            else
            {
                Configuration mergedConfig = accumulatedConfig.Clone();

                mergedConfig.Update(config);

                return mergedConfig;
            }
        }

        foreach (var block in section)
        {
            Configuration mergedConfig;

            if (block.GetData("plotter_config") is Configuration config)
            {
                var includedConfigs = block
                    .GetData("included_configs")
                    as IReadOnlyList<Configuration>
                    ?? [];

                var flattenConfig = new Configuration();

                foreach (var includedConfig in includedConfigs)
                {
                    flattenConfig.Update(includedConfig);
                }

                flattenConfig.Update(config);

                AccumulateVariables(accumulatedVariables, flattenConfig);

                mergedConfig = AccumulateConfig(flattenConfig);


                if (
                    config.Rows?.Value
                            is IEnumerable<ILengthWeight> newRowWeights
                )
                {
                    if (layoutDirection == null)
                    {
                        rowWeightsRaw.Add(newRowWeights);
                    }
                    else
                    {
                        layoutDirection.Switch(
                            () => blockGroupWeights.AddRange(newRowWeights),
                            () => blockWeights = newRowWeights.ToList()
                        );
                    }
                }

                if (
                    config.Columns?.Value
                        is IEnumerable<ILengthWeight> newColumnWeights
                )
                {
                    if (layoutDirection == null)
                    {
                        columnWeightsRaw.Add(newColumnWeights);
                    }
                    else
                    {
                        layoutDirection.Switch(
                            () => blockWeights = newColumnWeights.ToList(),
                            () => blockGroupWeights.AddRange(newColumnWeights)
                        );
                    }
                }
            }
            else
            {
                mergedConfig = accumulatedConfig;
            }

            if (!IsInvisibleBlock(block))
            {
                if (blockGroupIndex == 0 && blockIndex == 0)
                {
                    layoutDirectionRaw = mergedConfig.LayoutDirection;
                    layoutDirection = layoutDirectionRaw?.Value
                        ?? Plotance.Models.LayoutDirection.Row;

                    List<IEnumerable<ILengthWeight>> blockGroupWeightsRaw
                        = layoutDirection.Choose(
                            row: rowWeightsRaw,
                            column: columnWeightsRaw
                        );
                    List<IEnumerable<ILengthWeight>> blockWeightsRaw
                        = layoutDirection.Choose(
                            row: columnWeightsRaw,
                            column: rowWeightsRaw
                        );

                    blockGroupWeights = blockGroupWeightsRaw
                        .SelectMany(l => l).ToList();
                    blockWeights = blockWeightsRaw.Any()
                        ? blockWeightsRaw[^1].ToList()
                        : [
                            new RelativeLengthWeight(null, 0, 1)
                        ];
                }

                if (blockIndex == 0)
                {
                    blockGroupGap = layoutDirection!.Choose(
                        row: mergedConfig.RowGap,
                        column: mergedConfig.ColumnGap
                    );
                }

                var blockGap = layoutDirection!.Choose(
                    row: mergedConfig.ColumnGap,
                    column: mergedConfig.RowGap
                );
                var weight = blockIndex < blockWeights.Count
                    ? blockWeights[blockIndex]
                    : new RelativeLengthWeight(null, 0, 1);
                var column = new BlockContainer(
                    weight,
                    block,
                    new Dictionary<string, string>(accumulatedVariables),
                    blockGap,
                    mergedConfig.BodyFontScale,
                    mergedConfig.Language,
                    mergedConfig.BodyHorizontalAlign,
                    mergedConfig.BodyVerticalAlign,
                    mergedConfig.Clone().ChartOptions
                );

                currentBlockGroupContents.Add(column);
                blockIndex++;

                if (blockWeights.Count <= blockIndex)
                {
                    AddBlockGroup();
                }
            }
        }

        if (currentBlockGroupContents.Any())
        {
            AddBlockGroup();
        }

        return new ImplicitSection(
            headingBlock,
            blockGroups,
            new Dictionary<string, string>(accumulatedVariables),
            titleFontScale,
            titleLanguage,
            slideLevel,
            layoutDirectionRaw
        );
    }
}

/// <summary>
/// Represents a row (if LayoutDirection is "row") or column (if LayoutDirection
/// is "column") in an implicit section, containing blocks.
/// </summary>
/// <param name="Weight">The relative height of this block group.</param>
/// <param name="Blocks">The blocks in this block group.</param>
/// <param name="GapBefore">The gap before this block group.</param>
public record BlockGroup(
    ILengthWeight Weight,
    IReadOnlyList<BlockContainer> Blocks,
    ILength? GapBefore = null
)
{
    /// <summary>The path to the file containing the block group.</summary>
    public string? Path
        => Blocks.FirstOrDefault(block => block.Path != null)?.Path;

    /// <summary>
    /// The line number in the file containing the block group.
    /// </summary>
    public long Line
        => Blocks.FirstOrDefault(block => block.Path != null)?.Line ?? 0;
}

/// <summary>
/// Represents a block container in an implicit section block group, containing
/// a Markdown block with associated formatting options.
/// </summary>
/// <param name="Weight">The relative width of this block.</param>
/// <param name="Block">The Markdown block in this block.</param>
/// <param name="Variables">
/// Dictionary of variables to expand in the content.
/// </param>
/// <param name="BodyFontScale">
/// The scale factor for the font size of the body text, if any.
/// </param>
/// <param name="Language">The language for the content, if any.</param>
/// <param name="BodyHorizontalAlign">
/// The horizontal alignment of the body text, if any.
/// </param>
/// <param name="BodyVerticalAlign">
/// The vertical alignment of the body text, if any.
/// </param>
/// <param name="ChartOptions">
/// The chart options when rendering this block as a chart, if any.
/// </param>
public record BlockContainer(
    ILengthWeight Weight,
    Block Block,
    IReadOnlyDictionary<string, string> Variables,
    ILength? GapBefore = null,
    ValueWithLocation<decimal>? BodyFontScale = null,
    ValueWithLocation<string>? Language = null,
    ValueWithLocation<string>? BodyHorizontalAlign = null,
    ValueWithLocation<string>? BodyVerticalAlign = null,
    ChartOptions? ChartOptions = null
)
{
    /// <summary>The path to the file containing the block.</summary>
    public string? Path => Block.GetData("path") as string;

    /// <summary>The line number in the file containing the block.</summary>
    public long Line => Block.Line;
}
