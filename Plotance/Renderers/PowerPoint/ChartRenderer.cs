// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Numerics;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Plotance.Models;

using D = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Plotance.Renderers.PowerPoint;

/// <summary>Provides methods for rendering charts.</summary>
public abstract class ChartRenderer
{
    /// <summary>The set of numeric types.</summary>
    public static HashSet<Type> NumericTypes { get; } = [
        typeof(SByte),
        typeof(Byte),
        typeof(Decimal),
        typeof(Double),
        typeof(Single),
        typeof(Int16),
        typeof(UInt16),
        typeof(Int32),
        typeof(UInt32),
        typeof(Int64),
        typeof(UInt64),
        typeof(BigInteger)
    ];

    /// <summary>The default chart colors.</summary>
    // TODO more colors
    protected IReadOnlyList<OpenXmlCompositeElement> DefaultChartColors => [
        new D.SchemeColor() { Val = D.SchemeColorValues.Accent1 },
        new D.SchemeColor() { Val = D.SchemeColorValues.Accent2 },
        new D.SchemeColor() { Val = D.SchemeColorValues.Accent3 },
        new D.SchemeColor() { Val = D.SchemeColorValues.Accent4 },
        new D.SchemeColor() { Val = D.SchemeColorValues.Accent5 },
        new D.SchemeColor() { Val = D.SchemeColorValues.Accent6 },
    ];

    /// <summary>The URI of the chart.</summary>
    protected const string ChartUri
        = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    /// <summary>Render a query result to a slide as a chart.</summary>
    /// <param name="slidePart">The slide part to render to.</param>
    /// <param name="isFirst">Whether this is the first block.</param>
    /// <param name="x">The X coordinate of the chart.</param>
    /// <param name="y">The Y coordinate of the chart.</param>
    /// <param name="width">The width of the chart.</param>
    /// <param name="height">The height of the chart.</param>
    /// <param name="block">The block to render.</param>
    /// <param name="queryResults">The query results.</param>
    /// <exception cref="PlotanceException">
    /// Thrown if the chart format is not supported.
    /// </exception>
    public static void Render(
        SlidePart slidePart,
        bool isFirst,
        long x,
        long y,
        long width,
        long height,
        BlockContainer block,
        IReadOnlyList<QueryResultSet> queryResults
    )
    {
        if (!queryResults.Any())
        {
            return;
        }

        var queryResult = queryResults[^1];

        if (!queryResult.Columns.Any())
        {
            return;
        }

        var chartPart = slidePart.AddNewPart<ChartPart>();
        var embeddedSpreadsheetPart = chartPart.AddEmbeddedPackagePart(
            EmbeddedPackagePartType.Xlsx.ContentType
        );
        using var spreadsheetDocument = Spreadsheets
                      .GenerateSpreadsheetDocumentForChart(queryResult);

        spreadsheetDocument.Clone(embeddedSpreadsheetPart.GetStream());

        var relationShipId = chartPart.GetIdOfPart(embeddedSpreadsheetPart);

        var chartFormat = block.ChartOptions?.Format;
        ChartRenderer renderer = chartFormat?.Value switch
        {
            "none" => new NoneChartRenderer(),
            "bar" => new BarChartRenderer(),
            "line" => new ScatterChartRenderer(),
            "area" => new AreaChartRenderer(),
            "scatter" => new ScatterChartRenderer(),
            "bubble" => new BubbleChartRenderer(),
            null => new BarChartRenderer(),
            _ => throw new PlotanceException(
                chartFormat?.Path,
                chartFormat?.Line,
                $"Unknown chart format: {chartFormat?.Value}"
            )
        };

        renderer.RenderChart(chartPart, block, queryResult, relationShipId);

        var placeholderShape = new PlaceholderShape()
        {
            Type = PlaceholderValues.Chart
        };

        if (isFirst)
        {
            placeholderShape.Index = 1;
        }

        slidePart
            .Slide
            ?.CommonSlideData
            ?.ShapeTree
            ?.AppendChild(
                new GraphicFrame(
                    new NonVisualGraphicFrameProperties(
                        new NonVisualDrawingProperties()
                        {
                            Id = (UInt32Value)1U,
                            Name = ""
                        },
                        new NonVisualGraphicFrameDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties(
                            placeholderShape
                        )
                    ),
                    new Transform(
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
                    ),
                    new D.Graphic(
                        new D.GraphicData(
                            new C.ChartReference()
                            {
                                Id = slidePart.GetIdOfPart(chartPart)
                            }
                        )
                        {
                            Uri = ChartUri
                        }
                    )
                )
            );
    }

    /// <summary>Render a chart.</summary>
    /// <param name="chartPart">The chart part to render to.</param>
    /// <param name="block">The block to render.</param>
    /// <param name="queryResult">The query result.</param>
    /// <param name="relationShipId">
    /// The relationship ID of the chart part.
    /// </param>
    /// <exception cref="PlotanceException">
    /// Thrown if the chart configuration is invalid.
    /// </exception>
    protected abstract void RenderChart(
        ChartPart chartPart,
        BlockContainer block,
        QueryResultSet queryResult,
        string relationShipId
    );

    /// <summary>Creates a fill of a chart element.</summary>
    /// <param name="colors">The colors.</param>
    /// <param name="index">The index of the color.</param>
    /// <param name="opacity">The opacity.</param>
    /// <returns>The fill element.</returns>
    protected OpenXmlElement CreateFill(
        IReadOnlyList<OpenXmlCompositeElement>? colors,
        int index,
        ValueWithLocation<decimal>? opacity
    ) => CreateFill(colors, index, opacity?.Value ?? 1m);

    /// <summary>Creates a fill of a chart element.</summary>
    /// <param name="colors">The colors.</param>
    /// <param name="index">The index of the color.</param>
    /// <param name="opacity">The opacity.</param>
    /// <returns>The fill element.</returns>
    protected OpenXmlElement CreateFill(
        IReadOnlyList<OpenXmlCompositeElement>? colors,
        int index,
        decimal opacity
    )
    {
        if (opacity == 0m)
        {
            return new D.NoFill();
        }

        colors ??= DefaultChartColors;

        var colorIndex = index % colors.Count;
        var color = (OpenXmlCompositeElement)(
            colors[colorIndex].CloneNode(true)
        );

        color.AddChild(
            new D.Alpha() { Val = (int)decimal.Round(100000 * opacity) }
        );

        return new D.SolidFill(color);
    }

    /// <summary>Creates an outline of a chart element.</summary>
    /// <param name="colors">The colors.</param>
    /// <param name="index">The index of the color.</param>
    /// <param name="lineWidth">The line width.</param>
    /// <returns>The outline element.</returns>
    protected D.Outline CreateOutline(
        IReadOnlyList<OpenXmlCompositeElement>? colors,
        int index,
        ILength? lineWidth
    )
    {
        lineWidth ??= ILength.Zero;

        if (lineWidth.ToEmu() == 0)
        {
            return new D.Outline(new D.NoFill()) { Width = 0 };
        }
        else
        {
            colors ??= DefaultChartColors;

            return CreateOutline(
                colors[index % colors.Count].CloneNode(true),
                lineWidth
            );
        }
    }

    /// <summary>Creates an outline of a chart element.</summary>
    /// <param name="color">The color.</param>
    /// <param name="lineWidth">The line width.</param>
    /// <returns>The outline element.</returns>
    protected D.Outline CreateOutline(OpenXmlElement color, ILength lineWidth)
        => lineWidth.ToEmu() == 0
        ? new D.Outline(new D.NoFill()) { Width = 0 }
        : new D.Outline(new D.SolidFill(color))
        {
            Width = (int)lineWidth.ToEmu()
        };

    /// <summary>Creates a default run properties of a chart element.</summary>
    /// <param name="color">The color.</param>
    /// <param name="fontSize">The font size.</param>
    /// <returns>The default run properties element.</returns>
    protected D.DefaultRunProperties CreateDefaultRunProperties(
        OpenXmlElement? color,
        ILength fontSize
    )
    {
        var defaultRunProperties = new D.DefaultRunProperties(
            new D.LatinFont() { Typeface = "+mn-lt" },
            new D.EastAsianFont() { Typeface = "+mn-ea" },
            new D.ComplexScriptFont() { Typeface = "+mn-cs" }
        )
        {
            FontSize = fontSize.ToCentipoint()
        };

        if (color != null)
        {
            defaultRunProperties.AddChild(new D.SolidFill(color));
        }

        return defaultRunProperties;
    }

    /// <summary>Creates a scaling of an axis.</summary>
    /// <param name="reversed">Whether the axis is reversed.</param>
    /// <param name="logBase">The log base. Linear if null.</param>
    /// <param name="minValue">The minimum value.</param>
    /// <param name="maxValue">The maximum value.</param>
    /// <returns>The scaling element.</returns>
    protected C.Scaling CreateScaling(
        ValueWithLocation<bool>? reversed,
        ValueWithLocation<decimal>? logBase,
        IAxisRangeValue? minValue,
        IAxisRangeValue? maxValue
    )
    {
        var scaling = new C.Scaling(
            new C.Orientation
            {
                Val = (reversed?.Value ?? false)
                    ? C.OrientationValues.MaxMin
                    : C.OrientationValues.MinMax
            }
        );

        logBase.CallIfNotNull(
            b => scaling.AddChild(new C.LogBase() { Val = (double)b })
        );

        if (minValue?.ToDouble() is double min)
        {
            scaling.AddChild(new C.MinAxisValue() { Val = min });
        }

        if (maxValue?.ToDouble() is double max)
        {
            scaling.AddChild(new C.MaxAxisValue() { Val = max });
        }

        return scaling;
    }

    /// <summary>Creates an axis of the chart.</summary>
    /// <param name="block">The block.</param>
    /// <param name="valueType">The type of the column values.</param>
    /// <param name="id">The ID of the axis.</param>
    /// <param name="minValue">The minimum value.</param>
    /// <param name="maxValue">The maximum value.</param>
    /// <param name="position">The position of the axis.</param>
    /// <param name="crossingAxis">The ID of the crossing axis.</param>
    /// <param name="forceCategorical">
    /// Whether to make the axis categorical regardless of the type of the
    /// column values.
    /// </param>
    /// <returns>The axis element.</returns>
    protected OpenXmlCompositeElement CreateAxis(
        BlockContainer block,
        Type valueType,
        uint id,
        IAxisRangeValue? minValue,
        IAxisRangeValue? maxValue,
        C.AxisPositionValues position,
        uint crossingAxis,
        bool forceCategorical = false
    )
    {
        var dark1 = new D.SchemeColor() { Val = D.SchemeColorValues.Dark1 };
        var isXAxis = position == C.AxisPositionValues.Top
            || position == C.AxisPositionValues.Bottom;
        var axisOptions = isXAxis
            ? block.ChartOptions?.XAxisOptions
            : block.ChartOptions?.YAxisOptions;
        var title = axisOptions?.Title;
        var titleFontSize = block
            .ChartOptions
            ?.AxisTitleFontSize
            ?? ILength.FromPoint(14);
        var titleColor = ColorRenderer
            .ParseColor(block.ChartOptions?.AxisTitleColor);
        var labelFormat = axisOptions?.LabelFormat?.Value ?? "auto";
        var labelRotate = axisOptions?.LabelRotate?.Value ?? 0;
        var labelFontSize = block
            .ChartOptions
            ?.AxisLabelFontSize
            ?? ILength.FromPoint(14);
        var labelColor = ColorRenderer
            .ParseColor(block.ChartOptions?.AxisLabelColor);
        var lineWidth = axisOptions?.LineWidth ?? ILength.FromPoint(1);
        var lineColor = ColorRenderer
            .ParseColor(block.ChartOptions?.AxisLineColor)
            ?? dark1.CloneNode(true);
        var majorUnit = axisOptions?.MajorUnit;
        var minorUnit = axisOptions?.MinorUnit;
        var logBase = axisOptions?.LogBase;
        var reversed = axisOptions?.Reversed;
        var gridMajorWidth = axisOptions?.GridMajorWidth ?? ILength.Zero;
        var gridMinorWidth = axisOptions?.GridMinorWidth ?? ILength.Zero;
        var gridMajorColor = ColorRenderer
            .ParseColor(block.ChartOptions?.GridMajorColor)
            ?? dark1.CloneNode(true);
        var gridMinorColor = ColorRenderer
            .ParseColor(block.ChartOptions?.GridMinorColor)
            ?? dark1.CloneNode(true);

        minValue = axisOptions?.Minimum ?? minValue;
        maxValue = axisOptions?.Maximum ?? maxValue;

        Func<OpenXmlElement[], OpenXmlCompositeElement> constructor;

        if ((valueType == typeof(DateTime) || valueType == typeof(DateOnly)))
        {
            // Date axes are available only on stock charts, line charts,
            // column charts, bar charts, and area charts, so use ValueAxis
            // unless required.
            if (forceCategorical)
            {
                constructor = children => new C.DateAxis(children);
            }
            else
            {
                constructor = children => new C.ValueAxis(children);
            }
        }
        else if (
            !forceCategorical && (
                NumericTypes.Contains(valueType)
                    || valueType == typeof(Boolean)
                    || valueType == typeof(TimeOnly)
            )
        )
        {
            constructor = children => new C.ValueAxis(children);
        }
        else
        {
            constructor = children => new C.CategoryAxis(children);
        }

        var axis = constructor(
            [
                new C.AxisId() { Val = id },
                CreateScaling(reversed, logBase, minValue, maxValue),
                new C.Delete() { Val = false },
                new C.AxisPosition() { Val = position },
                new C.MajorTickMark()
                {
                    Val = valueType == typeof(string)
                        ? C.TickMarkValues.None
                        : C.TickMarkValues.Inside
                },
                new C.MinorTickMark()
                {
                    Val = valueType == typeof(string)
                        ? C.TickMarkValues.None
                        : C.TickMarkValues.Inside
                },
                new C.ChartShapeProperties(CreateOutline(lineColor, lineWidth)),
                new C.TextProperties(
                    new D.BodyProperties()
                    {
                        Rotation = (int)decimal.Round(
                            labelRotate * 21600000 / 360
                        )
                    },
                    new D.Paragraph(
                        new D.ParagraphProperties(
                            CreateDefaultRunProperties(
                                labelColor,
                                labelFontSize
                            )
                        )
                    )
                ),
                new C.CrossingAxis() { Val = crossingAxis }
            ]
        );

        if (title != null)
        {
            var rotation = position switch
            {
                _ when position == C.AxisPositionValues.Left => -5400000,
                _ when position == C.AxisPositionValues.Right => 5400000,
                _ => 0
            };

            axis.AddChild(
                new C.Title(
                    new C.ChartText(
                        new C.RichText(
                            new D.BodyProperties() { Rotation = rotation },
                            new D.Paragraph(
                                new D.ParagraphProperties(
                                    CreateDefaultRunProperties(
                                        titleColor,
                                        titleFontSize
                                    )
                                ),
                                new D.Run(new D.Text(title.Value))
                            )
                        )
                    ),
                    new C.Overlay() { Val = false }
                )
            );
        }

        if (labelFormat == "auto")
        {
            axis.AddChild(
                new C.NumberingFormat()
                {
                    FormatCode = ToFormatCode(valueType) ?? "General",
                    SourceLinked = true
                }
            );
        }
        else
        {
            axis.AddChild(
                new C.NumberingFormat()
                {
                    FormatCode = labelFormat,
                    SourceLinked = false
                }
            );
        }

        if (valueType != typeof(string))
        {
            if (gridMajorWidth.ToEmu() != 0)
            {
                axis.AddChild(
                    new C.MajorGridlines(
                        new C.ChartShapeProperties(
                            CreateOutline(gridMajorColor, gridMajorWidth)
                        )
                    )
                );
            }

            if (gridMinorWidth.ToEmu() != 0)
            {
                axis.AddChild(
                    new C.MinorGridlines(
                        new C.ChartShapeProperties(
                            CreateOutline(gridMinorColor, gridMinorWidth)
                        )
                    )
                );
            }

            if (majorUnit != null && axis is not C.CategoryAxis)
            {
                axis.AddChild(new C.MajorUnit() { Val = majorUnit.ToDouble() });
            }

            if (minorUnit != null && axis is not C.CategoryAxis)
            {
                axis.AddChild(new C.MinorUnit() { Val = minorUnit.ToDouble() });
            }
        }

        return axis;
    }

    /// <summary>Converts a type to a Excel format code.</summary>
    /// <param name="valueType">The type.</param>
    /// <returns>The format code or null for default.</returns>
    public static string? ToFormatCode(Type valueType)
    {
        if (NumericTypes.Contains(valueType) || valueType == typeof(Boolean))
        {
            return "General";
        }
        else if (
            valueType == typeof(DateTime) || valueType == typeof(DateTimeOffset)
        )
        {
            return "[$-F800]m/d/yyyy\\ [$-F400]hh:mm";
        }
        else if (valueType == typeof(DateOnly))
        {
            return "[$-F800]m/d/yyyy";
        }
        else if (valueType == typeof(TimeOnly))
        {
            return "[$-F400]hh:mm";
        }
        else
        {
            return null;
        }
    }

    /// <summary>Creates the legend of the chart if any.</summary>
    /// <param name="block">The block.</param>
    /// <returns>
    /// The legend element or null if the legend is not enabled.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the legend position is invalid.
    /// </exception>
    protected C.Legend? CreateLegend(BlockContainer block)
    {
        var position = block.ChartOptions?.LegendPosition?.Value ?? "right";

        if (position == "none")
        {
            return null;
        }

        var lineWidth = block.ChartOptions?.LegendLineWidth ?? ILength.Zero;
        var lineColor = ColorRenderer
            .ParseColor(block.ChartOptions?.LegendLineColor)
            ?? new D.SchemeColor() { Val = D.SchemeColorValues.Dark1 };
        var fontSize = block
            .ChartOptions
            ?.LegendFontSize
            ?? ILength.FromPoint(14);
        var color = ColorRenderer
            .ParseColor(block.ChartOptions?.LegendColor);
        var legendPosition = position switch
        {
            "bottom" => C.LegendPositionValues.Bottom,
            "top_right" => C.LegendPositionValues.TopRight,
            "left" => C.LegendPositionValues.Left,
            "right" => C.LegendPositionValues.Right,
            "top" => C.LegendPositionValues.Top,
            _ => throw new PlotanceException(
                block.ChartOptions?.LegendPosition?.Path,
                block.ChartOptions?.LegendPosition?.Line,
                $"Invalid legend position: {position}"
            )
        };

        return new C.Legend(
            new C.LegendPosition()
            {
                Val = legendPosition
            },
            new C.Overlay() { Val = false },
            new C.ChartShapeProperties(
                new D.NoFill(),
                CreateOutline(lineColor, lineWidth)
            ),
            new C.TextProperties(
                new D.BodyProperties(),
                new D.Paragraph(
                    new D.ParagraphProperties(
                        CreateDefaultRunProperties(color, fontSize)
                    )
                )
            )
        );
    }

    /// <summary>Creates the title of the chart if any.</summary>
    /// <param name="block">The block.</param>
    /// <returns>
    /// The title element or null if the title is not enabled.
    /// </returns>
    protected C.Title? CreateTitle(BlockContainer block)
    {
        var text = block.ChartOptions?.Title;

        if (text == null)
        {
            return null;
        }

        var fontSize = block
            .ChartOptions
            ?.TitleFontSize
            ?? ILength.FromPoint(18);
        var color = ColorRenderer
            .ParseColor(block.ChartOptions?.TitleColor);

        return new C.Title(
            new C.ChartText(
                new C.RichText(
                    new D.BodyProperties(),
                    new D.Paragraph(
                        new D.ParagraphProperties(
                            CreateDefaultRunProperties(color, fontSize)
                        ),
                        new D.Run(new D.Text(text.Value))
                    )
                )
            ),
            new C.Overlay() { Val = false }
        );
    }

    /// <summary>Creates the data labels of the chart if any.</summary>
    /// <param name="chartType">The type for the chart.</param>
    /// <param name="block">The block.</param>
    /// <param name="valueType">The type of the column values.</param>
    /// <returns>
    /// The data labels element or null if the data labels are not enabled.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the data label position is invalid.
    /// </exception>
    protected C.DataLabels? CreateDataLabels(
        ChartType chartType,
        BlockContainer block,
        Type valueType
    )
    {
        var position = block.ChartOptions?.DataLabelPosition;
        var contents = block
            .ChartOptions
            ?.DataLabelContents
            ?.Value
            ?.Select(v => v.Value)
            ?? new List<string> { "y_value" };

        if (
            position == null
                || position.Value == "none"
                || !contents.Any()
        )
        {
            return null;
        }

        var format = block.ChartOptions?.DataLabelFormat?.Value ?? "auto";
        var labelRotate = block.ChartOptions?.DataLabelRotate?.Value ?? 0;
        var fontSize = block
            .ChartOptions
            ?.DataLabelFontSize
            ?? ILength.FromPoint(14);
        var color = ColorRenderer
            .ParseColor(block.ChartOptions?.DataLabelColor);

        C.DataLabelPositionValues ThrowInvalidDataLabelPosition(
            bool forTheChartType
        )
        {
            throw new PlotanceException(
                position?.Path,
                position?.Line,
                string.Format(
                    "Invalid data label position{0}: {1}",
                    forTheChartType ? " for the chart type" : "",
                    position?.Value
                )
            );
        }

        C.DataLabelPositionValues? dataLabelPosition = chartType switch
        {
            ChartType.Bar => position.Value switch
            {
                "inside_end" => C.DataLabelPositionValues.InsideEnd,
                "outside_end" => C.DataLabelPositionValues.OutsideEnd,
                "inside_center" => C.DataLabelPositionValues.Center,
                "inside_base" => C.DataLabelPositionValues.InsideBase,
                "center" => C.DataLabelPositionValues.Center,

                "left" or "right" or "above" or "below"
                    => ThrowInvalidDataLabelPosition(true),

                _ => ThrowInvalidDataLabelPosition(false)
            },

            ChartType.Scatter or ChartType.Bubble => position.Value switch
            {
                "center" => C.DataLabelPositionValues.Center,
                "left" => C.DataLabelPositionValues.Left,
                "right" => C.DataLabelPositionValues.Right,
                "above" => C.DataLabelPositionValues.Top,
                "below" => C.DataLabelPositionValues.Bottom,

                "inside_end"
                or "outside_end"
                or "inside_center"
                or "inside_base"
                    => ThrowInvalidDataLabelPosition(true),

                _ => ThrowInvalidDataLabelPosition(false)
            },

            // Area chart cannot have DataLabelPosition.
            // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/7dd610ae-d805-40a4-ae3b-c171b6bf08e3
            ChartType.Area => position.Value switch
            {
                "inside_center" => null,
                "center" => null,

                "left"
                or "right"
                or "above"
                or "below"
                or "inside_end"
                or "outside_end"
                or "inside_base"
                    => ThrowInvalidDataLabelPosition(true),

                _ => ThrowInvalidDataLabelPosition(false)
            },

            _ => throw new ArgumentException(
                $"Invalid chart type {chartType}",
                nameof(chartType)
            )
        };

        var barDirection = block
            .ChartOptions
            ?.BarDirection
            ?.Value
            ?? "vertical";
        var swapAxis = chartType == ChartType.Bar
            && barDirection == "horizontal";

        var dataLabels = new C.DataLabels(
            new C.Delete() { Val = false },
            new C.ChartShapeProperties(new D.NoFill()),
            new C.TextProperties(
                new D.BodyProperties()
                {
                    Rotation = (int)decimal.Round(
                        labelRotate * 21600000 / 360
                    )
                },
                new D.Paragraph(
                    new D.ParagraphProperties(
                        CreateDefaultRunProperties(color, fontSize)
                    )
                )
            ),
            new C.ShowLegendKey() { Val = contents.Contains("legend_key") },
            new C.ShowValue()
            {
                Val = swapAxis
                    ? contents.Contains("x_value")
                    : contents.Contains("y_value")
            },
            new C.ShowCategoryName()
            {
                Val = swapAxis
                    ? contents.Contains("y_value")
                    : contents.Contains("x_value")
            },
            new C.ShowSeriesName() { Val = contents.Contains("series_name") },
            new C.ShowPercent() { Val = contents.Contains("percent") },
            new C.ShowBubbleSize() { Val = contents.Contains("bubble_size") }
        );

        if (format == "auto")
        {
            dataLabels.AddChild(
                new C.NumberingFormat()
                {
                    FormatCode = ToFormatCode(valueType) ?? "General",
                    SourceLinked = true
                }
            );
        }
        else
        {
            dataLabels.AddChild(
                new C.NumberingFormat()
                {
                    FormatCode = format,
                    SourceLinked = false
                }
            );
        }

        if (dataLabelPosition != null)
        {
            dataLabels.AddChild(
                new C.DataLabelPosition() { Val = dataLabelPosition }
            );
        }

        return dataLabels;
    }
}
