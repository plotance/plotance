// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Collections.Generic;
using System.Numerics;
using DocumentFormat.OpenXml.Packaging;
using Plotance.Models;

using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Plotance.Renderers.PowerPoint;

/// <summary>Converts a query result to a bar chart.</summary>
public class BarChartRenderer : CategoryToAbsoluteValueChartRenderer
{
    /// <inheritdoc/>
    protected override C.Chart CreateChart(
        ImplicitSectionColumn column,
        QueryResultSet queryResult,
        bool shouldUseAutoMinimum,
        bool shouldUseAutoMaximum
    )
    {
        var colors = ColorRenderer
            .ParseColors(column.ChartOptions?.SeriesColors);
        var lineWidth = column.ChartOptions?.LineWidth ?? ILength.Zero;
        var lineColor = ColorRenderer
            .ParseColor(column.ChartOptions?.LineColor);

        C.BarChartSeries CreateChartSeries(int columnIndex)
        {
            var spreadsheetColumnName = Spreadsheets.GetColumn(columnIndex);
            var index = columnIndex - 1;
            var outline = lineColor == null
                ? CreateOutline(colors, index, lineWidth)
                : CreateOutline(lineColor.CloneNode(true), lineWidth);

            return new C.BarChartSeries(
                new C.Index() { Val = (uint)index },
                new C.Order() { Val = (uint)index },
                new C.SeriesText(
                    ValueReferenceFactory.Create(
                        $"Sheet1!${spreadsheetColumnName}$1",
                        typeof(String),
                        [queryResult.Columns[columnIndex].Name]
                    )
                ),
                new C.ChartShapeProperties(
                    CreateFill(
                        colors,
                        index,
                        column.ChartOptions?.FillOpacity
                    ),
                    outline
                ),
                new C.InvertIfNegative() { Val = false },
                new C.CategoryAxisData(
                    ValueReferenceFactory.CreateFromQueryResultColumn(
                        queryResult,
                        0
                    )
                ),
                new C.Values(
                    ValueReferenceFactory.CreateFromQueryResultColumn(
                        queryResult,
                        columnIndex
                    )
                )
            );
        }

        var chartSeries = queryResult.Columns
            .Skip(1) // Skip the first column (category names)
            .Select((_, columnIndex) => CreateChartSeries(columnIndex + 1));
        var barDirection = ToBarDirectionValue(
            column.ChartOptions?.BarDirection
        );
        var barGrouping = ToBarGroupingValue(column.ChartOptions?.BarGrouping);
        var barGap = column.ChartOptions?.BarGap?.Value ?? 20;
        var barOverlap = column.ChartOptions?.BarOverlap?.Value ?? (
            barGrouping == C.BarGroupingValues.Clustered
                ? 0
                : 100
        );
        var categoryAxisPosition = barDirection == C.BarDirectionValues.Bar
            ? C.AxisPositionValues.Left
            : C.AxisPositionValues.Bottom;
        var valueAxisPosition = barDirection == C.BarDirectionValues.Bar
            ? C.AxisPositionValues.Bottom
            : C.AxisPositionValues.Left;
        var barChart = new C.BarChart([
            new C.BarDirection() { Val = barDirection },
            new C.BarGrouping() { Val = barGrouping },
            new C.VaryColors() { Val = false },
            .. chartSeries,
            new C.GapWidth() { Val = (ushort)barGap },
            new C.Overlap() { Val = (sbyte)barOverlap },
            new C.AxisId { Val = 1 },
            new C.AxisId { Val = 2 }
        ]);
        var valueType = queryResult.Columns.Count > 1
            ? queryResult.Columns[1].Type
            : typeof(Double);
        var dataLabels = CreateDataLabels(ChartType.Bar, column, valueType);

        if (dataLabels != null)
        {
            barChart.AddChild(dataLabels);
        }

        return new C.Chart(
            new C.PlotArea(
                barChart,
                CreateCategoryAxis(
                    column: column,
                    queryResult: queryResult,
                    position: categoryAxisPosition,
                    crossingAxis: 2
                ),
                CreateValueAxis(
                    column: column,
                    queryResult: queryResult,
                    shouldUseAutoMinimum: shouldUseAutoMinimum,
                    shouldUseAutoMaximum: shouldUseAutoMaximum,
                    position: valueAxisPosition,
                    crossingAxis: 1
                )
            )
        );
    }

    /// <summary>
    /// Converts a bar direction string to a bar direction value.
    /// </summary>
    /// <param name="barDirection">The bar direction string.</param>
    /// <returns>The bar direction value.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the bar direction string is invalid.
    /// </exception>
    private C.BarDirectionValues ToBarDirectionValue(
        ValueWithLocation<string>? barDirection
    ) => barDirection?.Value switch
    {
        "horizontal" => C.BarDirectionValues.Bar,
        "vertical" => C.BarDirectionValues.Column,
        null => C.BarDirectionValues.Column,
        _ => throw new PlotanceException(
            barDirection?.Path,
            barDirection?.Line,
            $"Invalid bar direction value: {barDirection?.Value}"
        )
    };

    /// <summary>
    /// Converts a bar grouping string to a bar grouping value.
    /// </summary>
    /// <param name="barGrouping">The bar grouping string.</param>
    /// <returns>The bar grouping value.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the bar grouping string is invalid.
    /// </exception>
    private C.BarGroupingValues ToBarGroupingValue(
        ValueWithLocation<string>? barGrouping
    ) => barGrouping?.Value switch
    {
        "stacked" => C.BarGroupingValues.Stacked,
        "percent_stacked" => C.BarGroupingValues.PercentStacked,
        "clustered" => C.BarGroupingValues.Clustered,
        null => C.BarGroupingValues.Clustered,
        _ => throw new PlotanceException(
            barGrouping?.Path,
            barGrouping?.Line,
            $"Invalid bar grouping value: {barGrouping?.Value}"
        )
    };
}
