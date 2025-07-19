// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Collections.Generic;
using System.Numerics;
using DocumentFormat.OpenXml.Packaging;
using Plotance.Models;

using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Plotance.Renderers.PowerPoint;

/// <summary>Converts a query result to an area chart.</summary>
public class AreaChartRenderer : CategoryToAbsoluteValueChartRenderer
{
    /// <inheritdoc/>
    protected override C.Chart CreateChart(
        BlockContainer block,
        QueryResultSet queryResult,
        bool shouldUseAutoMinimum,
        bool shouldUseAutoMaximum
    )
    {
        var colors = ColorRenderer
            .ParseColors(block.ChartOptions?.SeriesColors);
        var lineWidth = block.ChartOptions?.LineWidth ?? ILength.Zero;
        var lineColor = ColorRenderer
            .ParseColor(block.ChartOptions?.LineColor);

        C.AreaChartSeries CreateChartSeries(int columnIndex)
        {
            var spreadsheetColumnName = Spreadsheets.GetColumn(columnIndex);
            var index = columnIndex - 1;
            var outline = lineColor == null
                ? CreateOutline(colors, index, lineWidth)
                : CreateOutline(lineColor.CloneNode(true), lineWidth);

            return new C.AreaChartSeries(
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
                        block.ChartOptions?.FillOpacity
                    ),
                    outline
                ),
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
        var areaChart = new C.AreaChart([
            new C.Grouping() { Val = C.GroupingValues.Stacked },
            new C.VaryColors() { Val = false },
            .. chartSeries,
            new C.AxisId { Val = 1 },
            new C.AxisId { Val = 2 }
        ]);
        var valueType = queryResult.Columns.Count > 1
            ? queryResult.Columns[1].Type
            : typeof(Double);
        var dataLabels = CreateDataLabels(ChartType.Area, block, valueType);

        if (dataLabels != null)
        {
            areaChart.AddChild(dataLabels);
        }

        return new C.Chart(
            new C.PlotArea(
                areaChart,
                CreateCategoryAxis(
                    block: block,
                    queryResult: queryResult,
                    position: C.AxisPositionValues.Bottom,
                    crossingAxis: 2
                ),
                CreateValueAxis(
                    block: block,
                    queryResult: queryResult,
                    shouldUseAutoMinimum: shouldUseAutoMinimum,
                    shouldUseAutoMaximum: shouldUseAutoMaximum,
                    position: C.AxisPositionValues.Left,
                    crossingAxis: 1
                )
            )
        );
    }
}
