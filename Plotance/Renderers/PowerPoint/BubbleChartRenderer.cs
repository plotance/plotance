// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using Plotance.Models;

using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Plotance.Renderers.PowerPoint;

/// <summary>Converts a query result to a bubble chart.</summary>
public class BubbleChartRenderer : ChartRenderer
{
    /// <inheritdoc/>
    protected override void RenderChart(
        ChartPart chartPart,
        ImplicitSectionColumn column,
        QueryResultSet queryResult,
        string relationShipId
    )
    {
        var colors = ColorRenderer
            .ParseColors(column.ChartOptions?.SeriesColors);

        C.BubbleChartSeries CreateChartSeries(int columnIndex)
        {
            var spreadsheetColumnName = Spreadsheets.GetColumn(columnIndex);
            var index = (columnIndex - 1) / 2;

            return new C.BubbleChartSeries(
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
                        column.ChartOptions?.MarkerFillOpacity
                          ?? column.ChartOptions?.FillOpacity
                    ),
                    CreateOutline(
                        colors,
                        index,
                        column.ChartOptions?.MarkerLineWidth
                          ?? column.ChartOptions?.LineWidth
                    )
                ),
                new C.XValues(
                    ValueReferenceFactory.CreateFromQueryResultColumn(
                        queryResult,
                        0
                    )
                ),
                new C.YValues(
                    ValueReferenceFactory.CreateFromQueryResultColumn(
                        queryResult,
                        columnIndex
                    )
                ),
                new C.BubbleSize(
                    ValueReferenceFactory.CreateFromQueryResultColumn(
                        queryResult,
                        columnIndex + 1
                    )
                )
            );
        }

        // Skip the first column (category names)
        var seriesCount = (queryResult.Columns.Count - 1) / 2;
        var chartSeries = Enumerable
            .Range(0, seriesCount)
            .Select(i => CreateChartSeries(i * 2 + 1));
        var valueType = queryResult.Columns.Count > 1
            ? queryResult.Columns[1].Type
            : typeof(Double);
        var bubbleChart = new C.BubbleChart([
            new C.VaryColors() { Val = false },
            .. chartSeries,
            new C.AxisId { Val = 1 },
            new C.AxisId { Val = 2 }
        ]);
        var dataLabels = CreateDataLabels(ChartType.Bubble, column, valueType);

        if (dataLabels != null)
        {
            bubbleChart.AddChild(dataLabels);
        }

        var chart = new C.Chart(
            new C.PlotArea(
                bubbleChart,
                CreateAxis(
                    column: column,
                    valueType: queryResult.Columns[0].Type,
                    id: 1,
                    minValue: null,
                    maxValue: null,
                    position: C.AxisPositionValues.Bottom,
                    crossingAxis: 2
                ),
                CreateAxis(
                    column: column,
                    valueType: valueType,
                    id: 2,
                    minValue: null,
                    maxValue: null,
                    position: C.AxisPositionValues.Left,
                    crossingAxis: 1
                )
            )
        );

        if (CreateTitle(column) is C.Title title)
        {
            chart.AddChild(title);
        }

        if (CreateLegend(column) is C.Legend legend)
        {
            chart.AddChild(legend);
        }

        chartPart.ChartSpace = new C.ChartSpace(
            new C.Date1904() { Val = false },
            chart,
            new C.ExternalData() { Id = relationShipId }
        );
    }
}
