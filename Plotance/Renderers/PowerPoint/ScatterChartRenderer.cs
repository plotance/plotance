// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using DocumentFormat.OpenXml.Packaging;
using Plotance.Models;

using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Plotance.Renderers.PowerPoint;

/// <summary>Converts a query result to a scatter chart.</summary>
public class ScatterChartRenderer : ChartRenderer
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
        var isLineChart = column.ChartOptions?.Format?.Value == "line";
        var defaultLineWidth = isLineChart
            ? ILength.FromPoint(1)
            : ILength.Zero;
        var lineWidth = column.ChartOptions?.LineWidth ?? defaultLineWidth;
        var lineColor = ColorRenderer
            .ParseColor(column.ChartOptions?.LineColor);

        C.ScatterChartSeries CreateChartSeries(int columnIndex)
        {
            var spreadsheetColumnName = Spreadsheets.GetColumn(columnIndex);
            var index = columnIndex - 1;
            var outline = lineColor == null
                ? CreateOutline(colors, index, lineWidth)
                : CreateOutline(lineColor.CloneNode(true), lineWidth);

            return new C.ScatterChartSeries(
                new C.Index() { Val = (uint)index },
                new C.Order() { Val = (uint)index },
                new C.SeriesText(
                    ValueReferenceFactory.Create(
                        $"Sheet1!${spreadsheetColumnName}$1",
                        typeof(String),
                        [queryResult.Columns[columnIndex].Name]
                    )
                ),
                new C.ChartShapeProperties(outline),
                new C.Marker(
                    new C.Size()
                    {
                        Val = (byte)(
                            column.ChartOptions?.MarkerSize?.ToPoint()
                            ?? 5
                        )
                    },
                    new C.ChartShapeProperties(
                        CreateFill(
                            colors,
                            index,
                            column.ChartOptions?.MarkerFillOpacity
                        ),
                        CreateOutline(
                            colors,
                            index,
                            column.ChartOptions?.MarkerLineWidth
                        )
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
                new C.Smooth { Val = false }
            );
        }

        var chartSeries = queryResult.Columns
            .Skip(1) // Skip the first column (category names)
            .Select((_, columnIndex) => CreateChartSeries(columnIndex + 1));
        var valueType = queryResult.Columns.Count > 1
            ? queryResult.Columns[1].Type
            : typeof(Double);
        var scatterChart = new C.ScatterChart([
            new C.ScatterStyle
            {
                Val = C.ScatterStyleValues.LineMarker
            },
            new C.VaryColors() { Val = false },
            .. chartSeries,
            new C.AxisId { Val = 1 },
            new C.AxisId { Val = 2 }
        ]);
        var dataLabels = CreateDataLabels(ChartType.Scatter, column, valueType);

        if (dataLabels != null)
        {
            scatterChart.AddChild(dataLabels);
        }

        var (minX, maxX) = isLineChart
            ? ComputeValueRange(queryResult, 0)
            : (null, null);

        var chart = new C.Chart(
            new C.PlotArea(
                scatterChart,
                CreateAxis(
                    column: column,
                    valueType: queryResult.Columns[0].Type,
                    id: 1,
                    minValue: minX,
                    maxValue: maxX,
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

    private (IAxisRangeValue?, IAxisRangeValue?) ComputeValueRange(
        QueryResultSet queryResult,
        int columnIndex
    )
    {
        if (!queryResult.Rows.Any())
        {
            return (null, null);
        }

        var minimum = queryResult.Rows.Select(row => row[columnIndex]).Min();
        var maximum = queryResult.Rows.Select(row => row[columnIndex]).Max();

        return (
            new TextAxisRangeValue(
                null,
                0,
                Spreadsheets.FormatValue(minimum)
            ),
            new TextAxisRangeValue(
                null,
                0,
                Spreadsheets.FormatValue(maximum)
            )
        );
    }
}
