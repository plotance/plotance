// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Collections.Generic;
using System.Numerics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Plotance.Models;

using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Plotance.Renderers.PowerPoint;

/// <summary>
/// Base class for chart renderer for charts with category axis and value axis
/// (i.e. bar and area charts).
/// </summary>
public abstract class CategoryToAbsoluteValueChartRenderer : ChartRenderer
{
    /// <summary>
    /// Creates a chart.
    /// </summary>
    /// <param name="block">The block to render.</param>
    /// <param name="queryResult">The query result.</param>
    /// <param name="shouldUseAutoMinimum">
    /// Whether to use auto minimum or set minimum to 0.
    /// </param>
    /// <param name="shouldUseAutoMaximum">
    /// Whether to use auto maximum or set maximum to 0.
    /// </param>
    /// <returns>The chart.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the chart configuration is invalid.
    /// </exception>
    protected abstract C.Chart CreateChart(
        BlockContainer block,
        QueryResultSet queryResult,
        bool shouldUseAutoMinimum,
        bool shouldUseAutoMaximum
    );

    /// <inheritdoc/>
    protected override void RenderChart(
        ChartPart chartPart,
        BlockContainer block,
        QueryResultSet queryResult,
        string relationShipId
    )
    {
        var shouldUseAutoMaximum = ShouldUseAutoMaximum(queryResult.Rows);
        var shouldUseAutoMinimum = ShouldUseAutoMinimum(queryResult.Rows);
        var chart = CreateChart(
            block,
            queryResult,
            shouldUseAutoMinimum,
            shouldUseAutoMaximum
        );

        if (CreateTitle(block) is C.Title title)
        {
            chart.AddChild(title);
        }

        if (CreateLegend(block) is C.Legend legend)
        {
            chart.AddChild(legend);
        }

        chartPart.ChartSpace = new C.ChartSpace(
            new C.Date1904() { Val = false },
            chart,
            new C.ExternalData() { Id = relationShipId }
        );
    }

    /// <summary>
    /// Creates the category axis.
    /// </summary>
    /// <param name="block">The block to render.</param>
    /// <param name="queryResult">The query result.</param>
    /// <param name="position">The position of the axis.</param>
    /// <param name="crossingAxis">The ID of the crossing axis.</param>
    /// <returns>The category axis.</returns>
    protected OpenXmlCompositeElement CreateCategoryAxis(
        BlockContainer block,
        QueryResultSet queryResult,
        C.AxisPositionValues position,
        uint crossingAxis
    )
    {
        return CreateAxis(
            block: block,
            valueType: queryResult.Columns[0].Type,
            id: 1,
            minValue: null,
            maxValue: null,
            position: position,
            crossingAxis: crossingAxis,
            forceCategorical: true
        );
    }

    /// <summary>
    /// Creates the value axis.
    /// </summary>
    /// <param name="block">The block to render.</param>
    /// <param name="queryResult">The query result.</param>
    /// <param name="shouldUseAutoMinimum">
    /// Whether to use auto minimum or set minimum to 0.
    /// </param>
    /// <param name="shouldUseAutoMaximum">
    /// Whether to use auto maximum or set maximum to 0.
    /// </param>
    /// <param name="position">The position of the axis.</param>
    /// <param name="crossingAxis">The ID of the crossing axis.</param>
    /// <returns>The value axis.</returns>
    protected OpenXmlCompositeElement CreateValueAxis(
        BlockContainer block,
        QueryResultSet queryResult,
        bool shouldUseAutoMinimum,
        bool shouldUseAutoMaximum,
        C.AxisPositionValues position,
        uint crossingAxis
    )
    {
        if (queryResult.Columns.Count > 1)
        {
            return CreateAxis(
                block: block,
                valueType: queryResult.Columns[1].Type,
                id: 2,
                minValue: shouldUseAutoMinimum
                    ? null
                    : new DecimalAxisRangeValue(0),
                maxValue: shouldUseAutoMaximum
                    ? null
                    : new DecimalAxisRangeValue(0),
                position: position,
                crossingAxis: crossingAxis
            );
        }
        else
        {
            return CreateAxis(
                block: block,
                valueType: typeof(Double),
                id: 2,
                minValue: null,
                maxValue: null,
                position: position,
                crossingAxis: crossingAxis
            );
        }
    }

    /// <summary>
    /// Determines whether to use auto maximum or set maximum to 0. Basically,
    /// if any positive value exists, then auto maximum should be used. If all
    /// values are negative, then set maximum to 0. However, for some types
    /// (e.g. DateTime or TimeOnly), using zero as maximum value does not make
    /// sense, so we always use auto maximum.
    /// </summary>
    /// <param name="rows">The rows to check.</param>
    /// <returns>
    /// True if auto maximum should be used; otherwise, false.
    /// </returns>
    private bool ShouldUseAutoMaximum(IEnumerable<IEnumerable<object>> rows)
        => rows.Any(ShouldUseAutoMaximum);

    /// <summary>
    /// Determines whether to use auto maximum or set maximum to 0.
    /// </summary>
    /// <param name="row">The row to check.</param>
    /// <returns>
    /// True if auto maximum should be used; otherwise, false.
    /// </returns>
    private bool ShouldUseAutoMaximum(IEnumerable<object> row)
        // Skip the first column (category names)
        => row.Skip(1).Any(ShouldUseAutoMaximum);

    /// <summary>
    /// Determines whether to use auto maximum or set maximum to 0.
    /// </summary>
    /// <param name="obj">The object to check.</param>
    /// <returns>
    /// True if auto maximum should be used; otherwise, false.
    /// </returns>
    private bool ShouldUseAutoMaximum(object obj) => obj switch
    {
        Boolean => true,
        SByte sByte => IsPositive(sByte),
        Byte @byte => IsPositive(@byte),
        Decimal @decimal => IsPositive(@decimal),
        Double @double => IsPositive(@double),
        Single single => IsPositive(single),
        Int16 int16 => IsPositive(int16),
        UInt16 uint16 => IsPositive(uint16),
        Int32 int32 => IsPositive(int32),
        UInt32 uint32 => IsPositive(uint32),
        Int64 int64 => IsPositive(int64),
        UInt64 uint64 => IsPositive(uint64),
        BigInteger bigInteger => IsPositive(bigInteger),
        DateTime => true,
        DateOnly => true,
        TimeOnly => true,
        _ => true
    };

    /// <summary>Determines whether the specified number is positive.</summary>
    /// <typeparam name="T">The type of the number.</typeparam>
    /// <param name="number">The number to check.</param>
    /// <returns>True if the number is positive; otherwise, false.</returns>
    private bool IsPositive<T>(T number) where T : INumberBase<T>
        => T.IsPositive(number);

    /// <summary>
    /// Determines whether to use auto minimum or set minimum to 0. Basically,
    /// if any negative value exists, then auto minimum should be used. If all
    /// values are positive, then set minimum to 0. However, for some types
    /// (e.g. DateTime or TimeOnly), using zero as minimum value does not make
    /// sense, so we always use auto minimum.
    /// </summary>
    /// <param name="rows">The rows to check.</param>
    /// <returns>
    /// True if auto minimum should be used; otherwise, false.
    /// </returns>
    private bool ShouldUseAutoMinimum(IEnumerable<IEnumerable<object>> rows)
        => rows.Any(ShouldUseAutoMinimum);

    /// <summary>
    /// Determines whether to use auto minimum or set minimum to 0.
    /// </summary>
    /// <param name="row">The row to check.</param>
    /// <returns>
    /// True if auto minimum should be used; otherwise, false.
    /// </returns>
    private bool ShouldUseAutoMinimum(IEnumerable<object> row)
        // Skip the first column (category names)
        => row.Skip(1).Any(ShouldUseAutoMinimum);

    /// <summary>
    /// Determines whether to use auto minimum or set minimum to 0.
    /// </summary>
    /// <param name="obj">The object to check.</param>
    /// <returns>
    /// True if auto minimum should be used; otherwise, false.
    /// </returns>
    private bool ShouldUseAutoMinimum(object obj) => obj switch
    {
        Boolean => false,
        SByte sByte => IsNegative(sByte),
        Byte => false,
        Decimal @decimal => IsNegative(@decimal),
        Double @double => IsNegative(@double),
        Single single => IsNegative(single),
        Int16 int16 => IsNegative(int16),
        UInt16 => false,
        Int32 int32 => IsNegative(int32),
        UInt32 => false,
        Int64 int64 => IsNegative(int64),
        UInt64 => false,
        BigInteger bigInteger => IsNegative(bigInteger),
        DateTime => true,
        DateOnly => true,
        TimeOnly => false,
        _ => true
    };

    /// <summary>Determines whether the specified number is negative.</summary>
    /// <typeparam name="T">The type of the number.</typeparam>
    /// <param name="number">The number to check.</param>
    /// <returns>True if the number is negative; otherwise, false.</returns>
    private bool IsNegative<T>(T number) where T : INumberBase<T>
        => T.IsNegative(number);
}
