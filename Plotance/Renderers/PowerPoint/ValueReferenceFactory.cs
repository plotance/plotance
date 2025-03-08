// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using Plotance.Models;

using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace Plotance.Renderers.PowerPoint;

/// <summary>
/// Base factory class for creating chart value references in Open XML charts.
/// Provides methods for generating string or numeric reference elements based
/// on query results.
/// </summary>
public abstract class ValueReferenceFactory
{
    /// <summary>
    /// Creates a value reference from a column in a query result set.
    /// </summary>
    /// <param name="queryResult">
    /// The query result set containing the data.
    /// </param>
    /// <param name="columnIndex">
    /// The index of the column to create a reference for.
    /// </param>
    /// <returns>
    /// An Open XML composite element representing the value reference,
    /// (i.e. StringReference or NumberReference).
    /// </returns>
    public static OpenXmlCompositeElement CreateFromQueryResultColumn(
        QueryResultSet queryResult,
        int columnIndex
    )
    {
        var spreadsheetColumnName = Spreadsheets.GetColumn(columnIndex);
        var rowFrom = 2;
        var rowTo = queryResult.Rows.Count + 1;
        var rangeFrom = $"${spreadsheetColumnName}${rowFrom}";
        var rangeTo = $"${spreadsheetColumnName}${rowTo}";
        var valuesRange = $"Sheet1!{rangeFrom}:{rangeTo}";

        return Create(
            valuesRange,
            queryResult.Columns[columnIndex].Type,
            queryResult.Rows.Select(
                row => Spreadsheets.FormatValue(row[columnIndex])
            )
        );
    }

    /// <summary>
    /// Creates a value reference for a range of cells in a spreadsheet.
    /// </summary>
    /// <param name="rangeText">
    /// The text representation of the cell range of a column in Excel format
    /// (e.g., "Sheet1!$A$1:$A$10").
    /// </param>
    /// <param name="valueType">
    /// The type of the values in the range, used to determine whether to create
    /// a numeric or string reference and the format code.
    /// </param>
    /// <param name="values">
    /// The values to include in the reference cache.
    /// </param>
    /// <returns>
    /// An Open XML composite element representing the value reference,
    /// (i.e. StringReference or NumberReference).
    /// </returns>
    public static OpenXmlCompositeElement Create(
        string rangeText,
        Type valueType,
        IEnumerable<string> values
    )
    {
        var isNumeric = ChartRenderer.NumericTypes.Contains(valueType)
            || valueType == typeof(Boolean)
            || valueType == typeof(DateTime)
            || valueType == typeof(DateOnly)
            || valueType == typeof(TimeOnly);
        ValueReferenceFactory valueReferenceFactory = isNumeric
            ? new NumericReferenceFactory()
            : new StringReferenceFactory();

        return valueReferenceFactory.CreateReference(
            rangeText,
            valueType,
            values
        );
    }

    /// <summary>
    /// Creates a value reference for a range of cells in a spreadsheet.
    /// </summary>
    /// <param name="rangeText">
    /// The text representation of the cell range of a column in Excel format.
    /// </param>
    /// <param name="valueType">
    /// The type of the values in the range, used to determine the format code.
    /// </param>
    /// <param name="values">
    /// The values to include in the reference cache.
    /// </param>
    /// <returns>
    /// An OpenXml composite element representing the value reference,
    /// (i.e. StringReference or NumberReference).
    /// </returns>
    public OpenXmlCompositeElement CreateReference(
        string rangeText,
        Type valueType,
        IEnumerable<string> values
    )
    {
        var points = values.Select(
            (value, index) => CreatePoint(
                new C.NumericValue(value),
                (uint)index
            )
        );
        var cache = CreateCache([
            new C.PointCount() { Val = (uint)values.Count() },
            .. points
        ]);
        var formatCode = ChartRenderer.ToFormatCode(valueType);

        if (formatCode != null)
        {
            cache.AddChild(new C.FormatCode(formatCode));
        }

        return CreateReference([new C.Formula(rangeText), cache]);
    }

    /// <summary>
    /// Creates a reference element with the specified children.
    /// </summary>
    /// <param name="children">
    /// The child elements to include in the reference.
    /// </param>
    /// <returns>
    /// An OpenXml composite element representing the specific type of
    /// reference, (i.e. StringReference or NumberReference).
    /// </returns>
    protected abstract OpenXmlCompositeElement CreateReference(
        params OpenXmlElement[] children
    );

    /// <summary>
    /// Creates a cache element with the specified children.
    /// </summary>
    /// <param name="children">
    /// The child elements to include in the cache.
    /// </param>
    /// <returns>
    /// An OpenXml composite element representing the specific type of cache,
    /// (i.e. StringCache or NumberingCache).
    /// </returns>
    protected abstract OpenXmlCompositeElement CreateCache(
        params OpenXmlElement[] children
    );

    /// <summary>
    /// Creates a point element with the specified value and index.
    /// </summary>
    /// <param name="child">The numeric value to include in the point.</param>
    /// <param name="index">The index of the point in the series.</param>
    /// <returns>
    /// An OpenXml composite element representing the specific type of point,
    /// (i.e. StringPoint or NumericPoint).
    /// </returns>
    protected abstract OpenXmlCompositeElement CreatePoint(
        C.NumericValue child,
        uint index
    );
}

/// <summary>Creates string references in OpenXml charts.</summary>
public class StringReferenceFactory : ValueReferenceFactory
{
    /// <summary>
    /// Creates a string reference element with the specified children.
    /// </summary>
    /// <param name="children">
    /// The child elements to include in the reference.
    /// </param>
    /// <returns>A StringReference element.</returns>
    protected override OpenXmlCompositeElement CreateReference(
        params OpenXmlElement[] children
    ) => new C.StringReference(children);

    /// <summary>
    /// Creates a string cache element with the specified children.
    /// </summary>
    /// <param name="children">
    /// The child elements to include in the cache.
    /// </param>
    /// <returns>A StringCache element.</returns>
    protected override OpenXmlCompositeElement CreateCache(
        params OpenXmlElement[] children
    ) => new C.StringCache(children);

    /// <summary>
    /// Creates a string point element with the specified value and index.
    /// </summary>
    /// <param name="child">The numeric value to include in the point.</param>
    /// <param name="index">The index of the point in the series.</param>
    /// <returns>A StringPoint element with the specified index.</returns>
    protected override OpenXmlCompositeElement CreatePoint(
        C.NumericValue child,
        uint index
    ) => new C.StringPoint(child) { Index = index };
}

/// <summary>Creates numeric references in OpenXml charts.</summary>
public class NumericReferenceFactory : ValueReferenceFactory
{
    /// <summary>
    /// Creates a number reference element with the specified children.
    /// </summary>
    /// <param name="children">
    /// The child elements to include in the reference.
    /// </param>
    /// <returns>A NumberReference element.</returns>
    protected override OpenXmlCompositeElement CreateReference(
        params OpenXmlElement[] children
    ) => new C.NumberReference(children);

    /// <summary>
    /// Creates a numbering cache element with the specified children.
    /// </summary>
    /// <param name="children">
    /// The child elements to include in the cache.
    /// </param>
    /// <returns>A NumberingCache element.</returns>
    protected override OpenXmlCompositeElement CreateCache(
        params OpenXmlElement[] children
    ) => new C.NumberingCache(children);

    /// <summary>
    /// Creates a numeric point element with the specified value and index.
    /// </summary>
    /// <param name="child">The numeric value to include in the point.</param>
    /// <param name="index">The index of the point in the series.</param>
    /// <returns>A NumericPoint element with the specified index.</returns>
    protected override OpenXmlCompositeElement CreatePoint(
        C.NumericValue child,
        uint index
    ) => new C.NumericPoint(child) { Index = index };
}
