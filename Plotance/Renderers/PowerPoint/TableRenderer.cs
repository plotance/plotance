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

using static Plotance.Renderers.PowerPoint.OpenXmlElementUtilities;

using D = DocumentFormat.OpenXml.Drawing;

namespace Plotance.Renderers.PowerPoint;

/// <summary>Provides methods for rendering query result to a tables.</summary>
public static class TableRenderer
{
    /// <summary>Render a query result to a slide as a table.</summary>
    /// <param name="slidePart">The slide part to render to.</param>
    /// <param name="paragraphStyles">The paragraph styles.</param>
    /// <param name="isFirst">Whether this is the first block.</param>
    /// <param name="x">The X coordinate of the table.</param>
    /// <param name="y">The Y coordinate of the table.</param>
    /// <param name="width">The width of the table.</param>
    /// <param name="height">The height of the table.</param>
    /// <param name="block">The block to render.</param>
    /// <param name="queryResults">The query results.</param>
    /// <exception cref="PlotanceException">
    /// Thrown when the table configuration is invalid.
    /// </exception>
    public static void Render(
        SlidePart slidePart,
        ParagraphStyles paragraphStyles,
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

        var placeholderShape = new PlaceholderShape()
        {
            Type = PlaceholderValues.Table
        };

        if (isFirst)
        {
            placeholderShape.Index = 1;
        }

        var tableUri = "http://schemas.openxmlformats.org/drawingml/2006/table";

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
                            CreateTable(
                                block.Path,
                                block.Line,
                                block.ChartOptions,
                                paragraphStyles,
                                width,
                                height,
                                queryResult
                            )
                        )
                        {
                            Uri = tableUri
                        }
                    )
                )
            );
    }

    /// <summary>Creates a table from a query result.</summary>
    /// <param name="path">The path to the file containing the table.</param>
    /// <param name="line">
    /// The line number in the file containing the table.
    /// </param>
    /// <param name="options">The chart options.</param>
    /// <param name="paragraphStyles">The paragraph styles.</param>
    /// <param name="width">The width of the table.</param>
    /// <param name="height">The height of the table.</param>
    /// <param name="queryResult">The query result.</param>
    /// <returns>The table element.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the table configuration is invalid.
    /// </exception>
    private static D.Table CreateTable(
        string? path,
        long line,
        ChartOptions? options,
        ParagraphStyles paragraphStyles,
        long width,
        long height,
        QueryResultSet queryResult
    )
    {
        var cellWidth = width / queryResult.Columns.Count;
        var tableStyle = TableStyle.Create(
            options,
            queryResult.Rows.Count,
            queryResult.Columns.Count
        );
        var columnWidths = tableStyle.ComputeColumnWidths(
            path,
            line,
            width,
            queryResult.Columns.Count
        );
        var rowHeights = tableStyle.ComputeRowHeights(
            path,
            line,
            height,
            // +1 for header row.
            queryResult.Rows.Count + 1
        );

        return new D.Table([
            tableStyle.ToTableProperties(),
            new D.TableGrid(
                columnWidths.Select(
                    width => new D.GridColumn() { Width = width }
                )
            ),
            CreateHeaderRow(
                tableStyle,
                paragraphStyles,
                rowHeights[0],
                queryResult
            ),
            .. CreateRows(
                tableStyle,
                paragraphStyles,
                rowHeights.Skip(1).ToList(),
                queryResult
            )
        ]);
    }

    /// <summary>Creates the header row for a table.</summary>
    /// <param name="tableStyle">The table style.</param>
    /// <param name="paragraphStyles">The paragraph styles.</param>
    /// <param name="height">The height of the row.</param>
    /// <param name="queryResult">The query result.</param>
    /// <returns>The header row element.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the table style is invalid.
    /// </exception>
    private static D.TableRow CreateHeaderRow(
        TableStyle tableStyle,
        ParagraphStyles paragraphStyles,
        long height,
        QueryResultSet queryResult
    ) => new D.TableRow(
        queryResult.Columns.Select(
            (column, columnIndex) => CreateCell(
                tableStyle,
                paragraphStyles,
                queryResult.Rows.Count,
                queryResult.Columns.Count,
                0,
                columnIndex,
                column.Name
            )
        )
    )
    {
        Height = height
    };

    /// <summary>Creates the rows for a table.</summary>
    /// <param name="tableStyle">The table style.</param>
    /// <param name="paragraphStyles">The paragraph styles.</param>
    /// <param name="heights">The heights of the rows.</param>
    /// <param name="queryResult">The query result.</param>
    /// <returns>The row elements.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the table style is invalid.
    /// </exception>
    private static IEnumerable<D.TableRow> CreateRows(
        TableStyle tableStyle,
        ParagraphStyles paragraphStyles,
        IReadOnlyList<long> heights,
        QueryResultSet queryResult
    ) => queryResult.Rows.Zip(heights).Select(
        (tuple, rowIndex) => new D.TableRow(
            tuple.First.Select(
                (value, columnIndex) => CreateCell(
                    tableStyle,
                    paragraphStyles,
                    queryResult.Rows.Count + 1,  // +1 for header row
                    tuple.First.Count,
                    rowIndex + 1, // +1 for header row
                    columnIndex,
                    value
                )
            )
        )
        {
            Height = tuple.Second
        }
    );

    /// <summary>Creates a table cell.</summary>
    /// <param name="tableStyle">The table style.</param>
    /// <param name="paragraphStyles">The paragraph styles.</param>
    /// <param name="rowCount">The number of rows in the table.</param>
    /// <param name="columnCount">The number of columns in the table.</param>
    /// <param name="rowIndex">The index of the row.</param>
    /// <param name="columnIndex">The index of the column.</param>
    /// <param name="value">The value of the cell.</param>
    /// <returns>The cell element.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the table style is invalid.
    /// </exception>
    private static D.TableCell CreateCell(
        TableStyle tableStyle,
        ParagraphStyles paragraphStyles,
        int rowCount,
        int columnCount,
        int rowIndex,
        int columnIndex,
        object value
    )
    {
        var cellPosition = new CellPosition(
            rowCount,
            columnCount,
            rowIndex,
            columnIndex
        );
        var horizontalAlign = tableStyle.ComputeCellHorizontalAlign(
            cellPosition
        );

        if (horizontalAlign != null)
        {
            paragraphStyles = paragraphStyles.WithAlignment(
                horizontalAlign.Value
            );
        }

        var margins = tableStyle.ComputeCellMargins(cellPosition);
        var defaultCellMargin = (
            paragraphStyles
                .Level1
                .GetFirstChild<D.DefaultRunProperties>()
                ?.FontSize
                ?? 1800
        ) * ILength.EmuPerPt / 400;

        int GetMargin(CellMarginPosition position)
            => margins.TryGetValue(position, out var margin)
            ? (int)margin.ToEmu()
            : (int)defaultCellMargin;

        return new D.TableCell(
            new D.TextBody([
                new D.BodyProperties(new D.NoAutoFit()),
                paragraphStyles.ToListStyle(),
                .. (value.ToString() ?? "").Split("\n").Select(
                    line => new D.Paragraph(
                        new D.ParagraphProperties(new D.NoBullet())
                        {
                            Indent = 0,
                            LeftMargin = 0,
                            RightMargin = 0
                        },
                        new D.Run(new D.Text(line))
                    )
                )
            ]),
            new D.TableCellProperties()
            {
                LeftMargin = GetMargin(CellMarginPosition.Left),
                RightMargin = GetMargin(CellMarginPosition.Right),
                TopMargin = GetMargin(CellMarginPosition.Top),
                BottomMargin = GetMargin(CellMarginPosition.Bottom),
                Anchor = tableStyle.ComputeCellVerticalAlign(cellPosition)
            }
        );
    }
}

/// <summary>Represents the position of a table border.</summary>
public enum TableBorderPosition
{
    Left,
    Right,
    Top,
    Bottom,
    InsideHorizontal,
    InsideVertical
}

/// <summary>
/// Represents a configuration of a table border. This record is used for
/// resolving inheritance of table border configurations and generation of Open
/// XML ThemeableLineStyleType elements.
/// </summary>
/// <param name="Width">The width of the border.</param>
/// <param name="Color">The color of the border.</param>
/// <param name="Style">
/// The style of the border. Can be "none", "single", "double", "thick_thin",
/// "thin_thick", or "triple".
/// </param>
public record TableBorder(
    ILength? Width,
    Color? Color,
    ValueWithLocation<string>? Style
)
{
    /// <summary>Creates a table border configuration.</summary>
    /// <param name="options">The options for the table border.</param>
    /// <returns>The table border configuration.</returns>
    public static TableBorder Create(TableBorderOptions? options)
        => new TableBorder(options?.Width, options?.Color, options?.Style);

    /// <summary>Whether the table border configuration is null.</summary>
    public bool IsNull => Width == null && Color == null && Style == null;

    /// <summary>
    /// Combines this table border configuration with another.
    /// </summary>
    /// <param name="other">The other table border configuration.</param>
    /// <returns>The combined table border configuration.</returns>
    public TableBorder OrElse(TableBorder other) => new TableBorder(
        Width ?? other.Width,
        Color ?? other.Color,
        Style ?? other.Style
    );

    /// <summary>
    /// Converts this table border configuration to an Open XML table border.
    /// </summary>
    /// <typeparam name="T">The type of the Open XML table border.</typeparam>
    /// <returns>The Open XML table border.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the border style is not recognized.
    /// </exception>
    public T? ToOpenXml<T>() where T : D.ThemeableLineStyleType, new()
    {
        var width = Width ?? ILength.FromPoint(1);

        if (width.ToEmu() == 0 || Style?.Value == "none")
        {
            return null;
        }

        D.CompoundLineValues? compoundLineType = Style?.Value switch
        {
            "single" => D.CompoundLineValues.Single,
            "double" => D.CompoundLineValues.Double,
            "thick_thin" => D.CompoundLineValues.ThickThin,
            "thin_thick" => D.CompoundLineValues.ThinThick,
            "triple" => D.CompoundLineValues.Triple,
            null => null,
            _ => throw new PlotanceException(
                Style?.Path,
                Style?.Line,
                $"Unknown border style: {Style?.Value}"
            )
        };

        return CreateIfAnyChild<T>(
            CreateIfAnyChild<D.Outline>(
                CreateIfAnyChild<D.SolidFill>(
                    ColorRenderer.ParseColor(Color)
                )
            )
                .With((x, y) => x.Width = (int)y!, width.ToEmu())
                .With((x, y) => x.CompoundLineType = y!, compoundLineType)
        );
    }
}

/// <summary>Represents the position of a cell margin.</summary>
public enum CellMarginPosition
{
    Left,
    Right,
    Top,
    Bottom
}

/// <summary>
/// Represents a style of a table cells for a table region in TableStyle.
/// </summary>
/// <param name="BackgroundColor">The background color of the cells.</param>
/// <param name="Borders">The border styles of the cells.</param>
/// <param name="CellMargins">The margins of the cells.</param>
/// <param name="FontWeight">The font weight of the cells.</param>
/// <param name="FontColor">The font color of the cells.</param>
/// <param name="CellHorizontalAlign">
/// The horizontal alignment of texts in the cells.
/// </param>
/// <param name="CellVerticalAlign">
/// The vertical alignment of texts in the cells.
/// </param>
public record TableStyleElement(
    Color? BackgroundColor,
    IReadOnlyDictionary<TableBorderPosition, TableBorder> Borders,
    IReadOnlyDictionary<CellMarginPosition, ILength> CellMargins,
    ValueWithLocation<string>? FontWeight,
    Color? FontColor,
    ValueWithLocation<string>? CellHorizontalAlign,
    ValueWithLocation<string>? CellVerticalAlign
)
{
    /// <summary>
    /// The mapping from Open XML border element types to table border
    /// positions.
    /// </summary>
    private IReadOnlyDictionary<Type, TableBorderPosition> _typeToBorderPosition
        = new Dictionary<Type, TableBorderPosition>
        {
            {
                typeof(D.LeftBorder),
                TableBorderPosition.Left
            },
            {
                typeof(D.RightBorder),
                TableBorderPosition.Right
            },
            {
                typeof(D.TopBorder),
                TableBorderPosition.Top
            },
            {
                typeof(D.BottomBorder),
                TableBorderPosition.Bottom
            },
            {
                typeof(D.InsideHorizontalBorder),
                TableBorderPosition.InsideHorizontal
            },
            {
                typeof(D.InsideVerticalBorder),
                TableBorderPosition.InsideVertical
            }
        };

    /// <summary>Creates a table style from the specified options.</summary>
    /// <param name="options">The options for the table style.</param>
    /// <returns>The table style.</returns>
    public static TableStyleElement Create(TableOptions? options)
    {
        var cellMargins = new Dictionary<CellMarginPosition, ILength>();

        void AddMargin(CellMarginPosition position, ILength? margin)
        {
            if (margin != null)
            {
                cellMargins[position] = margin;
            }
        }

        AddMargin(CellMarginPosition.Left, options?.CellLeftMargin);
        AddMargin(CellMarginPosition.Right, options?.CellRightMargin);
        AddMargin(CellMarginPosition.Top, options?.CellTopMargin);
        AddMargin(CellMarginPosition.Bottom, options?.CellBottomMargin);

        return new TableStyleElement(
            options?.BackgroundColor,
            new Dictionary<TableBorderPosition, TableBorder>
            {
                {
                    TableBorderPosition.Left,
                    TableBorder.Create(options?.LeftBorderOptions)
                },
                {
                    TableBorderPosition.Right,
                    TableBorder.Create(options?.RightBorderOptions)
                },
                {
                    TableBorderPosition.Top,
                    TableBorder.Create(options?.TopBorderOptions)
                },
                {
                    TableBorderPosition.Bottom,
                    TableBorder.Create(options?.BottomBorderOptions)
                },
                {
                    TableBorderPosition.InsideHorizontal,
                    TableBorder.Create(options?.InsideHorizontalBorderOptions)
                },
                {
                    TableBorderPosition.InsideVertical,
                    TableBorder.Create(options?.InsideVerticalBorderOptions)
                }
            },
            cellMargins,
            options?.FontWeight,
            options?.FontColor,
            options?.CellHorizontalAlign,
            options?.CellVerticalAlign
        );
    }

    /// <summary>
    /// Creates a table style by inheriting properties from another table style.
    /// </summary>
    /// <param name="other">The other table style.</param>
    /// <param name="leftBorderPosition">
    /// The position of the left border to inherit.
    /// </param>
    /// <param name="rightBorderPosition">
    /// The position of the right border to inherit.
    /// </param>
    /// <param name="topBorderPosition">
    /// The position of the top border to inherit.
    /// </param>
    /// <param name="bottomBorderPosition">
    /// The position of the bottom border to inherit.
    /// </param>
    /// <param name="insideHorizontalBorderPosition">
    /// The position of the inside horizontal border to inherit.
    /// </param>
    /// <param name="insideVerticalBorderPosition">
    /// The position of the inside vertical border to inherit.
    /// </param>
    /// <returns>The combined table style.</returns>
    public TableStyleElement Inherit(
        TableStyleElement other,
        TableBorderPosition? leftBorderPosition,
        TableBorderPosition? rightBorderPosition,
        TableBorderPosition? topBorderPosition,
        TableBorderPosition? bottomBorderPosition,
        TableBorderPosition? insideHorizontalBorderPosition,
        TableBorderPosition? insideVerticalBorderPosition
    )
    {
        var mergedBorders = new Dictionary<TableBorderPosition, TableBorder>();

        void AddOtherBorder(
            TableBorderPosition position,
            TableBorderPosition? otherPosition
        )
        {
            // If border is not defined at all, do not inherit.
            if (!Borders.ContainsKey(position) || Borders[position].IsNull)
            {
                return;
            }

            if (
                otherPosition != null
                    && other.Borders.ContainsKey(otherPosition.Value)
            )
            {
                mergedBorders[position] = Borders[position]
                    .OrElse(other.Borders[otherPosition.Value]);
            }
            else
            {
                mergedBorders[position] = Borders[position];
            }
        }

        AddOtherBorder(
            TableBorderPosition.Left,
            leftBorderPosition
        );
        AddOtherBorder(
            TableBorderPosition.Right,
            rightBorderPosition
        );
        AddOtherBorder(
            TableBorderPosition.Top,
            topBorderPosition
        );
        AddOtherBorder(
            TableBorderPosition.Bottom,
            bottomBorderPosition
        );
        AddOtherBorder(
            TableBorderPosition.InsideHorizontal,
            insideHorizontalBorderPosition
        );
        AddOtherBorder(
            TableBorderPosition.InsideVertical,
            insideVerticalBorderPosition
        );

        var mergedCellMargins = new Dictionary<CellMarginPosition, ILength>(
            other.CellMargins
        );

        foreach (var pair in CellMargins)
        {
            mergedCellMargins[pair.Key] = pair.Value;
        }

        return new TableStyleElement(
            BackgroundColor ?? other.BackgroundColor,
            mergedBorders,
            mergedCellMargins,
            FontWeight ?? other.FontWeight,
            FontColor ?? other.FontColor,
            CellHorizontalAlign ?? other.CellHorizontalAlign,
            CellVerticalAlign ?? other.CellVerticalAlign
        );
    }

    /// <summary>Converts this table style to an Open XML table style.</summary>
    /// <typeparam name="T">The type of the Open XML table style.</typeparam>
    /// <returns>The Open XML table style.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the font weight is not recognized.
    /// </exception>
    public T? ToOpenXml<T>() where T : D.TablePartStyleType, new()
    {
        D.BooleanStyleValues? bold = FontWeight?.Value switch
        {
            null => null,
            "normal" => null,
            "bold" => D.BooleanStyleValues.On,
            _ => throw new PlotanceException(
                FontWeight?.Path,
                FontWeight?.Line,
                $"Unknown font weight: {FontWeight?.Value}"
            )
        };

        // TODO use monospace font?
        var textStyle = CreateIfAnyChild<D.TableCellTextStyle>(
            ColorRenderer.ParseColor(FontColor)
        ).With((e, v) => e.Bold = v, bold);

        var cellStyle = CreateIfAnyChild<D.TableCellStyle>(
            CreateIfAnyChild<D.TableCellBorders>(
                CreateBorder<D.LeftBorder>(),
                CreateBorder<D.RightBorder>(),
                CreateBorder<D.TopBorder>(),
                CreateBorder<D.BottomBorder>(),
                CreateBorder<D.InsideHorizontalBorder>(),
                CreateBorder<D.InsideVerticalBorder>()
            ),
            CreateIfAnyChild<D.FillProperties>(
                CreateIfAnyChild<D.SolidFill>(
                    ColorRenderer.ParseColor(BackgroundColor)
                )
            )
        );

        return CreateIfAnyChild<T>(textStyle, cellStyle);
    }

    /// <summary>
    /// Create an Open XML border element from this table style.
    /// </summary>
    /// <typeparam name="T">The type of the Open XML border element.</typeparam>
    /// <returns>The Open XML border element.</returns>
    private T? CreateBorder<T>() where T : D.ThemeableLineStyleType, new()
    {
        var key = _typeToBorderPosition[typeof(T)];

        return Borders.ContainsKey(key)
            ? Borders[key].ToOpenXml<T>()
            : null;
    }
}

/// <summary>The region where table style is applied.</summary>
public enum TableStyleElementType
{
    WholeTable,
    FirstRow,
    FirstColumn,
    LastRow,
    LastColumn,
    Band1Row,
    Band1Column,
    Band2Row,
    Band2Column,
    SoutheastCell,
    SouthwestCell,
    NortheastCell,
    NorthwestCell
}

/// <summary>Represents a table style.</summary>
public record TableStyle(
    ValueWithLocation<IReadOnlyList<ILengthWeight>>? Rows,
    ValueWithLocation<IReadOnlyList<ILengthWeight>>? Columns,
    IReadOnlyDictionary<TableStyleElementType, TableStyleElement>
        TableStyleElements
)
{
    /// <summary>
    /// The mapping from Open XML table style element types to table style keys.
    /// </summary>
    private IReadOnlyDictionary<Type, TableStyleElementType> _typeToKey
        = new Dictionary<Type, TableStyleElementType>
    {
        { typeof(D.WholeTable), TableStyleElementType.WholeTable },
        { typeof(D.FirstRow), TableStyleElementType.FirstRow },
        { typeof(D.FirstColumn), TableStyleElementType.FirstColumn },
        { typeof(D.LastRow), TableStyleElementType.LastRow },
        { typeof(D.LastColumn), TableStyleElementType.LastColumn },
        { typeof(D.Band1Horizontal), TableStyleElementType.Band1Row },
        { typeof(D.Band1Vertical), TableStyleElementType.Band1Column },
        { typeof(D.Band2Horizontal), TableStyleElementType.Band2Row },
        { typeof(D.Band2Vertical), TableStyleElementType.Band2Column },
        { typeof(D.SoutheastCell), TableStyleElementType.SoutheastCell },
        { typeof(D.SouthwestCell), TableStyleElementType.SouthwestCell },
        { typeof(D.NortheastCell), TableStyleElementType.NortheastCell },
        { typeof(D.NorthwestCell), TableStyleElementType.NorthwestCell }
    };

    /// <summary>Creates a table style from the specified options.</summary>
    /// <param name="options">
    /// The options to create the table style from.
    /// </param>
    /// <param name="rowCount">The number of rows in the table.</param>
    /// <param name="columnCount">The number of columns in the table.</param>
    /// <returns>The table style.</returns>
    public static TableStyle Create(
        ChartOptions? options,
        int rowCount,
        int columnCount
    )
    {
        var wholeTable = TableStyleElement.Create(options?.WholeTableOptions);

        var band1RowRaw = TableStyleElement.Create(options?.Band1RowOptions);
        var band1Row = band1RowRaw.Inherit(
            wholeTable,
            TableBorderPosition.Left,
            TableBorderPosition.Right,
            TableBorderPosition.InsideHorizontal,
            TableBorderPosition.InsideHorizontal,
            null,
            TableBorderPosition.InsideVertical
        );

        var band1ColumnRaw = TableStyleElement.Create(
            options?.Band1ColumnOptions
        );
        var band1Column = band1ColumnRaw.Inherit(
            wholeTable,
            TableBorderPosition.InsideVertical,
            TableBorderPosition.InsideVertical,
            TableBorderPosition.Top,
            TableBorderPosition.Bottom,
            TableBorderPosition.InsideHorizontal,
            null
        );

        var band2RowRaw = TableStyleElement.Create(options?.Band2RowOptions);
        var band2Row = band2RowRaw.Inherit(
            wholeTable,
            TableBorderPosition.Left,
            TableBorderPosition.Right,
            TableBorderPosition.InsideHorizontal,
            TableBorderPosition.InsideHorizontal,
            null,
            TableBorderPosition.InsideVertical
        );

        var band2ColumnRaw = TableStyleElement.Create(
            options?.Band2ColumnOptions
        );
        var band2Column = band2ColumnRaw.Inherit(
            wholeTable,
            TableBorderPosition.InsideVertical,
            TableBorderPosition.InsideVertical,
            TableBorderPosition.Top,
            TableBorderPosition.Bottom,
            TableBorderPosition.InsideHorizontal,
            null
        );

        var firstRowRaw = TableStyleElement.Create(options?.FirstRowOptions);

        firstRowRaw = firstRowRaw with
        {
            BackgroundColor = firstRowRaw.BackgroundColor
                ?? new Color(null, 0, "accent1"),
            FontWeight = firstRowRaw.FontWeight
                ?? new ValueWithLocation<string>(null, 0, "bold"),
            FontColor = firstRowRaw.FontColor
                ?? new Color(null, 0, "light1")
        };

        var firstRow = firstRowRaw.Inherit(
            band1RowRaw,
            TableBorderPosition.Left,
            TableBorderPosition.Right,
            TableBorderPosition.Top,
            TableBorderPosition.Bottom,
            null,
            TableBorderPosition.InsideVertical
        ).Inherit(
            wholeTable,
            TableBorderPosition.Left,
            TableBorderPosition.Right,
            TableBorderPosition.Top,
            TableBorderPosition.InsideHorizontal,
            null,
            TableBorderPosition.InsideVertical
        );

        var firstColumnRaw = TableStyleElement.Create(
            options?.FirstColumnOptions
        );
        var firstColumn = firstColumnRaw.Inherit(
            band1ColumnRaw,
            TableBorderPosition.Left,
            TableBorderPosition.Right,
            TableBorderPosition.Top,
            TableBorderPosition.Bottom,
            TableBorderPosition.InsideHorizontal,
            null
        ).Inherit(
            wholeTable,
            TableBorderPosition.Left,
            TableBorderPosition.InsideVertical,
            TableBorderPosition.Top,
            TableBorderPosition.Bottom,
            TableBorderPosition.InsideHorizontal,
            null
        );

        var lastRowRaw = TableStyleElement.Create(options?.LastRowOptions);
        var lastRow = lastRowRaw.Inherit(
            rowCount % 2 == 0 ? band2RowRaw : band1RowRaw,
            TableBorderPosition.Left,
            TableBorderPosition.Right,
            TableBorderPosition.Top,
            TableBorderPosition.Bottom,
            null,
            TableBorderPosition.InsideVertical
        ).Inherit(
            wholeTable,
            TableBorderPosition.Left,
            TableBorderPosition.Right,
            TableBorderPosition.InsideHorizontal,
            TableBorderPosition.Bottom,
            null,
            TableBorderPosition.InsideVertical
        );

        var lastColumnRaw = TableStyleElement.Create(
            options?.LastColumnOptions
        );
        var lastColumn = lastColumnRaw.Inherit(
            columnCount % 2 == 0 ? band2ColumnRaw : band1ColumnRaw,
            TableBorderPosition.Left,
            TableBorderPosition.Right,
            TableBorderPosition.Top,
            TableBorderPosition.Bottom,
            TableBorderPosition.InsideHorizontal,
            null
        ).Inherit(
            wholeTable,
            TableBorderPosition.InsideVertical,
            TableBorderPosition.Right,
            TableBorderPosition.Top,
            TableBorderPosition.Bottom,
            TableBorderPosition.InsideHorizontal,
            null
        );

        var southeastCell = TableStyleElement.Create(
            options?.SoutheastCellOptions
        ).Inherit(
            lastRowRaw,
            TableBorderPosition.InsideVertical,
            TableBorderPosition.Right,
            TableBorderPosition.Top,
            TableBorderPosition.Bottom,
            null,
            null
        ).Inherit(
            lastColumnRaw,
            TableBorderPosition.Left,
            TableBorderPosition.Right,
            TableBorderPosition.InsideHorizontal,
            TableBorderPosition.Bottom,
            null,
            null
        ).Inherit(
            columnCount % 2 == 0 ? band2ColumnRaw : band1ColumnRaw,
            TableBorderPosition.Left,
            TableBorderPosition.Right,
            TableBorderPosition.InsideHorizontal,
            TableBorderPosition.Bottom,
            null,
            null
        ).Inherit(
            rowCount % 2 == 0 ? band2RowRaw : band1RowRaw,
            TableBorderPosition.InsideVertical,
            TableBorderPosition.Right,
            TableBorderPosition.Top,
            TableBorderPosition.Bottom,
            null,
            null
        ).Inherit(
            wholeTable,
            TableBorderPosition.InsideVertical,
            TableBorderPosition.Right,
            TableBorderPosition.InsideHorizontal,
            TableBorderPosition.Bottom,
            null,
            null
        );

        var southwestCell = TableStyleElement.Create(
            options?.SouthwestCellOptions
        ).Inherit(
            lastRow,
            TableBorderPosition.Left,
            TableBorderPosition.InsideVertical,
            TableBorderPosition.Top,
            TableBorderPosition.Bottom,
            null,
            null
        ).Inherit(
            firstColumn,
            TableBorderPosition.Left,
            TableBorderPosition.Right,
            TableBorderPosition.InsideHorizontal,
            TableBorderPosition.Bottom,
            null,
            null
        ).Inherit(
            band1ColumnRaw,
            TableBorderPosition.Left,
            TableBorderPosition.Right,
            TableBorderPosition.InsideHorizontal,
            TableBorderPosition.Bottom,
            null,
            null
        ).Inherit(
            rowCount % 2 == 0 ? band2RowRaw : band1RowRaw,
            TableBorderPosition.Left,
            TableBorderPosition.InsideVertical,
            TableBorderPosition.Top,
            TableBorderPosition.Bottom,
            null,
            null
        ).Inherit(
            wholeTable,
            TableBorderPosition.Left,
            TableBorderPosition.InsideVertical,
            TableBorderPosition.InsideHorizontal,
            TableBorderPosition.Bottom,
            null,
            null
        );

        var northeastCell = TableStyleElement.Create(
            options?.NortheastCellOptions
        ).Inherit(
            firstRow,
            TableBorderPosition.InsideVertical,
            TableBorderPosition.Right,
            TableBorderPosition.Top,
            TableBorderPosition.Bottom,
            null,
            null
        ).Inherit(
            lastColumn,
            TableBorderPosition.Left,
            TableBorderPosition.Right,
            TableBorderPosition.Top,
            TableBorderPosition.InsideHorizontal,
            null,
            null
        ).Inherit(
            columnCount % 2 == 0 ? band2ColumnRaw : band1ColumnRaw,
            TableBorderPosition.Left,
            TableBorderPosition.Right,
            TableBorderPosition.Top,
            TableBorderPosition.InsideHorizontal,
            null,
            null
        ).Inherit(
            band1RowRaw,
            TableBorderPosition.InsideVertical,
            TableBorderPosition.Right,
            TableBorderPosition.Top,
            TableBorderPosition.Bottom,
            null,
            null
        ).Inherit(
            wholeTable,
            TableBorderPosition.InsideVertical,
            TableBorderPosition.Right,
            TableBorderPosition.Top,
            TableBorderPosition.InsideHorizontal,
            null,
            null
        );

        var northwestCell = TableStyleElement.Create(
            options?.NorthwestCellOptions
        ).Inherit(
            firstRow,
            TableBorderPosition.Left,
            TableBorderPosition.InsideVertical,
            TableBorderPosition.Top,
            TableBorderPosition.Bottom,
            null,
            null
        ).Inherit(
            firstColumn,
            TableBorderPosition.Left,
            TableBorderPosition.Right,
            TableBorderPosition.Top,
            TableBorderPosition.InsideHorizontal,
            null,
            null
        ).Inherit(
            band1ColumnRaw,
            TableBorderPosition.Left,
            TableBorderPosition.Right,
            TableBorderPosition.Top,
            TableBorderPosition.InsideHorizontal,
            null,
            null
        ).Inherit(
            band1RowRaw,
            TableBorderPosition.Right,
            TableBorderPosition.InsideVertical,
            TableBorderPosition.Top,
            TableBorderPosition.Bottom,
            null,
            null
        ).Inherit(
            wholeTable,
            TableBorderPosition.Left,
            TableBorderPosition.InsideVertical,
            TableBorderPosition.Top,
            TableBorderPosition.InsideHorizontal,
            null,
            null
        );

        return new TableStyle(
            options?.TableRows,
            options?.TableColumns,
            new Dictionary<TableStyleElementType, TableStyleElement>
            {
                { TableStyleElementType.WholeTable, wholeTable },
                { TableStyleElementType.FirstRow, firstRow },
                { TableStyleElementType.FirstColumn, firstColumn },
                { TableStyleElementType.LastRow, lastRow },
                { TableStyleElementType.LastColumn, lastColumn },
                { TableStyleElementType.Band1Row, band1Row },
                { TableStyleElementType.Band1Column, band1Column },
                { TableStyleElementType.Band2Row, band2Row },
                { TableStyleElementType.Band2Column, band2Column },
                { TableStyleElementType.SoutheastCell, southeastCell },
                { TableStyleElementType.SouthwestCell, southwestCell },
                { TableStyleElementType.NortheastCell, northeastCell },
                { TableStyleElementType.NorthwestCell, northwestCell }
            }
        );
    }

    /// <summary>
    /// Gets the table style element for the specified key.
    /// </summary>
    /// <param name="key">The key to get the table style element for.</param>
    /// <returns>The table style element.</returns>
    public TableStyleElement this[TableStyleElementType key]
        => TableStyleElements[key];

    /// <summary>
    /// Converts this table style to an Open XML table properties element.
    /// </summary>
    /// <returns>The Open XML table properties element.</returns>
    public D.TableProperties ToTableProperties()
    {
        var wholeTable = CreatePart<D.WholeTable>();
        var firstRow = CreatePart<D.FirstRow>();
        var firstColumn = CreatePart<D.FirstColumn>();
        var lastRow = CreatePart<D.LastRow>();
        var lastColumn = CreatePart<D.LastColumn>();
        var band1Row = CreatePart<D.Band1Horizontal>();
        var band1Column = CreatePart<D.Band1Vertical>();
        var band2Row = CreatePart<D.Band2Horizontal>();
        var band2Column = CreatePart<D.Band2Vertical>();
        var southeastCell = CreatePart<D.SoutheastCell>();
        var southwestCell = CreatePart<D.SouthwestCell>();
        var northeastCell = CreatePart<D.NortheastCell>();
        var northwestCell = CreatePart<D.NorthwestCell>();
        var tableStyle = CreateIfAnyChild<D.TableStyle>(
            wholeTable,
            firstRow,
            firstColumn,
            lastRow,
            lastColumn,
            band1Row,
            band1Column,
            band2Row,
            band2Column,
            southeastCell,
            southwestCell,
            northeastCell,
            northwestCell
        );

        if (tableStyle != null)
        {
            var guid = Guid.NewGuid().ToString().ToUpperInvariant();

            tableStyle.StyleId = "{" + guid + "}";
            tableStyle.StyleName = "Table Style";
        }

        var tableProperties = new D.TableProperties()
        {
            FirstRow = firstRow != null,
            FirstColumn = firstColumn != null,
            LastRow = lastRow != null,
            LastColumn = lastColumn != null,
            BandRow = band1Row != null || band2Row != null,
            BandColumn = band1Column != null || band2Column != null
        };

        if (tableStyle != null)
        {
            tableProperties.AddChild(tableStyle);
        }

        return tableProperties;
    }

    /// <summary>
    /// Creates a table part style element for the specified type.
    /// </summary>
    /// <typeparam name="T">
    /// The type of table part style element to create.
    /// </typeparam>
    /// <returns>The table part style element.</returns>
    private T? CreatePart<T>() where T : D.TablePartStyleType, new()
    {
        var key = _typeToKey[typeof(T)];

        return TableStyleElements.ContainsKey(key)
            ? TableStyleElements[key].ToOpenXml<T>()
            : null;
    }

    /// <summary>
    /// Computes the cell margins for the specified cell position,
    /// taking inheritance into account.
    /// </summary>
    /// <param name="cellPosition">
    /// The cell position to compute the cell margins for.
    /// </param>
    /// <returns>The cell margins.</returns>
    public IReadOnlyDictionary<CellMarginPosition, ILength> ComputeCellMargins(
        CellPosition cellPosition
    )
    {
        var cellMargins = new Dictionary<CellMarginPosition, ILength>();

        foreach (var tableStyleElementType in cellPosition.InheritanceList)
        {
            var tableStyleElement = this[tableStyleElementType];

            foreach (var pair in tableStyleElement.CellMargins)
            {
                cellMargins[pair.Key] = pair.Value;
            }
        }

        return cellMargins;
    }

    /// <summary>
    /// Computes the cell vertical alignment for the specified cell position,
    /// taking inheritance into account.
    /// </summary>
    /// <param name="cellPosition">
    /// The cell position to compute the cell vertical alignment for.
    /// </param>
    /// <returns>The cell vertical alignment.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the vertical alignment is not recognized.
    /// </exception>
    public D.TextAnchoringTypeValues? ComputeCellVerticalAlign(
        CellPosition cellPosition
    )
    {
        ValueWithLocation<string>? align = null;

        foreach (var tableStyleElementType in cellPosition.InheritanceList)
        {
            var tableStyleElement = this[tableStyleElementType];

            align = tableStyleElement.CellVerticalAlign ?? align;
        }

        return Geometries.ParseVerticalAlign(align);
    }

    /// <summary>
    /// Computes the cell horizontal alignment for the specified cell position,
    /// taking inheritance into account.
    /// </summary>
    /// <param name="cellPosition">
    /// The cell position to compute the cell horizontal alignment for.
    /// </param>
    /// <returns>The cell horizontal alignment.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the horizontal alignment is not recognized.
    /// </exception>
    public D.TextAlignmentTypeValues? ComputeCellHorizontalAlign(
        CellPosition cellPosition
    )
    {
        ValueWithLocation<string>? align = null;

        foreach (var tableStyleElementType in cellPosition.InheritanceList)
        {
            var tableStyleElement = this[tableStyleElementType];

            align = tableStyleElement.CellHorizontalAlign ?? align;
        }

        return Geometries.ParseHorizontalAlign(align);
    }

    /// <summary>
    /// Computes the column widths for the specified table.
    /// </summary>
    /// <param name="path">The path to the file containing the table.</param>
    /// <param name="line">
    /// The line number in the file containing the table.
    /// </param>
    /// <param name="totalWidth">The total width of the table in EMU.</param>
    /// <param name="columnCount">The number of columns in the table.</param>
    /// <returns>The column widths in EMU.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the table columns does not fit into the space.
    /// </exception>
    public IReadOnlyList<long> ComputeColumnWidths(
        string? path,
        long line,
        long totalWidth,
        int columnCount
    )
    {
        var columns = Columns?.Value ?? [];

        try
        {
            return ILengthWeight
                .Divide(
                    totalWidth,
                    Enumerable
                        .Range(0, columnCount)
                        .Select(
                            index => (
                                Weight: index < columns.Count
                                    ? columns[index]
                                    : new RelativeLengthWeight(null, 0, 1),
                                Gap: 0L
                            )
                        )
                        .ToList()
                )
                .Select(tuple => tuple.Length)
                .ToList();
        }
        catch (ArgumentException e)
        {
            throw new PlotanceException(
                path,
                line,
                "Table columns does not fit into the space.",
                e
            );
        }
    }

    /// <summary>
    /// Computes the row heights for the specified table.
    /// </summary>
    /// <param name="path">The path to the file containing the table.</param>
    /// <param name="line">
    /// The line number in the file containing the table.
    /// </param>
    /// <param name="totalHeight">The total height of the table in EMU.</param>
    /// <param name="rowCount">The number of rows in the table.</param>
    /// <returns>The row heights in EMU.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the table rows does not fit into the space.
    /// </exception>
    public IReadOnlyList<long> ComputeRowHeights(
        string? path,
        long line,
        long totalHeight,
        int rowCount
    )
    {
        var rows = Rows?.Value ?? [];

        try
        {
            return ILengthWeight
                .Divide(
                    totalHeight,
                    Enumerable
                        .Range(0, rowCount)
                        .Select(
                            index => (
                                Weight: index < rows.Count
                                    ? rows[index]
                                    : new RelativeLengthWeight(null, 0, 1),
                                Gap: 0L
                            )
                        )
                        .ToList()
                )
                .Select(tuple => tuple.Length)
                .ToList();
        }
        catch (ArgumentException e)
        {
            throw new PlotanceException(
                path,
                line,
                "Table rows does not fit into the space.",
                e
            );
        }
    }
}

/// <summary>Represents the position of a cell in a table.</summary>
public struct CellPosition(
    int RowCount,
    int ColumnCount,
    int RowIndex,
    int ColumnIndex
)
{
    public bool IsFirstRow => RowIndex == 0;
    public bool IsLastRow => RowIndex == RowCount - 1;
    public bool IsFirstColumn => ColumnIndex == 0;
    public bool IsLastColumn => ColumnIndex == ColumnCount - 1;
    public bool IsBand1Row => RowIndex % 2 == 0;
    public bool IsBand2Row => RowIndex % 2 == 1;
    public bool IsBand1Column => ColumnIndex % 2 == 0;
    public bool IsBand2Column => ColumnIndex % 2 == 1;
    public bool IsSoutheast => IsLastRow && IsLastColumn;
    public bool IsSouthwest => IsLastRow && IsFirstColumn;
    public bool IsNortheast => IsFirstRow && IsLastColumn;
    public bool IsNorthwest => IsFirstRow && IsFirstColumn;

    /// <summary>
    /// The list of table style element types that should be applied to this
    /// cell position, in order of inheritance from lower to higher.
    /// </summary>
    // https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/f7f85a73-5f60-4f7e-a725-4f1f06198388
    public IReadOnlyList<TableStyleElementType> InheritanceList
        => new List<TableStyleElementType?>([
            TableStyleElementType.WholeTable,
            IsBand1Row ? TableStyleElementType.Band1Row : null,
            IsBand2Row ? TableStyleElementType.Band2Row : null,
            IsBand1Column ? TableStyleElementType.Band1Column : null,
            IsBand2Column ? TableStyleElementType.Band2Column : null,
            IsFirstColumn ? TableStyleElementType.FirstColumn : null,
            IsLastColumn ? TableStyleElementType.LastColumn : null,
            IsFirstRow ? TableStyleElementType.FirstRow : null,
            IsLastRow ? TableStyleElementType.LastRow : null,
            IsNorthwest ? TableStyleElementType.NorthwestCell : null,
            IsNortheast ? TableStyleElementType.NortheastCell : null,
            IsSouthwest ? TableStyleElementType.SouthwestCell : null,
            IsSoutheast ? TableStyleElementType.SoutheastCell : null
        ]).OfType<TableStyleElementType>().ToList();
}
