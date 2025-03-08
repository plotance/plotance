// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Globalization;
using System.Numerics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Plotance.Models;

namespace Plotance.Renderers.PowerPoint;

/// <summary>
/// Provides methods for working with Open XML spreadsheets. Provides only
/// minimal functionality for creating embedded spreadsheets for use with
/// charts in PowerPoint.
/// </summary>
public static class Spreadsheets
{
    /// <summary>Whether to suppress messages.</summary>
    public static bool Quiet { get; set; }

    /// <summary>
    /// Generates a SpreadsheetDocument with data from a query result set for
    /// use with charts in PowerPoint.
    /// </summary>
    /// <param name="queryResult">
    /// The query result set containing the data for the spreadsheet.
    /// </param>
    /// <returns>
    /// A SpreadsheetDocument containing the data from the query result set.
    /// </returns>
    public static SpreadsheetDocument GenerateSpreadsheetDocumentForChart(
        QueryResultSet queryResult
    )
    {
        var spreadsheetDocument = SpreadsheetDocument.Create(
            new MemoryStream(),
            DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook
        );

        var workbookPart = ExtractWorkbookPart(spreadsheetDocument);
        var workbookStylesPart = ExtractWorkbookStylesPart(workbookPart);

        workbookStylesPart.Stylesheet = new Stylesheet(
            new Fonts(new Font()) { Count = 1 },
            new Fills(
                new Fill(
                    new PatternFill() { PatternType = PatternValues.None }
                )
            )
            {
                Count = 1
            },
            new Borders(
                new Border(
                    new LeftBorder(),
                    new RightBorder(),
                    new TopBorder(),
                    new BottomBorder(),
                    new DiagonalBorder()
                )
            )
            {
                Count = 1
            },
            new CellFormats(
                new CellFormat()
                {
                    NumberFormatId = 0 // General
                },
                new CellFormat()
                {
                    NumberFormatId = 14, // date
                    ApplyNumberFormat = true
                },
                new CellFormat()
                {
                    NumberFormatId = 21, // time
                    ApplyNumberFormat = true
                },
                new CellFormat()
                {
                    NumberFormatId = 22, // date and time
                    ApplyNumberFormat = true
                }
            )
            {
                Count = 4
            }
        );

        var worksheet = AppendSheet(spreadsheetDocument);
        var sheetData = worksheet.GetFirstChild<SheetData>()
            ?? throw new InvalidOperationException("Sheet data not found");

        sheetData.AppendChild(
            new Row(
                queryResult
                    .Columns
                    .Select(
                        (column, columnIndex) => new Cell(
                            ToCellChild(column.Name, CellValues.InlineString)
                        )
                        {
                            CellReference = $"{GetColumn(columnIndex)}1",
                            DataType = CellValues.InlineString
                        }
                    )
            )
            {
                RowIndex = 1
            }
        );

        Cell CreateCell(object cellValue, int rowIndex, int columnIndex)
        {
            var fieldType = queryResult.Columns[columnIndex].Type;
            var cellDataType = ToCellDataType(fieldType);

            return new Cell(ToCellChild(cellValue, cellDataType))
            {
                CellReference = $"{GetColumn(columnIndex)}{rowIndex + 2}",
                DataType = cellDataType,
                StyleIndex = fieldType switch
                {
                    _ when fieldType == typeof(DateOnly) => 1,
                    _ when fieldType == typeof(TimeOnly) => 2,
                    _ when fieldType == typeof(DateTime) => 3,
                    _ when fieldType == typeof(DateTimeOffset) => 3,
                    _ => 0
                }
            };
        }

        Row CreateRow(IReadOnlyList<object> rowData, int rowIndex)
            => new Row(
                rowData.Select(
                    (cellValue, columnIndex) => CreateCell(
                        cellValue,
                        rowIndex,
                        columnIndex
                    )
                )
            )
            {
                RowIndex = (uint)(rowIndex + 2)
            };

        sheetData.Append(queryResult.Rows.Select(CreateRow));

        if (!Quiet)
        {
            Validate(spreadsheetDocument);
        }

        return spreadsheetDocument;
    }

    /// <summary>Gets or creates the Workbook from a WorkbookPart.</summary>
    /// <param name="workbookPart">The workbook part to extract from.</param>
    /// <returns>The workbook, created if it doesn't exist.</returns>
    public static Workbook ExtractWorkbook(WorkbookPart workbookPart)
        => workbookPart.Workbook ??= new Workbook();

    /// <summary>Appends a new worksheet to a spreadsheet document.</summary>
    /// <param name="document">The spreadsheet document.</param>
    /// <param name="sheetName">
    /// Optional name for the sheet. If not provided, a default name will be
    /// generated.
    /// </param>
    /// <returns>The newly created worksheet.</returns>
    public static Worksheet AppendSheet(
        SpreadsheetDocument document,
        string? sheetName = null
    )
    {
        var workbookPart = ExtractWorkbookPart(document);
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

        worksheetPart.Worksheet = new Worksheet(new SheetData());

        var workbook = ExtractWorkbook(workbookPart);
        var sheets = (workbook.Sheets ??= new Sheets());
        var newSheetId = (
            sheets.Elements<Sheet>().Select(s => s.SheetId?.Value).Max() ?? 0
        ) + 1;
        var sheet = new Sheet()
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = (uint)newSheetId,
            Name = sheetName ?? $"Sheet{newSheetId}"
        };

        sheets.AppendChild(sheet);

        return worksheetPart.Worksheet;
    }

    /// <summary>
    /// Gets or creates the WorkbookPart from a SpreadsheetDocument.
    /// </summary>
    /// <param name="document">The spreadsheet document.</param>
    /// <returns>The workbook part, created if it doesn't exist.</returns>
    public static WorkbookPart ExtractWorkbookPart(SpreadsheetDocument document)
        => document.WorkbookPart ?? document.AddWorkbookPart();


    /// <summary>
    /// Gets or creates the WorkbookStylesPart from a WorkbookPart.
    /// </summary>
    /// <param name="workbookPart">The workbook part.</param>
    /// <returns>
    /// The workbook styles part, created if it doesn't exist.
    /// </returns>
    public static WorkbookStylesPart ExtractWorkbookStylesPart(
        WorkbookPart workbookPart
    ) => workbookPart.WorkbookStylesPart
        ?? workbookPart.AddNewPart<WorkbookStylesPart>();

    /// <summary>
    /// Converts a .NET type to the appropriate Open XML CellValues type.
    /// </summary>
    /// <param name="fieldType">The .NET type to convert.</param>
    /// <returns>The corresponding CellValues type.</returns>
    public static CellValues ToCellDataType(Type fieldType)
        => fieldType switch
        {
            var t when t == typeof(Boolean) => CellValues.Boolean,
            var t when
                t == typeof(SByte)
                || t == typeof(Byte)
                || t == typeof(Decimal)
                || t == typeof(Double)
                || t == typeof(Single)
                || t == typeof(Int16)
                || t == typeof(UInt16)
                || t == typeof(Int32)
                || t == typeof(UInt32)
                || t == typeof(Int64)
                || t == typeof(UInt64)
                || t == typeof(BigInteger)
                || t == typeof(TimeOnly)
                || t == typeof(DateTime)
                || t == typeof(DateTimeOffset)
                || t == typeof(DateOnly)
                => CellValues.Number,
            _ => CellValues.InlineString,
        };

    /// <summary>Formats a value for use in a spreadsheet cell.</summary>
    /// <param name="obj">The object to format.</param>
    /// <returns>
    /// A string representation of the object formatted for use in an Open XML
    /// spreadsheet.
    /// </returns>
    /// <remarks>
    /// Date and time values are converted to Excel serial date format, where
    /// dates are represented as days since December 30, 1899.
    /// </remarks>
    public static string FormatValue(object? obj)
        => obj switch
        {
            null => "",

            DateTime dateTime
                => (dateTime - new DateTime(1899, 12, 30))
                .TotalDays
                .ToString(CultureInfo.InvariantCulture),

            DateTimeOffset dateTimeOffset
                => (dateTimeOffset.DateTime - new DateTime(1899, 12, 30))
                .TotalDays
                .ToString(CultureInfo.InvariantCulture),

            DateOnly date
                => (date.DayNumber - new DateOnly(1899, 12, 30).DayNumber)
                .ToString(CultureInfo.InvariantCulture),

            TimeOnly time
                => time
                .ToTimeSpan()
                .TotalDays
                .ToString(CultureInfo.InvariantCulture),

            Boolean boolean
                => boolean ? "1" : "0",

            _ => Convert.ToString(obj, CultureInfo.InvariantCulture) ?? ""
        };

    /// <summary>
    /// Creates an appropriate Open XML element for a cell based on its data
    /// type.
    /// </summary>
    /// <param name="obj">The object to convert.</param>
    /// <param name="dataType">The Open XML cell data type.</param>
    /// <returns>
    /// An OpenXmlElement (either InlineString or CellValue) for the given
    /// object.
    /// </returns>
    public static OpenXmlElement ToCellChild(object obj, CellValues dataType)
    {
        var str = FormatValue(obj);

        if (dataType == CellValues.InlineString)
        {
            return new InlineString(new Text(str));
        }
        else
        {
            return new CellValue(str);
        }
    }

    /// <summary>
    /// Converts a zero-based column index to an Excel column letter.
    /// </summary>
    /// <param name="index">The zero-based column index.</param>
    /// <returns>
    /// The Excel column letter (e.g., A, B, ..., Z, AA, AB, etc.).
    /// </returns>
    public static string GetColumn(int index)
    {
        var column = "";

        while (index >= 0)
        {
            column = (char)('A' + (index % 26)) + column;
            index = index / 26 - 1;
        }

        return column;
    }

    /// <summary>
    /// Validates a SpreadsheetDocument against the Open XML schema and logs
    /// any errors.
    /// </summary>
    /// <param name="document">The spreadsheet document to validate.</param>
    /// <remarks>
    /// Validation errors are logged to the standard error output.
    /// </remarks>
    private static void Validate(SpreadsheetDocument document)
    {
        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);

        foreach (var error in validator.Validate(document))
        {
            Console.Error.WriteLine(
                string.Join(
                    "\n  ",
                    [
                        error.Id,
                        error.Part?.Uri,
                        error.Path?.XPath,
                        error.RelatedPart,
                        error.RelatedNode,
                        error.Description
                    ]
                )
            );
            Console.Error.WriteLine();
        }
    }
}
