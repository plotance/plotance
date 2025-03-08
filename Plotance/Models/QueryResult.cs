// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Data;

namespace Plotance.Models;

/// <summary>
/// Represents a set of results from a database query, including column metadata
/// and row data.
/// </summary>
/// <param name="Columns">
/// The list of columns in the result set, containing name and type information.
/// </param>
/// <param name="Rows">
/// The list of rows in the result set. Each row is a list of values
/// corresponding to the columns.
/// </param>
public sealed record QueryResultSet(
    IReadOnlyList<QueryResultColumn> Columns,
    IReadOnlyList<IReadOnlyList<object>> Rows
)
{
    /// <summary>Creates a QueryResultSet from an IDataReader.</summary>
    /// <param name="reader">
    /// The data reader to extract results from. The reader must be positioned
    /// before the first record.
    /// </param>
    /// <returns>
    /// A QueryResultSet containing the columns and rows from the data reader.
    /// </returns>
    public static QueryResultSet FromDataReader(IDataReader reader)
    {
        var columns = new List<QueryResultColumn>();

        for (var i = 0; i < reader.FieldCount; i++)
        {
            columns.Add(
                new QueryResultColumn(
                    Name: reader.GetName(i),
                    Type: reader.GetFieldType(i)
                )
            );
        }

        var rows = new List<IReadOnlyList<object>>();

        while (reader.Read())
        {
            var row = new List<object>();

            for (var i = 0; i < reader.FieldCount; i++)
            {
                row.Add(reader.GetValue(i));
            }

            rows.Add(row);
        }

        return new QueryResultSet(Columns: columns, Rows: rows);
    }
}

/// <summary>Represents metadata for a column in a query result set.</summary>
/// <param name="Name">The name of the column.</param>
/// <param name="Type">The data type of the column.</param>
public sealed record QueryResultColumn(string Name, Type Type);
