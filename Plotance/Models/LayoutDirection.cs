// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

namespace Plotance.Models;

/// <summary>The layout direction of the section.</summary>
public class LayoutDirection
{
    /// <summary>Left to right, top to bottom layout.</summary>
    public static readonly LayoutDirection Row = new LayoutDirection("row");

    /// <summary>Top to bottom, left to right layout.</summary>
    public static readonly LayoutDirection Column
        = new LayoutDirection("column");

    /// <summary>The name of the layout direction.</summary>
    public string Name { get; init; }

    /// <summary>
    /// Initializes a new layout direction with the specified name.
    /// </summary>
    private LayoutDirection(string name)
    {
        Name = name;
    }

    /// <summary>
    /// Parses the specified text into a layout direction.
    /// </summary>
    public static LayoutDirection Parse(string text) => text switch
    {
        "row" => Row,
        "column" => Column,
        _ => throw new FormatException("Unknown layout direction: " + text)
    };

    /// <summary>
    /// Chooses the specified value based on the layout direction.
    /// </summary>
    /// <param name="row">
    /// The value to return if the layout direction is row.
    /// </param>
    /// <param name="column">
    /// The value to return if the layout direction is column.
    /// </param>
    /// <returns>The chosen value.</returns>
    public T Choose<T>(T row, T column)
    {
        if (this == Row)
        {
            return row;
        }
        else if (this == Column)
        {
            return column;
        }
        else
        {
            throw new InvalidOperationException(
                "Unknown layout direction: " + Name
            );
        }
    }

    /// <summary>
    /// Invokes the specified function based on the layout direction.
    /// </summary>
    /// <param name="row">
    /// The function to invoke if the layout direction is row.
    /// </param>
    /// <param name="column">
    /// The function to invoke if the layout direction is column.
    /// </param>
    /// <returns>The result of the invoked function.</returns>
    public T Switch<T>(Func<T> row, Func<T> column)
        => Choose(row, column)();

    /// <summary>
    /// Invokes the specified function based on the layout direction.
    /// </summary>
    /// <param name="row">
    /// The function to invoke if the layout direction is row.
    /// </param>
    /// <param name="column">
    /// The function to invoke if the layout direction is column.
    /// </param>
    public void Switch(Action row, Action column)
    {
        Choose(row, column)();
    }

    /// <inheritdoc/>
    public override string ToString() => Name;
}
