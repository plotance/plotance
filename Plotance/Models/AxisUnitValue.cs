// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Globalization;
using System.Text.RegularExpressions;

namespace Plotance.Models;

/// <summary>Represents a value that can be used for chart axis units.</summary>
public interface IAxisUnitValue
{
    /// <summary>Converts the value to a decimal representation.</summary>
    /// <returns>The decimal value.</returns>
    public decimal ToDecimal();

    /// <summary>Converts the value to a double representation.</summary>
    /// <returns>The double value.</returns>
    public double ToDouble() => (double)ToDecimal();
}

/// <summary>
/// Represents a decimal value that can be used for chart axis units.
/// </summary>
/// <param name="Value">The decimal value.</param>
public record DecimalAxisUnitValue(decimal Value) : IAxisUnitValue
{
    /// <summary>Converts the value to a decimal representation.</summary>
    /// <returns>The decimal value.</returns>
    public decimal ToDecimal() => Value;
}

/// <summary>
/// Represents a text value that can be parsed into a chart axis unit.
/// </summary>
/// <param name="Path">
/// The path to the file containing the text, for error reporting.
/// </param>
/// <param name="Line">The line number in the file, for error reporting.</param>
/// <param name="Text">
/// The text to parse. Can be a number with optional time units
/// (days, hours, minutes/mins, seconds/secs, or corresponding singular forms).
/// </param>
public record TextAxisUnitValue(
    string? Path,
    long Line,
    string Text
) : IAxisUnitValue
{
    /// <summary>Converts the text value to a decimal representation.</summary>
    /// <returns>
    /// The decimal value, with appropriate conversion based on any specified
    /// time units.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the text cannot be parsed as a unit value.
    /// </exception>
    public decimal ToDecimal()
    {
        var text = Text.Trim();
        var units = "days?|hours?|minutes?|mins?|seconds?|secs?";
        var pattern = @$"^([0-9]*\.[0-9]+|[0-9]+)[ \t]*({units})?$";
        var match = Regex.Match(text, pattern);

        if (!match.Success)
        {
            throw new PlotanceException(
                Path,
                Line,
                $"Invalid value format: {Text}"
            );
        }

        var value = decimal.Parse(
            match.Groups[1].Value,
            CultureInfo.InvariantCulture
        );
        var unit = match.Groups[2].Value;

        return unit switch
        {
            "" => value,

            "day" or "days" => value,

            "hour" or "hours"
                => value * 1m / 24m,

            "minute" or "minutes" or "min" or "mins"
                => value * 1m / 24m / 60m,

            "second" or "seconds" or "sec" or "secs"
                => value * 1m / 24m / 60m / 60m,

            _ => throw new PlotanceException(
                Path,
                Line,
                $"Unknown unit: {unit}"
            )
        };
    }
}
