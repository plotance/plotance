// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Globalization;
using System.Text.RegularExpressions;

namespace Plotance.Models;

/// <summary>
/// Represents a value that can be used for chart axis range boundaries.
/// </summary>
public interface IAxisRangeValue
{
    /// <summary>Converts the value to a decimal representation.</summary>
    /// <returns>
    /// The decimal value, or null if the value represents an automatic range.
    /// </returns>
    decimal? ToDecimal();

    /// <summary>Converts the value to a double representation.</summary>
    /// <returns>
    /// The double value, or null if the value represents an automatic range.
    /// </returns>
    double? ToDouble()
    {
        var value = ToDecimal();

        return value == null ? null : (double)value;
    }
}

/// <summary>
/// Represents a decimal value that can be used for chart axis range boundaries.
/// </summary>
/// <param name="Value">The decimal value.</param>
public record DecimalAxisRangeValue(decimal Value) : IAxisRangeValue
{
    /// <summary>Converts the value to a decimal representation.</summary>
    /// <returns>The decimal value.</returns>
    public decimal? ToDecimal() => Value;
}

/// <summary>
/// Represents a text value that can be parsed into a chart axis range boundary.
/// </summary>
/// <param name="Path">
/// The path to the file containing the text, for error reporting.
/// </param>
/// <param name="Line">The line number in the file, for error reporting.</param>
/// <param name="Text">
/// The text to parse. Can be a number, a date/time in
/// yyyy-MM-dd[THH[:mm[:ss[.SSS]]]] format, a time in HH:mm[:ss[.SSS]] format,
/// or the special value "auto". "T" can be in lower case or " ".
/// </param>
public record TextAxisRangeValue(
    string? Path,
    long Line,
    string Text
) : IAxisRangeValue
{
    /// <summary>Converts the text value to a decimal representation.</summary>
    /// <returns>The decimal value, or null if the text is "auto".</returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the value cannot be parsed as a number, date, or time.
    /// </exception>
    public decimal? ToDecimal()
    {
        int ParseInt(string text)
            => Int32.Parse(text, CultureInfo.InvariantCulture);
        decimal ParseDecimal(string text)
            => decimal.Parse(text, CultureInfo.InvariantCulture);

        // Attempts to parse a date/time string in
        // yyyy-MM-dd[THH[:mm[:ss[.SSS]]]] format. "T" can be in lower case or
        // " ".
        // Returns the equivalent in days since December 30, 1899, or null if
        // the text is not in the expected format.
        decimal? TryParseDateTime(string text)
        {
            var pattern = new Regex(
                """
                  ^
                  (?<year>[0-9]{4})-(?<month>[0-9]{2})-(?<day>[0-9]{2})
                  (
                    [tT ]
                    (?<hour>[0-9]{2})
                    (
                      :
                      (?<minute>[0-9]{2})
                      (
                        :
                        (?<second>[0-9]{2})
                        (
                          \.
                          (?<fraction>[0-9]+)
                        )?
                      )?
                    )?
                  )?
                  $
                """,
                RegexOptions.IgnorePatternWhitespace
                    | RegexOptions.ExplicitCapture
            );

            var match = pattern.Match(text);

            if (!match.Success)
            {
                return null;
            }

            // Allow 2025-13-01 to be 2026-01-01.
            var date = new DateOnly(ParseInt(match.Groups["year"].Value), 1, 1);

            date.AddMonths(ParseInt(match.Groups["month"].Value) - 1);
            date.AddDays(ParseInt(match.Groups["day"].Value) - 1);

            var days = date.DayNumber - new DateOnly(1899, 12, 30).DayNumber;
            var hours = match.Groups["hour"].Success
                ? ParseInt(match.Groups["hour"].Value)
                : 0;
            var minutes = match.Groups["minute"].Success
                ? ParseInt(match.Groups["minute"].Value)
                : 0;
            var seconds = match.Groups["second"].Success
                ? ParseInt(match.Groups["second"].Value)
                : 0;
            var fraction = match.Groups["fraction"].Success
                ? ParseDecimal("0." + match.Groups["fraction"].Value)
                : 0;

            return days
                + (hours + (minutes + (seconds + fraction) / 60m) / 60m) / 24m;
        }

        // Attempts to parse a time string in HH:MM[:SS[.fraction]] format.
        // Returns the equivalent in days (fraction of a day), or null if the
        // text is not in the expected format.
        decimal? TryParseTime(string text)
        {
            var pattern = new Regex(
                """
                  ^
                  (?<hour>[0-9]{2})
                  :
                  (?<minute>[0-9]{2})
                  (
                    :
                    (?<second>[0-9]{2})
                    (
                      \.
                      (?<fraction>[0-9]+)
                    )?
                  )?
                  $
                """,
                RegexOptions.IgnorePatternWhitespace
                    | RegexOptions.ExplicitCapture
            );

            var match = pattern.Match(text);

            if (!match.Success)
            {
                return null;
            }

            var hours = ParseInt(match.Groups["hour"].Value);
            var minutes = ParseInt(match.Groups["minute"].Value);
            var seconds = match.Groups["second"].Success
                ? ParseInt(match.Groups["second"].Value)
                : 0;
            var fraction = match.Groups["fraction"].Success
                ? ParseDecimal("0." + match.Groups["fraction"].Value)
                : 0;

            return (hours + (minutes + (seconds + fraction) / 60m) / 60m) / 24m;
        }

        // Attempts to parse a numeric string.
        // Returns the decimal value, or null if the text is not a valid number.
        decimal? TryParseNumber(string text)
        {
            var pattern = @"^[+-]?(?:[0-9]*\.[0-9]+|[0-9]+)$";
            var match = Regex.Match(text, pattern);

            return match.Success ? ParseDecimal(text) : null;
        }

        var text = Text.Trim();

        try
        {
            return text == "auto"
                ? null
                : TryParseDateTime(text)
                ?? TryParseTime(text)
                ?? TryParseNumber(text)
                ?? throw new PlotanceException(
                    Path,
                    Line,
                    $"Invalid value format: {text}"
                );
        }
        catch (ArgumentOutOfRangeException e)
        {
            throw new PlotanceException(
                Path,
                Line,
                $"Invalid value format: {text}",
                e
            );
        }
    }
}
