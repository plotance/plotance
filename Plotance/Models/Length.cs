// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Globalization;
using System.Text.RegularExpressions;

namespace Plotance.Models;

/// <summary>Represents a length value that can be used in charts.</summary>
public interface ILength
{
    /// <summary>
    /// The number of English Metric Units (EMU) per millimeter.
    /// </summary>
    public const long EmuPerMm = 36000;
    /// <summary>
    /// The number of English Metric Units (EMU) per centimeter.
    /// </summary>
    public const long EmuPerCm = 360000;
    /// <summary>
    /// The number of English Metric Units (EMU) per inch.
    /// </summary>
    public const long EmuPerInch = 914400;
    /// <summary>
    /// The number of English Metric Units (EMU) per point (1/72 inch).
    /// </summary>
    public const long EmuPerPt = 12700;

    /// <summary>
    /// Converts the length to English Metric Units (EMU) as a decimal value.
    /// </summary>
    /// <returns>The length in EMU as a decimal.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the text cannot be parsed as a length with a valid unit.
    /// </exception>
    decimal ToEmuDecimal();

    /// <summary>
    /// Converts the length to English Metric Units (EMU) as a long integer.
    /// </summary>
    /// <returns>
    /// The length in EMU as a long, rounded to the nearest integer.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the text cannot be parsed as a length with a valid unit.
    /// </exception>
    long ToEmu() => (long)decimal.Round(ToEmuDecimal());

    /// <summary>Converts the length to points (1/72 inch).</summary>
    /// <returns>
    /// The length in points as an integer, rounded to the nearest point.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the text cannot be parsed as a length with a valid unit.
    /// </exception>
    int ToPoint() => (int)decimal.Round(ToEmuDecimal() / EmuPerPt);

    /// <summary>
    /// Converts the length to centipoints (1/100 of a point).
    /// </summary>
    /// <returns>
    /// The length in centipoints as an integer, rounded to the nearest
    /// centipoint.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the text cannot be parsed as a length with a valid unit.
    /// </exception>
    int ToCentipoint() => (int)decimal.Round(ToEmuDecimal() * 100 / EmuPerPt);

    /// <summary>Creates a length from a point value (1/72 inch).</summary>
    /// <param name="point">The length in points.</param>
    /// <returns>
    /// A length object representing the specified point value.
    /// </returns>
    public static ILength FromPoint(decimal point)
        => new EmuLength(point * EmuPerPt);

    /// <summary>Creates a length from an EMU value.</summary>
    /// <param name="emu">The length in English Metric Units.</param>
    /// <returns>A length object representing the specified EMU value.</returns>
    public static ILength FromEmu(decimal emu)
        => new EmuLength(emu);

    /// <summary>Gets a length object representing zero.</summary>
    public static ILength Zero => new EmuLength(0);
}

/// <summary>
/// Represents a length value stored directly in English Metric Units (EMU).
/// </summary>
/// <param name="Emu">The length in English Metric Units.</param>
public record EmuLength(decimal Emu) : ILength
{
    /// <summary>
    /// Converts the length to English Metric Units (EMU) as a decimal value.
    /// </summary>
    /// <returns>The length in EMU as a decimal.</returns>
    public decimal ToEmuDecimal() => Emu;
}

/// <summary>
/// Represents a length value defined as text in a configuration or data file.
/// The text can contain a number with a unit (mm, cm, in, pt, pc, px).
/// </summary>
/// <param name="Path">
/// The path to the file containing the length specification, for error
/// reporting.
/// </param>
/// <param name="Line">The line number in the file, for error reporting.</param>
/// <param name="Text">
/// The text representation of the length, which should include a number and a
/// unit (mm, cm, in, pt, pc, px).
/// </param>
public record TextLength(string? Path, long Line, string Text) : ILength
{
    /// <summary>
    /// Converts the text length to English Metric Units (EMU) as a decimal
    /// value.
    /// </summary>
    /// <returns>The length in EMU as a decimal.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the text cannot be parsed as a length with a valid unit.
    /// </exception>
    public decimal ToEmuDecimal()
    {
        var length = Text.Trim();

        if (string.IsNullOrEmpty(length))
        {
            throw new PlotanceException(
                Path,
                Line,
                $"Invalid length format: {Text}"
            );
        }

        if (length == "0")
        {
            return 0;
        }

        var units = "mm|cm|in|pt|pc|px";
        var pattern = @$"^([+-]?(?:[0-9]*\.[0-9]+|[0-9]+))[ \t]*({units})$";
        var match = Regex.Match(length, pattern);

        if (!match.Success)
        {
            throw new PlotanceException(
                Path,
                Line,
                $"Invalid length format: {Text}"
            );
        }

        var value = decimal.Parse(
            match.Groups[1].Value,
            CultureInfo.InvariantCulture
        );
        var unit = match.Groups[2].Value;

        return unit switch
        {
            "mm" => decimal.Round(value * ILength.EmuPerMm),
            "cm" => decimal.Round(value * ILength.EmuPerCm),
            "in" => decimal.Round(value * ILength.EmuPerInch),
            "pt" => decimal.Round(value * ILength.EmuPerPt),
            "pc" => decimal.Round(value * 12 * ILength.EmuPerPt), // 12 pt
            "px" => decimal.Round(value * ILength.EmuPerInch / 96), // 1/96 in
            _ => throw new PlotanceException(
                Path,
                Line,
                $"Unknown unit: {unit}"
            )
        };
    }
}
