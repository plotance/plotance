// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Plotance.Models;

public interface ILengthWeight
{
    /// <summary>
    /// The path to the file containing the length weight, for error reporting.
    /// </summary>
    string? Path { get; }

    /// <summary>The line number in the file, for error reporting.</summary>
    long Line { get; }

    /// <summary>
    /// Creates a length weight from a string representation.
    /// </summary>
    /// <param name="text">The text to parse.</param>
    /// <returns>
    /// A RelativeLengthWeight if the text contains only numbers,
    /// otherwise an AbsoluteLengthWeight.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown when the text cannot be parsed as a valid length weight.
    /// </exception>
    public static ILengthWeight FromString(ValueWithLocation<string> text)
    {
        var trimmed = text.Value.Trim();

        if (string.IsNullOrEmpty(trimmed))
        {
            throw new PlotanceException(
                text.Path,
                text.Line,
                "Invalid length or weight."
            );
        }

        if (Regex.IsMatch(trimmed, @"^(?:[0-9]*\.[0-9]+|[0-9]+)$"))
        {
            if (
                decimal.TryParse(
                    trimmed,
                    NumberStyles.AllowDecimalPoint,
                    CultureInfo.InvariantCulture,
                    out decimal weight
                )
            )
            {
                return new RelativeLengthWeight(text.Path, text.Line, weight);
            }
            else
            {
                throw new PlotanceException(
                    text.Path,
                    text.Line,
                    "Invalid length or weight."
                );
            }
        }
        else
        {
            return new AbsoluteLengthWeight(text.Path, text.Line, trimmed);
        }
    }

    /// <summary>
    /// Divides a total length into segments based on the given weights.
    /// </summary>
    /// <param name="total">The total length to divide.</param>
    /// <param name="weights">
    /// The list of weights and gaps. The gap is inserted before each weight.
    /// The first gap is ignored.
    /// </param>
    /// <returns>A list of (start, length) pairs.</returns>
    /// <exception cref="ArgumentException">
    /// Thrown when the total length is less than the sum of absolute lengths.
    /// </exception>
    /// <exception cref="PlotanceException">
    /// Thrown when an unknown length weight is encountered.
    /// </exception>
    public static IReadOnlyList<(long Start, long Length)> Divide(
        long total,
        IReadOnlyList<(ILengthWeight Weight, long Gap)> weights
    )
    {
        var totalRelativeLength = weights
            .Select(w => w.Weight)
            .OfType<RelativeLengthWeight>()
            .Sum(w => w.Weight);
        var totalAbsoluteLength = weights
            .Select(w => w.Weight)
            .OfType<ILength>()
            .Sum(l => l.ToEmu())
            + weights.Skip(1).Sum(w => w.Gap);
        var restLength = total - totalAbsoluteLength;

        if (restLength < 0)
        {
            throw new ArgumentException("Absolute length is too large.");
        }

        var result = new List<(long Start, long Length)>();
        var current = 0L;
        var isFirst = true;

        foreach (var (weight, gap) in weights)
        {
            var length = weight switch
            {
                RelativeLengthWeight relative => (long)decimal.Round(
                    relative.Weight * restLength / totalRelativeLength
                ),
                ILength absolute => absolute.ToEmu(),
                _ => throw new PlotanceException(
                    weight.Path,
                    weight.Line,
                    "Unknown length weight."
                )
            };

            if (!isFirst)
            {
                current += gap;
            }

            result.Add((current, length));

            current += length;
            isFirst = false;
        }

        return result;
    }
}

/// <summary>
/// Represents an absolute length weight, measured in a specific unit.
/// </summary>
/// <param name="Length">The length value.</param>
public record AbsoluteLengthWeight(TextLength Length) : ILength, ILengthWeight
{
    /// <inheritdoc/>
    public string? Path => Length.Path;

    /// <inheritdoc/>
    public long Line => Length.Line;

    /// <summary>
    /// Creates an absolute length weight from a string representation.
    /// </summary>
    /// <param name="path">The path to the file containing the length.</param>
    /// <param name="line">The line number in the file.</param>
    /// <param name="text">The text to parse.</param>
    public AbsoluteLengthWeight(string? path, long line, string text) : this(
        new TextLength(path, line, text)
    )
    {
    }

    /// <inheritdoc/>
    public decimal ToEmuDecimal() => Length.ToEmuDecimal();
};

/// <summary>
/// Represents a relative length weight, specified as a proportional number.
/// </summary>
/// <param name="Path">The path to the file containing the weight.</param>
/// <param name="Line">The line number in the file.</param>
/// <param name="Weight">The weight value.</param>
public record RelativeLengthWeight(
    string? Path,
    long Line,
    decimal Weight
) : ILengthWeight;
