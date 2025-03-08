// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Diagnostics.CodeAnalysis;

namespace Plotance.Models;

/// <summary>
/// Represents a value with its source location (file path and line number)
/// in a configuration or data file.
/// </summary>
/// <typeparam name="T">The type of the contained value.</typeparam>
/// <param name="Path">
/// The path to the file containing the value, for error reporting.
/// </param>
/// <param name="Line">The line number in the file, for error reporting.</param>
/// <param name="Value">The actual value.</param>
public record ValueWithLocation<T>(string? Path, long Line, T Value)
{
    /// <summary>
    /// Maps the contained value using the specified function, preserving the
    /// source location.
    /// </summary>
    /// <typeparam name="R">The return type of the mapping function.</typeparam>
    /// <param name="function">The function to apply to the value.</param>
    /// <returns>
    /// A new ValueWithLocation containing the result of applying the function
    /// to the value, with the same source location.
    /// </returns>
    public ValueWithLocation<R> Map<R>(Func<T, R> function)
        => new ValueWithLocation<R>(Path, Line, function(Value));

    /// <summary>
    /// Maps the contained value using the specified function, preserving the
    /// source location, and throws a PlotanceException if the function throws
    /// a FormatException.
    /// </summary>
    /// <typeparam name="R">The return type of the mapping function.</typeparam>
    /// <param name="name">The name of the value being parsed.</param>
    /// <param name="function">The function to apply to the value.</param>
    /// <returns>
    /// A new ValueWithLocation containing the result of applying the function
    /// to the value, with the same source location.
    /// </returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the function throws a FormatException.
    /// </exception>
    public ValueWithLocation<R> Parse<R>(string name, Func<T, R> function)
    {
        try
        {
            return Map(function);
        }
        catch (FormatException e)
        {
            throw new PlotanceException(
                Path,
                Line,
                $"Invalid {name}: {Value}",
                e
            );
        }
    }
}

/// <summary>
/// Provides extension methods for the ValueWithLocation class.
/// </summary>
public static class ValueWithLocationExtension
{
    /// <summary>
    /// Executes the specified action on the value if the ValueWithLocation is
    /// not null.
    /// </summary>
    /// <typeparam name="T">The type of the contained value.</typeparam>
    /// <param name="self">The ValueWithLocation instance.</param>
    /// <param name="action">
    /// The action to execute on the contained value.
    /// </param>
    public static void CallIfNotNull<T>(
        this ValueWithLocation<T>? self,
        Action<T> action
    )
    {
        if (self != null)
        {
            action(self.Value);
        }
    }

    /// <summary>
    /// Applies the specified function to the value if the ValueWithLocation is
    /// not null, otherwise returns the default value.
    /// </summary>
    /// <typeparam name="T">The type of the contained value.</typeparam>
    /// <typeparam name="R">The return type of the function.</typeparam>
    /// <param name="self">The ValueWithLocation instance.</param>
    /// <param name="function">
    /// The function to apply to the contained value.
    /// </param>
    /// <param name="defaultValue">
    /// The value to return if the ValueWithLocation is null.
    /// </param>
    /// <returns>
    /// The result of applying the function to the value, or the default value
    /// if the ValueWithLocation is null.
    /// </returns>
    public static R CallIfNotNull<T, R>(
        this ValueWithLocation<T>? self,
        Func<T, R> function,
        R defaultValue
    ) => self == null ? defaultValue : function(self.Value);
}
