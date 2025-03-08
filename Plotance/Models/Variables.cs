// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;

namespace Plotance.Models;

/// <summary>
/// Provides functionality for expanding variables in strings.
/// </summary>
public static class Variables
{
    /// <summary>
    /// Expands variables in a string using the specified dictionary.
    /// </summary>
    /// <remarks>
    /// Variables are enclosed in ${} syntax, for example ${varname}.
    /// If a variable is not found in the dictionary, the variable reference
    /// is left unchanged. Variables are not expanded recursively.
    /// </remarks>
    /// <param name="text">The text containing variable references.</param>
    /// <param name="variables">
    /// Dictionary mapping variable names to their values.
    /// </param>
    /// <returns>
    /// A new string with variables expanded, or null if the input is null.
    /// </returns>
    [return: NotNullIfNotNull(nameof(text))]
    public static string? ExpandVariables(
        string? text,
        IReadOnlyDictionary<string, string> variables
    )
    {
        if (text == null)
        {
            return null;
        }

        return Regex.Replace(
            text,
            @"\$\{(?<name>[^}]*)}",
            match =>
            {
                string name = match.Groups["name"].Value;

                return variables.TryGetValue(name, out var value)
                    ? value
                    : match.Value;
            }
        );
    }
}
