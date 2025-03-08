// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

namespace Plotance.Models;

/// <summary>
/// Represents a color value defined in a configuration or data file.
/// </summary>
/// <param name="Path">
/// The path to the file containing the color specification, for error
/// reporting.
/// </param>
/// <param name="Line">The line number in the file, for error reporting.</param>
/// <param name="Text">
/// The text representation of the color. This could be a named color, hex code.
/// </param>
public record Color(string? Path, long Line, string Text);
