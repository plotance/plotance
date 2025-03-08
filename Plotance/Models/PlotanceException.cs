// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

namespace Plotance.Models;

/// <summary>Exception thrown by Plotance.</summary>
public class PlotanceException : Exception
{
    /// <summary>
    /// The path to the Markdown/YAML file the exception occurred in.
    /// </summary>
    public string? Path { get; init; }

    /// <summary>
    /// The line number in the Markdown/YAML file the exception occurred in.
    /// </summary>
    public long? Line { get; init; }

    /// <summary>The message with the path and line number.</summary>
    public string MessageWithLocation => (Path, Line) switch
    {
        (null, null) => Message,
        (null, 0) => Message,
        (_, null) => $"{Path}: {Message}",
        (_, 0) => $"{Path}: {Message}",
        _ => $"{Path}:{Line}: {Message}"
    };

    /// <summary>Initializes a new instance of the exception.</summary>
    public PlotanceException() : base()
    {
    }

    /// <summary>Initializes a new instance of the exception.</summary>
    /// <param name="path">
    /// The path to the Markdown/YAML file the exception occurred in.
    /// </param>
    /// <param name="line">
    /// The line number in the Markdown/YAML file the exception occurred in.
    /// </param>
    /// <param name="message">The message to display.</param>
    public PlotanceException(string? path, long? line, string message)
        : base(message)
    {
        Path = path;
        Line = line;
    }

    /// <summary>Initializes a new instance of the exception.</summary>
    /// <param name="path">
    /// The path to the Markdown/YAML file the exception occurred in.
    /// </param>
    /// <param name="line">
    /// The line number in the Markdown/YAML file the exception occurred in.
    /// </param>
    /// <param name="message">The message to display.</param>
    /// <param name="innerException">The inner exception.</param>
    public PlotanceException(
        string? path,
        long? line,
        string message,
        Exception innerException
    ) : base(message, innerException)
    {
        Path = path;
        Line = line;
    }
}
