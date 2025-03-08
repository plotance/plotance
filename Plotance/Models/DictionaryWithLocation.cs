// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace Plotance.Models;

/// <summary>
/// Represents a dictionary that preserves the location (file path and line
/// number) of both keys and values in a configuration or data file.
/// </summary>
/// <typeparam name="TKey">The type of keys in the dictionary.</typeparam>
/// <typeparam name="TValue">The type of values in the dictionary.</typeparam>
/// <param name="Path">
/// The path to the file containing the dictionary, for error reporting.
/// </param>
/// <param name="Line">
/// The line number in the file where the dictionary starts, for error
/// reporting.
/// </param>
/// <param name="KeyValues">
/// Dictionary mapping keys to values with their locations.
/// </param>
/// <param name="KeyLocations">
/// Dictionary mapping keys to their source locations (file path and line
/// number) for error reporting.
/// </param>
public record DictionaryWithLocation<TKey, TValue>(
    string? Path,
    long Line,
    IReadOnlyDictionary<TKey, ValueWithLocation<TValue>> KeyValues,
    IReadOnlyDictionary<TKey, (string?, long)> KeyLocations
) : IReadOnlyDictionary<TKey, ValueWithLocation<TValue>>
{
    /// <summary>
    /// Returns the number of key/value pairs in the dictionary.
    /// </summary>
    public int Count => KeyValues.Count;

    /// <summary>
    /// Returns the value with location associated with the specified key.
    /// </summary>
    /// <param name="key">The key of the value to get.</param>
    /// <returns>
    /// The value with location associated with the specified key.
    /// </returns>
    /// <exception cref="KeyNotFoundException">
    /// The key does not exist in the dictionary.
    /// </exception>
    public ValueWithLocation<TValue> this[TKey key] => KeyValues[key];

    /// <summary>
    /// Returns a collection containing the keys in the dictionary.
    /// </summary>
    public IEnumerable<TKey> Keys => KeyValues.Keys;

    /// <summary>
    /// Returns a collection containing the values with location in the
    /// dictionary.
    /// </summary>
    public IEnumerable<ValueWithLocation<TValue>> Values => KeyValues.Values;

    /// <summary>
    /// Determines whether the dictionary contains the specified key.
    /// </summary>
    /// <param name="key">The key to locate in the dictionary.</param>
    /// <returns>
    /// true if the dictionary contains an element with the specified key;
    /// otherwise, false.
    /// </returns>
    public bool ContainsKey(TKey key) => KeyValues.ContainsKey(key);

    /// <summary>
    /// Returns an enumerator that iterates through the dictionary.
    /// </summary>
    /// <returns>An enumerator for the dictionary.</returns>
    System.Collections.IEnumerator
        System.Collections.IEnumerable.GetEnumerator()
        => KeyValues.GetEnumerator();

    /// <summary>
    /// Returns an enumerator that iterates through the dictionary.
    /// </summary>
    /// <returns>An enumerator for the dictionary.</returns>
    public IEnumerator<KeyValuePair<TKey, ValueWithLocation<TValue>>>
        GetEnumerator()
        => KeyValues.GetEnumerator();

    /// <summary>
    /// Gets the value with location associated with the specified key.
    /// </summary>
    /// <param name="key">The key of the value to get.</param>
    /// <param name="value">
    /// When this method returns, contains the value with location associated
    /// with the specified key, if the key is found; otherwise, the default
    /// value for the type of the value parameter.
    /// </param>
    /// <returns>
    /// true if the dictionary contains an element with the specified key;
    /// otherwise, false.
    /// </returns>
    public bool TryGetValue(
        TKey key,
        [MaybeNullWhen(false)] out ValueWithLocation<TValue> value
    ) => KeyValues.TryGetValue(key, out value);
}
