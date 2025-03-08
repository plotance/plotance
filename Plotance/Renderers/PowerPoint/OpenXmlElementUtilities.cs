// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using DocumentFormat.OpenXml;

namespace Plotance.Renderers.PowerPoint;

/// <summary>Provides methods for Open XML elements.</summary>
public static class OpenXmlElementUtilities
{
    /// <summary>
    /// Invokes an action on an Open XML element if a value is provided.
    /// If value is provided and self is null, a new instance of T is created.
    /// </summary>
    /// <typeparam name="T">The type of the Open XML element.</typeparam>
    /// <typeparam name="V">
    /// The type of the value to pass to the action.
    /// </typeparam>
    /// <param name="self">The Open XML element to invoke the action on.</param>
    /// <param name="action">The action to invoke on the element.</param>
    /// <param name="value">The value to pass to the action.</param>
    /// <returns>The self.</returns>
    public static T? With<T, V>(this T? self, Action<T, V> action, V? value)
        where T : new()
    {
        if (value != null)
        {
            self ??= new T();

            action(self, value);
        }

        return self;
    }

    /// <summary>
    /// Creates an Open XML element if any of the given children are not null.
    /// Children are added to the element using AddChild method. Be aware that
    /// AddChild overrides existing child of the same type, so you cannot pass
    /// multiple children of the same type.
    /// </summary>
    /// <typeparam name="T">The type of the Open XML element.</typeparam>
    /// <param name="children">The children to add to the element.</param>
    /// <returns>
    /// The created Open XML element, or null if all children are null.
    /// </returns>
    public static T? CreateIfAnyChild<T>(params OpenXmlElement?[] children)
        where T : OpenXmlCompositeElement, new()
    {
        if (children.All(child => child == null))
        {
            return null;
        }

        var result = new T();

        foreach (var child in children)
        {
            if (child != null)
            {
                result.AddChild(child);
            }
        }

        return result;
    }
}
