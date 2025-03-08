// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using DocumentFormat.OpenXml;

using D = DocumentFormat.OpenXml.Drawing;

namespace Plotance.Renderers.PowerPoint;

/// <summary>
/// Represents child elements and attributes of RunProperties.
/// </summary>
public class RunPropertiesChildren
{
    /// <summary>
    /// The stack of actions which adds child elements and attributes to
    /// the RunProperties.
    /// </summary>
    private LinkedList<Action<D.RunProperties>> _actions = new();

    /// <summary>
    /// Creates a RunProperties element with the childrens and attributes.
    /// </summary>
    /// <returns>
    /// The RunProperties element, or null if no actions are specified.
    /// </returns>
    public D.RunProperties? CreateRunProperties()
    {
        if (_actions.Any())
        {
            var runProperties = new D.RunProperties();

            foreach (var action in _actions)
            {
                action(runProperties);
            }

            return runProperties;
        }
        else
        {
            return null;
        }
    }

    /// <summary>
    /// Pushes an action that adds a child element or an attribute.
    /// </summary>
    /// <param name="action">The action to push.</param>
    public void Push(Action<D.RunProperties> action)
    {
        _actions.AddLast(action);
    }

    /// <summary>Pushes a child element to the stack.</summary>
    /// <param name="child">The child element to push.</param>
    public void Push(OpenXmlElement child)
    {
        _actions.AddLast(
            runProperties => runProperties.AddChild(child.CloneNode(true))
        );
    }

    /// <summary>Removes the action from the stack top.</summary>
    public void Pop()
    {
        _actions.RemoveLast();
    }

    /// <summary>
    /// Temporarily pushes an action and runs the body with the stack.
    /// </summary>
    /// <param name="action">The action to push.</param>
    /// <param name="body">The body to run.</param>
    /// <returns>The result of the body.</returns>
    public R WithValue<R>(
        Action<D.RunProperties> action,
        Func<RunPropertiesChildren, R> body
    )
    {
        Push(action);

        try
        {
            return body(this);
        }
        finally
        {
            Pop();
        }
    }

    /// <summary>
    /// Temporarily pushes actions and runs the body with the stack.
    /// </summary>
    /// <param name="actions">The actions to push.</param>
    /// <param name="body">The body to run.</param>
    /// <returns>The result of the body.</returns>
    public R WithValues<R>(
        IEnumerable<Action<D.RunProperties>> actions,
        Func<RunPropertiesChildren, R> body
    )
    {
        foreach (var action in actions)
        {
            Push(action);
        }

        try
        {
            return body(this);
        }
        finally
        {
            foreach (var _ in actions)
            {
                Pop();
            }
        }
    }

    /// <summary>
    /// Temporarily pushes a child element and runs the body with the stack.
    /// </summary>
    /// <param name="child">The child element to push.</param>
    /// <param name="body">The body to run.</param>
    /// <returns>The result of the body.</returns>
    public R WithValue<R>(
        OpenXmlElement child,
        Func<RunPropertiesChildren, R> body
    )
    {
        Push(child);

        try
        {
            return body(this);
        }
        finally
        {
            Pop();
        }
    }

    /// <summary>
    /// Temporarily pushes child elements and runs the body with the stack.
    /// </summary>
    /// <param name="children">The child elements to push.</param>
    /// <param name="body">The body to run.</param>
    /// <returns>The result of the body.</returns>
    public R WithValues<R>(
        IEnumerable<OpenXmlElement> children,
        Func<RunPropertiesChildren, R> body
    )
    {
        foreach (var child in children)
        {
            Push(child);
        }

        try
        {
            return body(this);
        }
        finally
        {
            foreach (var _ in children)
            {
                Pop();
            }
        }
    }
}
