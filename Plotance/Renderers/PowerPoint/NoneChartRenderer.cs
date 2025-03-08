// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using DocumentFormat.OpenXml.Packaging;
using Plotance.Models;

namespace Plotance.Renderers.PowerPoint;

/// <summary>A chart renderer that does nothing.</summary>
public class NoneChartRenderer : ChartRenderer
{
    /// <inheritdoc/>
    protected override void RenderChart(
        ChartPart chartPart,
        ImplicitSectionColumn column,
        QueryResultSet queryResult,
        string relationShipId
    )
    {
        // Do nothing
    }
}
