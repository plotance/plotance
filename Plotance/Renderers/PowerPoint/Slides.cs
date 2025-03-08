// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Plotance.Models;

using D = DocumentFormat.OpenXml.Drawing;

namespace Plotance.Renderers.PowerPoint;

/// <summary>Provides methods for manipulating slides.</summary>
public static class Slides
{
    /// <summary>
    /// Adds a new slide part to the presentation with the specified layout
    /// type.
    /// </summary>
    /// <param name="presentationPart">
    /// The presentation part to add the slide to.
    /// </param>
    /// <param name="layoutType">The layout type for the new slide.</param>
    /// <returns>The newly added slide part.</returns>
    /// <exception cref="InvalidOperationException">
    /// Thrown if the slide layout does not have common slide data.
    /// </exception>
    public static SlidePart AddSlidePart(
        PresentationPart presentationPart,
        SlideLayoutValues layoutType
    )
    {
        var slidePart = presentationPart.AddNewPart<SlidePart>();
        var slideMasterPart = SlideMasters
            .ExtractSlideMasterPart(presentationPart);
        var slideLayoutPart = SlideLayouts.ExtractSlideLayoutPart(
            slideMasterPart,
            layoutType
        );

        slidePart.AddPart(slideLayoutPart);

        slidePart.Slide = CreateSlideFromSlideLayout(
            SlideLayouts.ExtractSlideLayout(slideLayoutPart, layoutType)
        );

        var presentation = Presentations
            .ExtractPresentation(presentationPart);
        var slideIdList = (presentation.SlideIdList ??= new SlideIdList());
        var nextId = (
            slideIdList.Elements<SlideId>().Max(slideId => slideId.Id) ?? 256
        ) + 1;

        slideIdList.AppendChild(
            new SlideId()
            {
                Id = nextId,
                RelationshipId = presentationPart.GetIdOfPart(slidePart)
            }
        );

        return slidePart;
    }

    /// <summary>
    /// Creates a slide from a slide layout.  Shapes that are not placeholders
    /// are removed. Text body of placeholders are cleared.
    /// </summary>
    /// <param name="slideLayout">The slide layout to base the slide on.</param>
    /// <returns>A new Slide instance.</returns>
    /// <exception cref="InvalidOperationException">
    /// Thrown if the slide layout does not have common slide data.
    /// </exception>
    public static Slide CreateSlideFromSlideLayout(SlideLayout slideLayout)
    {
        var commonSlideData = slideLayout
            .CommonSlideData
            ?.CloneNode(true)
            as CommonSlideData;

        if (commonSlideData == null)
        {
            throw new InvalidOperationException("Cannot get common slide data");
        }

        commonSlideData.Name = null;

        List<PlaceholderValues?> headerFooterTypes = [
            PlaceholderValues.SlideNumber,
            PlaceholderValues.DateAndTime,
            PlaceholderValues.Header,
            PlaceholderValues.Footer
        ];

        var shapes = commonSlideData
            .GetFirstChild<ShapeTree>()
            ?.Elements()
            ?.ToList()
            ?? [];

        foreach (var element in shapes)
        {
            if (
                element
                    is NonVisualGroupShapeProperties or GroupShapeProperties
            )
            {
                continue;
            }
            else if (
                SlideLayouts.ExtractPlaceholderShape(element)
                    is PlaceholderShape placeholderShape
            )
            {
                var placeholderType = placeholderShape.Type?.Value;

                if (!headerFooterTypes.Contains(placeholderShape.Type?.Value))
                {
                    Shapes.CleanPlaceholderShape(element);
                }
            }
            else
            {
                element.Remove();
            }
        }

        return new Slide(commonSlideData);
    }

    /// <summary>Removes all slides from the presentation part.</summary>
    /// <param name="presentationPart">
    /// The presentation part to remove slides from.
    /// </param>
    public static void RemoveAllSlides(PresentationPart presentationPart)
    {
        var slideIdList = presentationPart.Presentation?.SlideIdList;

        if (slideIdList == null)
        {
            return;
        }

        var slideParts = slideIdList
            .Elements<SlideId>()
            .Select(slideId => slideId.RelationshipId?.Value)
            .OfType<string>()
            .Select(presentationPart.GetPartById)
            .OfType<SlidePart>();


        foreach (var slidePart in slideParts)
        {
            presentationPart.DeletePart(slidePart);
        }

        slideIdList.RemoveAllChildren<SlideId>();
    }
}
