// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

using D = DocumentFormat.OpenXml.Drawing;

namespace Plotance.Renderers.PowerPoint;

/// <summary>Provides methods for manipulating presentations.</summary>
public static class Presentations
{
    /// <summary>
    /// Extracts or adds the presentation part from a document.
    /// </summary>
    /// <param name="document">The presentation document.</param>
    /// <returns>The extracted presentation part.</returns>
    public static PresentationPart ExtractPresentationPart(
        PresentationDocument document
    )
    {
        var presentationPart = document.PresentationPart
            ?? document.AddPresentationPart();

        if (presentationPart.ThemePart == null)
        {
            var themePart = presentationPart.AddNewPart<ThemePart>();

            themePart.Theme = CreateTheme();
        }

        ExtractPresentation(presentationPart);

        return presentationPart;
    }

    /// <summary>
    /// Extracts or adds the presentation from a presentation part.
    /// </summary>
    /// <param name="presentationPart">The presentation part.</param>
    /// <returns>The extracted presentation.</returns>
    public static Presentation ExtractPresentation(
        PresentationPart presentationPart
    )
    {
        var presentation = presentationPart.Presentation;

        if (presentation == null)
        {
            presentation = CreatePresentation(presentationPart);

            presentationPart.Presentation = presentation;
        }

        return presentation;
    }

    /// <summary>
    /// Creates a new presentation with default settings.
    /// </summary>
    /// <param name="presentationPart">The presentation part.</param>
    /// <returns>A new Presentation instance.</returns>
    public static Presentation CreatePresentation(
        PresentationPart presentationPart
    )
    {
        var slideMasterPart = SlideMasters
            .ExtractSlideMasterPart(presentationPart);
        var slideMasterRelationshipId = presentationPart
            .GetIdOfPart(slideMasterPart);

        return new Presentation(
            new SlideMasterIdList(
                new SlideMasterId()
                {
                    Id = (UInt32Value)2147483648U,
                    RelationshipId = slideMasterRelationshipId
                }
            ),
            new SlideSize()
            {
                Cx = (Int32Value)9144000,
                Cy = (Int32Value)6858000,
                Type = SlideSizeValues.Screen4x3
            },
            new NotesSize()
            {
                Cx = (Int64Value)6858000L,
                Cy = (Int64Value)9144000L
            },
            new DefaultTextStyle()
        )
        {
            ShowSpecialPlaceholderOnTitleSlide = false
        };
    }

    /// <summary>Creates a theme with default settings.</summary>
    /// <returns>A new theme.</returns>
    private static D.Theme CreateTheme()
        => new D.Theme(
            new D.ThemeElements(
                new D.ColorScheme(
                    new D.Dark1Color(
                        new D.SystemColor()
                        {
                            Val = D.SystemColorValues.WindowText
                        }
                    ),
                    new D.Light1Color(
                        new D.SystemColor()
                        {
                            Val = D.SystemColorValues.Window
                        }
                    ),
                    new D.Dark2Color(
                        new D.RgbColorModelHex() { Val = "44546A" }
                    ),
                    new D.Light2Color(
                        new D.RgbColorModelHex() { Val = "E7E6E6" }
                    ),
                    new D.Accent1Color(
                        new D.RgbColorModelHex() { Val = "4472C4" }
                    ),
                    new D.Accent2Color(
                        new D.RgbColorModelHex() { Val = "ED7D31" }
                    ),
                    new D.Accent3Color(
                        new D.RgbColorModelHex() { Val = "A5A5A5" }
                    ),
                    new D.Accent4Color(
                        new D.RgbColorModelHex() { Val = "FFC000" }
                    ),
                    new D.Accent5Color(
                        new D.RgbColorModelHex() { Val = "5B9BD5" }
                    ),
                    new D.Accent6Color(
                        new D.RgbColorModelHex() { Val = "70AD47" }
                    ),
                    new D.Hyperlink(
                        new D.RgbColorModelHex() { Val = "0563C1" }
                    ),
                    new D.FollowedHyperlinkColor(
                        new D.RgbColorModelHex() { Val = "954F72" }
                    )
                )
                {
                    Name = "Office Theme"
                },
                new D.FontScheme(
                    new D.MajorFont(
                        new D.LatinFont() { Typeface = "Verdana" },
                        new D.EastAsianFont() { Typeface = "" },
                        new D.ComplexScriptFont() { Typeface = "" },
                        new D.SupplementalFont()
                        {
                            Script = "Jpan",
                            Typeface = "BIZ UDPゴシック"
                        },
                        new D.SupplementalFont()
                        {
                            Script = "Hang",
                            Typeface = "맑은 고딕"
                        },
                        new D.SupplementalFont()
                        {
                            Script = "Hans",
                            Typeface = "微软雅黑"
                        },
                        new D.SupplementalFont()
                        {
                            Script = "Hant",
                            Typeface = "微軟正黑體"
                        }

                    ),
                    new D.MinorFont(
                        new D.LatinFont() { Typeface = "Verdana" },
                        new D.EastAsianFont() { Typeface = "" },
                        new D.ComplexScriptFont() { Typeface = "" },
                        new D.SupplementalFont()
                        {
                            Script = "Jpan",
                            Typeface = "BIZ UDPゴシック"
                        },
                        new D.SupplementalFont()
                        {
                            Script = "Hang",
                            Typeface = "맑은 고딕"
                        },
                        new D.SupplementalFont()
                        {
                            Script = "Hans",
                            Typeface = "微软雅黑"
                        },
                        new D.SupplementalFont()
                        {
                            Script = "Hant",
                            Typeface = "微軟正黑體"
                        }

                    )
                )
                {
                    Name = "Universal Design"
                },
                new D.FormatScheme(
                    new D.FillStyleList(
                        new D.NoFill(),
                        new D.NoFill(),
                        new D.NoFill()
                    ),
                    new D.LineStyleList(
                        new D.Outline(),
                        new D.Outline(),
                        new D.Outline()
                    ),
                    new D.EffectStyleList(
                        new D.EffectStyle(new D.EffectList()),
                        new D.EffectStyle(new D.EffectList()),
                        new D.EffectStyle(new D.EffectList())
                    ),
                    new D.BackgroundFillStyleList(
                        new D.NoFill(),
                        new D.NoFill(),
                        new D.NoFill()
                    )
                )
            )
        )
        {
            Name = "Sample"
        };

    /// <summary>Find an item in a collection of OpenXml parts.</summary>
    /// <typeparam name="T">Type of the OpenXml part in collection.</typeparam>
    /// <typeparam name="R">Type of the return value.</typeparam>
    /// <param name="items">Collection of items to search.</param>
    /// <param name="extractor">Function to evaluate on each item.</param>
    /// <returns>
    /// The first result of the extractor function which is not null, or default
    /// when all results are null.
    /// </returns>
    public static R? FindOrDefault<T, R>(
        IEnumerable<T> items,
        Func<T, R?> extractor
    )
    {
        foreach (var item in items)
        {
            if (extractor(item) is R result)
            {
                return result;
            }
        }

        return default;
    }
}
