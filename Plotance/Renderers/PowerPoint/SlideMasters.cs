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

/// <summary>Provides methods for manipulating slide masters.</summary>
public static class SlideMasters
{
    /// <summary>
    /// Extracts or adds the slide master part from a presentation part.
    /// </summary>
    /// <param name="presentationPart">
    /// The presentation part to extract the slide master part from.
    /// </param>
    /// <returns>The first slide master part found.</returns>
    public static SlideMasterPart ExtractSlideMasterPart(
        PresentationPart presentationPart
    )
    {
        var slideMasterParts = presentationPart.SlideMasterParts;
        var slideMasterPart = slideMasterParts.FirstOrDefault()
            ?? presentationPart.AddNewPart<SlideMasterPart>();

        if (
            slideMasterPart.ThemePart == null
                && presentationPart.ThemePart != null
        )
        {
            slideMasterPart.AddPart(presentationPart.ThemePart);
        }

        ExtractSlideMaster(slideMasterPart);

        return slideMasterPart;
    }

    /// <summary>
    /// Extracts or adds the slide master from a slide master part.
    /// </summary>
    /// <param name="slideMasterPart">
    /// The slide master part to extract the slide master from.
    /// </param>
    /// <returns>The slide master.</returns>
    public static SlideMaster ExtractSlideMaster(
        SlideMasterPart slideMasterPart
    )
    {
        var slideMaster = slideMasterPart.SlideMaster;

        if (slideMaster == null)
        {
            List<int> fontSizes = [
                2800,
                2400,
                2000,
                1800,
                1800,
                1800,
                1800,
                1800,
                1800
            ];
            var emuPerPoint = 914400 / 72;
            var leftMargins = fontSizes
                .Select(size => size * emuPerPoint / 100)
                .Aggregate(
                    (Margin: 0, Result: Enumerable.Empty<int>()),
                    (tuple, size) => (
                        Margin: tuple.Margin + size,
                        Result: tuple.Result.Append(tuple.Margin + size)
                    ),
                    tuple => tuple.Result
                )
                .ToList();
            var indents = fontSizes
                .Select(size => -size * emuPerPoint / 100)
                .ToList();
            var language = CultureInfo.CurrentCulture.Name;

            if (string.IsNullOrEmpty(language))
            {
                language = "en-US";
            }

            T CreateParagraphProperties<T>(
                string bulletCharacter,
                int fontSize,
                int leftMargin,
                int indent
            )
                where T : D.TextParagraphPropertiesType, new()
            {
                var paragraphProperties = new T();

                paragraphProperties.AddChild(
                    new D.CharacterBullet() { Char = bulletCharacter }
                );
                paragraphProperties.AddChild(
                    new D.DefaultRunProperties(
                        new D.SolidFill(
                            new D.SchemeColor()
                            {
                                Val = D.SchemeColorValues.Dark1
                            }
                        ),
                        new D.LatinFont()
                        {
                            Typeface = "+mn-lt"
                        },
                        new D.EastAsianFont()
                        {
                            Typeface = "+mn-ea"
                        },
                        new D.ComplexScriptFont()
                        {
                            Typeface = "+mn-cs"
                        }
                    )
                    {
                        Language = language,
                        AlternativeLanguage = "en-US",
                        FontSize = fontSize
                    }
                );
                paragraphProperties.LeftMargin = leftMargin;
                paragraphProperties.Indent = indent;

                return paragraphProperties;
            }

            slideMaster = new SlideMaster(
                SlideLayouts.CreateMainCommonSlideData("‹#›"),
                CreateColorMap(),
                new TextStyles(
                    new TitleStyle(
                        new D.DefaultParagraphProperties(
                            new D.DefaultRunProperties(
                                new D.SolidFill(
                                    new D.SchemeColor()
                                    {
                                        Val = D.SchemeColorValues.Dark1
                                    }
                                )
                            )
                            {
                                Language = language,
                                AlternativeLanguage = "en-US",
                            }
                        ),
                        new D.Level1ParagraphProperties(
                            new D.DefaultRunProperties(
                                new D.SolidFill(
                                    new D.SchemeColor()
                                    {
                                        Val = D.SchemeColorValues.Dark1
                                    }
                                ),
                                new D.LatinFont()
                                {
                                    Typeface = "+mj-lt"
                                },
                                new D.EastAsianFont()
                                {
                                    Typeface = "+mj-ea"
                                },
                                new D.ComplexScriptFont()
                                {
                                    Typeface = "+mj-cs"
                                }
                            )
                            {
                                Language = language,
                                AlternativeLanguage = "en-US",
                                FontSize = 4400
                            }
                        )
                    ),
                    new BodyStyle(
                        new D.DefaultParagraphProperties(
                            new D.DefaultRunProperties(
                                new D.SolidFill(
                                    new D.SchemeColor()
                                    {
                                        Val = D.SchemeColorValues.Dark1
                                    }
                                )
                            )
                            {
                                Language = language,
                                AlternativeLanguage = "en-US",
                            }
                        ),
                        CreateParagraphProperties<D.Level1ParagraphProperties>(
                            "•",
                            fontSizes[0],
                            leftMargins[0],
                            indents[0]
                        ),
                        CreateParagraphProperties<D.Level2ParagraphProperties>(
                            "⁃",
                            fontSizes[1],
                            leftMargins[1],
                            indents[1]
                        ),
                        CreateParagraphProperties<D.Level3ParagraphProperties>(
                            "*",
                            fontSizes[2],
                            leftMargins[2],
                            indents[2]
                        ),
                        CreateParagraphProperties<D.Level4ParagraphProperties>(
                            "‣",
                            fontSizes[3],
                            leftMargins[3],
                            indents[3]
                        ),
                        CreateParagraphProperties<D.Level5ParagraphProperties>(
                            "○",
                            fontSizes[4],
                            leftMargins[4],
                            indents[4]
                        ),
                        CreateParagraphProperties<D.Level6ParagraphProperties>(
                            "○",
                            fontSizes[5],
                            leftMargins[5],
                            indents[5]
                        ),
                        CreateParagraphProperties<D.Level7ParagraphProperties>(
                            "○",
                            fontSizes[6],
                            leftMargins[6],
                            indents[6]
                        ),
                        CreateParagraphProperties<D.Level8ParagraphProperties>(
                            "○",
                            fontSizes[7],
                            leftMargins[7],
                            indents[7]
                        ),
                        CreateParagraphProperties<D.Level9ParagraphProperties>(
                            "○",
                            fontSizes[8],
                            leftMargins[8],
                            indents[8]
                        )
                    )
                )
            );

            slideMasterPart.SlideMaster = slideMaster;
        }

        return slideMaster;
    }

    /// <summary>Creates a color map with default color mappings.</summary>
    /// <returns>A new ColorMap instance.</returns>
    private static ColorMap CreateColorMap()
        => new ColorMap
        {
            Background1 = D.ColorSchemeIndexValues.Light1,
            Text1 = D.ColorSchemeIndexValues.Dark1,
            Background2 = D.ColorSchemeIndexValues.Light2,
            Text2 = D.ColorSchemeIndexValues.Dark2,
            Accent1 = D.ColorSchemeIndexValues.Accent1,
            Accent2 = D.ColorSchemeIndexValues.Accent2,
            Accent3 = D.ColorSchemeIndexValues.Accent3,
            Accent4 = D.ColorSchemeIndexValues.Accent4,
            Accent5 = D.ColorSchemeIndexValues.Accent5,
            Accent6 = D.ColorSchemeIndexValues.Accent6,
            Hyperlink = D.ColorSchemeIndexValues.Hyperlink,
            FollowedHyperlink = D.ColorSchemeIndexValues.FollowedHyperlink,
        };
}
