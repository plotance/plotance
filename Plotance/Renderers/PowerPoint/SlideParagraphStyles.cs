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

/// <summary>
/// Contains paragraph styles for different elements in a PowerPoint slide.
/// </summary>
/// <param name="Title">Paragraph styles for title shapes.</param>
/// <param name="Body">Paragraph styles for body shapes.</param>
public record SlideParagraphStyles(
    ParagraphStyles Title,
    ParagraphStyles Body
)
{
    /// <summary>
    /// Extracts paragraph styles from a slide part and its ancestors.
    /// </summary>
    /// <param name="slidePart">
    /// The slide part to extract paragraph styles from.
    /// </param>
    /// <returns>
    /// The combined paragraph styles with title and body styles extracted from
    /// the slide part and its ancestors.
    /// </returns>
    public static SlideParagraphStyles ExtractParagraphStyles(
        SlidePart slidePart
    )
    {
        ParagraphStyles DoExtract(
            Func<OpenXmlPart, Shape?> shapeExtractor,
            Func<TextStyles?, TextListStyleType?> textStyleExtractor,
            int defaultFontSize
        )
        {
            var textStyles = ExtractTextStylesHierarchy(
                slidePart,
                shapeExtractor,
                textStyleExtractor
            );

            T ExtractProperties<T>()
                where T : D.TextParagraphPropertiesType, new()
            {
                var result = new T();

                result.AddChild(
                    new D.DefaultRunProperties() { FontSize = defaultFontSize }
                );

                foreach (var textStyle in textStyles)
                {
                    MergeProperties(
                        result,
                        textStyle
                            ?.GetFirstChild<T>()
                            ?.CloneNode(true)
                            as T
                    );
                }

                return result;
            }

            return new ParagraphStyles(
                ExtractProperties<D.DefaultParagraphProperties>(),
                ExtractProperties<D.Level1ParagraphProperties>(),
                ExtractProperties<D.Level2ParagraphProperties>(),
                ExtractProperties<D.Level3ParagraphProperties>(),
                ExtractProperties<D.Level4ParagraphProperties>(),
                ExtractProperties<D.Level5ParagraphProperties>(),
                ExtractProperties<D.Level6ParagraphProperties>(),
                ExtractProperties<D.Level7ParagraphProperties>(),
                ExtractProperties<D.Level8ParagraphProperties>(),
                ExtractProperties<D.Level9ParagraphProperties>()
            );
        }

        return new SlideParagraphStyles(
            Title: DoExtract(
                Shapes.ExtractTitleShape,
                textStyles => textStyles?.TitleStyle,
                4400
            ),
            Body: DoExtract(
                Shapes.ExtractBodyShape,
                textStyles => textStyles?.BodyStyle,
                1800
            )
        );
    }

    /// <summary>
    /// Extracts a list of text styles from a slide part and its ancestors.
    /// </summary>
    /// <param name="slidePart">
    /// The slide part to extract text styles from.
    /// </param>
    /// <param name="shapeExtractor">
    /// A function that extracts a shape from a part to get text styles from.
    /// </param>
    /// <param name="textStyleExtractor">
    /// A function that extracts a text list style from text styles.
    /// </param>
    /// <returns>
    /// A list of text styles in order from lowest priority to highest priority.
    /// </returns>
    private static List<OpenXmlElement?> ExtractTextStylesHierarchy(
        SlidePart slidePart,
        Func<OpenXmlPart, Shape?> shapeExtractor,
        Func<TextStyles?, TextListStyleType?> textStyleExtractor
    )
    {
        D.ListStyle? ExtractListStyle(OpenXmlPart? part)
            => part == null
            ? null
            : shapeExtractor(part)?.TextBody?.ListStyle;

        var slideLayoutPart = slidePart.SlideLayoutPart;
        var slideMasterPart = slideLayoutPart?.SlideMasterPart;

        // Return in order from lowest priority to highest priority
        return [
            ((PresentationDocument)slidePart.OpenXmlPackage)
                .PresentationPart
                ?.Presentation
                ?.DefaultTextStyle,
            textStyleExtractor(slideMasterPart?.SlideMaster?.TextStyles),
            ExtractListStyle(slideMasterPart),
            ExtractListStyle(slideLayoutPart),
            ExtractListStyle(slidePart)
        ];
    }

    /// <summary>
    /// Copies a child element of given type from source to target if any.
    /// </summary>
    /// <typeparam name="T">The type of the child element.</typeparam>
    /// <param name="target">The target element to copy to.</param>
    /// <param name="source">The source element to copy from.</param>
    private static void CopyChildIfAny<T>(
        OpenXmlCompositeElement target,
        OpenXmlCompositeElement source
    )
        where T : OpenXmlElement
    {
        var child = source.GetFirstChild<T>();

        if (child != null)
        {
            target.AddChild(child.CloneNode(true));
        }
    }

    /// <summary>
    /// Merges a child element of given type from source to target using the
    /// specified merger function.
    /// </summary>
    /// <typeparam name="T">The type of the child element.</typeparam>
    /// <param name="target">The target element to merge to.</param>
    /// <param name="source">The source element to merge from.</param>
    /// <param name="merger">
    /// The function to use for merging the child elements. The function takes
    /// two arguments: the target child and the source child. The function
    /// should merge the source child into the target child.
    /// </param>
    private static void MergeChild<T>(
        OpenXmlCompositeElement target,
        OpenXmlCompositeElement source,
        Action<T, T> merger
    )
        where T : OpenXmlElement
    {
        var sourceChild = source.GetFirstChild<T>();

        if (sourceChild == null)
        {
            return;
        }

        var targetChild = target.GetFirstChild<T>();

        if (targetChild == null)
        {
            target.AddChild((T)sourceChild.CloneNode(true));

            return;
        }

        merger(targetChild, sourceChild);
    }

    /// <summary>
    /// Merges text paragraph properties of given type from source to target.
    /// </summary>
    /// <typeparam name="T">The type of text paragraph properties.</typeparam>
    /// <param name="target">The target properties to merge to.</param>
    /// <param name="source">The source properties to merge from.</param>
    private static void MergeProperties<T>(T target, T? source)
        where T : D.TextParagraphPropertiesType
    {
        if (source == null) return;

        // Attributes
        target.LeftMargin = source.LeftMargin ?? target.LeftMargin;
        target.RightMargin = source.RightMargin ?? target.RightMargin;
        target.Level = source.Level ?? target.Level;
        target.Indent = source.Indent ?? target.Indent;
        target.Alignment = source.Alignment ?? target.Alignment;
        target.DefaultTabSize = source.DefaultTabSize ?? target.DefaultTabSize;
        target.RightToLeft = source.RightToLeft ?? target.RightToLeft;
        target.EastAsianLineBreak = source.EastAsianLineBreak
            ?? target.EastAsianLineBreak;
        target.FontAlignment = source.FontAlignment ?? target.FontAlignment;
        target.LatinLineBreak = source.LatinLineBreak ?? target.LatinLineBreak;
        target.Height = source.Height ?? target.Height;

        // Children
        CopyChildIfAny<D.LineSpacing>(target, source);
        CopyChildIfAny<D.SpaceBefore>(target, source);
        CopyChildIfAny<D.SpaceAfter>(target, source);

        CopyChildIfAny<D.BulletColorText>(target, source);
        CopyChildIfAny<D.BulletColor>(target, source);

        CopyChildIfAny<D.BulletSizeText>(target, source);
        CopyChildIfAny<D.BulletSizePercentage>(target, source);
        CopyChildIfAny<D.BulletSizePoints>(target, source);

        CopyChildIfAny<D.NoBullet>(target, source);
        CopyChildIfAny<D.CharacterBullet>(target, source);
        CopyChildIfAny<D.AutoNumberedBullet>(target, source);
        CopyChildIfAny<D.PictureBullet>(target, source);

        CopyChildIfAny<D.TabStopList>(target, source);

        MergeRunProperties(target, source);
        MergeExtensionList<D.ExtensionList>(target, source);
    }

    /// <summary>Merges run properties from source to target.</summary>
    /// <param name="target">The target properties to merge to.</param>
    /// <param name="source">The source properties to merge from.</param>
    private static void MergeRunProperties(
        D.TextParagraphPropertiesType target,
        D.TextParagraphPropertiesType source
    )
    {
        MergeChild<D.DefaultRunProperties>(target, source, (target, source) =>
        {
            // Attributes
            target.Kumimoji = source.Kumimoji ?? target.Kumimoji;
            target.Language = source.Language ?? target.Language;
            target.AlternativeLanguage = source.AlternativeLanguage
                ?? target.AlternativeLanguage;
            target.FontSize = source.FontSize ?? target.FontSize;
            target.Bold = source.Bold ?? target.Bold;
            target.Italic = source.Italic ?? target.Italic;
            target.Underline = source.Underline ?? target.Underline;
            target.Strike = source.Strike ?? target.Strike;
            target.Kerning = source.Kerning ?? target.Kerning;
            target.Capital = source.Capital ?? target.Capital;
            target.Spacing = source.Spacing ?? target.Spacing;
            target.NormalizeHeight = source.NormalizeHeight
                ?? target.NormalizeHeight;
            target.Baseline = source.Baseline ?? target.Baseline;
            target.NoProof = source.NoProof ?? target.NoProof;
            target.Dirty = source.Dirty ?? target.Dirty;
            target.SpellingError = source.SpellingError ?? target.SpellingError;
            target.SmartTagClean = source.SmartTagClean ?? target.SmartTagClean;
            target.SmartTagId = source.SmartTagId ?? target.SmartTagId;
            target.Bookmark = source.Bookmark ?? target.Bookmark;

            // Children
            MergeLinePropertiesType<D.Outline>(target, source);

            CopyChildIfAny<D.NoFill>(target, source);
            CopyChildIfAny<D.SolidFill>(target, source);
            CopyChildIfAny<D.GradientFill>(target, source);
            CopyChildIfAny<D.BlipFill>(target, source);
            CopyChildIfAny<D.PatternFill>(target, source);
            CopyChildIfAny<D.GroupFill>(target, source);

            CopyChildIfAny<D.EffectList>(target, source);
            CopyChildIfAny<D.EffectDag>(target, source);

            CopyChildIfAny<D.Highlight>(target, source);

            CopyChildIfAny<D.UnderlineFollowsText>(target, source);
            MergeLinePropertiesType<D.Underline>(target, source);

            CopyChildIfAny<D.UnderlineFillText>(target, source);
            CopyChildIfAny<D.UnderlineFill>(target, source);

            CopyChildIfAny<D.LatinFont>(target, source);
            CopyChildIfAny<D.EastAsianFont>(target, source);
            CopyChildIfAny<D.ComplexScriptFont>(target, source);
            CopyChildIfAny<D.SymbolFont>(target, source);

            MergeHyperlinkType<D.HyperlinkOnClick>(target, source);
            MergeHyperlinkType<D.HyperlinkOnMouseOver>(target, source);

            CopyChildIfAny<D.RightToLeft>(target, source);

            MergeExtensionList<D.ExtensionList>(target, source);
        });
    }

    /// <summary>
    /// Merges line properties of given type from source to target.
    /// </summary>
    /// <param name="target">The target properties to merge to.</param>
    /// <param name="source">The source properties to merge from.</param>
    private static void MergeLinePropertiesType<T>(
        OpenXmlCompositeElement target,
        OpenXmlCompositeElement source
    )
        where T : D.LinePropertiesType
    {
        MergeChild<T>(target, source, (target, source) =>
        {
            CopyChildIfAny<D.NoFill>(target, source);
            CopyChildIfAny<D.SolidFill>(target, source);
            CopyChildIfAny<D.GradientFill>(target, source);
            CopyChildIfAny<D.PatternFill>(target, source);

            CopyChildIfAny<D.PresetDash>(target, source);
            CopyChildIfAny<D.CustomDash>(target, source);

            CopyChildIfAny<D.Round>(target, source);
            CopyChildIfAny<D.LineJoinBevel>(target, source);
            CopyChildIfAny<D.Miter>(target, source);

            MergeLineEndPropertiesType<D.HeadEnd>(target, source);
            MergeLineEndPropertiesType<D.TailEnd>(target, source);

            MergeExtensionList<D.LinePropertiesExtensionList>(target, source);
        });
    }

    /// <summary>
    /// Merges line end properties of given type from source to target.
    /// </summary>
    /// <param name="target">The target properties to merge to.</param>
    /// <param name="source">The source properties to merge from.</param>
    private static void MergeLineEndPropertiesType<T>(
        OpenXmlCompositeElement target,
        OpenXmlCompositeElement source
    )
        where T : D.LineEndPropertiesType
    {
        MergeChild<T>(target, source, (target, source) =>
        {
            target.Type = source.Type ?? target.Type;
            target.Width = source.Width ?? target.Width;
            target.Length = source.Length ?? target.Length;
        });
    }

    /// <summary>
    /// Merges hyperlink properties of given type from source to target.
    /// </summary>
    /// <param name="target">The target properties to merge to.</param>
    /// <param name="source">The source properties to merge from.</param>
    private static void MergeHyperlinkType<T>(
        OpenXmlCompositeElement target,
        OpenXmlCompositeElement source
    )
        where T : D.HyperlinkType
    {
        MergeChild<T>(target, source, (target, source) =>
        {
            CopyChildIfAny<D.HyperlinkSound>(target, source);
            MergeExtensionList<D.HyperlinkExtensionList>(target, source);
        });
    }

    /// <summary>
    /// Merges extension list of given type from source to target.
    /// </summary>
    /// <param name="target">The target properties to merge to.</param>
    /// <param name="source">The source properties to merge from.</param>
    private static void MergeExtensionList<T>(
        OpenXmlCompositeElement target,
        OpenXmlCompositeElement source
    )
        where T : OpenXmlCompositeElement
    {
        MergeChild<T>(target, source, (target, source) =>
        {
            foreach (var extension in source)
            {
                target.AppendChild(extension.CloneNode(true));
            }
        });
    }
}

/// <summary>
/// Contains a set of paragraph properties for different heading levels in a
/// PowerPoint slide.
/// </summary>
/// <param name="Default">Default paragraph properties.</param>
/// <param name="Level1">Level 1 paragraph properties.</param>
/// <param name="Level2">Level 2 paragraph properties.</param>
/// <param name="Level3">Level 3 paragraph properties.</param>
/// <param name="Level4">Level 4 paragraph properties.</param>
/// <param name="Level5">Level 5 paragraph properties.</param>
/// <param name="Level6">Level 6 paragraph properties.</param>
/// <param name="Level7">Level 7 paragraph properties.</param>
/// <param name="Level8">Level 8 paragraph properties.</param>
/// <param name="Level9">Level 9 paragraph properties.</param>
public record ParagraphStyles(
    D.DefaultParagraphProperties Default,
    D.Level1ParagraphProperties Level1,
    D.Level2ParagraphProperties Level2,
    D.Level3ParagraphProperties Level3,
    D.Level4ParagraphProperties Level4,
    D.Level5ParagraphProperties Level5,
    D.Level6ParagraphProperties Level6,
    D.Level7ParagraphProperties Level7,
    D.Level8ParagraphProperties Level8,
    D.Level9ParagraphProperties Level9
)
{
    /// <summary>
    /// Creates a new paragraph styles with all font sizes and margins scaled by
    /// the specified factor.
    /// </summary>
    /// <param name="scale">The scale factor with source location.</param>
    /// <returns>
    /// A new paragraph styles with scaled font sizes and margins.
    /// </returns>
    public ParagraphStyles Scaled(ValueWithLocation<decimal>? scale)
        => Scaled(scale?.Value ?? 1m);

    /// <summary>
    /// Creates a new paragraph styles with all font sizes and margins scaled by
    /// the specified factor.
    /// </summary>
    /// <param name="scale">The scale factor.</param>
    /// <returns>
    /// A new paragraph styles with scaled font sizes and margins.
    /// </returns>
    public ParagraphStyles Scaled(decimal scale)
        => scale == 1m
        ? this
        : new ParagraphStyles(
            Scaled(Default, scale),
            Scaled(Level1, scale),
            Scaled(Level2, scale),
            Scaled(Level3, scale),
            Scaled(Level4, scale),
            Scaled(Level5, scale),
            Scaled(Level6, scale),
            Scaled(Level7, scale),
            Scaled(Level8, scale),
            Scaled(Level9, scale)
        );

    /// <summary>
    /// Scales font sizes and margins in paragraph properties by the specified
    /// factor.
    /// </summary>
    /// <typeparam name="T">The type of paragraph properties.</typeparam>
    /// <param name="paragraphProperties">
    /// The paragraph properties to scale.
    /// </param>
    /// <param name="scale">The scale factor.</param>
    /// <returns>
    /// A new paragraph properties with scaled font sizes and margins.
    /// </returns>
    private T Scaled<T>(T paragraphProperties, decimal scale)
        where T : D.TextParagraphPropertiesType
    {
        var cloned = (T)paragraphProperties.CloneNode(true);
        var defaultRunProperties = cloned
            .GetFirstChild<D.DefaultRunProperties>();

        if (defaultRunProperties == null)
        {
            defaultRunProperties = new D.DefaultRunProperties();
            cloned.AddChild(defaultRunProperties);
        }

        int fontSize = defaultRunProperties.FontSize ?? 1800;

        defaultRunProperties.FontSize = (int)(fontSize * scale);

        cloned.LeftMargin = (int)((cloned.LeftMargin ?? 0) * scale);
        cloned.RightMargin = (int)((cloned.RightMargin ?? 0) * scale);
        cloned.Indent = (int)((cloned.Indent ?? 0) * scale);

        return cloned;
    }

    /// <summary>
    /// Creates a new paragraph styles with the specified language.
    /// </summary>
    /// <param name="language">
    /// The language with source location to set for all text.
    /// </param>
    /// <returns>
    /// A new paragraph styles with the specified language.
    /// </returns>
    public ParagraphStyles WithLanguage(ValueWithLocation<string>? language)
        => language == null ? this : WithLanguage(language.Value);

    /// <summary>
    /// Creates a new paragraph styles with the specified language.
    /// </summary>
    /// <param name="language">The language to set for all text.</param>
    /// <returns>
    /// A new paragraph styles with the specified language.
    /// </returns>
    public ParagraphStyles WithLanguage(string language)
        => new ParagraphStyles(
            WithLanguage(Default, language),
            WithLanguage(Level1, language),
            WithLanguage(Level2, language),
            WithLanguage(Level3, language),
            WithLanguage(Level4, language),
            WithLanguage(Level5, language),
            WithLanguage(Level6, language),
            WithLanguage(Level7, language),
            WithLanguage(Level8, language),
            WithLanguage(Level9, language)
        );

    /// <summary>
    /// Creates a new cloned paragraph properties with the specified language.
    /// </summary>
    /// <typeparam name="T">The type of paragraph properties.</typeparam>
    /// <param name="paragraphProperties">
    /// The paragraph properties to clone.
    /// </param>
    /// <param name="language">The language to set.</param>
    /// <returns>
    /// A new cloned paragraph properties with the specified language.
    /// </returns>
    private T WithLanguage<T>(T paragraphProperties, string language)
        where T : D.TextParagraphPropertiesType, new()
    {
        var cloned = (T)paragraphProperties.CloneNode(true);
        var defaultRunProperties = cloned
            .GetFirstChild<D.DefaultRunProperties>();

        if (defaultRunProperties == null)
        {
            defaultRunProperties = new D.DefaultRunProperties();
            cloned.AddChild(defaultRunProperties);
        }

        defaultRunProperties.Language = language;
        defaultRunProperties.AlternativeLanguage = "en-US";

        return cloned;
    }

    /// <summary>
    /// Creates a new paragraph styles with the specified text alignment.
    /// </summary>
    /// <param name="alignment">The text alignment to set.</param>
    /// <returns>
    /// A new paragraph styles with the specified text alignment.
    /// </returns>
    public ParagraphStyles WithAlignment(D.TextAlignmentTypeValues alignment)
        => new ParagraphStyles(
            WithAlignment(Default, alignment),
            WithAlignment(Level1, alignment),
            WithAlignment(Level2, alignment),
            WithAlignment(Level3, alignment),
            WithAlignment(Level4, alignment),
            WithAlignment(Level5, alignment),
            WithAlignment(Level6, alignment),
            WithAlignment(Level7, alignment),
            WithAlignment(Level8, alignment),
            WithAlignment(Level9, alignment)
        );

    /// <summary>
    /// Creates a new cloned paragraph properties with the specified text
    /// alignment.
    /// </summary>
    /// <typeparam name="T">The type of paragraph properties.</typeparam>
    /// <param name="paragraphProperties">
    /// The paragraph properties to clone.
    /// </param>
    /// <param name="alignment">The text alignment to set.</param>
    /// <returns>
    /// A new cloned paragraph properties with the specified text alignment.
    /// </returns>
    private T WithAlignment<T>(
        T paragraphProperties,
        D.TextAlignmentTypeValues alignment
    )
        where T : D.TextParagraphPropertiesType, new()
    {
        var cloned = (T)paragraphProperties.CloneNode(true);

        cloned.Alignment = alignment;

        return cloned;
    }

    /// <summary>
    /// Creates a new paragraph styles with only font sizes, removing other
    /// styling.
    /// </summary>
    /// <returns>
    /// A new paragraph styles with only font sizes.
    /// </returns>
    public ParagraphStyles OnlySizes()
        => new ParagraphStyles(
            OnlySizes(Default),
            OnlySizes(Level1),
            OnlySizes(Level2),
            OnlySizes(Level3),
            OnlySizes(Level4),
            OnlySizes(Level5),
            OnlySizes(Level6),
            OnlySizes(Level7),
            OnlySizes(Level8),
            OnlySizes(Level9)
        );

    /// <summary>
    /// Creates a copy of paragraph properties with only font sizes, removing
    /// other styling.
    /// </summary>
    /// <typeparam name="T">The type of paragraph properties.</typeparam>
    /// <param name="paragraphProperties">
    /// The paragraph properties to extract font sizes from.
    /// </param>
    /// <returns>
    /// A new paragraph properties with only font sizes.
    /// </returns>
    private T OnlySizes<T>(T paragraphProperties)
        where T : D.TextParagraphPropertiesType, new()
    {
        var newParagraphProperties = new T();
        var defaultRunProperties = paragraphProperties
            .GetFirstChild<D.DefaultRunProperties>();

        if (defaultRunProperties != null)
        {
            newParagraphProperties.AddChild(
                new D.DefaultRunProperties()
                {
                    FontSize = defaultRunProperties.FontSize
                }
            );
        }

        newParagraphProperties.LeftMargin = paragraphProperties.LeftMargin;
        newParagraphProperties.RightMargin = paragraphProperties.RightMargin;
        newParagraphProperties.Indent = paragraphProperties.Indent;

        return newParagraphProperties;
    }

    /// <summary>
    /// Converts these paragraph styles to an OpenXml ListStyle element.
    /// </summary>
    /// <returns>A ListStyle element for use in slide shapes.</returns>
    public D.ListStyle ToListStyle()
        => new D.ListStyle(
            Default.CloneNode(true),
            Level1.CloneNode(true),
            Level2.CloneNode(true),
            Level3.CloneNode(true),
            Level4.CloneNode(true),
            Level5.CloneNode(true),
            Level6.CloneNode(true),
            Level7.CloneNode(true),
            Level8.CloneNode(true),
            Level9.CloneNode(true)
        );
}
