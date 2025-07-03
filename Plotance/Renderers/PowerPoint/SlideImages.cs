// SPDX-FileCopyrightText: 2025 Plotance contributors <https://plotance.github.io/>
//
// SPDX-License-Identifier: MIT

using System.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using Plotance.Image;
using Plotance.Models;

using D = DocumentFormat.OpenXml.Drawing;

namespace Plotance.Renderers.PowerPoint;

/// <summary>Provides methods for handling images.</summary>
public static class SlideImages
{
    /// <summary>Extracts an image link from a block.</summary>
    /// <remarks>
    /// The image link is a image object that is:
    /// <list type="bullet">
    ///   <item>
    ///     the sole child of a paragraph block, or
    ///   </item>
    ///   <item>
    ///     the sole child of a link that is the sole child of a paragraph
    ///     block.
    ///   </item>
    /// </list>
    /// </remarks>
    /// <param name="variables">
    /// The variables to use for variable expansion.
    /// </param>
    /// <param name="block">The block to extract the image link from.</param>
    /// <returns>
    /// The extracted image link, or null if no image link is found.
    /// </returns>
    public static ImageLink? ExtractImageLink(
        IReadOnlyDictionary<string, string> variables,
        Block block
    )
    {
        if (block is ParagraphBlock paragraphBlock)
        {
            var paragraphInline = paragraphBlock.Inline;

            if (
                paragraphInline?.Count() == 1
                    && paragraphInline.FirstChild is LinkInline link
                    && link.IsImage
            )
            {
                return new ImageLink(
                    PowerPointRenderer.ExtractPath(block),
                    link.Line,
                    link.Url,
                    Variables.ExpandVariables(link.Title, variables)
                        ?? Paragraphs.ToPlainText(variables, link),
                    null
                );
            }
            else if (
                paragraphInline?.Count() == 1
                    && paragraphInline.FirstChild is LinkInline parentLink
                    && !parentLink.IsImage
                    && parentLink.Count() == 1
                    && parentLink.FirstChild is LinkInline childLink
                    && childLink.IsImage
            )
            {
                return new ImageLink(
                    PowerPointRenderer.ExtractPath(block),
                    childLink.Line,
                    childLink.Url,
                    Variables.ExpandVariables(childLink.Title, variables)
                        ?? Paragraphs.ToPlainText(variables, childLink),
                    parentLink.Url
                );
            }
            else
            {
                return null;
            }
        }
        else
        {
            return null;
        }
    }

    /// <summary>
    /// Embeds an image into a slide. The image is fetched from the URL and
    /// embedded into the presentation. Only http, https, and file URLs are
    /// supported. Shape size is adjusted to preserve aspect ratio of the image.
    /// </summary>
    /// <param name="slidePart">The slide part to embed the image into.</param>
    /// <param name="imageLink">The image link to embed.</param>
    /// <param name="x">The X coordinate of the image.</param>
    /// <param name="y">The Y coordinate of the image.</param>
    /// <param name="width">The width of the image.</param>
    /// <param name="height">The height of the image.</param>
    /// <param name="isFirst">
    /// Whether the image is the first image in the slide.
    /// </param>
    /// <param name="horizontalAlignment">
    /// The horizontal alignment of the image.
    /// </param>
    /// <param name="verticalAlignment">
    /// The vertical alignment of the image.
    /// </param>
    /// <exception cref="PlotanceException">
    /// Thrown if the URL is invalid, the image format is not supported, or the
    /// image data is invalid.
    /// </exception>
    public static void EmbedImage(
        SlidePart slidePart,
        ImageLink imageLink,
        long x,
        long y,
        long width,
        long height,
        bool isFirst,
        D.TextAlignmentTypeValues horizontalAlignment,
        D.TextAnchoringTypeValues verticalAlignment
    )
    {
        if (imageLink.ImageUrl == null)
        {
            return;
        }

        var imagePart = CreateImagePart(slidePart, imageLink);
        long adjustedWidth;
        long adjustedHeight;

        // Adjust shape size to preserve aspect ratio of the image.
        try
        {
            var imageSize = ExtractImageSize(imagePart);
            var imageWidth = imageSize.Width;
            var imageHeight = imageSize.Height;

            if (imageWidth * height < width * imageHeight)
            {
                adjustedWidth = height * imageWidth / imageHeight;
                adjustedHeight = height;
            }
            else
            {
                adjustedWidth = width;
                adjustedHeight = width * imageHeight / imageWidth;
            }
        }
        catch (ArgumentException e)
        {
            throw new PlotanceException(
                imageLink.Path,
                imageLink.Line,
                e.Message,
                e
            );
        }

        if (horizontalAlignment == D.TextAlignmentTypeValues.Center)
        {
            x = x + (width - adjustedWidth) / 2;
        }
        else if (horizontalAlignment == D.TextAlignmentTypeValues.Right)
        {
            x = x + width - adjustedWidth;
        }

        if (verticalAlignment == D.TextAnchoringTypeValues.Center)
        {
            y = y + (height - adjustedHeight) / 2;
        }
        else if (verticalAlignment == D.TextAnchoringTypeValues.Bottom)
        {
            y = y + height - adjustedHeight;
        }

        var picture = CreatePictureElement(
            slidePart,
            slidePart.GetIdOfPart(imagePart),
            imageLink.AlternativeText,
            imageLink.LinkUrl,
            x,
            y,
            adjustedWidth,
            adjustedHeight,
            isFirst
        );

        slidePart
            .Slide
            ?.CommonSlideData
            ?.ShapeTree
            ?.AppendChild(picture);
    }

    /// <summary>
    /// Creates an image part from an image link. Fetches images from the URL.
    /// </summary>
    /// <param name="slidePart">
    /// The slide part to create the image part in.
    /// </param>
    /// <param name="imageLink">The image link.</param>
    /// <returns>The created image part.</returns>
    /// <exception cref="PlotanceException">
    /// Thrown if the URL is invalid or the image format is not supported.
    /// </exception>
    private static ImagePart CreateImagePart(
        SlidePart slidePart,
        ImageLink imageLink
    )
    {
        var imageUrl = imageLink.ImageUrl;

        if (!Uri.TryCreate(imageUrl, UriKind.RelativeOrAbsolute, out var uri))
        {
            throw new PlotanceException(
                imageLink.Path,
                imageLink.Line,
                $"Invalid URL: {imageUrl}"
            );
        }

        try
        {
            if (
                uri.IsAbsoluteUri && (
                    uri.Scheme == Uri.UriSchemeHttp
                    || uri.Scheme == Uri.UriSchemeHttps
                )
            )
            {
                using var httpClient = new HttpClient();
                using var response = httpClient.GetAsync(uri).Result;

                response.EnsureSuccessStatusCode();

                var contentType = response
                    .Content
                    .Headers.ContentType
                    ?.MediaType;
                var imagePartType = DetermineImagePartType(
                    imageUrl,
                    contentType
                );
                var imagePart = slidePart.AddImagePart(imagePartType);
                using var stream = response.Content.ReadAsStreamAsync().Result;

                imagePart.FeedData(stream);

                return imagePart;
            }
            else if (uri.IsAbsoluteUri && uri.IsFile || !uri.IsAbsoluteUri)
            {
                var path = uri.IsAbsoluteUri
                    ? uri.LocalPath
                    : Uri.UnescapeDataString(imageUrl);
                var imagePartType = DetermineImagePartTypeFromExtension(path);
                var imagePart = slidePart.AddImagePart(imagePartType);
                using var stream = File.OpenRead(path);

                imagePart.FeedData(stream);

                return imagePart;
            }
            else
            {
                throw new PlotanceException(
                    imageLink.Path,
                    imageLink.Line,
                    $"Unsupported URL scheme: {uri.Scheme}"
                );
            }
        }
        catch (HttpRequestException e)
        {
            throw new PlotanceException(
                imageLink.Path,
                imageLink.Line,
                $"Cannot get image: {imageUrl}",
                e
            );
        }
        catch (IOException e)
        {
            throw new PlotanceException(
                imageLink.Path,
                imageLink.Line,
                $"Cannot read image: {imageUrl}",
                e
            );
        }
        catch (ArgumentException e)
        {
            throw new PlotanceException(
                imageLink.Path,
                imageLink.Line,
                e.Message,
                e
            );
        }
    }

    /// <summary>
    /// Determines the image part type from the Content-Type or the file
    /// extension.
    /// </summary>
    /// <param name="imageUrl">The URL of the image.</param>
    /// <param name="contentType">The content type of the image.</param>
    /// <returns>The determined image part type.</returns>
    /// <exception cref="ArgumentException">
    /// Thrown if the content type is not recognized.
    /// </exception>
    private static PartTypeInfo DetermineImagePartType(
        string imageUrl,
        string? contentType
    )
    {
        return contentType?.ToLowerInvariant() switch
        {
            "image/bmp" => ImagePartType.Bmp,
            "image/gif" => ImagePartType.Gif,
            "image/png" => ImagePartType.Png,
            "image/jp2" => ImagePartType.Jp2,
            "image/tiff" => ImagePartType.Tiff,
            "image/x-icon" => ImagePartType.Icon,
            "image/x-pcx" => ImagePartType.Pcx,
            "image/jpeg" => ImagePartType.Jpeg,
            "image/svg+xml" => ImagePartType.Svg,
            "image/emf" => ImagePartType.Emf,
            "image/wmf" => ImagePartType.Wmf,
            _ => DetermineImagePartTypeFromExtension(imageUrl)
        };
    }

    /// <summary>
    /// Determines the image part type from the file extension.
    /// </summary>
    /// <param name="path">The path of the image.</param>
    /// <returns>The determined image part type.</returns>
    /// <exception cref="ArgumentException">
    /// Thrown if the file extension is not recognized.
    /// </exception>
    private static PartTypeInfo DetermineImagePartTypeFromExtension(string path)
    {
        var extension = Path.GetExtension(path).ToLowerInvariant();

        return extension switch
        {
            ".bmp" => ImagePartType.Bmp,
            ".gif" => ImagePartType.Gif,
            ".png" => ImagePartType.Png,
            ".jp2" => ImagePartType.Jp2,
            ".tif" => ImagePartType.Tif,
            ".tiff" => ImagePartType.Tiff,
            ".ico" => ImagePartType.Icon,
            ".pcx" => ImagePartType.Pcx,
            ".jpg" => ImagePartType.Jpeg,
            ".jpeg" => ImagePartType.Jpeg,
            ".emf" => ImagePartType.Emf,
            ".wmf" => ImagePartType.Wmf,
            ".svg" => ImagePartType.Svg,
            _ => throw new ArgumentException(
                $"Unknown image extension: {extension}",
                nameof(path)
            )
        };
    }

    /// <summary>Extracts the size of an image.</summary>
    /// <param name="imagePart">The image part to extract the size from.</param>
    /// <returns>The size of the image.</returns>
    /// <exception cref="ArgumentException">
    /// Thrown if the image format is not supported.
    /// </exception>
    /// <exception cref="InvalidDataException">
    /// Thrown if the image data is invalid.
    /// </exception>
    private static Size ExtractImageSize(ImagePart imagePart)
    {
        using var stream = imagePart.GetStream();
        var contentType = imagePart.ContentType;

        return contentType switch
        {
            _ when contentType == ImagePartType.Png.ContentType
                => Images.GetPngImageSize(stream),
            _ when contentType == ImagePartType.Jpeg.ContentType
                => Images.GetJpegImageSize(stream),
            _ when contentType == ImagePartType.Svg.ContentType
                => Images.GetSvgImageSize(stream),
            _ => throw new ArgumentException(
                $"Unsupported image format: {contentType}",
                nameof(imagePart)
            )
        };
    }

    /// <summary>Creates a picture element.</summary>
    /// <param name="slidePart">
    /// The slide part to create the picture element in.
    /// </param>
    /// <param name="relationshipId">
    /// The relationship ID of the image part.
    /// </param>
    /// <param name="alternativeText">
    /// The alternative text of the picture.
    /// </param>
    /// <param name="linkUrl">The URL to link the picture to.</param>
    /// <param name="x">The X coordinate of the picture.</param>
    /// <param name="y">The Y coordinate of the picture.</param>
    /// <param name="width">The width of the picture.</param>
    /// <param name="height">The height of the picture.</param>
    /// <param name="isFirst">
    /// Whether the picture is the first picture in the slide.
    /// </param>
    /// <returns>The created picture element.</returns>
    private static Picture CreatePictureElement(
        SlidePart slidePart,
        string relationshipId,
        string? alternativeText,
        string? linkUrl,
        long x,
        long y,
        long width,
        long height,
        bool isFirst
    )
    {
        var nonVisualDrawingProperties = new NonVisualDrawingProperties()
        {
            Id = (UInt32Value)1U,
            Name = "",
            Description = alternativeText
        };

        if (!string.IsNullOrEmpty(linkUrl))
        {
            nonVisualDrawingProperties.AddChild(
                Paragraphs.CreateHyperlinkOnClick(
                    slidePart,
                    linkUrl,
                    null
                )
            );
        }

        return new Picture(
            new NonVisualPictureProperties(
                nonVisualDrawingProperties,
                new NonVisualPictureDrawingProperties(
                    new D.PictureLocks() { NoChangeAspect = true }
                )
                {
                    PreferRelativeResize = true
                },
                new ApplicationNonVisualDrawingProperties(
                    new PlaceholderShape()
                    {
                        Type = PlaceholderValues.Picture,
                        Index = isFirst ? 1 : null
                    }
                )
            ),
            new BlipFill(
                new D.Blip() { Embed = relationshipId },
                new D.Stretch(new D.FillRectangle())
            ),
            new ShapeProperties(
                new D.Transform2D(
                    new D.Offset() { X = x, Y = y },
                    new D.Extents() { Cx = width, Cy = height }
                ),
                new D.PresetGeometry(new D.AdjustValueList())
                {
                    Preset = D.ShapeTypeValues.Rectangle
                }
            )
        );
    }
}

/// <summary>Represents an image link in a Markdown.</summary>
/// <param name="Path">
/// The path to the file containing the value, for error reporting.
/// </param>
/// <param name="Line">The line number in the file, for error reporting.</param>
/// <param name="ImageUrl">The URL of the image.</param>
/// <param name="AlternativeText">The alternative text of the image.</param>
/// <param name="LinkUrl">The URL to link the image to.</param>
public record ImageLink(
    string? Path,
    long Line,
    string? ImageUrl,
    string? AlternativeText,
    string? LinkUrl
);
