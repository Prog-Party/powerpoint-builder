using DocumentFormat.OpenXml.Packaging;

namespace PowerPoint.Builder.Utility;

internal static class ImageUtility
{
    internal static PartTypeInfo GetImagePartTypeByPath(string extension)
    {
        if (string.IsNullOrEmpty(extension))
            throw new ArgumentException("Path cannot be null or empty");

        return extension.ToLower() switch
        {
            ".bmp" => ImagePartType.Bmp,
            ".emf" => ImagePartType.Emf,
            ".gif" => ImagePartType.Gif,
            ".ico" => ImagePartType.Icon,
            ".jpg" => ImagePartType.Jpeg,
            ".jpeg" => ImagePartType.Jpeg,
            ".pcx" => ImagePartType.Pcx,
            ".png" => ImagePartType.Png,
            ".tiff" => ImagePartType.Tiff,
            ".wmf" => ImagePartType.Wmf,
            _ => throw new ArgumentException("Unsupported image format")
        };
    }
}