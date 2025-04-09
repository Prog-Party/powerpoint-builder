namespace PowerPoint.Builder.Core;

public class PartSize
{
    internal int Width { get; init; } = 0;
    internal int Height { get; init; } = 0;

    private PartSize(int width, int height)
    {
        Width = width;
        Height = height;
    }

    public static PartSize Construct(
        int? width = null,
        int? widthPercentage = null,
        int? height = null,
        int? heightPercentage = null
        )
    {
        return new PartSize(
            CalculateSize(width ?? Constants.SlideWidth, widthPercentage, Constants.SlideWidth),
            CalculateSize(height ?? Constants.SlideHeight, heightPercentage, Constants.SlideHeight)
        );
    }

    private static int CalculateSize(int size, int? sizePercentage, int maxSize)
    {
        if (sizePercentage.HasValue)
            return GetOffsetForPercentage(sizePercentage.Value, maxSize);

        return size;
    }

    private static int GetOffsetForPercentage(int sizePercentage, int maxSize)
    {
        if (sizePercentage < 0 || sizePercentage > 100)
            throw new ArgumentOutOfRangeException("Percentage value must be between 0 and 100.");

        return maxSize / 100 * sizePercentage;
    }
}