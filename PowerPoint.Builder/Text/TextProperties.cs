namespace PowerPoint.Builder.Text;

internal class TextProperties
{
    internal int XOffset { get; init; } = 0;
    internal int YOffset { get; init; } = 0;
    internal int Width { get; init; } = 0;
    internal int Height { get; init; } = 0;
    internal string Text { get; init; } = "";

    internal TextProperties(
        string text,
        int xOffset = 0,
        int? xOffsetPercentage = null,
        int yOffset = 0,
        int? yOffsetPercentage = null,
        int? width = null,
        int? widthPercentage = null,
        int? height = null,
        int? heightPercentage = null
        )
    {
        XOffset = GetOffset(xOffset, xOffsetPercentage);
        YOffset = GetOffset(yOffset, yOffsetPercentage);

        Width = GetOffset(width ?? Constants.SlideWidth, widthPercentage);
        Height = GetOffset(height ?? Constants.SlideHeight, heightPercentage);

        Text = text;
    }

    private int GetOffset(int offset, int? OffsetPercentage)
    {
        if (OffsetPercentage is not null)
        {
            if (OffsetPercentage < 0 || OffsetPercentage > 100)
            {
                throw new ArgumentOutOfRangeException("xOffsetPercentage", "Value must be between 0 and 100.");
            }

            return Constants.SlideWidth / 100 * OffsetPercentage.Value;
        }

        return offset;
    }
}