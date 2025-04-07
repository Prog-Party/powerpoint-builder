namespace PowerPoint.Builder.Core;

public class PartPosition
{
    internal int XOffset { get; init; } = 0;
    internal int YOffset { get; init; } = 0;

    private PartPosition(int xOffset, int yOffset)
    {
        XOffset = xOffset;
        YOffset = yOffset;
    }

    public static PartPosition Construct(
        int x = 0,
        int? xPercentage = null,
        int y = 0,
        int? yPercentage = null
        )
    {
        return new PartPosition(
            GetOffset(x, xPercentage),
            GetOffset(y, yPercentage)
        );
    }

    private static int GetOffset(int offset, int? OffsetPercentage)
    {
        if (OffsetPercentage is not null)
        {
            if (OffsetPercentage < 0 || OffsetPercentage > 100)
            {
                throw new ArgumentOutOfRangeException("Percentage Value must be between 0 and 100.");
            }
            return Constants.SlideWidth / 100 * OffsetPercentage.Value;
        }
        return offset;
    }
}