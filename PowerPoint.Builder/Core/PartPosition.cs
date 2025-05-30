﻿namespace PowerPoint.Builder.Core;

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
            GetOffset(x, xPercentage, Constants.SlideWidth),
            GetOffset(y, yPercentage, Constants.SlideHeight)
        );
    }

    private static int GetOffset(int offset, int? OffsetPercentage, int maxSize)
    {
        if (OffsetPercentage.HasValue)
            return GetOffsetForPercentage(OffsetPercentage.Value, maxSize);

        return offset;
    }

    private static int GetOffsetForPercentage(int OffsetPercentage, int maxSize)
    {
        if (OffsetPercentage < 0 || OffsetPercentage > 100)
            throw new ArgumentOutOfRangeException("Percentage Value must be between 0 and 100.");

        return maxSize / 100 * OffsetPercentage;
    }
}