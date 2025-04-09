using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace PowerPoint.Builder.Slides;

public abstract class SlidePartBuilder
{
    internal abstract void Build(SlidePart slidePart, ShapeTree tree);
}