using DocumentFormat.OpenXml;

namespace PowerPoint.Builder.Slides;

public abstract class SlidePartBuilder
{
    internal abstract OpenXmlElement Build();
}