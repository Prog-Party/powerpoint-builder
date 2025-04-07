using PowerPoint.Builder.PowerPointSlide;

namespace PowerPoint.Builder;

internal record BuilderProperties
{
    internal BuilderPropertiesSource Source { get; set; }
    internal List<SlideBuilder> Slides { get; set; } = new();

    internal BuilderProperties(string? filePath = null, Stream? stream = null, System.IO.Packaging.Package? package = null)
    {
        Source = new BuilderPropertiesSource(filePath, stream, package);
    }
}