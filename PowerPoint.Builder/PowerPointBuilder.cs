using PowerPoint.Builder.Slides;
using PowerPoint.Builder.Presentation;

namespace PowerPoint.Builder;

public class PowerPointBuilder
{
    private BuilderProperties _properties;

    public PowerPointBuilder(string? filePath = null, Stream? stream = null)
    {
        _properties = new BuilderProperties(filePath);
    }

    public PowerPointBuilder AddSlide(Action<SlideBuilder> slideAction)
    {
        var slideBuilder = new SlideBuilder();
        slideAction(slideBuilder);
        _properties.Slides.Add(slideBuilder);
        return this;
    }

    public void Build()
    {
        new PresentationUtility(_properties).Build();
    }
}