using PowerPoint.Builder.Slides;
using PowerPoint.Builder.Presentation;
using PowerPoint.Builder.Template;
using DocumentFormat.OpenXml.Presentation;

namespace PowerPoint.Builder;

public class PowerPointBuilder
{
    private BuilderProperties _properties;

    public PowerPointBuilder(string? filePath = null, Stream? stream = null)
    {
        _properties = new BuilderProperties(filePath);
    }

    public PowerPointBuilder AddSlide(Action<SlideBuilder>? slideAction = null, TemplateLayoutBuilder? layout = null, Slide? slide = null)
    {
        var slideBuilder = new SlideBuilder(slide, layout);
        slideAction?.Invoke(slideBuilder);
        _properties.Slides.Add(slideBuilder);
        return this;
    }

    public void Build()
    {
        new PresentationUtility(_properties).Build();
    }
}