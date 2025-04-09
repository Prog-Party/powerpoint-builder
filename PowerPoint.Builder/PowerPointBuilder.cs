using DocumentFormat.OpenXml.Presentation;
using PowerPoint.Builder.Presentation;
using PowerPoint.Builder.Slides;
using PowerPoint.Builder.Template;

namespace PowerPoint.Builder;

/// <summary>
/// Provides functionality to build PowerPoint presentations programmatically.
/// For more details, see the
/// <see href="https://github.com/Prog-Party/powerpoint-builder/wiki/PowerPointBuilder-Class-Documentation">
/// PowerPointBuilder Class Documentation
/// </see>.
/// </summary>
public class PowerPointBuilder
{
    private BuilderProperties _properties;

    /// <summary>
    /// Initializes a new instance of the <see cref="PowerPointBuilder"/> class.
    /// One of the parameters <paramref name="filePath"/>, <paramref name="stream"/> must be provided.
    /// </summary>
    /// <param name="filePath">The file path where the PowerPoint presentation will be saved.</param>
    /// <param name="stream">The stream to write the PowerPoint presentation to.</param>
    public PowerPointBuilder(string? filePath = null, Stream? stream = null)
    {
        _properties = new BuilderProperties(filePath);
    }

    /// <summary>
    /// Adds a slide to the PowerPoint presentation.
    /// </summary>
    /// <param name="slideAction">An optional action to configure the slide using a <see cref="SlideBuilder"/>.</param>
    /// <param name="layout">An optional layout to apply to the slide using a <see cref="TemplateLayoutBuilder"/>.</param>
    /// <param name="slide">An optional existing Document.FormatXml slide when you want to use the full functionality of the base package.</param>
    /// <returns>The current instance of <see cref="PowerPointBuilder"/> for method chaining.</returns>
    public PowerPointBuilder AddSlide(Action<SlideBuilder>? slideAction = null, TemplateLayoutBuilder? layout = null, Slide? slide = null)
    {
        var slideBuilder = new SlideBuilder(slide, layout);
        slideAction?.Invoke(slideBuilder);
        _properties.Slides.Add(slideBuilder);
        return this;
    }

    /// <summary>
    /// Builds the PowerPoint presentation and writes it to the specified file or stream.
    /// </summary>
    public void Build()
    {
        new PresentationUtility(_properties).Build();
    }
}