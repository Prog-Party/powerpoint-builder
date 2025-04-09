using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PowerPoint.Builder.Slides.Parts;
using PowerPoint.Builder.Template;

using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPoint.Builder.Slides;

/// <summary>
/// Provides functionality to add PowerPoint slides.
/// For more details, see the
/// <see href="https://github.com/Prog-Party/powerpoint-builder/wiki/SlideBuilder-Class">
/// SlideBuilder Class Documentation
/// </see>.
/// </summary>
public class SlideBuilder
{
    private readonly Slide? _slide;
    private readonly TemplateLayoutBuilder? _layout;
    private List<SlidePartBuilder> _elements = new();

    /// <summary>
    /// Initializes a new instance of the <see cref="SlideBuilder"/> class.
    /// </summary>
    /// <param name="slide">An optional DocumentFormat.OpenXml slide to use for full customization.</param>
    /// <param name="layout">An optional layout template to apply to the slide.</param>
    internal SlideBuilder(Slide? slide = null, TemplateLayoutBuilder? layout = null)
    {
        _slide = null;
        _layout = layout;
    }

    /// <summary>
    /// Adds a text element to the slide.
    /// A text element consist of paragraphs with their own properties.
    /// </summary>
    /// <param name="action">An optional action to configure the <see cref="TextBuilder"/>.</param>
    /// <returns>The current <see cref="SlideBuilder"/> instance for method chaining.</returns>
    public SlideBuilder AddText(Action<TextBuilder>? action = null)
    {
        var builder = new TextBuilder();
        action?.Invoke(builder);
        _elements.Add(builder);
        return this;
    }

    /// <summary>
    /// Adds an image element to the slide.
    /// </summary>
    /// <param name="action">An optional action to configure the <see cref="ImageBuilder"/>.</param>
    /// <returns>The current <see cref="SlideBuilder"/> instance for method chaining.</returns>
    public SlideBuilder AddImage(Action<ImageBuilder>? action = null)
    {
        var builder = new ImageBuilder();
        action?.Invoke(builder);
        _elements.Add(builder);
        return this;
    }

    /// <summary>
    /// Builds the slide and adds all configured elements to the specified <see cref="SlidePart"/>.
    /// When there is a layout available, the first elements will be added to the layout parts.
    /// </summary>
    /// <param name="slidePart">The <see cref="SlidePart"/> to which the slide and its elements will be added.</param>
    internal void Build(SlidePart slidePart)
    {
        if (_slide != null)
        {
            slidePart.Slide = _slide;
            return;
        }
        var hasLayout = _layout != null;

        var shapeTree = new ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new TransformGroup())
                        );

        var nonImageElements = _elements.Skip(hasLayout ? _layout!.GetCount() : 0).Where(e => e is not ImageBuilder);

        for (var i = 0; i < _elements.Count; i++)
        {
            var element = _elements[i];
            if (element is ImageBuilder)
                continue; // Skip image elements, these will be processed later

            // If there is a layout, add the items to the layout first
            var hasLayoutToApply = hasLayout && i < _layout!.GetCount();
            if (hasLayoutToApply)
                _layout!.Build(i, element, slidePart, shapeTree);
            else
                element.Build(slidePart, shapeTree);
        }

        var slide = new Slide(
                new CommonSlideData(shapeTree),
                new ColorMapOverride(new MasterColorMapping())
            );

        slidePart.Slide = slide;

        for (var i = 0; i < _elements.Count; i++)
        {
            var element = _elements[i];
            if (element is not ImageBuilder)
                continue; // Skip non-image elements, these are added earlier already

            // If there is a layout, add the items to the layout first
            var hasLayoutToApply = hasLayout && i < _layout!.GetCount();
            if (hasLayoutToApply)
                _layout!.Build(i, element, slidePart, shapeTree);
            else
                element.Build(slidePart, shapeTree);
        }

        // Add the empty layout parts
        if (hasLayout)
        {
            for (var i = _elements.Count; i < _layout!.GetCount(); i++)
            {
                _layout.Build(i, slidePart, shapeTree);
            }
        }
    }
}