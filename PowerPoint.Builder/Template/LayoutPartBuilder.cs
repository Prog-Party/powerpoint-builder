using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PowerPoint.Builder.Core;
using PowerPoint.Builder.Slides;
using PowerPoint.Builder.Slides.Parts;

namespace PowerPoint.Builder.Template;

public class LayoutPartBuilder
{
    private string _placeholderText = "Empty";
    private PartPosition _position;
    private PartSize _size;

    internal LayoutPartBuilder()
    {
        _position = PartPosition.Construct(0, 0);
        _size = PartSize.Construct(widthPercentage: 100, heightPercentage: 100);
    }

    public LayoutPartBuilder SetPosition(PartPosition position)
        => Execute(builder => builder._position = position);

    public LayoutPartBuilder SetSize(PartSize size)
        => Execute(builder => builder._size = size);

    public LayoutPartBuilder SetPlaceholderText(string text)
        => Execute(builder => builder._placeholderText = text);

    /// <summary>
    /// Parts of templates could be empty, then we create a placeholder text.
    /// </summary>
    /// <param name="slidePart"></param>
    /// <param name="tree"></param>
    internal void Build(SlidePart slidePart, ShapeTree tree)
    {
        var slidePartBuilder = new TextBuilder()
            .AddParagraph(_placeholderText);
        Build(slidePartBuilder, slidePart, tree);
    }

    internal void Build(SlidePartBuilder slidePartBuilder, SlidePart slidePart, ShapeTree tree)
    {
        var type = slidePartBuilder.GetType();

        // Check and invoke SetSize method
        var setSizeMethod = type.GetMethod("SetSize");
        if (setSizeMethod != null)
            setSizeMethod.Invoke(slidePartBuilder, [_size]);

        // Check and invoke SetPosition method
        var setPositionMethod = type.GetMethod("SetPosition");
        if (setPositionMethod != null)
            setPositionMethod.Invoke(slidePartBuilder, [_position]);

        slidePartBuilder.Build(slidePart, tree);
    }

    private LayoutPartBuilder Execute(Action<LayoutPartBuilder> action)
    {
        action(this);
        return this;
    }
}