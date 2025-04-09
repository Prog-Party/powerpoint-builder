using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PowerPoint.Builder.Slides;

namespace PowerPoint.Builder.Template;

public class TemplateLayoutBuilder
{
    private Slide? _slide;
    private List<LayoutPartBuilder> _layoutParts = new();

    public TemplateLayoutBuilder AddLayoutPart(Action<LayoutPartBuilder>? action = null)
    {
        var builder = new LayoutPartBuilder();
        action?.Invoke(builder);
        _layoutParts.Add(builder);
        return this;
    }

    /// <summary>
    ///
    /// </summary>
    /// <param name="index"></param>
    /// <param name="slidePartBuilder"></param>
    /// <param name="slidePart"></param>
    /// <param name="tree"></param>
    /// <exception cref="ArgumentOutOfRangeException"></exception>
    internal void Build(int index, SlidePartBuilder slidePartBuilder, SlidePart slidePart, ShapeTree tree)
    {
        ValidateIndex(index);

        var layoutPart = _layoutParts[index];
        layoutPart.Build(slidePartBuilder, slidePart, tree);
    }

    /// <summary>
    /// A template can be created without slideparts to be added, in this case we will add a textbox with a placeholder text
    /// </summary>
    /// <param name="index"></param>
    /// <param name="slidePart"></param>
    /// <param name="tree"></param>
    /// <exception cref="ArgumentOutOfRangeException"></exception>
    internal void Build(int index, SlidePart slidePart, ShapeTree tree)
    {
        ValidateIndex(index);

        var layoutPart = _layoutParts[index];
        layoutPart.Build(slidePart, tree);
    }

    private void ValidateIndex(int index)
    {
        if (index < 0 || index >= _layoutParts.Count)
            throw new ArgumentOutOfRangeException(nameof(index), $"Index {index} is out of range.");
    }

    internal int GetCount()
        => _layoutParts.Count;
}