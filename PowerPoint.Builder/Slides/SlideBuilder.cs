using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PowerPoint.Builder.Slides.Parts;
using PowerPoint.Builder.Template;

using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPoint.Builder.Slides;

public class SlideBuilder
{
    private readonly Slide? _slide;
    private readonly TemplateLayoutBuilder? _layout;
    private List<SlidePartBuilder> _elements = new();

    internal SlideBuilder(Slide? slide = null, TemplateLayoutBuilder? layout = null)
    {
        _slide = null;
        _layout = layout;
    }

    public SlideBuilder AddText(Action<TextBuilder>? action = null)
    {
        var builder = new TextBuilder();
        action?.Invoke(builder);
        _elements.Add(builder);
        return this;
    }

    public SlideBuilder AddImage(Action<ImageBuilder>? action = null)
    {
        var builder = new ImageBuilder();
        action?.Invoke(builder);
        _elements.Add(builder);
        return this;
    }

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
                continue; //skip image elements, these will be processed later

            //if there is a layout, we need to add the items to the layout first
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
                continue; //skip non image elements, these are added earlier already

            //if there is a layout, we need to add the items to the layout first
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