using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PowerPoint.Builder.Slides.Parts;

using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPoint.Builder.Slides;

public class SlideBuilder
{
    private List<SlidePartBuilder> _elements = new();
    private List<ImageBuilder> _images = new();

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
        _images.Add(builder);
        return this;
    }

    internal void Build(SlidePart slidePart)
    {
        var shapeTree = new ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new TransformGroup())
                        );

        foreach (var element in _elements)
            shapeTree.Append(element.Build());

        var slide = new Slide(
                new CommonSlideData(shapeTree),
                new ColorMapOverride(new MasterColorMapping())
            );

        slidePart.Slide = slide;

        foreach (var image in _images)
            image.Build(slidePart, shapeTree);
    }
}