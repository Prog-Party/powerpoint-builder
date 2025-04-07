using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;

using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

using PowerPoint.Builder.Slides.Parts;

namespace PowerPoint.Builder.Slides;

public class SlideBuilder
{
    private List<SlidePartBuilder> _elements = new();

    public SlideBuilder AddText(string text, Action<TextBuilder> action)
    {
        var builder = new TextBuilder(text);
        action(builder);
        _elements.Add(builder);
        return this;
    }

    internal Slide Build()
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
        return slide;
    }
}