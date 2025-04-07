using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;
using PowerPoint.Builder.Text;

using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace PowerPoint.Builder.PowerPointSlide;

public class SlideBuilder
{
    private List<OpenXmlElement> _elements = new();

    public SlideBuilder AddText(
        string text,
        int x = 0,
        int y = 0,
        int? width = null,
        int? height = null,
        int? xPercent = null,
        int? yPercent = null,
        int? widthPercent = null,
        int? heightPercent = null)
    {
        var properties = new TextProperties(text,
            xOffset: x, yOffset: y, width: width, height: height,
            xOffsetPercentage: xPercent, yOffsetPercentage: yPercent,
            widthPercentage: widthPercent, heightPercentage: heightPercent);

        _elements.Add(TextBuilder.Build(properties));
        return this;
    }

    public Slide Build()
    {
        var shapeTree = new ShapeTree(
                        new P.NonVisualGroupShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
                            new P.NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new GroupShapeProperties(new TransformGroup())
                        );

        foreach (var element in _elements)
            shapeTree.Append(element);

        var slide = new Slide(
                new CommonSlideData(shapeTree),
                new ColorMapOverride(new MasterColorMapping())
            );
        return slide;
    }
}