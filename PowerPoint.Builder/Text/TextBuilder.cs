using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;

using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPoint.Builder.Text;

internal class TextBuilder
{
    internal static OpenXmlElement Build(TextProperties properties)
    {
        var shapeProperties = new P.ShapeProperties(
                     new D.Transform2D(
                         new D.Offset() { X = properties.XOffset, Y = properties.YOffset },
                         new D.Extents() { Cx = properties.Width, Cy = properties.Height }
                     ));

        int randomId = new Random().Next(0, 1000000);

        var shape = new P.Shape(
                            new P.NonVisualShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = UInt32Value.FromUInt32((uint)randomId), Name = "" },
                                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                            shapeProperties,
                            new P.TextBody(
                            new BodyProperties(),
                            new ListStyle(),
                            new Paragraph(
                                new Run(
                                    new D.Text(properties.Text))

                                )));

        return shape;
    }
}