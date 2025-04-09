using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PowerPoint.Builder.Core;
using PowerPoint.Builder.Slides;

using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPoint.Builder.Slides.Parts;

public class TextBuilder : SlidePartBuilder
{
    private List<Paragraph> _paragraphs = new List<Paragraph>();
    private PartPosition _position;
    private PartSize _size;

    internal TextBuilder()
    {
        _position = PartPosition.Construct(0, 0);
        _size = PartSize.Construct(widthPercentage: 100, heightPercentage: 100);
    }

    public TextBuilder AddParagraph(Action<ParagraphBuilder> builder)
    {
        var paragraphBuilder = new ParagraphBuilder();
        builder(paragraphBuilder);
        _paragraphs.Add(paragraphBuilder.Build());
        return this;
    }

    public TextBuilder AddParagraph(string text)
        => AddParagraph(paragraph => paragraph.AddText(text));

    public TextBuilder AddParagraph(List<string> texts)
        => AddParagraph(string.Join(Environment.NewLine, texts));

    public TextBuilder SetPosition(PartPosition position)
        => Execute(builder => builder._position = position);

    public TextBuilder SetSize(PartSize size)
        => Execute(builder => builder._size = size);

    internal override void Build(SlidePart slidePart, ShapeTree tree)
    {
        var shapeProperties = new P.ShapeProperties(
                     new Transform2D(
                         new Offset() { X = _position.XOffset, Y = _position.YOffset },
                         new Extents() { Cx = _size.Width, Cy = _size.Height }
                     ));

        var randomId = new Random().Next(0, 1000000);

        var bodyChildren = new List<OpenXmlElement>
        {
            new BodyProperties(),
            new ListStyle()
        };
        bodyChildren.AddRange(_paragraphs.Cast<OpenXmlElement>());

        var shape = new P.Shape(
                            new P.NonVisualShapeProperties(
                                new P.NonVisualDrawingProperties() { Id = UInt32Value.FromUInt32((uint)randomId), Name = "" },
                                new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                                new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
                            shapeProperties,
                            new P.TextBody(bodyChildren));

        tree.Append(shape);
    }

    private TextBuilder Execute(Action<TextBuilder> action)
    {
        action(this);
        return this;
    }
}