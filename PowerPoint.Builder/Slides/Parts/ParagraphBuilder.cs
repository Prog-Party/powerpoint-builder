using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Diagrams;

using D = DocumentFormat.OpenXml.Drawing;

namespace PowerPoint.Builder.Slides.Parts;

public class ParagraphBuilder
{
    private List<(string, RunProperties?)> _paragraphParts = new();
    private ParagraphProperties? _properties = null;

    public ParagraphBuilder AddText(string text, RunProperties? properties = null)
    {
        _paragraphParts.Add((text, properties));
        return this;
    }

    public ParagraphBuilder AddTexts(params string[] texts)
        => AddText(string.Join(Environment.NewLine, texts));

    public ParagraphBuilder AddBoldText(string text)
        => AddText(text, new RunProperties() { Bold = true });

    public ParagraphBuilder AddItalicText(string text)
        => AddText(text, new RunProperties() { Italic = true });

    public ParagraphBuilder AddUnderlineText(string text, TextUnderlineValues? style = null)
        => AddText(text, new RunProperties() { Underline = style ?? TextUnderlineValues.Single });

    public ParagraphBuilder SetBulletList(string character = "•")
    {
        if (_properties == null)
            _properties = new ParagraphProperties();
        _properties.LeftMargin = 342900; // 24pt in EMUs
        _properties.Indent = -285750;     // -22.5pt hanging indent
        _properties.AddChild(new BulletFont { Typeface = "Arial" });
        _properties.AddChild(new CharacterBullet { Char = character });

        return this;
    }

    public ParagraphBuilder SetNumberedList(int startAt = 1)
    {
        if (_properties == null)
            _properties = new ParagraphProperties();

        _properties.LeftMargin = 342900; // 24pt in EMUs
        _properties.Indent = -285750;     // -22.5pt hanging indent
        _properties.AddChild(new AutoNumberedBullet()
        {
            Type = TextAutoNumberSchemeValues.ArabicPeriod, // e.g., 1.
            StartAt = 1
        });
        return this;
    }

    public ParagraphBuilder SetProperties(ParagraphProperties properties)
        => Execute(builder => builder._properties = properties);

    internal Paragraph Build()
    {
        var paragraph = new Paragraph();

        if (_properties != null)
        {
            paragraph.ParagraphProperties = _properties;
        }
        foreach (var (text, runProperties) in _paragraphParts)
        {
            var run = new Run(new D.Text(text));
            if (runProperties != null)
            {
                run.AddChild(runProperties);
            }
            paragraph.Append(run);
        }

        return paragraph;
    }

    private ParagraphBuilder Execute(Action<ParagraphBuilder> action)
    {
        action(this);
        return this;
    }
}