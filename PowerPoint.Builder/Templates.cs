using PowerPoint.Builder.Core;
using PowerPoint.Builder.Template;

namespace PowerPoint.Builder;

public static class Templates
{
    /// <summary>
    /// This template is used to create a title and content layout.
    /// It contains a title part that is on top and a content part that is below the title.
    /// </summary>
    public static TemplateLayoutBuilder TitleContentTemplate
        => new TemplateLayoutBuilder()
        .AddLayoutPart(part => part
            .SetPlaceholderText("Title")
            .SetPosition(PartPosition.Construct(xPercentage: 10, yPercentage: 10))
            .SetSize(PartSize.Construct(widthPercentage: 80, heightPercentage: 10)))
        .AddLayoutPart(part => part
            .SetPlaceholderText("Content")
            .SetPosition(PartPosition.Construct(xPercentage: 10, yPercentage: 30))
            .SetSize(PartSize.Construct(widthPercentage: 80, heightPercentage: 50)));
}