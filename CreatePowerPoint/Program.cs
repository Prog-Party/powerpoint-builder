// See https://aka.ms/new-console-template for more information
using PowerPoint.Builder;
using PowerPoint.Builder.Core;

string path = $"test_{DateTime.Now.ToString("yyyyMMddHHmmss")}.pptx";

Stream imageStream = File.OpenRead("lightbulb.jpg");

new PowerPointBuilder(path)
    .AddSlide(slide => slide
        .AddText(text => text
            .AddParagraph("Text1")
            .AddParagraph(paragraph => paragraph
                .AddText("Text2")
                .AddBoldText("Bold")
                .AddItalicText("Italic")
                .AddUnderlineText("Underline"))
            .AddParagraph(paragraph => paragraph
                .AddTexts("Bullet 1", "Bullet 2")
                .SetBulletList())
            .AddParagraph(paragraph => paragraph
                .AddTexts("Dash list 1", "Dash list 2")
                .SetBulletList("-"))
            .AddParagraph(paragraph => paragraph
                .AddTexts("Nr 3", "Nr 4")
                .SetNumberedList(3))
            .SetPosition(PartPosition.Construct(xPercentage: 10))
            .SetSize(PartSize.Construct(widthPercentage: 80, heightPercentage: 50)))
        .AddText(text => text
            .AddParagraph("Test 2")
            .SetPosition(PartPosition.Construct(xPercentage: 90, yPercentage: 50))
            .SetSize(PartSize.Construct(width: 200000, height: 1000000))))
    .AddSlide(slide => slide
        .AddImage(image => image
            .SetImage(imageStream, ".jpg"))
        .AddImage(image => image
            .SetImage("lightbulb.jpg")
            .SetPosition(PartPosition.Construct(xPercentage: 70, yPercentage: 50)))
        )
    .Build();