// See https://aka.ms/new-console-template for more information
using PowerPoint.Builder;
using PowerPoint.Builder.Core;

string path = $"test_{DateTime.Now.ToString("yyyyMMddHHmmss")}.pptx";

Stream imageStream = File.OpenRead("lightbulb.jpg");

new PowerPointBuilder(path)
    .AddSlide(slide => slide
        .AddImage(image => image
            .SetImage(imageStream, ".jpg"))
        .AddImage(image => image
            .SetImage("lightbulb.jpg")
            .SetPosition(PartPosition.Construct(xPercentage: 70, yPercentage: 50)))
        .AddText("Test", text => text
            .SetPosition(PartPosition.Construct(xPercentage: 50))
            .SetSize(PartSize.Construct(widthPercentage: 20, heightPercentage: 20)))
        .AddText("Test 2", text => text
            .SetPosition(PartPosition.Construct(xPercentage: 90, yPercentage: 50))
            .SetSize(PartSize.Construct(width: 200000, height: 1000000)))
        .AddText("Test 3", text => text
            .SetPosition(PartPosition.Construct(x: 500000, y: 600000))
            .SetSize(PartSize.Construct(width: 200000, height: 1000000))))
    .AddSlide(slide => slide
        .AddText("Twee", text => text
            .SetPosition(PartPosition.Construct(xPercentage: 50)))
        .AddText("Twee 2", text => text
            .SetPosition(PartPosition.Construct(xPercentage: 90)))
        .AddText("Twee 3", text => text
            .SetPosition(PartPosition.Construct(x: 500000)))
        .AddImage(image => image
            .SetImage("lightbulb.jpg")
            .SetPosition(PartPosition.Construct(xPercentage: 70, yPercentage: 30))))
    .Build();