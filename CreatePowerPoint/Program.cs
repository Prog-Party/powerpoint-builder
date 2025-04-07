// See https://aka.ms/new-console-template for more information
using PowerPoint.Builder;

string path = $"test_{DateTime.Now.ToString("yyyyMMddHHmmss")}.pptx";

new PowerPointBuilder(path)
    .AddSlide(slide => slide
        .AddText("Test", xPercent: 50, yPercent: 0, widthPercent: 20, heightPercent: 20)
        .AddText("Test", xPercent: 50, yPercent: 0, widthPercent: 20, heightPercent: 20)
        .AddText("Test 2", xPercent: 90, yPercent: 50, width: 200000, height: 1000000)
        .AddText("Test 3", x: 500000, y: 600000, widthPercent: 80, heightPercent: 50))
    .AddSlide(slide => slide
        .AddText("Twee", xPercent: 50)
        .AddText("Twee 2", xPercent: 90)
        .AddText("Twee 3", x: 500000))
    .Build();