using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Moq;
using PowerPoint.Builder.Slides;
using PowerPoint.Builder.Template;

namespace PowerPoint.Builder.Tests.Template.TemplateLayoutBuilderTests;

public static class TestHelpers
{
    public static SlidePart GetSlidePartForTest()
    {
        // Create an in-memory presentation document
        var stream = new MemoryStream();
        var presentation = PresentationDocument.Create(stream, PresentationDocumentType.Presentation, true);

        // Add the required parts
        var presentationPart = presentation.AddPresentationPart();
        presentationPart.Presentation = new DocumentFormat.OpenXml.Presentation.Presentation();

        var slidePart = presentationPart.AddNewPart<SlidePart>();
        slidePart.Slide = new Slide(new CommonSlideData(new ShapeTree()));

        // Important: DO NOT dispose the presentation here, or the SlidePart will become invalid!
        // Just return the SlidePart, and keep the stream/presentation open as long as needed
        return slidePart;
    }
}

public class BuildTests
{
    public readonly SlidePart DummySlidePart = TestHelpers.GetSlidePartForTest();
    public readonly ShapeTree DummyShapeTree = new ShapeTree();
    public readonly Mock<SlidePartBuilder> SlidePartBuilderMock = new Mock<SlidePartBuilder>();

    [Fact]
    public void Build_WithIndexOutOfRange_ThrowsArgumentOutOfRangeException()
    {
        // Arrange
        var builder = new TemplateLayoutBuilder();

        // Add 1 layout part, so valid index = 0
        builder.AddLayoutPart();

        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            builder.Build(1, SlidePartBuilderMock.Object, DummySlidePart, DummyShapeTree)
        );
    }

    [Fact]
    public void Build_WithCorrectIndex_Works()
    {
        // Arrange
        var builder = new TemplateLayoutBuilder();

        // Add 1 layout part, so valid index = 0
        builder.AddLayoutPart();

        // Act & Assert
        builder.Build(0, SlidePartBuilderMock.Object, DummySlidePart, DummyShapeTree);
    }

    [Fact]
    public void Build_WithNegativeIndex_ThrowsArgumentOutOfRangeException()
    {
        // Arrange
        var builder = new TemplateLayoutBuilder();

        // Act & Assert
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            builder.Build(-1, SlidePartBuilderMock.Object, DummySlidePart, DummyShapeTree)
        );
    }
}