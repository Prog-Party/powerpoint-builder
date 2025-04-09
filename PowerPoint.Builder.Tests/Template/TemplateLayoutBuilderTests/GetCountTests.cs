using PowerPoint.Builder.Template;

namespace PowerPoint.Builder.Tests.Template.TemplateLayoutBuilderTests;

public class GetCountTests
{
    [Fact]
    public void GetCount_ShouldReturnZero_WhenNoElementsAreAdded()
    {
        // Arrange
        var templateLayoutBuilder = new TemplateLayoutBuilder();

        // Act
        var count = templateLayoutBuilder.GetCount();

        // Assert
        Assert.Equal(0, count);
    }

    [Fact]
    public void GetCount_ShouldReturnCorrectCount_WhenElementsAreAdded()
    {
        // Arrange
        var templateLayoutBuilder = new TemplateLayoutBuilder();
        templateLayoutBuilder.AddLayoutPart();
        templateLayoutBuilder.AddLayoutPart();

        // Act
        var count = templateLayoutBuilder.GetCount();

        // Assert
        Assert.Equal(2, count);
    }
}