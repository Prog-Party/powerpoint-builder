using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ÒpenXmlPresentation = DocumentFormat.OpenXml.Presentation.Presentation;

namespace PowerPoint.Builder.Presentation;

internal class PresentationUtility
{
    private BuilderProperties _properties { get; init; }

    internal PresentationUtility(BuilderProperties properties)
    {
        _properties = properties;
    }

    internal void Build()
    {
        using (var presentationDoc = CreatePresentationDocument())
        {
            var presentationPart = presentationDoc.AddPresentationPart();
            presentationPart.Presentation = new ÒpenXmlPresentation();

            foreach (var slideBuilder in _properties.Slides)
            {
                var slide = slideBuilder.Build();
                AddSlide(presentationDoc, presentationPart, slide);
            }
        }
    }

    private static void AddSlide(PresentationDocument presentationDoc, PresentationPart presentationPart, Slide slide)
    {
        var length = presentationPart.GetPartsOfType<SlidePart>().Count();

        if (length == 0)
            SlideUtility.ConstructFirstSlide(presentationPart, slide);
        else
            SlideUtility.InsertNewSlide(presentationDoc, length, slide);
    }

    private PresentationDocument CreatePresentationDocument()
    {
        var source = _properties.Source;
        if (source.Filepath is not null)
            return PresentationDocument.Create(source.Filepath, PresentationDocumentType.Presentation);
        else if (source.Stream is not null)
            return PresentationDocument.Create(source.Stream, PresentationDocumentType.Presentation);
        else if (source.Package is not null)
            return PresentationDocument.Create(source.Package, PresentationDocumentType.Presentation);
        else
            throw new ArgumentException("Invalid source.");
    }
}