using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PowerPoint.Builder.Slides;

using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPoint.Builder.Presentation;

internal class SlideUtility
{
    internal static void ConstructFirstSlide(PresentationPart presentationPart, SlideBuilder slideBuilder)
    {
        var slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
        var slideIdList1 = new SlideIdList(
            new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" }
            );
        var slideSize1 = new SlideSize() { Cx = Constants.SlideWidth, Cy = Constants.SlideHeight, Type = SlideSizeValues.Screen4x3 };
        var notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
        var defaultTextStyle1 = new DefaultTextStyle();

        presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);

        var slidePart1 = presentationPart.AddNewPart<SlidePart>("rId2");
        slideBuilder.Build(slidePart1);

        var slideLayoutPart1 = CreateSlideLayoutPart(slidePart1);
        var slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1);
        var themePart1 = new ThemeUtility(slideMasterPart1).Build();
        slideMasterPart1.AddPart(slideLayoutPart1, "rId1");
        presentationPart.AddPart(slideMasterPart1, "rId1");
        presentationPart.AddPart(themePart1, "rId5");
    }

    private static SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1)
    {
        var slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>("rId1");
        var slideLayout = new SlideLayout(
        new CommonSlideData(new ShapeTree(
          new P.NonVisualGroupShapeProperties(
          new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
          new P.NonVisualGroupShapeDrawingProperties(),
          new ApplicationNonVisualDrawingProperties()),
          new GroupShapeProperties(new TransformGroup()),
          new P.Shape(
          new P.NonVisualShapeProperties(
            new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "" },
            new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
            new ApplicationNonVisualDrawingProperties(new PlaceholderShape())),
          new P.ShapeProperties(),
          new P.TextBody(
            new BodyProperties(),
            new ListStyle(),
            new Paragraph(new EndParagraphRunProperties()))))),
        new ColorMapOverride(new MasterColorMapping()));
        slideLayoutPart1.SlideLayout = slideLayout;
        return slideLayoutPart1;
    }

    public static SlidePart InsertNewSlide(PresentationDocument presentationDocument, int position, SlideBuilder slideBuilder)
    {
        var presentationPart = presentationDocument.PresentationPart;

        // Verify that the presentation is not empty.
        if (presentationPart is null)
            throw new InvalidOperationException("The presentation document is empty.");

        // Create the slide part for the new slide.
        var slidePart = presentationPart.AddNewPart<SlidePart>();

        slideBuilder.Build(slidePart);

        // Modify the slide ID list in the presentation part.
        // The slide ID list should not be null.
        var slideIdList = presentationPart.Presentation.SlideIdList;

        // Find the highest slide ID in the current list.
        uint maxSlideId = 1;
        SlideId? prevSlideId = null;

        var slideIds = slideIdList?.ChildElements ?? default;

        foreach (SlideId slideId in slideIds)
        {
            if (slideId.Id is not null && slideId.Id > maxSlideId)
                maxSlideId = slideId.Id;

            position--;
            if (position == 0)
                prevSlideId = slideId;
        }

        maxSlideId++;

        // Get the ID of the previous slide.
        SlidePart lastSlidePart;

        if (prevSlideId is not null && prevSlideId.RelationshipId is not null)
        {
            lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId!);
        }
        else
        {
            string? firstRelId = ((SlideId)slideIds[0]).RelationshipId;
            // If the first slide does not contain a relationship ID, throw an exception.
            if (firstRelId is null)
                throw new ArgumentNullException(nameof(firstRelId));

            lastSlidePart = (SlidePart)presentationPart.GetPartById(firstRelId);
        }

        // Use the same slide layout as that of the previous slide.
        if (lastSlidePart.SlideLayoutPart is not null)
            slidePart.AddPart(lastSlidePart.SlideLayoutPart);

        // Insert the new slide into the slide list after the previous slide.
        var newSlideId = slideIdList!.InsertAfter(new SlideId(), prevSlideId);
        newSlideId.Id = maxSlideId;
        newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

        return slidePart;
    }

    private static SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1)
    {
        var slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
        var slideMaster = new SlideMaster(
        new CommonSlideData(new ShapeTree(
          new P.NonVisualGroupShapeProperties(
          new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
          new P.NonVisualGroupShapeDrawingProperties(),
          new ApplicationNonVisualDrawingProperties()),
          new GroupShapeProperties(new TransformGroup()),
          new P.Shape(
          new P.NonVisualShapeProperties(
            new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Title Placeholder 1" },
            new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
            new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Title })),
          new P.ShapeProperties(),
          new P.TextBody(
            new BodyProperties(),
            new ListStyle(),
            new Paragraph())))),
        new P.ColorMap() { Background1 = ColorSchemeIndexValues.Light1, Text1 = ColorSchemeIndexValues.Dark1, Background2 = ColorSchemeIndexValues.Light2, Text2 = ColorSchemeIndexValues.Dark2, Accent1 = ColorSchemeIndexValues.Accent1, Accent2 = ColorSchemeIndexValues.Accent2, Accent3 = ColorSchemeIndexValues.Accent3, Accent4 = ColorSchemeIndexValues.Accent4, Accent5 = ColorSchemeIndexValues.Accent5, Accent6 = ColorSchemeIndexValues.Accent6, Hyperlink = ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = ColorSchemeIndexValues.FollowedHyperlink },
        new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }),
        new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));
        slideMasterPart1.SlideMaster = slideMaster;

        return slideMasterPart1;
    }
}