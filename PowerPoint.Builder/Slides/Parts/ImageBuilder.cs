using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PowerPoint.Builder.Core;
using PowerPoint.Builder.Slides;
using PowerPoint.Builder.Utility;

using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace PowerPoint.Builder.Slides.Parts;

public class ImageBuilder
{
    private string? _imageLocation;
    private Stream? _imageStream;
    private string _imageExtension;
    private PartPosition _position;
    private PartSize _size;

    internal ImageBuilder()
    {
        _position = PartPosition.Construct(0, 0);
        _size = PartSize.Construct(widthPercentage: 10, heightPercentage: 10);
        _imageExtension = string.Empty;
    }

    public ImageBuilder SetImage(string imageLocation)
    {
        _imageLocation = imageLocation;
        _imageExtension = System.IO.Path.GetExtension(imageLocation);
        return this;
    }

    /// <summary>
    /// Set the image from a stream. The stream must be seekable.
    /// </summary>
    /// <param name="imageStream"></param>
    /// <param name="extension">The extension (or filename) of the file, example: image.png, or .png</param>
    /// <returns></returns>
    public ImageBuilder SetImage(Stream imageStream, string extension)
    {
        _imageStream = imageStream;
        _imageExtension = System.IO.Path.GetExtension(extension);
        return this;
    }

    public ImageBuilder SetPosition(PartPosition position)
    {
        _position = position;
        return this;
    }

    public ImageBuilder SetSize(PartSize size)
    {
        _size = size;
        return this;
    }

    internal void Build(SlidePart slidePart, ShapeTree tree)
    {
        int randomId = new Random().Next(0, 1000000);

        var part = slidePart.AddImagePart(ImageUtility.GetImagePartTypeByPath(_imageExtension));

        if (_imageStream != null)
        {
            part.FeedData(_imageStream);
        }
        else if (_imageLocation != null)
        {
            using (var stream = File.OpenRead(_imageLocation))
            {
                part.FeedData(stream);
            }
        }
        else
        {
            throw new ArgumentException("Image location or stream must be provided.");
        }

        var picture = new P.Picture();

        picture.NonVisualPictureProperties = new P.NonVisualPictureProperties();
        picture.NonVisualPictureProperties.Append(new P.NonVisualDrawingProperties
        {
            Name = "My Shape",
            Id = UInt32Value.FromUInt32((uint)randomId)
        });

        var nonVisualPictureDrawingProperties = new P.NonVisualPictureDrawingProperties();
        nonVisualPictureDrawingProperties.Append(new D.PictureLocks()
        {
            NoChangeAspect = true
        });
        picture.NonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);
        picture.NonVisualPictureProperties.Append(new P.ApplicationNonVisualDrawingProperties());

        var blipFill = new P.BlipFill();
        var blip1 = new D.Blip()
        {
            Embed = slidePart.GetIdOfPart(part)
        };
        var blipExtensionList1 = new D.BlipExtensionList();
        var blipExtension1 = new D.BlipExtension()
        {
            Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
        };
        var useLocalDpi1 = new DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi()
        {
            Val = false
        };
        useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
        blipExtension1.Append(useLocalDpi1);
        blipExtensionList1.Append(blipExtension1);
        blip1.Append(blipExtensionList1);
        var stretch = new D.Stretch();
        stretch.Append(new D.FillRectangle());
        blipFill.Append(blip1);
        blipFill.Append(stretch);
        picture.Append(blipFill);

        picture.ShapeProperties = new P.ShapeProperties();
        picture.ShapeProperties.Transform2D = new D.Transform2D();
        picture.ShapeProperties.Transform2D.Append(new D.Offset
        {
            X = _position.XOffset,
            Y = _position.YOffset,
        });
        picture.ShapeProperties.Transform2D.Append(new D.Extents
        {
            Cx = _size.Width,
            Cy = _size.Height,
        });
        picture.ShapeProperties.Append(new D.PresetGeometry
        {
            Preset = D.ShapeTypeValues.Rectangle
        });

        tree.Append(picture);
    }
}