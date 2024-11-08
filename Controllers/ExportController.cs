using System;
using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Pictures;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Test.Data;
using Test.Models;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace Test.Controllers
{
    [Authorize]
    public class ExportController : Controller
    {
        private readonly ApplicationDbContext _context;

        public ExportController(ApplicationDbContext context)
        {
            _context = context;
        }

        public async Task<IActionResult> ExportToWord(int id)
        {
            var letter = await _context.Letters
                .Include(l => l.Category)
                .Include(l => l.User)
                .FirstOrDefaultAsync(m => m.Id == id);

            if (letter == null)
            {
                return NotFound();
            }

            using (var stream = new MemoryStream())
            {
                using (var wordDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
                {
                    // Add a main document part
                    var mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    var body = mainPart.Document.AppendChild(new Body());

                    // Add the logo image
                    AddImage(mainPart, body, "/logo/logo.png");

                    // Add letter details
                    body.AppendChild(CreateParagraph("شماره نامه: " + letter.Id));
                    body.AppendChild(CreateParagraph("تاریخ: " + DateTime.Now.ToShortDateString()));
                    body.AppendChild(CreateParagraph("موضوع: " + letter.Subject));
                    body.AppendChild(CreateParagraph("از: " + letter.Sender));
                    body.AppendChild(CreateParagraph("به: " + letter.Receiver));
                    body.AppendChild(CreateParagraph("توضیحات: " + letter.Description));

                    // Add the signature image
                    AddImage(mainPart, body, "/logo/signature.png");

                    // Add the footer image
                    AddImage(mainPart, body, "/logo/footer.png");
                }

                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "OfficialLetter.docx");
            }
        }

        private void AddImage(MainDocumentPart mainPart, Body body, string imagePath)
        {
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (var imageStream = new FileStream($"wwwroot{imagePath}", FileMode.Open, FileAccess.Read))
            {
                imagePart.FeedData(imageStream);
            }

            AddImageToBody(mainPart.GetIdOfPart(imagePart), body);
        }

        private void AddImageToBody(string relationshipId, Body body)
        {
            var element =
                new Drawing(
                    new DW.Inline(
                        new DW.Extent() { Cx = 990000L, Cy = 792000L },
                        new DW.EffectExtent()
                        {
                            LeftEdge = 19050L,
                            TopEdge = 0L,
                            RightEdge = 9525L,
                            BottomEdge = 0L
                        },
                        new DW.DocProperties()
                        {
                            Id = (UInt32Value)1U,
                            Name = "Picture 1"
                        },
                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new A.GraphicFrameLocks() { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties()
                                        {
                                            Id = (UInt32Value)0U,
                                            Name = "New Bitmap Image.png"
                                        },
                                        new PIC.NonVisualPictureDrawingProperties()),
                                    new PIC.BlipFill(
                                        new A.Blip(
                                            new A.BlipExtensionList(
                                                new A.BlipExtension()
                                                {
                                                    Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                })
                                        )
                                        {
                                            Embed = relationshipId,
                                            CompressionState = A.BlipCompressionValues.Print
                                        },
                                        new A.Stretch(new A.FillRectangle())),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset() { X = 0L, Y = 0L },
                                            new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                        new A.PresetGeometry(new A.AdjustValueList())
                                        { Preset = A.ShapeTypeValues.Rectangle }))
                            )
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
                    {
                        DistanceFromTop = (UInt32Value)0U,
                        DistanceFromBottom = (UInt32Value)0U,
                        DistanceFromLeft = (UInt32Value)0U,
                        DistanceFromRight = (UInt32Value)0U,
                        EditId = "50D07946"
                    });

            body.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(element)));
        }

        private DocumentFormat.OpenXml.Wordprocessing.Paragraph CreateParagraph(string text)
        {
            return new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(text)));
        }
    }
}
