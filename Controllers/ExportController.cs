using System;
using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Test.Data;
using Test.Models;
using DocumentFormat.OpenXml.Drawing;
using Word = DocumentFormat.OpenXml.Wordprocessing;
using Draw = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Globalization;


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
                    mainPart.Document = new Word.Document();
                    var body = mainPart.Document.AppendChild(new Word.Body());

                    // Set document styles
                    SetDocumentStyles(mainPart);

                    // Add images and letter details
                    AddImage(mainPart, body, "/logo/logo.png", 100, 100);
                    AddLetterDetails(body, letter);
                    AddImage(mainPart, body, "/logo/signature.png", 50, 50);
                    AddFooter(mainPart, body, "/logo/footer.png");
                }

                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "OfficialLetter.docx");
            }
        }

        private void AddFooter(MainDocumentPart mainPart, Word.Body body, string footerImagePath)
        {
            AddImage(mainPart, body, footerImagePath, 800, 100);  // Adjust size as necessary
            var footerParagraph = CreateParagraph("Confidential", true);
            body.AppendChild(footerParagraph);
        }

        // Helper method to add document styles (font, alignment, RTL)
        private void SetDocumentStyles(MainDocumentPart mainPart)
        {
            var styleDefinitionsPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            styleDefinitionsPart.Styles = new Word.Styles(
                new Word.Style()
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Normal",
                    StyleName = new Word.StyleName() { Val = "Normal" },
                    NextParagraphStyle = new Word.NextParagraphStyle() { Val = "Normal" },
                    StyleParagraphProperties = new Word.StyleParagraphProperties()
                    {
                        Justification = new Word.Justification() { Val = JustificationValues.Right },
                        BiDi = new Word.BiDi() { Val = OnOffValue.FromBoolean(true) }
                    },
                    StyleRunProperties = new Word.StyleRunProperties()
                    {
                        RunFonts = new Word.RunFonts() { Ascii = "Arial", ComplexScript = "Arial" },
                        FontSize = new Word.FontSize() { Val = "24" }
                    }
                }
            );
        }

        // Add letter details with proper styling and RTL direction
        private void AddLetterDetails(Word.Body body, Letter letter)
        {
            var persianDate = GetPersianDate();
            body.AppendChild(CreateParagraph($"شماره نامه: {letter.Id}", true));
            body.AppendChild(CreateParagraph($"تاریخ: {persianDate}", true));
            body.AppendChild(CreateParagraph($"موضوع: {letter.Subject}", true));
            body.AppendChild(CreateParagraph($"از: {letter.Sender}", true));
            body.AppendChild(CreateParagraph($"به: {letter.Receiver}", true));
            body.AppendChild(CreateParagraph($"توضیحات: {letter.Description}", true));
        }

        // Method to add images with size parameters
        private void AddImage(MainDocumentPart mainPart, Word.Body body, string imagePath, int width, int height)
        {
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (var imageStream = new FileStream($"wwwroot{imagePath}", FileMode.Open, FileAccess.Read))
            {
                imagePart.FeedData(imageStream);
            }

            var relationshipId = mainPart.GetIdOfPart(imagePart);
            var element = new Drawing(
                new DW.Inline(
                    new DW.Extent() { Cx = width * 9525L, Cy = height * 9525L },  // Adjust size
                    new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties() { Id = (UInt32Value)1U, Name = "Picture" },
                    new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "Picture" },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip() { Embed = relationshipId, CompressionState = A.BlipCompressionValues.Print },
                                    new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset() { X = 0L, Y = 0L },
                                        new A.Extents() { Cx = width * 9525L, Cy = height * 9525L }),
                                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
                { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U });

            body.AppendChild(new Word.Paragraph(new Word.Run(element)));
        }

        // Method to generate Persian date in format YYYY/MM/DD
        private string GetPersianDate()
        {
            var persianCalendar = new PersianCalendar();
            int year = persianCalendar.GetYear(DateTime.Now);
            int month = persianCalendar.GetMonth(DateTime.Now);
            int day = persianCalendar.GetDayOfMonth(DateTime.Now);
            return $"{year}/{month:D2}/{day:D2}";
        }

        // Helper method to create paragraphs with right alignment and RTL
        private Word.Paragraph CreateParagraph(string text, bool rightToLeft = false)
        {
            var paragraph = new Word.Paragraph(
                new Word.ParagraphProperties(
                    new Word.Justification()
                    {
                        Val = rightToLeft ? Word.JustificationValues.Right : Word.JustificationValues.Left
                    },
                    new Word.BiDi()
                    {
                        Val = OnOffValue.FromBoolean(rightToLeft)
                    }
                ),
                new Word.Run(new Word.Text(text))
            );
            return paragraph;
        }
    }
}
