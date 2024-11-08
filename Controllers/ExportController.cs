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

                    // Create a table for the header
                    var table = new Word.Table(
                        new Word.TableProperties(
                            new Word.TableStyle() { Val = "TableGrid" },
                            new Word.TableWidth() { Type = Word.TableWidthUnitValues.Auto, Width = "5000" } // Use Auto and specify width
                        ),
                        new Word.TableRow(
                            new Word.TableCell(new Word.Paragraph()),
                            new Word.TableCell(new Word.Paragraph()),
                            new Word.TableCell(
                                new Word.Paragraph(
                                    new Word.ParagraphProperties(
                                        new Word.SpacingBetweenLines() { Before = "240" }
                                    ),
                                    new Word.Run(new Word.Text($"شماره نامه: {letter.Id}")),
                                    new Word.Run(new Word.Break()),
                                    new Word.Run(new Word.Text($"تاریخ: {GetPersianDate()}"))
                                )
                            )
                        )
                    );

                    body.AppendChild(table);

                    // Add the logo and god_name image in the first two cells
                    AddImageToCell(mainPart, table.Elements<Word.TableRow>().First().Elements<Word.TableCell>().First(), "/logo/logo.png", 100, 100);
                    AddImageToCell(mainPart, table.Elements<Word.TableRow>().First().Elements<Word.TableCell>().ElementAt(1), "/logo/god_name.png", 100, 100);

                    // Add title
                    var titleParagraph = CreateBoldTitleParagraph(letter.Title);
                    body.AppendChild(titleParagraph);

                    // Add letter details
                    AddLetterDetails(body, letter);

                    // Add signature and footer
                    AddImage(mainPart, body, "/logo/signature.png", 50, 50);
                    AddFooter(mainPart, body, "/logo/footer.png");
                }

                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "OfficialLetter.docx");
            }
        }

        // Helper method to add images to specific table cells
        private void AddImageToCell(MainDocumentPart mainPart, Word.TableCell cell, string imagePath, int width, int height)
        {
            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
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

            cell.AppendChild(new Word.Paragraph(new Word.Run(element)));
        }

        // Create a bold title paragraph
        private Word.Paragraph CreateBoldTitleParagraph(string text)
        {
            var paragraphProperties = new Word.ParagraphProperties(
                new Word.Justification() { Val = Word.JustificationValues.Center },
                new Word.SpacingBetweenLines() { After = "240" } // Adjust the spacing after the title
            );

            var runProperties = new Word.RunProperties(
                new Word.Bold(),  // Make the title bold
                new Word.FontSize() { Val = "32" } // Adjust font size as needed
            );

            var run = new Word.Run(runProperties, new Word.Text(text));

            return new Word.Paragraph(paragraphProperties, run);
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
                },
                new Word.Style()
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Title",
                    StyleName = new Word.StyleName() { Val = "Title" },
                    StyleParagraphProperties = new Word.StyleParagraphProperties()
                    {
                        Justification = new Word.Justification() { Val = JustificationValues.Center }
                    },
                    StyleRunProperties = new Word.StyleRunProperties()
                    {
                        RunFonts = new Word.RunFonts() { Ascii = "Arial", ComplexScript = "Arial" },
                        FontSize = new Word.FontSize() { Val = "32" },
                        Bold = new Word.Bold()
                    }
                }
            );
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

        private void AddHeaderRow(MainDocumentPart mainPart, Word.Body body, string logoPath, string godNameImagePath, int letterId, string persianDate)
        {
            // Create a table for the header
            var table = new Word.Table(
                new Word.TableProperties(
                    new Word.TableStyle() { Val = "TableGrid" },
                    new Word.TableWidth() { Type = Word.TableWidthUnitValues.Auto }
                ),
                new Word.TableRow(
                    new Word.TableCell(new Word.Paragraph(new Word.Run(new Word.Text("Logo")))),
                    new Word.TableCell(new Word.Paragraph(new Word.Run(new Word.Text("God Name")))),
                    new Word.TableCell(new Word.Paragraph(new Word.Run(new Word.Text($"Letter No: {letterId}, Date: {persianDate}"))))
                )
            );

            body.AppendChild(table);

            // Add the logo and god_name image in one row (using table cell contents)
            AddImage(mainPart, body, logoPath, 100, 100);
            AddImage(mainPart, body, godNameImagePath, 100, 100);
        }



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

      

        // Create a Title Paragraph with Spacing After
        private Word.Paragraph CreateTitleParagraph(string text)
        {
            var paragraphProperties = new Word.ParagraphProperties(
                new Word.Justification() { Val = Word.JustificationValues.Center },
                new Word.SpacingBetweenLines() { After = "240" } // Adjust the spacing after the title
            );

            var paragraph = new Word.Paragraph(paragraphProperties,
                new Word.Run(new Word.Text(text))
            );
            return paragraph;
        }


        

    }
}
