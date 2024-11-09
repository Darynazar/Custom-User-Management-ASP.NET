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
                    var mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Word.Document();
                    var body = mainPart.Document.AppendChild(new Word.Body());

                    SetDocumentStyles(mainPart);

                    // Define width for 75% of standard page width in twips (12000 twips is 100%)
                    var contentWidthTwips = 9000;

                    // Create and add header
                    var headerTable = CreateHeaderTable(mainPart, letter, contentWidthTwips);
                    body.AppendChild(headerTable);

                    // Add title
                    var titleParagraph = CreateBoldTitleParagraph(letter.Title);
                    body.AppendChild(titleParagraph);

                    // Add letter details
                    AddLetterDetails(body, letter);

                    // Add signature and footer
                    AddImage(mainPart, body, "/logo/signature.png", 50, 50);
                    AddFooter(mainPart, body, "/logo/footer.png", contentWidthTwips);
                }

                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "OfficialLetter.docx");
            }
        }

        private Word.Table CreateHeaderTable(MainDocumentPart mainPart, Letter letter, int contentWidthTwips)
        {
            var headerTable = new Word.Table(
                new Word.TableProperties(
                    new Word.TableWidth() { Type = Word.TableWidthUnitValues.Dxa, Width = contentWidthTwips.ToString() },
                    new Word.TableStyle() { Val = "TableGrid" }
                ),
                new Word.TableRow(
                    new Word.TableCell(new Word.Paragraph(new Word.Run())),
                    new Word.TableCell(new Word.Paragraph(new Word.Run())),
                    new Word.TableCell(new Word.Paragraph(new Word.Run()))
                )
            );

            // Add images and text in respective cells
            AddImageToCell(mainPart, headerTable.Elements<Word.TableRow>().First().Elements<Word.TableCell>().ElementAt(0), "/logo/logo.png", 80, 80);
            AddImageToCell(mainPart, headerTable.Elements<Word.TableRow>().First().Elements<Word.TableCell>().ElementAt(1), "/logo/god_name.png", 80, 80);

            var letterInfoParagraph = new Word.Paragraph(
                new Word.ParagraphProperties(
                    new Word.Justification() { Val = Word.JustificationValues.Right },
                    new Word.SpacingBetweenLines() { Before = "240" }
                ),
                new Word.Run(new Word.Text($"شماره نامه  {letter.Id} : ")),
                new Word.Run(new Word.Break()),
                new Word.Run(new Word.Text($"تاریخ  {GetPersianDate()} : "))
            );
            headerTable.Elements<Word.TableRow>().First().Elements<Word.TableCell>().ElementAt(2).AppendChild(letterInfoParagraph);

            return headerTable;
        }

        private void AddFooter(MainDocumentPart mainPart, Word.Body body, string footerImagePath, int contentWidthTwips)
        {
            AddImage(mainPart, body, footerImagePath, contentWidthTwips / 1440 * 96, 100);
            var footerParagraph = CreateParagraph("Confidential", true);
            body.AppendChild(footerParagraph);
        }

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
                        RunFonts = new Word.RunFonts() { Ascii = "B Nazanin", ComplexScript = "B Nazanin" },
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
                        RunFonts = new Word.RunFonts() { Ascii = "B Nazanin", ComplexScript = "B Nazanin" },
                        FontSize = new Word.FontSize() { Val = "32" },
                        Bold = new Word.Bold()
                    }
                }
            );
        }
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
                    new DW.Extent() { Cx = width * 9525L, Cy = height * 9525L },
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
        private void AddImage(MainDocumentPart mainPart, Word.Body body, string imagePath, int width, int height)
        {
            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (var imageStream = new FileStream($"wwwroot{imagePath}", FileMode.Open, FileAccess.Read))
            {
                imagePart.FeedData(imageStream);
            }

            var relationshipId = mainPart.GetIdOfPart(imagePart);
            var element = new Drawing(
                new DW.Inline(
                    new DW.Extent() { Cx = width * 9525L, Cy = height * 9525L },
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

            var paragraph = new Word.Paragraph(new Word.Run(element));
            body.AppendChild(paragraph);
        }
        private Word.Paragraph CreateBoldTitleParagraph(string text)
        {
            var paragraphProperties = new Word.ParagraphProperties(
                new Word.Justification() { Val = Word.JustificationValues.Center },
                new Word.SpacingBetweenLines() { After = "240" }
            );

            var runProperties = new Word.RunProperties(
                new Word.Bold(),
                new Word.FontSize() { Val = "32" }
            );

            var run = new Word.Run(runProperties, new Word.Text(text));
            return new Word.Paragraph(paragraphProperties, run);
        }
        private string GetPersianDate()
        {
            var persianCalendar = new PersianCalendar();
            var now = DateTime.Now;
            return $"{persianCalendar.GetYear(now)}/{persianCalendar.GetMonth(now):D2}/{persianCalendar.GetDayOfMonth(now):D2}";
        }
        // Helper method to create paragraphs with right alignment and RTL
        private Word.Paragraph CreateParagraph(string text, bool rightToLeft = false)
        {
            // Paragraph Properties
            var paragraphProperties = new Word.ParagraphProperties();
            if (rightToLeft)
            { 
                // Set BiDi for RTL
                paragraphProperties.BiDi = new Word.BiDi() { Val = OnOffValue.FromBoolean(true) };
                // Set text direction to Right to Left
                paragraphProperties.TextDirection = new Word.TextDirection()
                {
                    Val = Word.TextDirectionValues.TopToBottomRightToLeft
                };
                // Ensure Right alignment
                paragraphProperties.Justification = new Word.Justification()
                {
                    Val = Word.JustificationValues.Right
                };
                // Set the language (usually Arabic or Persian for RTL)
                var lang = new Word.Languages() { Val = "ar-SA" }; // "ar-SA" for Arabic; "fa-IR" for Persian (Farsi)
                paragraphProperties.Append(lang);
            }
            else
            {
                // Left alignment if not RTL
                paragraphProperties.Justification = new Word.Justification()
                {
                    Val = Word.JustificationValues.Left
                };
            }
            // Run Properties (applies to the text itself)
            var runProperties = new Word.RunProperties();
            if (rightToLeft)
            {
                // Set an RTL font (B Nazanin or any other font you need)
                runProperties.Append(new Word.RunFonts() { Ascii = "B Nazanin", ComplexScript = "B Nazanin" });
                // Apply BiDi to the run
                runProperties.Append(new Word.BiDi() { Val = OnOffValue.FromBoolean(true) });
            }

            // Create the Run and Text objects
            var run = new Word.Run(runProperties, new Word.Text(text));

            // Return the paragraph with the run
            return new Word.Paragraph(paragraphProperties, run);
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
            body.AppendChild(CreateParagraph($"موضوع: {letter.Subject}", true));
            body.AppendChild(CreateParagraph($"از: {letter.Sender}", true));
            body.AppendChild(CreateParagraph($"به: {letter.Receiver}", true));
            body.AppendChild(CreateParagraph($"کارشناس: {letter.CurrentOrganization}", true));
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
