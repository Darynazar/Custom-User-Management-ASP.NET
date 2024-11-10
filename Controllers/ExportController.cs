using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Test.Data;
using Test.Models;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
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
                    var mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    var body = mainPart.Document.AppendChild(new Body());

                    SetDocumentStyles(mainPart);

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
                    AddSignature(mainPart, body);
                    AddFooter(mainPart, body, "/logo/footer.png", contentWidthTwips);
                }

                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "OfficialLetter.docx");
            }
        }

        private Table CreateHeaderTable(MainDocumentPart mainPart, Letter letter, int contentWidthTwips)
        {
            var headerTable = new Table(
                new TableProperties(
                    new TableWidth { Type = TableWidthUnitValues.Dxa, Width = contentWidthTwips.ToString() },
                    new TableStyle { Val = "TableGrid" },
                    new TableLook { FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true }
                ),
                new TableRow(
                    new TableCell(new Paragraph(new Run())), // Left cell for logo
                    new TableCell(new Paragraph(new Run())), // Center cell for god_name
                    new TableCell(new Paragraph(new Run()))  // Right cell for letter info
                )
            );

            // Logo on the left
            AddImageToCell(mainPart, headerTable.Elements<TableRow>().First().Elements<TableCell>().ElementAt(0), "/logo/logo.png", 80, 80);

            // God name image in the center
            AddImageToCell(mainPart, headerTable.Elements<TableRow>().First().Elements<TableCell>().ElementAt(1), "/logo/god_name.png", 80, 80);

            // Letter info on the right with RTL
            var letterInfoParagraph = new Paragraph(
                new ParagraphProperties(
                    new Justification { Val = JustificationValues.Left },
                    new BiDi { Val = OnOffValue.FromBoolean(true) },
                    new SpacingBetweenLines { Before = "240" },
                    new ParagraphMarkRunProperties(new RunProperties(new BiDi { Val = OnOffValue.FromBoolean(true) }))
                ),
                new Run(
                    new RunProperties(new BiDi { Val = OnOffValue.FromBoolean(true) }),
                    new Text($"شماره نامه {letter.Id} : ")
                ),
                new Run(new Break()),
                new Run(
                    new RunProperties(new BiDi { Val = OnOffValue.FromBoolean(true) }),
                    new Text($"تاریخ {GetPersianDate()} : ")
                )
            );
            headerTable.Elements<TableRow>().First().Elements<TableCell>().ElementAt(2).AppendChild(letterInfoParagraph);

            return headerTable;
        }

        private void AddFooter(MainDocumentPart mainPart, Body body, string footerImagePath, int contentWidthTwips)
        {
            AddImage(mainPart, body, footerImagePath, contentWidthTwips / 1440 * 96, 100);
            var footerParagraph = CreateParagraph("محرمانه", false);
            body.AppendChild(footerParagraph);
        }

        private void SetDocumentStyles(MainDocumentPart mainPart)
        {
            var styleDefinitionsPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            styleDefinitionsPart.Styles = new Styles(
                new Style
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Normal",
                    StyleName = new StyleName { Val = "Normal" },
                    NextParagraphStyle = new NextParagraphStyle { Val = "Normal" },
                    StyleParagraphProperties = new StyleParagraphProperties
                    {
                        Justification = new Justification { Val = JustificationValues.Right },
                        // Removed TextDirection
                    },
                    StyleRunProperties = new StyleRunProperties
                    {
                        RunFonts = new RunFonts { Ascii = "B Nazanin", ComplexScript = "B Nazanin" },
                        FontSize = new FontSize { Val = "24" }
                        // Removed BiDi
                    }
                },
                new Style
                {
                    Type = StyleValues.Paragraph,
                    StyleId = "Title",
                    StyleName = new StyleName { Val = "Title" },
                    StyleParagraphProperties = new StyleParagraphProperties
                    {
                        Justification = new Justification { Val = JustificationValues.Center }
                        // Removed TextDirection
                    },
                    StyleRunProperties = new StyleRunProperties
                    {
                        RunFonts = new RunFonts { Ascii = "B Nazanin", ComplexScript = "B Nazanin" },
                        FontSize = new FontSize { Val = "32" },
                        Bold = new Bold()
                        // Removed BiDi
                    }
                }
            );
        }


        private void AddImageToCell(MainDocumentPart mainPart, TableCell cell, string imagePath, int width, int height)
        {
            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (var imageStream = new FileStream($"wwwroot{imagePath}", FileMode.Open, FileAccess.Read))
            {
                imagePart.FeedData(imageStream);
            }

            var relationshipId = mainPart.GetIdOfPart(imagePart);
            var element = CreateDrawingElement(relationshipId, width, height);

            cell.AppendChild(new Paragraph(new Run(element)));
        }

        private void AddImage(MainDocumentPart mainPart, Body body, string imagePath, int width, int height)
        {
            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (var imageStream = new FileStream($"wwwroot{imagePath}", FileMode.Open, FileAccess.Read))
            {
                imagePart.FeedData(imageStream);
            }

            var relationshipId = mainPart.GetIdOfPart(imagePart);
            var element = CreateDrawingElement(relationshipId, width, height);

            body.AppendChild(new Paragraph(new Run(element)));
        }

        private Drawing CreateDrawingElement(string relationshipId, int width, int height)
        {
            return new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = width * 9525L, Cy = height * 9525L },
                    new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties { Id = (UInt32Value)1U, Name = "Picture" },
                    new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties { Id = (UInt32Value)0U, Name = "Picture" },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip { Embed = relationshipId, CompressionState = A.BlipCompressionValues.Print },
                                    new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset { X = 0L, Y = 0L },
                                        new A.Extents { Cx = width * 9525L, Cy = height * 9525L }),
                                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }))
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
                { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U });
        }

        private Paragraph CreateBoldTitleParagraph(string text)
        {
            var paragraphProperties = new ParagraphProperties(
                new Justification { Val = JustificationValues.Center },
                new SpacingBetweenLines { After = "240" },
                new BiDi { Val = OnOffValue.FromBoolean(true) }
            );

            var runProperties = new RunProperties(
                new Bold(),
                new FontSize { Val = "32" },
                new BiDi { Val = OnOffValue.FromBoolean(true) }
            );

            var run = new Run(runProperties, new Text(text));
            return new Paragraph(paragraphProperties, run);
        }

        private string GetPersianDate()
        {
            var persianCalendar = new PersianCalendar();
            var now = DateTime.Now;
            return $"{persianCalendar.GetYear(now)}/{persianCalendar.GetMonth(now):D2}/{persianCalendar.GetDayOfMonth(now):D2}";
        }


        private Paragraph CreateParagraph(string text, bool rightToLeft = true)
        {
            var paragraphProperties = new ParagraphProperties
            {
                Justification = new Justification { Val = JustificationValues.Right }
                // Removed TextDirection
            };

            paragraphProperties.Append(new Languages { Val = "fa-IR" });

            var runProperties = new RunProperties
            {
                RunFonts = new RunFonts { Ascii = "B Nazanin", ComplexScript = "B Nazanin" }
                // Removed BiDi
            };

            var run = new Run(runProperties, new Text(text));
            return new Paragraph(paragraphProperties, run);
        }

        private void AddLetterDetails(Body body, Letter letter)
        {
            body.AppendChild(CreateParagraph($"{letter.Subject}  : موضوع", true));
            body.AppendChild(CreateParagraph($"{letter.Sender} : از", true));
            body.AppendChild(CreateParagraph($"{letter.Receiver} : به", true));
            body.AppendChild(CreateParagraph($"{letter.CurrentOrganization} : کارشناس", true));
            body.AppendChild(CreateParagraph($"{letter.Description} : توضیحات", true));
        }

        private void AddSignature(MainDocumentPart mainPart, Body body)
        {
            // Add "Signature:" text
            var signatureParagraph = new Paragraph(
                new ParagraphProperties(
                    new Justification { Val = JustificationValues.Left },
                    new Indentation { Left = "800" }  // 20px padding (approx. 400 twips)
                ),
                new Run(
                    new RunProperties(new Bold()),
                    new Text(": امضا")
                )
            );
            body.AppendChild(signatureParagraph);

            // Create a new paragraph for the signature image with padding
            var signatureImageParagraph = new Paragraph(
                new ParagraphProperties(
                    new Justification { Val = JustificationValues.Left },
                    new Indentation { Left = "800" }  // 20px padding (approx. 400 twips)
                )
            );

            var imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (var imageStream = new FileStream($"wwwroot/logo/signature.png", FileMode.Open, FileAccess.Read))
            {
                imagePart.FeedData(imageStream);
            }

            var relationshipId = mainPart.GetIdOfPart(imagePart);
            var element = CreateDrawingElement(relationshipId, 50, 50);

            signatureImageParagraph.AppendChild(new Run(element));
            body.AppendChild(signatureImageParagraph);
        }
    }
}