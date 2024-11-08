using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Test.Data;
using Test.Models;
using Microsoft.AspNetCore.Authorization;
using SelectPdf;
using Microsoft.AspNetCore.Mvc.ViewEngines;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using HtmlToOpenXml;
using System.Text;
using System.IO;

namespace Test.Controllers
{
    [Authorize]
    public class LettersController : Controller
    {
        private readonly ApplicationDbContext _context;
        private readonly ICompositeViewEngine _viewEngine;
        private readonly IServiceProvider _serviceProvider;

        public LettersController(ApplicationDbContext context, ICompositeViewEngine viewEngine, IServiceProvider serviceProvider)
        {
            _context = context;
            _viewEngine = viewEngine;
            _serviceProvider = serviceProvider;
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

            // Render the view to a string
            var htmlContent = await RenderViewToStringAsync("OfficialLetter", letter);

            // Convert HTML to Word
            using (var stream = new MemoryStream())
            {
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                    mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();
                    var converter = new HtmlConverter(mainPart);
                    var body = mainPart.Document.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Body());

                    converter.ParseHtml(htmlContent);

                    mainPart.Document.Save();
                }

                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "OfficialLetter.docx");
            }
        }


        public async Task<IActionResult> ExportToPdf(int id)
        {
            var letter = await _context.Letters
                .Include(l => l.Category)
                .Include(l => l.User)
                .FirstOrDefaultAsync(m => m.Id == id);

            if (letter == null)
            {
                return NotFound();
            }

            // Render the view to a string
            var htmlContent = await RenderViewToStringAsync("OfficialLetter", letter);

            // Convert HTML to PDF
            HtmlToPdf converter = new HtmlToPdf();
            PdfDocument doc = converter.ConvertHtmlString(htmlContent);

            // Save to a byte array
            byte[] pdf = doc.Save();

            // Close the document
            doc.Close();

            // Return the PDF file
            return File(pdf, "application/pdf", "OfficialLetter.pdf");
        }

        private async Task<string> RenderViewToStringAsync(string viewName, object model)
        {
            ViewData.Model = model;
            using (var writer = new StringWriter())
            {
                var viewResult = _viewEngine.FindView(ControllerContext, viewName, false);
                if (viewResult.View == null)
                {
                    throw new ArgumentNullException($"{viewName} does not match any available view");
                }
                var viewContext = new ViewContext(
                    ControllerContext,
                    viewResult.View,
                    ViewData,
                    TempData,
                    writer,
                    new HtmlHelperOptions()
                );
                await viewResult.View.RenderAsync(viewContext);
                return writer.GetStringBuilder().ToString();
            }
        }

        // GET: Letters/Display/5
        public async Task<IActionResult> Display(int id)
        {
            var letter = await _context.Letters
                .Include(l => l.Category)
                .Include(l => l.User)
                .FirstOrDefaultAsync(m => m.Id == id);

            if (letter == null)
            {
                return NotFound();
            }

            return View(letter);
        }

        // GET: Letters
        public async Task<IActionResult> Index()
        {
            var applicationDbContext = _context.Letters.Include(l => l.Category).Include(l => l.User);
            return View(await applicationDbContext.ToListAsync());
        }

        // GET: Letters/Details/5
        public async Task<IActionResult> Details(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var letter = await _context.Letters
                .Include(l => l.Category)
                .Include(l => l.User)
                .FirstOrDefaultAsync(m => m.Id == id);
            if (letter == null)
            {
                return NotFound();
            }

            return View(letter);
        }

        // GET: Letters/Create
        public IActionResult Create()
        {
            ViewData["CategoryId"] = new SelectList(_context.Categories, "Id", "Name");
            ViewData["UserId"] = new SelectList(_context.Users, "Id", "Id");
            return View();
        }

        // POST: Letters/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("Id,Title,Subject,Sender,Receiver,Description,Status,CurrentOrganization,CategoryId,UserId")] Letter letter, IFormFile imageFile)
        {
            if (ModelState.IsValid)
            {
                if (imageFile != null && imageFile.Length > 0)
                {
                    // Generate a unique file name
                    var fileName = Path.GetFileName(imageFile.FileName);
                    var filePath = Path.Combine("wwwroot/images", fileName);

                    // Save the file
                    using (var stream = new FileStream(filePath, FileMode.Create))
                    {
                        await imageFile.CopyToAsync(stream);
                    }

                    // Store the file path in the database
                    letter.Image = $"/images/{fileName}";
                }

                _context.Add(letter);
                await _context.SaveChangesAsync();
                return RedirectToAction(nameof(Index));
            }

            ViewData["CategoryId"] = new SelectList(_context.Categories, "Id", "Name", letter.CategoryId);
            ViewData["UserId"] = new SelectList(_context.Users, "Id", "Id", letter.UserId);
            return View(letter);
        }

        // GET: Letters/Edit/5
        public async Task<IActionResult> Edit(int? id)
        {
          if (id == null) return NotFound();

          var letter = await _context.Letters.FindAsync(id);
          if (letter == null) return NotFound();

          ViewData["CategoryId"] = new SelectList(_context.Categories, "Id", "Name", letter.CategoryId);
          ViewData["UserId"] = new SelectList(_context.Users, "Id", "Id", letter.UserId); // Set current UserId
          return View(letter);
        }


        // POST: Letters/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
      [HttpPost]
      [ValidateAntiForgeryToken]
      public async Task<IActionResult> Edit(int id, [Bind("Id,Title,Subject,Sender,Receiver,Description,Status,CurrentOrganization,CategoryId,UserId,Image")] Letter letter, IFormFile? imageFile)
      {
          if (id != letter.Id)
          {
              return NotFound();
          }

          if (ModelState.IsValid)
          {
              try
              {
                  var existingLetter = await _context.Letters.AsNoTracking().FirstOrDefaultAsync(l => l.Id == id);

                  // Remove the old image if a new one is uploaded
                  if (imageFile != null && imageFile.Length > 0)
                  {
                      if (!string.IsNullOrEmpty(existingLetter.Image))
                      {
                          var oldFilePath = Path.Combine("wwwroot", existingLetter.Image.TrimStart('/'));
                          if (System.IO.File.Exists(oldFilePath))
                          {
                              System.IO.File.Delete(oldFilePath);
                          }
                      }

                      // Save new image
                      var fileName = Path.GetFileName(imageFile.FileName);
                      var filePath = Path.Combine("wwwroot/images", fileName);

                      using (var stream = new FileStream(filePath, FileMode.Create))
                      {
                          await imageFile.CopyToAsync(stream);
                      }

                      letter.Image = $"/images/{fileName}";
                  }
                  else
                  {
                      letter.Image = existingLetter.Image;
                  }

                  _context.Update(letter);
                  await _context.SaveChangesAsync();
              }
              catch (DbUpdateConcurrencyException)
              {
                  if (!LetterExists(letter.Id))
                  {
                      return NotFound();
                  }
                  else
                  {
                      throw;
                  }
              }
              return RedirectToAction(nameof(Index));
          }
          ViewData["CategoryId"] = new SelectList(_context.Categories, "Id", "Name", letter.CategoryId);
          return View(letter);
      }

        // GET: Letters/Delete/5
        public async Task<IActionResult> Delete(int? id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var letter = await _context.Letters
                .Include(l => l.Category)
                .Include(l => l.User)
                .FirstOrDefaultAsync(m => m.Id == id);
            if (letter == null)
            {
                return NotFound();
            }

            return View(letter);
        }

        // POST: Letters/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            var letter = await _context.Letters.FindAsync(id);
            if (letter != null)
            {
                _context.Letters.Remove(letter);
            }

            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        private bool LetterExists(int id)
        {
            return _context.Letters.Any(e => e.Id == id);
        }
    }
}
