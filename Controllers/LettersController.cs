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

namespace Test.Controllers
{
    [Authorize]
    public class LettersController : Controller
    {
        private readonly ApplicationDbContext _context;

        public LettersController(ApplicationDbContext context)
        {
            _context = context;
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
