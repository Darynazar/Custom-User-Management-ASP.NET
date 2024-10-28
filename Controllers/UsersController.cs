using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Test.Data;
using Test.Models;
using Microsoft.AspNetCore.Identity;
using System.Threading.Tasks;

namespace Test.Controllers
{
    public class UsersController : Controller
    {
       private readonly UserManager<IdentityUser> _userManager;

         public UsersController(UserManager<IdentityUser> userManager)
         {
              _userManager = userManager;
         }

         // GET: Users
          public async Task<IActionResult> Index()
          {
                var users = await _userManager.Users.ToListAsync();
                return View(users);
          }

        // GET: Users/Create
        public IActionResult Create()
        {
            return View(new IdentityUser());
        }

        // POST: Users/Create
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("UserName,Email,PasswordHash")] IdentityUser user)
        {
            if (ModelState.IsValid)
            {
                var result = await _userManager.CreateAsync(user, user.PasswordHash);

                if (result.Succeeded)
                {
                    return RedirectToAction(nameof(Index));
                }

                foreach (var error in result.Errors)
                {
                    ModelState.AddModelError(string.Empty, error.Description);
                }
            }
            return View(user);
        }
    }
}

