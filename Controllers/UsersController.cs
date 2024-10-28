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
    }
}

