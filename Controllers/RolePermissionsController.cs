using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Test.Data;
using Test.Models;

namespace Test.Controllers
{
    public class RolePermissionsController : Controller
    {
        private readonly ApplicationDbContext _context;

        public RolePermissionsController(ApplicationDbContext context)
        {
            _context = context;
        }

        // GET: RolePermissions
        public async Task<IActionResult> Index()
        {
            var applicationDbContext = _context.RolePermissions.Include(r => r.Permission).Include(r => r.Role);
            return View(await applicationDbContext.ToListAsync());
        }

        // GET: RolePermissions/Details/5
        public async Task<IActionResult> Details(string id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var rolePermission = await _context.RolePermissions
                .Include(r => r.Permission)
                .Include(r => r.Role)
                .FirstOrDefaultAsync(m => m.RoleId == id);
            if (rolePermission == null)
            {
                return NotFound();
            }

            return View(rolePermission);
        }

        // GET: RolePermissions/Create
        public IActionResult Create()
        {
            ViewData["PermissionId"] = new SelectList(_context.Permissions, "Id", "Name");
            ViewData["RoleId"] = new SelectList(_context.Roles, "Id", "Name");
            return View();
        }


        // POST: RolePermissions/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(string roleId, string permissionId)
        {
            if (string.IsNullOrEmpty(roleId) || string.IsNullOrEmpty(permissionId))
            {
                ModelState.AddModelError("", "Role and Permission must be selected.");
            }
            else
            {
                // Check if the role and permission exist
                var role = await _context.Roles.FindAsync(roleId);

                // Convert permissionId to int before finding the permission
                if (int.TryParse(permissionId, out int permId))
                {
                    var permission = await _context.Permissions.FindAsync(permId);

                    if (role == null)
                    {
                        ModelState.AddModelError("RoleId", "Selected role does not exist.");
                    }
                    if (permission == null)
                    {
                        ModelState.AddModelError("PermissionId", "Selected permission does not exist.");
                    }

                    if (ModelState.IsValid) // Proceed only if ModelState is valid
                    {
                        var rolePermission = new RolePermission
                        {
                            RoleId = roleId,
                            PermissionId = permId // Use the parsed permission ID
                        };

                        _context.Add(rolePermission);
                        await _context.SaveChangesAsync();
                        return RedirectToAction(nameof(Index));
                    }
                }
                else
                {
                    ModelState.AddModelError("PermissionId", "Invalid permission ID.");
                }
            }

            // If we got this far, something failed, redisplay form
            ViewData["PermissionId"] = new SelectList(_context.Permissions, "Id", "Name");
            ViewData["RoleId"] = new SelectList(_context.Roles, "Id", "Name");
            return View();
        }

        // GET: RolePermissions/Edit/5
        public async Task<IActionResult> Edit(string id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var rolePermission = await _context.RolePermissions.FindAsync(id);
            if (rolePermission == null)
            {
                return NotFound();
            }
            ViewData["PermissionId"] = new SelectList(_context.Permissions, "Id", "Id", rolePermission.PermissionId);
            ViewData["RoleId"] = new SelectList(_context.Roles, "Id", "Id", rolePermission.RoleId);
            return View(rolePermission);
        }

        // POST: RolePermissions/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to.
        // For more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(string id, [Bind("RoleId,PermissionId")] RolePermission rolePermission)
        {
            if (id != rolePermission.RoleId)
            {
                return NotFound();
            }

            if (ModelState.IsValid)
            {
                try
                {
                    _context.Update(rolePermission);
                    await _context.SaveChangesAsync();
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!RolePermissionExists(rolePermission.RoleId))
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
            ViewData["PermissionId"] = new SelectList(_context.Permissions, "Id", "Id", rolePermission.PermissionId);
            ViewData["RoleId"] = new SelectList(_context.Roles, "Id", "Id", rolePermission.RoleId);
            return View(rolePermission);
        }

        // GET: RolePermissions/Delete/5
        public async Task<IActionResult> Delete(string id)
        {
            if (id == null)
            {
                return NotFound();
            }

            var rolePermission = await _context.RolePermissions
                .Include(r => r.Permission)
                .Include(r => r.Role)
                .FirstOrDefaultAsync(m => m.RoleId == id);
            if (rolePermission == null)
            {
                return NotFound();
            }

            return View(rolePermission);
        }

        // POST: RolePermissions/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(string id)
        {
            var rolePermission = await _context.RolePermissions.FindAsync(id);
            if (rolePermission != null)
            {
                _context.RolePermissions.Remove(rolePermission);
            }

            await _context.SaveChangesAsync();
            return RedirectToAction(nameof(Index));
        }

        private bool RolePermissionExists(string id)
        {
            return _context.RolePermissions.Any(e => e.RoleId == id);
        }
    }
}
