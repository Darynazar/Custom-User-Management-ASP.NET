using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Identity;

using Test.Models;
using Test.Data.ModelPropert;
using System.Diagnostics.Metrics;
using System.Reflection.Emit;

namespace Test.Data
{
    public class ApplicationDbContext : IdentityDbContext<IdentityUser>
    {
        public DbSet<Permission> Permissions { get; set; }
        public DbSet<RolePermission> RolePermissions { get; set; }
        public DbSet<Category> Categories { get; set; }
        public DbSet<Letter> Letters { get; set; }
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {
        }
        protected override void OnModelCreating(ModelBuilder builder)
        {
            base.OnModelCreating(builder);

            new CategoryConfiguration().Configure(builder.Entity<Category>());
            new LetterConfiguration().Configure(builder.Entity<Letter>());

            builder.Entity<RolePermission>()
                .HasKey(rp => new { rp.RoleId, rp.PermissionId });

            builder.Entity<RolePermission>()
                .HasOne(rp => rp.Role)
                .WithMany()
                .HasForeignKey(rp => rp.RoleId);

            builder.Entity<RolePermission>()
                .HasOne(rp => rp.Permission)
                .WithMany()
                .HasForeignKey(rp => rp.PermissionId);
        }
        public DbSet<Test.Models.Category> Category { get; set; } = default!;
    }
}

//using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
//using Microsoft.EntityFrameworkCore;
//using Test.Data.ModelPropert;
//using Test.Models;
//
//namespace Test.Data
//{
//    public class ApplicationDbContext : IdentityDbContext<ApplicationUser>
//    {
//        public DbSet<Permission> Permissions { get; set; }
//        public DbSet<RolePermission> RolePermissions { get; set; }
//        public DbSet<Category> Categories { get; set; }
//        public DbSet<Letter> Letters { get; set; }
//        public DbSet<UserCategory> UserCategories { get; set; }
//
//        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
//            : base(options)
//        {
//        }
//
//        protected override void OnModelCreating(ModelBuilder builder)
//        {
//            base.OnModelCreating(builder);
//
//            // Many-to-many relationship setup between User and Category
//            builder.Entity<UserCategory>()
//                .HasKey(uc => new { uc.UserId, uc.CategoryId });
//
//            builder.Entity<UserCategory>()
//                .HasOne(uc => uc.ApplicationUser)
//                .WithMany(u => u.UserCategories)
//                .HasForeignKey(uc => uc.UserId);
//
//            builder.Entity<UserCategory>()
//                .HasOne(uc => uc.Category)
//                .WithMany(c => c.UserCategories)
//                .HasForeignKey(uc => uc.CategoryId);
//
//            // RolePermission configuration
//            builder.Entity<RolePermission>()
//                .HasKey(rp => new { rp.RoleId, rp.PermissionId });
//
//            builder.Entity<RolePermission>()
//                .HasOne(rp => rp.Role)
//                .WithMany()
//                .HasForeignKey(rp => rp.RoleId);
//
//            builder.Entity<RolePermission>()
//                .HasOne(rp => rp.Permission)
//                .WithMany()
//                .HasForeignKey(rp => rp.PermissionId);
//        }
//    }
//
//}
