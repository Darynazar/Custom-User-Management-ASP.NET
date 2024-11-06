using Microsoft.AspNetCore.Identity;
using Test.Models;

namespace Test.Models
{
    public class UserCategory
    {
        public int Id { get; set; }  // Primary Key

        public string UserId { get; set; }
        public ApplicationUser ApplicationUser { get; set; }

        public int CategoryId { get; set; }
        public Category Category { get; set; }
    }
}