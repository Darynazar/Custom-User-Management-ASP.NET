using Microsoft.AspNetCore.Identity;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Test.Models
{
    public class Letter
    {
        [Key]
        public int Id { get; set; }

        [Required]
        public string Title { get; set; }

        [Required]
        public string Subject { get; set; }

        [Required]
        public string Sender { get; set; }

        [Required]
        public string Receiver { get; set; }

        public string Description { get; set; }

        [Required]
        public bool Status { get; set; } = false;  // Default value is false (0)

        [Required]
        public string CurrentOrganization { get; set; }  // Tracks which organization the letter is in

        [ForeignKey("Category")]
        public int CategoryId { get; set; }
        public virtual Category? Category { get; set; }

        //[ForeignKey("User")]
        //public string UserId { get; set; }  // ForeignKey to AspNetUsers table
        //public virtual ApplicationUser User { get; set; }  // Assuming ApplicationUser class is the user model

        public string? UserId { get; set; }
        [ForeignKey("UserId")]
        public IdentityUser? User { get; set; }

        public string? Image { get; set; }  // Nullable image field
    }
}
