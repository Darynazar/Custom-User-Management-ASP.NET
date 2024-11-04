
using global::Test.Models;
using Test.Models;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;

namespace Test.Data.ModelPropert
{
    public class LetterConfiguration : IEntityTypeConfiguration<Letter>
    {
        public void Configure(EntityTypeBuilder<Letter> builder)
        {
            builder.HasKey(l => l.Id);
            builder.Property(l => l.Title).IsRequired();
            builder.Property(l => l.Subject).IsRequired();
            builder.Property(l => l.Sender).IsRequired();
            builder.Property(l => l.Receiver).IsRequired();
            builder.Property(l => l.Status).HasDefaultValue(false);  // Set default value for Status
            builder.Property(l => l.CurrentOrganization).IsRequired();

            builder.HasOne(l => l.Category)
                .WithMany()
                .HasForeignKey(l => l.CategoryId);

            builder.HasOne(l => l.User)
                .WithMany()
                .HasForeignKey(l => l.UserId);

            builder.Property(l => l.Image)
                .IsRequired(false);  // Nullable image
        }
    }
}
