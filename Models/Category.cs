namespace Test.Models;

using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Collections.Generic;

public class Category
{
    [Key]
    public int Id { get; set; }


    public string Name { get; set; }

    public int? ParentId { get; set; }  // Nullable ParentId for nested categories


    public virtual Category? ParentCategory { get; set; }  // Self-referencing relationship

    public virtual ICollection<Category> ChildCategories { get; set; } = new List<Category>();  // Collection for child categories
}
