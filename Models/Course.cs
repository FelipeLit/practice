using System.ComponentModel.DataAnnotations;

namespace practice2.Models
{
    public class Course
    {
         public int Id { get; set; }
        [Required]
        public string Name { get; set; }
        public ICollection<Student> Students { get; set; }
    }
}