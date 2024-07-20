using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace practice2.Models
{
    public class Student
    {
        public int Id { get; set; }
        [Required]
        public string Name { get; set; }
        [Required]
        public int Age { get; set; }
        
        [ForeignKey("Curso")]
        public int CourseId { get; set; }
        public Course Course { get; set; }
    }
}