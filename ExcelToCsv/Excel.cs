using System.ComponentModel.DataAnnotations;

namespace ExcelToCsv
{
    public class Excel
    {
        [Key]
        public int PID { get; set; }
        [Required]
        public string ProductId { get; set; }
        [Required]
        public string MfrName { get; set; }
        [Required]
        public string MfrPN { get; set; }
        [Required]
        public double Cost { get; set; }
        [Required]
        public string COO { get; set; }
        [Required]
        public string ShortDescription { get; set; }
        public string UPC { get; set; }
        [Required]
        public string UOM { get; set; }
    }
}
