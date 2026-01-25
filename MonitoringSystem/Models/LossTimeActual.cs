using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace MonitoringSystem.Models
{
    public class LossTimeActual
    {
        [Key]
        public int Id { get; set; }

        [Required]
        public string Category { get; set; } = string.Empty;

        [Required]
        public string MachineLine { get; set; } = string.Empty;

        [Required]
        public int Day { get; set; } // Kolom Day yang Anda minta

        [Required]
        public int Month { get; set; }

        [Required]
        public int Year { get; set; }

        [Required]
        public double Minutes { get; set; }

        public DateTime CreatedAt { get; set; } = DateTime.Now;
    }
}