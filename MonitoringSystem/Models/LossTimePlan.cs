using System;
using System.ComponentModel.DataAnnotations;

namespace MonitoringSystem.Models
{
    public class LossTimePlan
    {
        [Key]
        public int Id { get; set; }

        [Required]
        public string Category { get; set; }

        [Required]
        public string MachineLine { get; set; }

        public int Month { get; set; }
        public int Year { get; set; }

        public double TargetMinutes { get; set; }

        public DateTime UploadedAt { get; set; } = DateTime.Now;
    }
}