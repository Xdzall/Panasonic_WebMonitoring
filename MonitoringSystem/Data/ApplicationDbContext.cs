using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;
using MonitoringSystem.Models;
using static MonitoringSystem.Pages.Shared.ApplyBreakFilterModel;
using static MonitoringSystem.Pages.Summary.SummaryModel;

namespace MonitoringSystem.Data
{
    public class ApplicationDbContext : IdentityDbContext<ApplicationUser>
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
            : base(options)
        {
        }

        public DbSet<LossTimePlan> LossTimePlans { get; set; }
        public DbSet<AdditionalBreakTime> AdditionalBreakTimes { get; set; }
        public DbSet<HourlyPlanData> HourlyPlanData { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            // Existing configurations...

            modelBuilder.Entity<AdditionalBreakTime>()
                .HasKey(b => b.Id);
        }
    }

}
