using Backend.Models;
using Microsoft.EntityFrameworkCore;

namespace Backend.Context
{
    public class ExcelPreviewContext : DbContext
    {
        public ExcelPreviewContext(DbContextOptions options) : base(options)
        {
        }

        protected ExcelPreviewContext()
        {
        }

        // DBSets
        public virtual DbSet<ExcelData> ExcelDatas { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<ExcelData>(entity =>
            {
                entity.HasKey(e => e.Id);
                entity.Property(e => e.Name).IsRequired().HasMaxLength(100);
                entity.Property(e => e.Value).IsRequired().HasMaxLength(500);
            });

            
            base.OnModelCreating(modelBuilder);
        }
    }
}
