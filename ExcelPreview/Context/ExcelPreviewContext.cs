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

                // Seed Data
                entity.HasData(GenerateSeedData());
            });

            
            base.OnModelCreating(modelBuilder);
        }

        private static ExcelData[] GenerateSeedData()
        {
            var random = new Random(42); // Fixed seed for consistent data across environments
            var seedData = new ExcelData[10]; // Generate 10 seed records

            for (int i = 0; i < seedData.Length; i++)
            {
                seedData[i] = new ExcelData
                {
                    Id = i + 1, // Auto-increment IDs starting from 1
                    Name = Ulid.NewUlid().ToString(), // Generate ULID as string
                    Value = random.Next(1, 101).ToString() // Random number between 1-100
                };
            }

            return seedData;
        }
    }
}
