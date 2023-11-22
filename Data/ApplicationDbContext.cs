using ExcelUpload.Models;
using Microsoft.EntityFrameworkCore;

namespace ExcelUpload.Data
{
    public class ApplicationDbContext : DbContext
    {
        public DbSet<ExcelData> ExcelData { get; set; }

        public ApplicationDbContext(DbContextOptions options) : base(options)
        {

        }
    }
}
