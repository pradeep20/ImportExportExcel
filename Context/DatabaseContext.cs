
using Microsoft.EntityFrameworkCore;
using SampleExcel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SampleExcel.Context
{
    public class DatabaseContext : DbContext
    {
        protected override void OnConfiguring(DbContextOptionsBuilder options)
        {
            var connectionString = string.Format(@"Data Source=DESKTOP-772IDC;Initial Catalog= UserDB; Integrated Security=true");
            options.UseSqlServer(connectionString);
        }

        public DbSet<User> Users { get; set; }
    }
}
