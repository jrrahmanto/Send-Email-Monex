using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using static WorkerEmail.Model;

namespace WorkerEmail
{
    public class Data : DbContext
    {
        private readonly IConfiguration _iConfiguration;
        private readonly string _connectionString;
        public Data()
        {
            IConfigurationBuilder builder = new ConfigurationBuilder()
                        .SetBasePath(Directory.GetCurrentDirectory())
                        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
            _iConfiguration = builder.Build();
            _connectionString = _iConfiguration.GetConnectionString("TINConnection");
        }
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            _ = optionsBuilder.UseSqlServer(_connectionString, providerOptions => providerOptions.CommandTimeout(1800))
                          .UseQueryTrackingBehavior(QueryTrackingBehavior.NoTracking);
        }

        public DbSet<TradeProgress> TradeProgresses { get; set; }
        public DbSet<ClearingMember> clearingMembers { get; set; }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<TradeProgress>().ToTable("EODTradeProgress", "SKD");
            modelBuilder.Entity<ClearingMember>().ToTable("CMProfile", "SKD");
        }
    }
}
