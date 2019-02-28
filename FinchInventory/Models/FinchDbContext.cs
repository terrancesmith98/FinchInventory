namespace FinchInventory.Models
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class FinchDbContext : DbContext
    {
        public FinchDbContext()
            : base("name=FinchDbContext")
        {
        }

        public virtual DbSet<Clothing> Clothings { get; set; }
        public virtual DbSet<Goal> Goals { get; set; }
        public virtual DbSet<Location> Locations { get; set; }
        public virtual DbSet<Machine> Machines { get; set; }
        public virtual DbSet<Manufacturer> Manufacturers { get; set; }
        public virtual DbSet<Position> Positions { get; set; }
        public virtual DbSet<Role> Roles { get; set; }
        public virtual DbSet<RollType> RollTypes { get; set; }
        public virtual DbSet<Status> Status { get; set; }
        public virtual DbSet<Type> Types { get; set; }
        public virtual DbSet<UserRole> UserRoles { get; set; }
        public virtual DbSet<User> Users { get; set; }
        public virtual DbSet<Admin> Admins { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Clothing>()
                .Property(e => e.Dimensions)
                .IsUnicode(false);

            modelBuilder.Entity<Clothing>()
                .Property(e => e.CurrentDia)
                .HasPrecision(5, 3);

            modelBuilder.Entity<Clothing>()
                .Property(e => e.MinDia)
                .HasPrecision(5, 3);

            modelBuilder.Entity<Clothing>()
                .Property(e => e.Crown)
                .HasPrecision(5, 3);

            modelBuilder.Entity<Machine>()
                .HasMany(e => e.Clothings)
                .WithRequired(e => e.Machine)
                .HasForeignKey(e => e.PM_Number)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Machine>()
                .HasMany(e => e.Goals)
                .WithRequired(e => e.Machine)
                .HasForeignKey(e => e.PM_ID)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Position>()
                .HasMany(e => e.Clothings)
                .WithRequired(e => e.Position)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Position>()
                .HasMany(e => e.Goals)
                .WithRequired(e => e.Position)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Role>()
                .HasMany(e => e.UserRoles)
                .WithRequired(e => e.Role)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<RollType>()
                .HasMany(e => e.Clothings)
                .WithRequired(e => e.RollType)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Status>()
                .HasMany(e => e.Clothings)
                .WithRequired(e => e.Status)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Type>()
                .HasMany(e => e.Clothings)
                .WithRequired(e => e.Type)
                .WillCascadeOnDelete(false);
        }
    }
}
