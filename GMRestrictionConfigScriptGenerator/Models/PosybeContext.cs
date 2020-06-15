using System;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;

namespace GMRestrictionConfigScriptGenerator.Models
{
    public partial class PosybeContext : DbContext
    {
        public PosybeContext()
        {
        }

        public PosybeContext(DbContextOptions<PosybeContext> options)
            : base(options)
        {
        }

        public virtual DbSet<FSkladPohybySposoby> FSkladPohybySposoby { get; set; }
        public virtual DbSet<UplSubcategories> UplSubcategories { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
#warning To protect potentially sensitive information in your connection string, you should move it out of source code. See http://go.microsoft.com/fwlink/?LinkId=723263 for guidance on storing connection strings.
                optionsBuilder.UseSqlServer("Data Source=10.10.3.114;Initial Catalog=Posybe;Persist Security Info=True;User ID=sa;Password=StarForce4");
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<FSkladPohybySposoby>(entity =>
            {
                entity.HasKey(e => new { e.Operacia, e.Typ, e.Sposob });

                entity.ToTable("F_SKLAD_POHYBY_SPOSOBY");

                entity.Property(e => e.Operacia).HasColumnName("OPERACIA");

                entity.Property(e => e.Typ).HasColumnName("TYP");

                entity.Property(e => e.Sposob).HasColumnName("SPOSOB");

                entity.Property(e => e.MapovaniePohyb).HasColumnName("MAPOVANIE_POHYB");

                entity.Property(e => e.Nazov)
                    .IsRequired()
                    .HasColumnName("NAZOV")
                    .HasMaxLength(64)
                    .IsUnicode(false);

                entity.Property(e => e.Povoleny).HasColumnName("POVOLENY");

                entity.Property(e => e.Skratka)
                    .IsRequired()
                    .HasColumnName("SKRATKA")
                    .HasMaxLength(6)
                    .IsUnicode(false);

                entity.Property(e => e.SqlFilter)
                    .HasColumnName("SQL_FILTER")
                    .HasMaxLength(2048)
                    .IsUnicode(false);

                entity.Property(e => e.SqlFilterPartner)
                    .HasColumnName("SQL_FILTER_PARTNER")
                    .HasMaxLength(1024)
                    .IsUnicode(false);
            });

            modelBuilder.Entity<UplSubcategories>(entity =>
            {
                entity.ToTable("UPL_SUBCATEGORIES");

                entity.Property(e => e.Id)
                    .HasColumnName("ID")
                    .ValueGeneratedNever();

                entity.Property(e => e.IdCtCategories).HasColumnName("ID_CT_CATEGORIES");

                entity.Property(e => e.Notes)
                    .HasColumnName("NOTES")
                    .HasMaxLength(256)
                    .IsUnicode(false);

                entity.Property(e => e.Number).HasColumnName("NUMBER");

                entity.Property(e => e.Title)
                    .IsRequired()
                    .HasColumnName("TITLE")
                    .HasMaxLength(128)
                    .IsUnicode(false)
                    .IsFixedLength();
            });

            OnModelCreatingPartial(modelBuilder);
        }

        partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
    }
}
