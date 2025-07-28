
using Holdtime.Models;
using Info.Models;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc.Infrastructure;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.ChangeTracking;
using Newtonsoft.Json;
using Vue.Areas.V1.Models;
using Vue.Models;

namespace Vue.Data
{
    public class QLSXContext : DbContext
    {
        private IActionContextAccessor actionAccessor;
        private UserManager<UserModel> UserManager;
        public QLSXContext(DbContextOptions<QLSXContext> options, UserManager<UserModel> UserMgr, IActionContextAccessor ActionAccessor) : base(options)
        {
            actionAccessor = ActionAccessor;
            UserManager = UserMgr;

        }

        public DbSet<DTA_CHANGECONTROL> DTA_CHANGECONTROL { get; set; }
        public DbSet<DTA_CHANGECONTROL_A> DTA_CHANGECONTROL_A { get; set; }
        public DbSet<DTA_CHANGECONTROL_B> DTA_CHANGECONTROL_B { get; set; }
        public DbSet<DTA_CHANGECONTROL_C> DTA_CHANGECONTROL_C { get; set; }
        public DbSet<DTA_CHANGECONTROL_D> DTA_CHANGECONTROL_D { get; set; }
        public DbSet<DTA_CHANGECONTROL_E> DTA_CHANGECONTROL_E { get; set; }
        public DbSet<DTA_CHANGECONTROL_F> DTA_CHANGECONTROL_F { get; set; }
        public DbSet<DTA_CHANGECONTROL_G> DTA_CHANGECONTROL_G { get; set; }
        public DbSet<DTA_CHANGECONTROL_H> DTA_CHANGECONTROL_H { get; set; }
        public DbSet<DTA_CHANGECONTROL_I> DTA_CHANGECONTROL_I { get; set; }
        public DbSet<DTA_CHANGECONTROL_J> DTA_CHANGECONTROL_J { get; set; }
        public DbSet<DTA_CHANGECONTROL_K> DTA_CHANGECONTROL_K { get; set; }




        public DbSet<CHANGECONTROL_ANBAN_IN> CHANGECONTROL_ANBAN_IN { get; set; }
        public DbSet<CHANGECONTROL_CAPA> CHANGECONTROL_CAPA { get; set; }
        public DbSet<CHANGECONTROL_PHANLOAI> CHANGECONTROL_PHANLOAI { get; set; }
        public DbSet<CHANGECONTROL_THANHVIEN> CHANGECONTROL_THANHVIEN { get; set; }
        public DbSet<DIEUTRASUCO> DIEUTRASUCO { get; set; }
        public DbSet<DIEUTRASUCO_PARTA> DIEUTRASUCO_PARTA { get; set; }
        public DbSet<DIEUTRASUCO_PARTB> DIEUTRASUCO_PARTB { get; set; }
        public DbSet<DIEUTRASUCO_CAPA> DIEUTRASUCO_CAPA { get; set; }
        public DbSet<SUCO> SUCO { get; set; }
        public DbSet<SUCO_DANHGIA> SUCO_DANHGIA { get; set; }
        public DbSet<SUCO_HANHDONG> SUCO_HANHDONG { get; set; }
        public DbSet<COAModel> COAModel { get; set; }
        public DbSet<LenhXuatXuongModel> LenhXuatXuongModel { get; set; }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {

            modelBuilder.Entity<CHANGECONTROL_ANBAN_IN>().HasKey(table => new
            {
                table.makhuvuc,
                table.anban,
                table.ngayhieuluc
            });


            modelBuilder.Entity<DTA_CHANGECONTROL_A>()
                .HasKey(c => new { c.sochange, c.ngaydenghi });

            modelBuilder.Entity<DTA_CHANGECONTROL_B>()
                .HasKey(c => new { c.sochange, c.ngaydenghi });

            modelBuilder.Entity<DTA_CHANGECONTROL_C>()
                .HasKey(c => new { c.sochange, c.ngaydenghi });

            modelBuilder.Entity<DTA_CHANGECONTROL_D>()
                .HasKey(c => new { c.sochange, c.ngaydenghi });

            modelBuilder.Entity<DTA_CHANGECONTROL_E>()
                .HasKey(c => new { c.sochange, c.ngaydenghi });

            modelBuilder.Entity<DTA_CHANGECONTROL_F>()
                .HasKey(c => new { c.sochange, c.ngaydenghi });

            modelBuilder.Entity<DTA_CHANGECONTROL_G>()
                .HasKey(c => new { c.sochange, c.ngaydenghi });
            modelBuilder.Entity<DTA_CHANGECONTROL_H>()
                .HasKey(c => new { c.sochange, c.ngaydenghi });
            modelBuilder.Entity<DTA_CHANGECONTROL_I>()
                .HasKey(c => new { c.sochange, c.ngaydenghi });
            modelBuilder.Entity<DTA_CHANGECONTROL_J>()
                .HasKey(c => new { c.sochange, c.ngaydenghi });

            modelBuilder.Entity<DTA_CHANGECONTROL_K>()
                .HasKey(c => new { c.sochange, c.ngaydenghi });
        }
        protected override void ConfigureConventions(ModelConfigurationBuilder builder)
        {
        }


    }
}
