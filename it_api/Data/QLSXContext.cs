
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore.ChangeTracking;
using Microsoft.AspNetCore.Mvc.Infrastructure;
using Newtonsoft.Json;
using Holdtime.Models;
using Vue.Models;
using Info.Models;

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
        }
        protected override void ConfigureConventions(ModelConfigurationBuilder builder)
        {
        }


    }
}
