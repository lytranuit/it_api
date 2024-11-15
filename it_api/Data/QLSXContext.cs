
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


        public DbSet<COAModel> COAModel { get; set; }
        public DbSet<LenhXuatXuongModel> LenhXuatXuongModel { get; set; }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {


        }
        protected override void ConfigureConventions(ModelConfigurationBuilder builder)
        {
        }


    }
}
