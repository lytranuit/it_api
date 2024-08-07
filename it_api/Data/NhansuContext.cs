﻿
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore.ChangeTracking;
using Microsoft.AspNetCore.Mvc.Infrastructure;
using Newtonsoft.Json;
using Holdtime.Models;
using Vue.Models;

namespace Vue.Data
{
    public class NhansuContext : DbContext
    {
        private IActionContextAccessor actionAccessor;
        private UserManager<UserModel> UserManager;
        public NhansuContext(DbContextOptions<NhansuContext> options, UserManager<UserModel> UserMgr, IActionContextAccessor ActionAccessor) : base(options)
        {
            actionAccessor = ActionAccessor;
            UserManager = UserMgr;

        }


        public DbSet<Info.Models.NewsModel> NewsModel { get; set; }
        public DbSet<Info.Models.HotNewsModel> HotNewsModel { get; set; }
        public DbSet<Info.Models.CategoryModel> CategoryModel { get; set; }
        public DbSet<Info.Models.NewsCategoryModel> NewsCategoryModel { get; set; }
        public DbSet<Info.Models.PersonnelModel> PersonnelModel { get; set; }
        public DbSet<Info.Models.PositionModel> PositionModel { get; set; }
        public DbSet<Info.Models.PhongModel> DepartmentModel { get; set; }
        public DbSet<Info.Models.ChuyenmonModel> ChuyenmonModel { get; set; }
        public DbSet<Info.Models.KhoiModel> KhoiModel { get; set; }
        public DbSet<Info.Models.NganhangModel> NganhangModel { get; set; }
        public DbSet<Info.Models.TrinhdoModel> TrinhdoModel { get; set; }
        public DbSet<Info.Models.LoaiHDModel> LoaiHDModel { get; set; }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {

        }
        protected override void ConfigureConventions(ModelConfigurationBuilder builder)
        {
        }
    }
}
