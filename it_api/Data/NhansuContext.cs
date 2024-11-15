
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
        public DbSet<Info.Models.TinhModel> TinhModel { get; set; }
        public DbSet<Info.Models.PhongModel> DepartmentModel { get; set; }
        public DbSet<Info.Models.ChuyenmonModel> ChuyenmonModel { get; set; }
        public DbSet<Info.Models.KhoiModel> KhoiModel { get; set; }
        public DbSet<Info.Models.NganhangModel> NganhangModel { get; set; }
        public DbSet<Info.Models.TrinhdoModel> TrinhdoModel { get; set; }
        public DbSet<Info.Models.LoaiHDModel> LoaiHDModel { get; set; }

        public DbSet<Info.Models.ShiftModel> ShiftModel { get; set; }
        public DbSet<Info.Models.ShiftHolidayModel> ShiftHolidayModel { get; set; }
        public DbSet<Info.Models.ShiftUserModel> ShiftUserModel { get; set; }
        public DbSet<Info.Models.SalaryModel> SalaryModel { get; set; }
        public DbSet<Info.Models.SalaryUserModel> SalaryUserModel { get; set; }

        public DbSet<Info.Models.CalendarModel> CalendarModel { get; set; }
        public DbSet<Info.Models.CalendarHolidayModel> CalendarHolidayModel { get; set; }
        public DbSet<Info.Models.ChamanModel> ChamanModel { get; set; }
        public DbSet<Info.Models.ChamanKhachModel> ChamanKhachModel { get; set; }
        public DbSet<Info.Models.ChamcongModel> ChamcongModel { get; set; }
        public DbSet<Info.Models.HikModel> HikModel { get; set; }
        public DbSet<Info.Models.HolidayModel> HolidayModel { get; set; }
        public DbSet<Info.Models.OptionModel> OptionModel { get; set; }
        public DbSet<Info.Models.OrderletterModel> OrderletterModel { get; set; }
        public DbSet<Info.Models.OrderletterDetailsModel> OrderletterDetailsModel { get; set; }
        public DbSet<UserModel> UserModel { get; set; }
        public DbSet<EmailModel> EmailModel { get; set; }
        public DbSet<AuditTrailsModel> AuditTrailsModel { get; set; }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Info.Models.HikModel>().HasKey(table => new
            {
                table.id,
                table.device,
                table.datetime
            });

        }
        protected override void ConfigureConventions(ModelConfigurationBuilder builder)
        {
        }
        public override int SaveChanges()
        {
            OnBeforeSaveChanges();
            return base.SaveChanges();
        }
        private void OnBeforeSaveChanges()
        {
            ChangeTracker.DetectChanges();
            var auditEntries = new List<AuditEntry>();
            var user_http = actionAccessor.ActionContext.HttpContext.User;
            var user_id = UserManager.GetUserId(user_http);
            var changes = ChangeTracker.Entries();
            foreach (var entry in changes)
            {
                if (entry.Entity is AuditTrailsModel || entry.Entity is EmailModel || entry.State == EntityState.Detached || entry.State == EntityState.Unchanged)
                    continue;

                var auditEntry = new AuditEntry(entry);
                auditEntry.TableName = entry.Entity.GetType().Name;
                auditEntry.UserId = user_id;
                auditEntries.Add(auditEntry);
                foreach (var property in entry.Properties)
                {

                    string propertyName = property.Metadata.Name;
                    if (property.Metadata.IsPrimaryKey())
                    {
                        auditEntry.KeyValues[propertyName] = property.CurrentValue;
                        continue;
                    }
                    switch (entry.State)
                    {
                        case EntityState.Added:
                            auditEntry.AuditType = AuditType.Create;
                            auditEntry.NewValues[propertyName] = property.CurrentValue;
                            break;
                        case EntityState.Deleted:
                            auditEntry.AuditType = AuditType.Delete;
                            auditEntry.OldValues[propertyName] = property.OriginalValue;
                            break;
                        case EntityState.Modified:
                            if (property.IsModified)
                            {
                                var Original = entry.GetDatabaseValues().GetValue<object>(propertyName);
                                var Current = property.CurrentValue;
                                if (JsonConvert.SerializeObject(Original) == JsonConvert.SerializeObject(Current))
                                    continue;
                                auditEntry.ChangedColumns.Add(propertyName);
                                auditEntry.AuditType = AuditType.Update;
                                auditEntry.OldValues[propertyName] = Original;
                                auditEntry.NewValues[propertyName] = Current;

                            }
                            break;
                    }

                }
            }
            foreach (var auditEntry in auditEntries)
            {
                AuditTrailsModel.Add(auditEntry.ToAudit());
            }
        }

    }
}
