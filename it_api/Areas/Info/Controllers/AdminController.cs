using it_template.Areas.Trend.Controllers;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PH.WorkingDaysAndTimeUtility.Configuration;
using PH.WorkingDaysAndTimeUtility;
using System.Collections;
using System.Text.Json.Serialization;
using Vue.Data;
using Vue.Models;
using Vue.Services;
using Point85.ShiftSharp.Schedule;
using System.Dynamic;
using Info.Models;
using static iText.StyledXmlParser.Jsoup.Select.Evaluator;

namespace it_template.Areas.Info.Controllers
{
    public class AdminController : BaseController
    {
        private UserManager<UserModel> UserManager;
        private readonly TinhCong _tinhcong;
        public AdminController(NhansuContext context, AesOperation aes, TinhCong tinhcong, UserManager<UserModel> UserMgr) : base(context, aes)
        {
            UserManager = UserMgr;
            _tinhcong = tinhcong;
        }
        public async Task<JsonResult> HomeBadge()
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user = await UserManager.GetUserAsync(currentUser);
            var user_id = user.Id;
            var email = user.Email;


            var data = _context.HotNewsModel.Where(d => d.id == 1).FirstOrDefault();
            var tin_moi = _context.NewsModel.Where(d => d.deleted_at == null && d.is_publish == true).OrderByDescending(d => d.created_at).Take(7).ToList();
            var highlight = _context.NewsModel.Where(d => d.deleted_at == null && d.is_publish == true && d.is_highlight == true).OrderByDescending(d => d.created_at).Take(4).ToList();
            var cate = _context.CategoryModel.Where(d => d.deleted_at == null).ToList();
            foreach (var item in cate)
            {
                item.list_news = _context.NewsCategoryModel.Where(d => d.category_id == item.id).Include(d => d.news).Where(d => d.news.deleted_at == null).Select(d => d.news).Take(10).ToList();
            }

            /////Công
            ///
            decimal tong = 0;
            decimal phep = 0;
            decimal conglam = 0;
            decimal phepconlai = 0;
            var record = _context.PersonnelModel.Where(d => d.EMAIL == email).ToList();
            if (record.Count() > 0)
            {

                var now = DateTime.Now;
                if (now.Day >= 26)
                {
                    now = now.AddMonths(1);
                }
                var date_from = new DateTime(now.Year, now.Month - 1, 26);
                var date_to = new DateTime(now.Year, now.Month, 25);

                var list_data_cong = _tinhcong.cong(record, date_from, date_to);
                IDictionary<string, dynamic> data_cong = list_data_cong.Find(x => x["EMAIL"] == email);
                conglam = (decimal)data_cong["tong"];
                phep = (decimal)data_cong["tongphep"];
                tong = (decimal)data_cong["tongcong"];
                phepconlai = (decimal)data_cong["phepnamconlai"];
            }


            var tongthunhap = record.FirstOrDefault().tong_thunhap;
            //var noibat = _context.
            return Json(new { message = data.message, tin_moi, highlight, cate, phep, tong, conglam, tongthunhap, phepconlai }, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        private WorkingDaysAndTimeUtility GetSchedule(string id, TimeSpan start, TimeSpan end)
        {
            var wts = new List<WorkTimeSpan>() { new WorkTimeSpan()
                { Start = start, End = end } };

            var week = new WeekDaySpan()
            {
                WorkDays = new Dictionary<DayOfWeek, WorkDaySpan>()
                {
                    {DayOfWeek.Monday, new WorkDaySpan() {TimeSpans = wts}}
                    ,
                    {DayOfWeek.Tuesday, new WorkDaySpan() {TimeSpans = wts}}
                    ,
                    {DayOfWeek.Wednesday, new WorkDaySpan() {TimeSpans = wts}}
                    ,
                    {DayOfWeek.Thursday, new WorkDaySpan() {TimeSpans = wts}}
                    ,
                    {DayOfWeek.Friday, new WorkDaySpan() {TimeSpans = wts}}
                    ,
                    {DayOfWeek.Saturday, new WorkDaySpan() {TimeSpans = wts}}
                    ,
                    {DayOfWeek.Sunday, new WorkDaySpan() {TimeSpans = wts}}
                }
            };
            var context = _context.ShiftHolidayModel.Where(d => d.shift_id == id);

            var holidays = context.ToList();
            //this is the configuration for holidays: 
            //in Italy we have this list of Holidays plus 1 day different on each province,
            //for mine is 1 Dec (see last element of the List<AHolyDay>).
            var italiansHoliDays = new List<AHolyDay>()
            {

            };
            italiansHoliDays.Add(new WHolidays(holidays));
            //instantiate with configuration
            var utility = new WorkingDaysAndTimeUtility(week, italiansHoliDays);
            return utility;
        }
    }

}
