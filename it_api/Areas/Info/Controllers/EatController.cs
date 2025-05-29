


using Holdtime.Models;
using Info.Models;
using it_template.Areas.Trend.Controllers;
using iText.Commons.Actions.Contexts;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using PH.WorkingDaysAndTimeUtility.Configuration;
using PH.WorkingDaysAndTimeUtility;
using System.Collections;
using System.Data;
using System.Text.Json.Serialization;
using Vue.Data;
using Vue.Models;
using Vue.Services;
using System.Globalization;
using System.Net.WebSockets;
using System.Dynamic;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Spire.Xls;
using iText.StyledXmlParser.Node;

namespace it_template.Areas.Info.Controllers
{

    public class EatController : BaseController
    {
        private readonly IConfiguration _configuration;
        private UserManager<UserModel> UserManager;
        public EatController(NhansuContext context, AesOperation aes, IConfiguration configuration, UserManager<UserModel> UserMgr) : base(context, aes)
        {
            _configuration = configuration;
            UserManager = UserMgr;
        }

        public async Task<JsonResult> holidays(int year, string type)
        {
            var data = _context.CalendarHolidayModel.Where(d => d.calendar_id == type && d.date.Value.Year == year).ToList();
            return Json(data, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        [HttpPost]
        public async Task<JsonResult> Save(List<CalendarHolidayModel> list_add, List<string> list_remove)
        {
            if (list_remove != null && list_remove.Count() > 0)
            {
                var list = _context.CalendarHolidayModel.Where(d => list_remove.Contains(d.id)).ToList();
                _context.RemoveRange(list);
                _context.SaveChanges();
            }
            if (list_add != null && list_add.Count() > 0)
            {
                foreach (var item in list_add)
                {
                    item.id = Guid.NewGuid().ToString();

                }
                _context.AddRange(list_add);
                _context.SaveChanges();
            }
            return Json(new { success = true });
        }
        [HttpPost]
        public async Task<JsonResult> SaveChaman(List<ChamanModel> list_add, List<string> list_remove)
        {
            if (list_remove != null && list_remove.Count() > 0)
            {
                var list = _context.ChamanModel.Where(d => list_remove.Contains(d.id)).ToList();
                _context.RemoveRange(list);
                _context.SaveChanges();
            }
            if (list_add != null && list_add.Count() > 0)
            {
                foreach (var item in list_add)
                {
                    item.id = Guid.NewGuid().ToString();

                }
                _context.AddRange(list_add);
                _context.SaveChanges();
            }
            return Json(new { success = true });
        }
        [HttpPost]
        public async Task<JsonResult> SaveChamanKhach(List<ChamanKhachModel> list_add, List<ChamanKhachModel> list_update, List<string> list_remove)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user = await UserManager.GetUserAsync(currentUser);
            var user_id = user.Id;
            if (list_remove != null && list_remove.Count() > 0)
            {
                var list = _context.ChamanKhachModel.Where(d => list_remove.Contains(d.id)).ToList();
                _context.RemoveRange(list);
                _context.SaveChanges();
            }
            if (list_add != null && list_add.Count() > 0)
            {
                foreach (var item in list_add)
                {
                    item.id = Guid.NewGuid().ToString();
                    item.created_by = user_id;
                    item.created_at = DateTime.Now;
                }
                _context.AddRange(list_add);
                _context.SaveChanges();
            }
            if (list_update != null && list_update.Count() > 0)
            {
                _context.UpdateRange(list_update);
                _context.SaveChanges();
            }
            return Json(new { success = true });
        }

        [HttpPost]
        public async Task<JsonResult> Table()
        {
            var draw = Request.Form["draw"].FirstOrDefault();
            var start = Request.Form["start"].FirstOrDefault();
            var length = Request.Form["length"].FirstOrDefault();
            int pageSize = length != null ? Convert.ToInt32(length) : 0;
            var search = Request.Form["filters[search]"].FirstOrDefault();
            var id = Request.Form["filters[id]"].FirstOrDefault();
            var type = Request.Form["filters[type]"].FirstOrDefault();
            var department = Request.Form["filters[department]"].FirstOrDefault();
            var date_from_string = Request.Form["filters[date_from]"].FirstOrDefault();
            var date_to_string = Request.Form["filters[date_to]"].FirstOrDefault();
            DateTime date_from = date_from_string != null ? date_from = DateTime.ParseExact(date_from_string, "yyyy-MM-dd",
                                           System.Globalization.CultureInfo.InvariantCulture) : date_from = DateTime.Now;
            DateTime date_to = date_to_string != null ? date_to = DateTime.ParseExact(date_to_string, "yyyy-MM-dd",
                                           System.Globalization.CultureInfo.InvariantCulture) : date_to = DateTime.Now;
            var calendar = _context.CalendarModel.Where(d => d.id == type).FirstOrDefault();
            TimeSpan start_time = calendar.time_from.Value;
            TimeSpan end_time = calendar.time_to.Value;
            //var tenhh = Request.Form["filters[tenhh]"].FirstOrDefault();
            int skip = start != null ? Convert.ToInt32(start) : 0;
            var customerData = _context.PersonnelModel.Where(d => d.NGAYNGHIVIEC == null);


            /// CHECK PHAN QUYEN
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user = await UserManager.GetUserAsync(currentUser);
            var user_id = user.Id;
            var email = user.Email;
            var is_admin = await UserManager.IsInRoleAsync(user, "Administrator");
            var is_manager = await UserManager.IsInRoleAsync(user, "Manager HR");
            var is_hr = await UserManager.IsInRoleAsync(user, "HR");

            if (is_manager)
            {
                var person = _context.PersonnelModel.Where(d => d.EMAIL == email).FirstOrDefault();
                if (person != null)
                {
                    var maphong = person.MAPHONG;
                    customerData = customerData.Where(d => d.MAPHONG == maphong);
                }
                else
                {
                    customerData = customerData.Where(d => 0 == 1);
                }
            }
            else if (!is_admin && !is_hr)
            {
                customerData = customerData.Where(d => d.EMAIL == email);
            }

            ////



            int recordsTotal = customerData.Count();

            if (search != null && search != "")
            {
                customerData = customerData.Where(d => d.MANV.Contains(search) || d.HOVATEN.Contains(search));
            }
            if (department != null)
            {
                customerData = customerData.Where(d => d.MAPHONG == department);
            }
            if (id != null)
            {
                customerData = customerData.Where(d => d.id == id);
            }

            int recordsFiltered = customerData.Count();
            var datapost = customerData.OrderByDescending(d => d.NGAYNHANVIEC).ThenBy(d => d.MANV).Skip(skip).Take(pageSize).ToList();
            var data = new ArrayList();
            var utility = GetSchedule(type);
            var date_working = utility.GetWorkingDaysBetweenTwoWorkingDateTimes(date_from, date_to, false);
            if (utility.IsAWorkDay(date_from))
                date_working.Add(date_from);
            if (utility.IsAWorkDay(date_to))
                date_working.Add(date_to);
            var list_nv = datapost.Select(d => d.MANV).ToList();
            var list_chaman = _context.ChamanModel.Where(d => d.date.Value.Date >= date_from && d.date.Value.Date <= date_to && list_nv.Contains(d.MANV) && d.calendar_id == type).ToList();
            var list_hik = _context.HikModel.Where(d => d.date.Value.Date >= date_from && d.date.Value.Date <= date_to && d.time >= start_time && d.time <= end_time && d.device == "A.CT.1").OrderBy(d => d.date).ThenBy(d => d.time).ToList();

            foreach (var record in datapost)
            {
                var d = new ExpandoObject() as IDictionary<string, dynamic>;


                d.Add("MANV", record.MANV);
                d.Add("HOVATEN", record.HOVATEN);
                var tong = 0;
                foreach (var date in date_working)
                {
                    var chaman = list_chaman.Where(d => d.MANV == record.MANV && d.date.Value.Date == date.Date).FirstOrDefault();
                    if (chaman == null)
                    {
                        chaman = new ChamanModel()
                        {
                            date = date,
                            MANV = record.MANV,
                            NV_id = record.id,
                            calendar_id = type
                        };
                    }
                    else
                    {
                        tong++;
                    }


                    chaman.first_hik = list_hik.Where(d => d.id == record.MACC && d.date.Value.Date == date.Date && d.time.Value >= new TimeSpan(10, 30, 0) && d.time.Value <= new TimeSpan(14, 0, 0)).FirstOrDefault();

                    d.Add(date.ToString("yyyyMMdd"), chaman);
                }
                d.Add("tong", tong);
                data.Add(d);
            }
            var tong_date = new ExpandoObject() as IDictionary<string, dynamic>;
            foreach (var date in date_working)
            {
                var tong = list_chaman.Where(d => d.date.Value.Date == date.Date).Count();

                tong_date.Add(date.ToString("yyyyMMdd"), tong);
            }
            var now = DateTime.Now;
            var deadline = DateTime.Now.Date;
            if (now > DateTime.Now.Date.AddHours(17))
            {
                deadline = deadline.AddDays(1);
            }
            var person1 = _context.PersonnelModel.Where(d => d.EMAIL == email).FirstOrDefault();
            bool? autoeat = false;
            if (person1 != null)
            {
                autoeat = person1.autoeat;
            }
            var jsonData = new
            {
                draw = draw,
                recordsFiltered = recordsFiltered,
                recordsTotal = recordsTotal,
                data = data,
                tong_date,
                list_chaman,
                deadline,
                auto = autoeat
            };
            return Json(jsonData, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }

        [HttpPost]
        public async Task<JsonResult> TableKhach(string type, DateTime date)
        {

            var count = _context.ChamanModel.Where(d => d.date.Value.Date == date.Date && d.calendar_id == type).Count();
            var list_chaman_khach = _context.ChamanKhachModel.Where(d => d.date.Value.Date == date.Date && d.calendar_id == type).ToList();

            list_chaman_khach.Insert(0, new ChamanKhachModel()
            {
                title = "Suất ăn nhân viên",
                soluong = count,
                date = date,
                ignore = true
            });

            var jsonData = new { data = list_chaman_khach };
            return Json(jsonData, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        [HttpPost]
        public async Task<JsonResult> excel()
        {
            //try
            //{

            ////Lấy 
            var draw = Request.Form["draw"].FirstOrDefault();
            var start = Request.Form["start"].FirstOrDefault();
            var length = Request.Form["length"].FirstOrDefault();
            int pageSize = length != null ? Convert.ToInt32(length) : 0;
            var search = Request.Form["filters[search]"].FirstOrDefault();
            var id = Request.Form["filters[id]"].FirstOrDefault();
            var type = Request.Form["filters[type]"].FirstOrDefault();
            var department = Request.Form["filters[department]"].FirstOrDefault();
            var date_from_string = Request.Form["filters[date_from]"].FirstOrDefault();
            var date_to_string = Request.Form["filters[date_to]"].FirstOrDefault();
            DateTime date_from = date_from_string != null ? date_from = DateTime.ParseExact(date_from_string, "yyyy-MM-dd",
                                           System.Globalization.CultureInfo.InvariantCulture) : date_from = DateTime.Now;
            DateTime date_to = date_to_string != null ? date_to = DateTime.ParseExact(date_to_string, "yyyy-MM-dd",
                                           System.Globalization.CultureInfo.InvariantCulture) : date_to = DateTime.Now;
            TimeSpan start_time = new TimeSpan(10, 30, 0);
            TimeSpan end_time = new TimeSpan(14, 0, 0);
            var customerData = _context.PersonnelModel.Where(d => d.NGAYNGHIVIEC == null);
            /// CHECK PHAN QUYEN
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user = await UserManager.GetUserAsync(currentUser);
            var user_id = user.Id;
            var email = user.Email;
            var is_admin = await UserManager.IsInRoleAsync(user, "Administrator");
            var is_manager = await UserManager.IsInRoleAsync(user, "Manager HR");
            var is_hr = await UserManager.IsInRoleAsync(user, "HR");

            if (is_manager)
            {
                var person = _context.PersonnelModel.Where(d => d.EMAIL == email).FirstOrDefault();
                if (person != null)
                {
                    var maphong = person.MAPHONG;
                    customerData = customerData.Where(d => d.MAPHONG == maphong);
                }
                else
                {
                    customerData = customerData.Where(d => 0 == 1);
                }
            }
            else if (!is_admin && !is_hr)
            {
                customerData = customerData.Where(d => d.EMAIL == email);
            }




            int recordsFiltered = customerData.Count();
            var datapost = customerData.OrderBy(d => d.MANV).ToList();
            var data = new ArrayList();
            var utility = GetSchedule(type);
            var date_working = utility.GetWorkingDaysBetweenTwoWorkingDateTimes(date_from, date_to, false);
            //date_working = date_working.OrderBy(d => d).ToList();
            if (utility.IsAWorkDay(date_from))
                date_working.Add(date_from);
            if (utility.IsAWorkDay(date_to))
                date_working.Add(date_to);
            var list_nv = datapost.Select(d => d.MANV).ToList();
            var list_chaman = _context.ChamanModel.Where(d => d.date.Value.Date >= date_from && d.date.Value.Date <= date_to && list_nv.Contains(d.MANV)).ToList();
            var list_hik = _context.HikModel.Where(d => d.date.Value.Date >= date_from && d.date.Value.Date <= date_to && d.time >= start_time && d.time <= end_time && d.device == "A.CT.1").ToList();

            ///EXCEL
            /// 
            var viewPath = "wwwroot/report/excel/Danh sách đăng ký cơm.xlsx";
            var documentPath = "/tmp/" + DateTime.Now.ToFileTimeUtc() + ".xlsx";
            string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(viewPath);
            if (datapost.Count > 0)
            {
                Worksheet sheet = workbook.Worksheets[0];
                int stt = 0;
                var start_r = 4;
                DataTable dt = new DataTable();
                dt.Columns.Add("stt", typeof(int));
                dt.Columns.Add("MANV", typeof(string));
                dt.Columns.Add("HOVATEN", typeof(string));
                var start_c = 3;

                var date_check = date_from;
                while (date_check <= date_to)
                {
                    dt.Columns.Add(date_check.ToString("yyyyMMdd"), typeof(string));
                    var nRow = sheet.Rows[2];
                    nRow.Cells[start_c++].DateTimeValue = date_check;

                    date_check = date_check.AddDays(1);
                }

                dt.Columns.Add("tong", typeof(int));
                sheet.InsertRow(5, datapost.Count());

                foreach (var record in datapost)
                {
                    DataRow dr1 = dt.NewRow();
                    dr1["stt"] = (++stt);
                    dr1["MANV"] = record.MANV;
                    dr1["HOVATEN"] = record.HOVATEN;
                    var tong = 0;

                    date_check = date_from;
                    while (date_check <= date_to)
                    {
                        var chaman = list_chaman.Where(d => d.MANV == record.MANV && d.date.Value.Date == date_check.Date).FirstOrDefault();
                        if (chaman != null)
                        {
                            tong++;
                            dr1[date_check.ToString("yyyyMMdd")] = "x";
                        }
                        date_check = date_check.AddDays(1);
                    }

                    dr1["tong"] = tong;
                    dt.Rows.Add(dr1);
                    start_r++;
                    CellRange originDataRang = sheet.Range["A4:AZ4"];
                    CellRange targetDataRang = sheet.Range["A" + start_r + ":AZ" + start_r];
                    sheet.Copy(originDataRang, targetDataRang, true);
                }
                DataRow dr2 = dt.NewRow();
                dr2["stt"] = (++stt);
                dr2["MANV"] = "Tổng";
                date_check = date_from;
                while (date_check <= date_to)
                {
                    var tong = list_chaman.Where(d => d.date.Value.Date == date_check.Date).Count();

                    dr2[date_check.ToString("yyyyMMdd")] = tong;


                    date_check = date_check.AddDays(1);
                }

                dt.Rows.Add(dr2);
                sheet.InsertDataTable(dt, false, 4, 1);
            }

            workbook.SaveToFile("./wwwroot" + documentPath, ExcelVersion.Version2013);
            //var congthuc_ct = _QLSXcontext.Congthuc_CTModel.Where()
            var jsonData = new { success = true, link = documentPath };
            return Json(jsonData);

            //}
            //catch (Exception ex)
            //{
            //    return Json(new { success = false, message = ex.Message });
            //}
        }
        private WorkingDaysAndTimeUtility GetSchedule(string type, int? year = null)
        {
            var wts = new List<WorkTimeSpan>() { new WorkTimeSpan()
                { Start = new TimeSpan(10, 30, 0), End = new TimeSpan(14, 0, 0) } };

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
            var context = _context.CalendarHolidayModel.Where(d => d.calendar_id == type);
            if (year != null)
            {
                context = context.Where(d => d.date.Value.Year == year);
            }
            var holidays = context.ToList();
            //this is the configuration for holidays: 
            //in Italy we have this list of Holidays plus 1 day different on each province,
            //for mine is 1 Dec (see last element of the List<AHolyDay>).
            var italiansHoliDays = new List<AHolyDay>()
            {

            };
            italiansHoliDays.Add(new Holidays(holidays));
            //instantiate with configuration
            var utility = new WorkingDaysAndTimeUtility(week, italiansHoliDays);
            return utility;
        }
    }
    public class Holidays : MultiCalculatedHoliDay
    {
        private List<CalendarHolidayModel> _holidays;
        internal Holidays(List<CalendarHolidayModel> holidays) : base(0, 0)
        {
            _holidays = holidays;
        }


        public override Type GetHolyDayType() => typeof(Holidays);

        /// <summary>Calculates the list of MultiCalculatedHoliDays for the given year.</summary>
        /// <param name="year">The year.</param>
        /// <returns></returns>
        public override List<DateTime> CalculateList(int year)
        {
            var baseHolidays = new List<DateTime>();
            baseHolidays = _holidays.Select(d => d.date.Value).ToList();
            return baseHolidays;
        }
    }
}
