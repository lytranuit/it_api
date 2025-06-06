
using Vue.Models;
using Microsoft.AspNetCore.Mvc;
using Vue.Data;
using System.Net.Mail;
using Vue.Services;
using System.Net.Mime;
using Microsoft.EntityFrameworkCore.Infrastructure;
using System.Diagnostics;
using Microsoft.EntityFrameworkCore;
using PH.WorkingDaysAndTimeUtility.Configuration;
using PH.WorkingDaysAndTimeUtility;
using Info.Models;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Spire.Xls;
using Hik.Api;
using System.Net;
using System.Security.Policy;
using Microsoft.VisualStudio.Web.CodeGenerators.Mvc.Templates.BlazorIdentity.Pages.Manage;
using Azure.Core;
using Humanizer;
using System.Collections.Generic;
using System.Text;
using System;
using System.Text.Json;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Net.Http;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Net.Http.Headers;
using System.ComponentModel.DataAnnotations;
using CertificateManager.Models;
using CertificateManager;
using System.Security.Cryptography.X509Certificates;
using NodaTime;

namespace Vue.Controllers
{

    public class HomeController : Controller
    {
        protected readonly ItContext _context;
        private readonly ViewRender _view;
        protected readonly HoldTimeContext _holdtimecontext;
        protected readonly NhansuContext _nhansuContext;

        private readonly IConfiguration _configuration;

        public HomeController(ItContext context, HoldTimeContext holdtimecontext, NhansuContext nhansucontext, IConfiguration configuration, ViewRender view)
        {
            _configuration = configuration;
            _context = context;
            _holdtimecontext = holdtimecontext;
            _nhansuContext = nhansucontext;
            _view = view;
            var listener = _context.GetService<DiagnosticSource>();
            (listener as DiagnosticListener).SubscribeWithAdapter(new CommandInterceptor());
        }

        public JsonResult Index()
        {
            var wts = new List<WorkTimeSpan>() { new WorkTimeSpan()
                { Start = new TimeSpan(7, 30, 0), End = new TimeSpan(17, 0, 0) } };

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
                }
            };

            //this is the configuration for holidays: 
            //in Italy we have this list of Holidays plus 1 day different on each province,
            //for mine is 1 Dec (see last element of the List<AHolyDay>).
            var italiansHoliDays = new List<AHolyDay>()
            {
                new HoliDay(1, 1),new HoliDay(6, 1),
                new HoliDay(25, 4),new HoliDay(1, 5),new HoliDay(2, 6),
                new HoliDay(15, 8),new HoliDay(1, 11),new HoliDay(8, 12),
                new HoliDay(25, 12),new HoliDay(26, 12)
                , new HoliDay(1, 12)
            };

            //instantiate with configuration
            var utility = new WorkingDaysAndTimeUtility(week, italiansHoliDays);

            //lets-go: add 3 working-days to Jun 1, 2015
            var result = utility.AddWorkingDays(new DateTime(2015, 6, 1), 100);
            var start = new DateTime(2015, 12, 31, 9, 0, 0);
            var end = new DateTime(2016, 1, 7, 9, 0, 0);
            var r = utility.GetWorkingDaysBetweenTwoWorkingDateTimes(start, end);
            TimeSpan workedTime = TimeSpan.FromSeconds(0);

            while (start < end)
            {
                var day = new DateTime(2021, 1, 1);
                var check0 = utility.IsAWorkDay(day);
                workedTime += TimeSpan.FromMinutes(1);
            }

            //result is Jun 5, 2015 (see holidays list) 
            //utility.c
            //var u = new PH.WorkingDaysAndTimeUtility.WorkingDateTimeExtension();
            return Json(new { test = 1, message = DateTime.Now, result, r, workedTime, workedTime.TotalDays });

        }


        public async Task<JsonResult> cronjobHoldtimeDaily()
        {
            var timecheck = DateTime.Now.Date.Date;
            var timecheck1 = DateTime.Now.AddDays(1).Date;
            var timecheck5 = DateTime.Now.AddDays(5).Date;
            var timecheck2 = DateTime.Now.Date;
            //var tasks = _context.TaskModel.Where(d => d.deleted_at == null && (d.finished_at == null || d.progress != 100) && d.endDate != null && d.endDate < DateTime.Now.AddDays(1)).Include(d => d.assignees).ToList();
            var query1 = _holdtimecontext.HoldTimeModel
                .Where(d => d.deleted_at == null && d.date_reality == null && d.is_remind == true && ((d.date_theory >= timecheck && d.date_theory <= timecheck1) || d.date_theory == timecheck5));

            var sql = query1.ToQueryString();
            var holdtime = query1.Include(d => d.Hold).ThenInclude(d => d.stage)
                .ToList();
            foreach (var item in holdtime)
            {
                var user_emails = _holdtimecontext.HoldAlertModel.Where(d => d.hold_id == item.hold_id).Include(d => d.user).Select(d => d.user.Email).ToList();
                var mail_string = string.Join(",", user_emails);
                string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;
                var body = _view.Render("Emails/DueHoldtime", new
                {
                    link_logo = Domain + "/images/clientlogo_astahealthcare.com_f1800.png",
                    link = _configuration["Application:Holdtime:link"] + "hold/edit/" + item.hold_id,
                    data = item
                });
                var email = new EmailModel
                {
                    email_to = mail_string,
                    subject = "[Nhắc nhở] Lấy mẫu thời gian chờ.",
                    body = body,
                    email_type = "DueHoldtime",
                    status = 1
                };
                _context.Add(email);

            }

            ////Quá hạn
            var query = _holdtimecontext.HoldTimeModel.Where(d => d.deleted_at == null && d.date_reality == null && d.is_remind == true && d.date_theory < timecheck2).ToList();

            var hold_overdue = query.Select(d => d.hold_id).ToList();
            var user_overdue = _holdtimecontext.HoldAlertModel.Where(d => hold_overdue.Contains(d.hold_id))
                .Include(d => d.Hold)
                .ThenInclude(d => d.stage)
                .Include(d => d.Hold)
                .ThenInclude(d => d.times.Where(e => e.date_reality == null && e.date_theory < timecheck2))
                .ToList();
            var all_overdue = user_overdue.GroupBy(d => d.user_id, (x, y) => new
            {
                num_sign = y.Count(),
                data = y.Select(d => d.Hold).ToList(),
                userId = x
            }).ToList();
            foreach (var item in all_overdue)
            {

                var user = _context.UserModel.Where(d => d.Id == item.userId).FirstOrDefault();
                if (user == null)
                    continue;

                if (user.deleted_at != null || (user.LockoutEnd != null && user.LockoutEnd >= DateTime.Now))
                    continue;
                ///Xóa user nếu user 1 tháng chưa đăng nhập
                //var last_login = user.last_login != null ? user.last_login : user.created_at;
                //if (last_login < DateTime.Now.AddMonths(-1))
                //{
                //    user.LockoutEnd = DateTime.Now.AddDays(360);
                //    _context.Update(user);
                //    _context.SaveChanges();
                //    continue;
                //}
                //foreach (var da in item.data)
                //{
                //    var type = _context.DocumentTypeModel.Where(d => d.id == da.type_id).FirstOrDefault();
                //    da.type = type;
                //}
                var mail_string = user.Email;
                string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;
                var body = _view.Render("Emails/OverDueHoldtime", new
                {
                    link_logo = Domain + "/images/clientlogo_astahealthcare.com_f1800.png",
                    link = _configuration["Application:Holdtime:link"],
                    data = item.data
                });
                var email = new EmailModel
                {
                    email_to = mail_string,
                    subject = "[Quá hạn] Mẫu thời gian chờ đã quá hạn lấy mẫu.",
                    body = body,
                    email_type = "OverDueHoldtime",
                    status = 1
                };
                _context.Add(email);

            }
            _context.SaveChanges();
            return Json(new { success = true });
        }
        public async Task<JsonResult> cronjobDangkyantrua()
        {
            var date = DateTime.Now.AddDays(1);
            var is_holiday = _nhansuContext.CalendarHolidayModel.Where(d => d.calendar_id == "Buổi trưa" && d.date.Value.Date == date.Date).Count();
            if (is_holiday > 0)
            {

                return Json(new { success = true });
            }
            var person = _nhansuContext.PersonnelModel.Where(d => d.NGAYNGHIVIEC == null && d.autoeat == true).ToList();
            foreach (var record in person)
            {
                var find = _nhansuContext.ChamanModel.Where(d => d.MANV == record.MANV && d.date.Value.Date == date.Date && d.calendar_id == "Buổi trưa").FirstOrDefault();
                if (find == null)
                {
                    var chaman = new ChamanModel()
                    {
                        id = Guid.NewGuid().ToString(),
                        date = date,
                        MANV = record.MANV,
                        NV_id = record.id,
                        calendar_id = "Buổi trưa",
                        value = true
                    };
                    _nhansuContext.Add(chaman);
                }

            }
            _nhansuContext.SaveChanges();
            return Json(new { success = true });
        }

        public async Task<JsonResult> syncTknganhang()
        {
            return Json(new { });
            // Khởi tạo workbook để đọc
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            book.LoadFromFile("./wwwroot/data/info/BANG LUONG T09.2024 (Finish).xlsx", ExcelVersion.Version2013);

            Spire.Xls.Worksheet sheet = book.Worksheets[0];
            var lastrow = sheet.LastDataRow;
            // nếu vẫn chưa gặp end thì vẫn lấy data
            Console.WriteLine(lastrow);
            var list_Result = new List<ResultModel>();
            for (int rowIndex = 12; rowIndex < lastrow; rowIndex++)
            {
                // lấy row hiện tại
                var nowRow = sheet.Rows[rowIndex];
                if (nowRow == null)
                    continue;
                // vì ta dùng 3 cột A, B, C => data của ta sẽ như sau
                //int numcount = nowRow.Cells.Count;
                //for(int y = 0;y<numcount - 1 ;y++)
                var code = nowRow.Cells[2] != null ? nowRow.Cells[2].Value.Trim() : null;
                // Xuất ra thông tin lên màn hình
                Console.WriteLine("MS: {0} ", code);
                Console.WriteLine("rowIndex: {0} ", rowIndex);

                if (code == null || code == "")
                    continue;
                string? stk = nowRow.Cells[37] != null && nowRow.Cells[37].Value != "NA" && nowRow.Cells[37].Value != "" ? nowRow.Cells[37].Value.Trim() : null;
                if (stk == null || stk == "")
                    continue;
                var split = stk.Split(" - ").ToList();
                if (split.Count() < 2)
                {
                    continue;
                }
                var taikhoan = split[0].Trim();
                var nganhang = split[1].Trim();

                //double? luongcb = nowRow.Cells[6] != null && nowRow.Cells[6].Value != "NA" && nowRow.Cells[6].Value != "" ? nowRow.Cells[6].NumberValue : null;
                //double? luongkpi = nowRow.Cells[16] != null && nowRow.Cells[16].Value != "NA" && nowRow.Cells[16].Value != "" ? (double)nowRow.Cells[16].FormulaValue : 0;
                //double? tongtrocap = nowRow.Cells[15] != null && nowRow.Cells[15].Value != "NA" && nowRow.Cells[15].Value != "" ? (double)nowRow.Cells[15].FormulaValue : 0;
                //double? khoancong = nowRow.Cells[18] != null && nowRow.Cells[18].Value != "NA" && nowRow.Cells[18].Value != "" ? (double)nowRow.Cells[18].NumberValue : 0;
                //double? tongthunhap = nowRow.Cells[19] != null && nowRow.Cells[19].Value != "NA" && nowRow.Cells[19].Value != "" ? (double)nowRow.Cells[19].FormulaValue : null;
                //int? nguoiphuthuoc = nowRow.Cells[24] != null && nowRow.Cells[24].Value != "NA" && nowRow.Cells[24].Value != "" ? (int)nowRow.Cells[24].NumberValue : null;
                //double? pc_trachnhiem = nowRow.Cells[14] != null && nowRow.Cells[14].Value != "NA" && nowRow.Cells[14].Value != "" ? nowRow.Cells[14].NumberValue : 0;
                //double? pc_xangxe = nowRow.Cells[10] != null && nowRow.Cells[10].Value != "NA" && nowRow.Cells[10].Value != "" ? nowRow.Cells[10].NumberValue : 0;
                //double? tamungdot1 = 0;
                //if (nowRow.Cells[34] != null && nowRow.Cells[34].Value != "NA" && nowRow.Cells[34].Value != "" && (double)nowRow.Cells[34].FormulaValue > 0)
                //{
                //    tamungdot1 = nowRow.Cells[33] != null && nowRow.Cells[33].Value != "NA" && nowRow.Cells[33].Value != "" ? nowRow.Cells[33].NumberValue : null;
                //}
                //string? stk = nowRow.Cells[35] != null && nowRow.Cells[35].Value != "NA" && nowRow.Cells[35].Value != "" ? nowRow.Cells[35].Value.Trim() : 0;


                var findP = _nhansuContext.PersonnelModel.Where(d => d.MANV == code).FirstOrDefault();
                if (findP == null)
                {
                    continue;
                }
                findP.sotk_icb = taikhoan;
                findP.sotk_vba = nganhang;
                //if (code == "NMK170962")
                //{
                //    Console.WriteLine("MS: {0} ", code);
                //}
                _nhansuContext.Update(findP);
                _nhansuContext.SaveChanges();
            }
            //_context.AddRange(list_Result);
            //_context.SaveChanges();

            return Json(new { success = true });
        }
        public async Task<JsonResult> syncLuong()
        {
            return Json(new { });
            // Khởi tạo workbook để đọc
            Spire.Xls.Workbook book = new Spire.Xls.Workbook();
            book.LoadFromFile("./wwwroot/data/info/BANG LUONG T09.2024 (Finish).xlsx", ExcelVersion.Version2013);

            Spire.Xls.Worksheet sheet = book.Worksheets[0];
            var lastrow = sheet.LastDataRow;
            // nếu vẫn chưa gặp end thì vẫn lấy data
            Console.WriteLine(lastrow);
            var list_Result = new List<ResultModel>();
            for (int rowIndex = 12; rowIndex < lastrow; rowIndex++)
            {
                // lấy row hiện tại
                var nowRow = sheet.Rows[rowIndex];
                if (nowRow == null)
                    continue;
                // vì ta dùng 3 cột A, B, C => data của ta sẽ như sau
                //int numcount = nowRow.Cells.Count;
                //for(int y = 0;y<numcount - 1 ;y++)
                var code = nowRow.Cells[2] != null ? nowRow.Cells[2].Value.Trim() : null;
                // Xuất ra thông tin lên màn hình
                Console.WriteLine("MS: {0} ", code);
                Console.WriteLine("rowIndex: {0} ", rowIndex);

                if (code == null || code == "")
                    continue;


                double? luongcb = nowRow.Cells[6] != null && nowRow.Cells[6].Value != "NA" && nowRow.Cells[6].Value != "" ? nowRow.Cells[6].NumberValue : null;
                double? luongkpi = nowRow.Cells[16] != null && nowRow.Cells[16].Value != "NA" && nowRow.Cells[16].Value != "" ? (double)nowRow.Cells[16].FormulaValue : 0;
                double? tongtrocap = nowRow.Cells[15] != null && nowRow.Cells[15].Value != "NA" && nowRow.Cells[15].Value != "" ? (double)nowRow.Cells[15].FormulaValue : 0;
                double? khoancong = nowRow.Cells[18] != null && nowRow.Cells[18].Value != "NA" && nowRow.Cells[18].Value != "" ? (double)nowRow.Cells[18].NumberValue : 0;
                double? tongthunhap = nowRow.Cells[19] != null && nowRow.Cells[19].Value != "NA" && nowRow.Cells[19].Value != "" ? (double)nowRow.Cells[19].FormulaValue : null;
                int? nguoiphuthuoc = nowRow.Cells[24] != null && nowRow.Cells[24].Value != "NA" && nowRow.Cells[24].Value != "" ? (int)nowRow.Cells[24].NumberValue : null;
                double? pc_trachnhiem = nowRow.Cells[14] != null && nowRow.Cells[14].Value != "NA" && nowRow.Cells[14].Value != "" ? nowRow.Cells[14].NumberValue : 0;
                double? pc_xangxe = nowRow.Cells[10] != null && nowRow.Cells[10].Value != "NA" && nowRow.Cells[10].Value != "" ? nowRow.Cells[10].NumberValue : 0;
                double? tamungdot1 = 0;
                if (nowRow.Cells[34] != null && nowRow.Cells[34].Value != "NA" && nowRow.Cells[34].Value != "" && (double)nowRow.Cells[34].FormulaValue > 0)
                {
                    tamungdot1 = nowRow.Cells[33] != null && nowRow.Cells[33].Value != "NA" && nowRow.Cells[33].Value != "" ? nowRow.Cells[33].NumberValue : null;
                }
                //string? stk = nowRow.Cells[35] != null && nowRow.Cells[35].Value != "NA" && nowRow.Cells[35].Value != "" ? nowRow.Cells[35].Value.Trim() : 0;


                var findP = _nhansuContext.PersonnelModel.Where(d => d.MANV == code).FirstOrDefault();
                if (findP == null)
                {
                    continue;
                }
                findP.tien_luong = luongcb;
                findP.tien_luong_kpi = luongkpi;
                findP.tong_thunhap = luongcb + tongtrocap + luongkpi;
                findP.tien_luong_dot1 = tamungdot1;
                findP.nguoiphuthuoc = nguoiphuthuoc;
                findP.pc_trachnhiem = pc_trachnhiem;
                findP.pc_khac = pc_xangxe;
                //if (code == "NMK170962")
                //{
                //    Console.WriteLine("MS: {0} ", code);
                //}
                _nhansuContext.Update(findP);
                _nhansuContext.SaveChanges();
            }
            //_context.AddRange(list_Result);
            //_context.SaveChanges();

            return Json(new { success = true });
        }

        public async Task<JsonResult> syncPhep()
        {
            return Json(new { });
            // Khởi tạo workbook để đọc
            string folderPath = @"./wwwroot/data/info/phep";

            // Lấy tất cả các tệp trong thư mục
            string[] files = Directory.GetFiles(folderPath);
            var list_no = new List<string>();
            // Duyệt qua và in ra danh sách tệp
            foreach (string file in files)
            {
                Spire.Xls.Workbook book = new Spire.Xls.Workbook();
                book.LoadFromFile(file, ExcelVersion.Version2013);

                Spire.Xls.Worksheet sheet = book.Worksheets[0];
                var lastrow = sheet.LastDataRow;
                // nếu vẫn chưa gặp end thì vẫn lấy data
                Console.WriteLine(lastrow);
                var list_Result = new List<ResultModel>();
                for (int rowIndex = 6; rowIndex < lastrow; rowIndex++)
                {
                    // lấy row hiện tại
                    var nowRow = sheet.Rows[rowIndex];
                    if (nowRow == null)
                        continue;
                    // vì ta dùng 3 cột A, B, C => data của ta sẽ như sau
                    //int numcount = nowRow.Cells.Count;
                    //for(int y = 0;y<numcount - 1 ;y++)
                    var tennv = nowRow.Cells[1] != null ? nowRow.Cells[1].Value.Trim() : null;
                    // Xuất ra thông tin lên màn hình
                    Console.WriteLine("MS: {0} ", tennv);
                    Console.WriteLine("rowIndex: {0} ", rowIndex);

                    if (tennv == null || tennv == "")
                        continue;


                    double dauky = nowRow.Cells[74] != null ? nowRow.Cells[74].NumberValue : 0;
                    dauky = dauky > 0 ? dauky - 1 : 0;

                    var findP = _nhansuContext.PersonnelModel.Where(d => d.HOVATEN.Contains(tennv)).FirstOrDefault();
                    if (findP == null)
                    {
                        list_no.Add(tennv);
                        continue;
                    }
                    findP.ngayphep = dauky;
                    findP.ngayphep_date = new DateTime(2024, 8, 25);

                    _nhansuContext.Update(findP);
                    _nhansuContext.SaveChanges();
                }
            }
            return Json(list_no);

            //_context.AddRange(list_Result);
            //_context.SaveChanges();

            return Json(new { success = true });
        }
        public async Task<JsonResult> syncChamcong()
        {

            var list_person = _nhansuContext.PersonnelModel.ToList();


            var totalMatches = 152;
            var maxResults = 30;
            var start = 0;
            var failed = new List<string>();
            do
            {
                var employees = await getPerson(maxResults, start);
                totalMatches = employees.UserInfoSearch.totalMatches;
                foreach (var employee in employees.UserInfoSearch.UserInfo)
                {
                    var is_enable = employee.Valid.enable;
                    if (!is_enable)
                        continue;
                    var name = employee.name;
                    var macc = employee.employeeNo;
                    var Normalize = NormalizeName(name);
                    var person = list_person.Where(d => d.NormalizeName == Normalize).FirstOrDefault();
                    if (person == null)
                    {
                        failed.Add(name + " - " + Normalize + " - " + macc);
                        continue;
                    }
                    if (person.MACC != null)
                    {
                        continue;
                    }
                    person.MACC = macc;
                    _nhansuContext.Update(person);
                    _nhansuContext.SaveChanges();
                }
                start += employees.UserInfoSearch.UserInfo.Count();
            }
            while (start < totalMatches);



            return Json(new { success = true, failed, list_person });
        }
        public async Task<JsonResult> syncOtherLetterDetails()
        {

            var list_chamcong = _nhansuContext.ChamcongModel.Where(d => d.orderletter_id != null).ToList();

            foreach (var d in list_chamcong)
            {
                ///Copy
                var details = new OrderletterDetailsModel()
                {
                    id = d.id,
                    MANV = d.MANV,
                    NV_id = d.NV_id,
                    date = d.date,
                    shift_id = d.shift_id,
                    orderletter_id = d.orderletter_id,
                    value = d.value,
                    value_new = d.value_new,
                };
                _nhansuContext.Add(details);

            }
            _nhansuContext.SaveChanges();


            return Json(new { success = true });
        }

        public async Task<JsonResult> taovacapnhatEmail()
        {
            var list_email_update = _nhansuContext.PersonnelModel.Where(d => d.is_email_update == true && d.EMAIL.Contains("@astahealthcare.com") && d.NGAYNGHIVIEC == null).ToList();
            if (list_email_update.Count > 0)
            {

                var user = _configuration["Sync:username"];
                var pass = _configuration["Sync:password"];
                ////CHECK TẠO TÀI KHOẢN EMAIL
                ///



                var options = new ChromeOptions();
                //options.AddArgument("--headless"); // Chạy không cần giao diện (tuỳ chọn)

                // Khởi tạo WebDriver
                IWebDriver driver = new ChromeDriver(options);

                // Điều hướng đến một URL
                driver.Navigate().GoToUrl("https://mail.astahealthcare.com:2083/cpsess4532993050/frontend/jupiter/email_accounts/index.html#/list");
                driver.Manage().Window.Maximize();

                System.Threading.Thread.Sleep(2000);



                driver.FindElement(By.Name("user")).Clear();
                driver.FindElement(By.Name("user")).SendKeys(user);

                driver.FindElement(By.Name("pass")).Clear();
                driver.FindElement(By.Name("pass")).SendKeys(pass);


                driver.FindElement(By.Id("login_submit")).Click();

                Thread.Sleep(2000);


                var cookies = driver.Manage().Cookies.AllCookies;
                string cookieHeader = string.Join("; ", cookies.Select(c => $"{c.Name}={c.Value}"));
                var current_url10 = driver.Url;
                var uri10 = new Uri(current_url10);
                var AbsolutePath10 = uri10.AbsolutePath.Split('/');
                var cpress10 = AbsolutePath10.Count() > 1 ? AbsolutePath10[1] : "";
                var desiredUrl10 = $"{uri10.Scheme}://{uri10.Host}:{uri10.Port}/{cpress10}/execute/Email/list_pops_with_disk";
                var client = new HttpClient();
                var authToken = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{user}:{pass}"));
                client.DefaultRequestHeaders.Add("Cookie", cookieHeader);
                // Thêm Authorization header
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", authToken);

                var response = await client.GetAsync(desiredUrl10);
                if (response.IsSuccessStatusCode)
                {
                    LoginAdminResponse responseJson1 = await response.Content.ReadFromJsonAsync<LoginAdminResponse>();
                    var list = responseJson1.data.Select(d => d.email).ToList();
                    foreach (var item in list_email_update)
                    {
                        var email = item.EMAIL.ToLower();
                        var hovaten = item.HOVATEN;
                        var user_person = "";
                        string[] words = email.Split('@');
                        var is_asta = false;
                        if (words.Length > 1)
                        {
                            is_asta = words[1].Trim() == "astahealthcare.com" ? true : false;
                            user_person = words[0];
                        }

                        if (!is_asta)
                        {
                            continue;
                        }
                        if (list.Contains(email))
                        {
                            continue;
                        }
                        ///Tạo email mới
                        var client1 = new HttpClient();
                        var authToken1 = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{user}:{pass}"));
                        client1.DefaultRequestHeaders.Add("Cookie", cookieHeader);

                        // Thêm Authorization header
                        client1.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", authToken1);
                        string generatedPassword = $"!{user_person}@123";
                        string url1 = $"{uri10.Scheme}://{uri10.Host}:{uri10.Port}/{cpress10}/execute/Email/add_pop?email={email}&password={generatedPassword}";
                        item.mk_email = generatedPassword;
                        HttpResponseMessage response1 = await client1.GetAsync(url1);
                        if (response1.IsSuccessStatusCode)
                        {
                            string responseBody = await response1.Content.ReadAsStringAsync();
                            Console.WriteLine("Tạo tài khoản email thành công:");
                            Console.WriteLine(responseBody);
                        }
                        else
                        {
                            Console.WriteLine("Lỗi khi tạo tài khoản email:");
                            Console.WriteLine(await response1.Content.ReadAsStringAsync());
                        }

                        /////Thêm vào esign
                        UserModel user_esign = new UserModel
                        {
                            Email = user_person + "@astahealthcare.com",
                            UserName = user_person + "@astahealthcare.com",
                            EmailConfirmed = true,
                            FullName = hovaten,
                            msnv = item.MANV,
                            image_sign = "/private/images/tick.png",
                            image_url = "/private/images/user.webp",
                        };
                        user_esign.deleted_at = null;
                        user_esign.LockoutEnd = null;

                        _context.Add(user_esign);

                        _context.SaveChanges();
                        var user_role = new UserRoleModel
                        {
                            UserId = user_esign.Id,
                            RoleId = "fa63ccaa-1f31-4d24-8344-20eca0410141"
                        };
                        _context.Add(user_role);
                        _context.SaveChanges();

                        CreatePfx(user_esign);

                    }
                }

                foreach (var person in list_email_update)
                {
                    // Đợi một chút để trang tải (tối ưu bằng WebDriverWait nếu cần)
                    var email = person.EMAIL.ToLower();
                    var hovaten = person.HOVATEN;
                    var user_person = "";
                    string[] words = email.Split('@');
                    var is_asta = false;
                    if (words.Length > 1)
                    {
                        is_asta = words[1].Trim() == "astahealthcare.com" ? true : false;
                        user_person = words[0];
                    }

                    if (!is_asta)
                    {
                        continue;
                    }

                    driver.FindElement(By.Id("email_table_search_input")).Clear();
                    driver.FindElement(By.Id("email_table_search_input")).SendKeys(user_person);

                    Thread.Sleep(2000);

                    driver.FindElement(By.Id("email_table_menu_webmail_" + email)).Click();



                    ////SỬA HỌ VÀ TÊN.

                    // Lưu lại cửa sổ hiện tại
                    string originalWindow = driver.CurrentWindowHandle;

                    // Chuyển sang cửa sổ mới
                    foreach (string window in driver.WindowHandles)
                    {
                        if (window != originalWindow)
                        {
                            driver.SwitchTo().Window(window);
                            break;
                        }
                    }

                    // Đợi cho đến khi trang mới tải hoàn toàn (nếu cần)
                    System.Threading.Thread.Sleep(2000);


                    //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                    //IWebElement element = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.Id("element-id")));

                    //System.Threading.Thread.Sleep(5000);
                    //driver.FindElement(By.Id("rcmbtn110")).Click();
                    //System.Threading.Thread.Sleep(2000);
                    //driver.FindElement(By.Id("rcmbtn112")).Click();

                    var current_url = driver.Url;
                    // Phân tích URL
                    Uri uri = new Uri(current_url);

                    // Lấy base URL (không bao gồm query string)
                    string baseUrl = $"{uri.Scheme}://{uri.Host}:{uri.Port}{uri.AbsolutePath}";

                    // Thay đổi đường dẫn và query string
                    string desiredPath = uri.AbsolutePath.Replace("webmail/jupiter/index.html", "3rdparty/roundcube/");


                    string desiredQuery = "?_task=mail&_mbox=INBOX";

                    // Tái tạo URL
                    string desiredUrl = $"{uri.Scheme}://{uri.Host}:{uri.Port}{desiredPath}?{desiredQuery}";
                    //string desiredUrl = current_url.Replace("_action=identities", "_action=edit-identity&_iid=1&_framed=1");

                    driver.Navigate().GoToUrl(desiredUrl);


                    System.Threading.Thread.Sleep(2000);

                    desiredQuery = "_task=settings&_action=edit-identity&_iid=1&_framed=1";
                    desiredUrl = $"{uri.Scheme}://{uri.Host}:{uri.Port}{desiredPath}?{desiredQuery}";

                    driver.Navigate().GoToUrl(desiredUrl);


                    System.Threading.Thread.Sleep(2000);

                    driver.FindElement(By.Id("rcmfd_name")).Clear();
                    driver.FindElement(By.Id("rcmfd_name")).SendKeys(hovaten);

                    driver.FindElement(By.Id("rcmbtnfrm102")).Click();
                    driver.Close();

                    // Quay lại cửa sổ ban đầu
                    driver.SwitchTo().Window(originalWindow);






                    person.is_email_update = false;

                    _nhansuContext.Update(person);
                    _nhansuContext.SaveChanges();
                }


                /////Add vào asta.all
                var list_email = list_email_update.Select(d => d.EMAIL.ToLower()).ToList();
                var current_url1 = driver.Url;
                var uri1 = new Uri(current_url1);
                var AbsolutePath = uri1.AbsolutePath.Split('/');
                var cpress = AbsolutePath.Count() > 1 ? AbsolutePath[1] : "";
                var desiredUrl1 = $"{uri1.Scheme}://{uri1.Host}:{uri1.Port}/{cpress}/frontend/jupiter/mail/lists.html";

                driver.Navigate().GoToUrl(desiredUrl1);

                System.Threading.Thread.Sleep(2000);
                driver.FindElement(By.Id("datarow_asta_all_astahealthcare_com_manage_btn")).Click();
                // Lưu lại cửa sổ hiện tại
                var originalWindow1 = driver.CurrentWindowHandle;

                // Chuyển sang cửa sổ mới
                foreach (string window in driver.WindowHandles)
                {
                    if (window != originalWindow1)
                    {
                        driver.SwitchTo().Window(window);
                        break;
                    }
                }
                // Đợi cho đến khi trang mới tải hoàn toàn (nếu cần)
                System.Threading.Thread.Sleep(2000);

                current_url1 = driver.Url;
                uri1 = new Uri(current_url1);
                var AbsolutePath1 = uri1.AbsolutePath.Split('/');
                var cpress1 = AbsolutePath1.Count() > 1 ? AbsolutePath1[1] : "";
                desiredUrl1 = $"{uri1.Scheme}://{uri1.Host}:{uri1.Port}/{cpress1}/3rdparty/mailman/admin/asta.all_astahealthcare.com/members/add";

                driver.Navigate().GoToUrl(desiredUrl1);

                System.Threading.Thread.Sleep(2000);
                driver.FindElement(By.Name("subscribees")).Clear();
                driver.FindElement(By.Name("subscribees")).SendKeys(string.Join("\n", list_email));
                driver.FindElement(By.Name("setmemberopts_btn")).Click();

                System.Threading.Thread.Sleep(2000);
                driver.Close();
                driver.SwitchTo().Window(originalWindow1);




                //////Add vào sharecontact
                // Điều hướng đến một URL

                Thread.Sleep(2000);
                current_url1 = driver.Url;
                uri1 = new Uri(current_url1);
                AbsolutePath1 = uri1.AbsolutePath.Split('/');
                cpress1 = AbsolutePath1.Count() > 1 ? AbsolutePath1[1] : "";
                desiredUrl1 = $"{uri1.Scheme}://{uri1.Host}:{uri1.Port}/{cpress}/frontend/jupiter/email_accounts/index.html#/list";

                driver.Navigate().GoToUrl(desiredUrl1);
                Thread.Sleep(2000);


                driver.FindElement(By.Id("email_table_search_input")).Clear();
                driver.FindElement(By.Id("email_table_search_input")).SendKeys("sharecontact");

                Thread.Sleep(2000);

                driver.FindElement(By.Id("email_table_menu_webmail_sharecontact@astahealthcare.com")).Click();

                // Lưu lại cửa sổ hiện tại
                originalWindow1 = driver.CurrentWindowHandle;

                // Chuyển sang cửa sổ mới
                foreach (string window in driver.WindowHandles)
                {
                    if (window != originalWindow1)
                    {
                        driver.SwitchTo().Window(window);
                        break;
                    }
                }

                // Đợi cho đến khi trang mới tải hoàn toàn (nếu cần)
                System.Threading.Thread.Sleep(2000);


                current_url1 = driver.Url;
                uri1 = new Uri(current_url1);
                AbsolutePath1 = uri1.AbsolutePath.Split('/');
                cpress1 = AbsolutePath1.Count() > 1 ? AbsolutePath1[1] : "";
                foreach (var person in list_email_update)
                {
                    desiredUrl1 = $"{uri1.Scheme}://{uri1.Host}:{uri1.Port}/{cpress1}/3rdparty/roundcube/?_task=addressbook&_framed=1&_action=add&_source=carddav_1";

                    driver.Navigate().GoToUrl(desiredUrl1);

                    System.Threading.Thread.Sleep(2000);

                    var email = person.EMAIL.ToLower();
                    var hovaten = person.HOVATEN;
                    driver.FindElement(By.Id("ff_email0")).Clear();
                    driver.FindElement(By.Id("ff_email0")).SendKeys(email);

                    //driver.FindElement(By.ClassName("displayname")).Click();
                    IWebElement element = driver.FindElement(By.ClassName("displayname"));
                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                    js.ExecuteScript("arguments[0].style.display='block';", element);

                    driver.FindElement(By.Id("ff_name")).Clear();
                    driver.FindElement(By.Id("ff_name")).SendKeys(hovaten);

                    driver.FindElement(By.Id("rcmbtnfrm102")).Click();

                    System.Threading.Thread.Sleep(2000);
                }
                driver.Close();

                // Quay lại cửa sổ ban đầu
                driver.SwitchTo().Window(originalWindow1);

                driver.Close();
            }



            return Json(new { success = true });
        }
        private async Task<UserInfoSearchResponse> getPerson(int maxResults, int start)
        {
            CredentialCache credentialCache = new CredentialCache();
            var url = "https://192.168.40.35/ISAPI/AccessControl/UserInfo/Search?format=json";
            var username = "admin";
            var password = "1234567a";
            credentialCache.Add(new Uri(url), "Digest", new NetworkCredential(username, password));

            using var httpClientHandler = new HttpClientHandler
            {
                Credentials = credentialCache,
                ServerCertificateCustomValidationCallback = (message, cert, chain, sslPolicyErrors) => true  // Bypass certificate validation

            };
            var client = new HttpClient(httpClientHandler);
            // Define the JSON payload
            var jsonPayload = new
            {
                UserInfoSearchCond = new
                {
                    searchID = "9691a9581fd34b83bb2e93d8923897bb",
                    maxResults = maxResults,
                    searchResultPosition = start
                }
            };

            // Serialize the JSON object to a string
            var jsonContent = new StringContent(JsonSerializer.Serialize(jsonPayload), Encoding.UTF8, "application/json");
            // Send the POST request with JSON content
            var response = await client.PostAsync(url, jsonContent);

            // Ensure the response is successful
            response.EnsureSuccessStatusCode();

            // Optionally read the response content
            var responseContent = await response.Content.ReadAsStringAsync();
            var employees = JsonSerializer.Deserialize<UserInfoSearchResponse>(responseContent);
            return employees;
        }
        // Hàm loại bỏ dấu tiếng Việt
        public static string RemoveDiacritics(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return text;
            text = text.Replace('đ', 'd').Replace('Đ', 'D');
            text = text.Normalize(NormalizationForm.FormD);
            var chars = text.ToCharArray();
            StringBuilder sb = new StringBuilder();

            foreach (var c in chars)
            {
                UnicodeCategory unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    sb.Append(c);
                }
            }

            return sb.ToString().Normalize(NormalizationForm.FormC);
        }

        // Hàm chuẩn hóa tên
        public static string NormalizeName(string name)
        {
            // Chuyển về chữ thường
            string normalized = name.ToLower();

            // Loại bỏ dấu tiếng Việt
            normalized = RemoveDiacritics(normalized);

            // Xóa khoảng trắng thừa
            normalized = Regex.Replace(normalized, @"\s+", " ").Trim();

            return normalized;
        }
        static string GenerateStrongPassword(int length = 12)
        {
            const string uppercase = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            const string lowercase = "abcdefghijklmnopqrstuvwxyz";
            const string digits = "0123456789";
            const string specialChars = "!@#$%^&*()-_=+[]{}|;:,.<>?";
            const string allChars = uppercase + lowercase + digits + specialChars;

            Random random = new Random();
            char[] password = new char[length];

            // Đảm bảo mật khẩu chứa ít nhất một ký tự từ mỗi nhóm
            password[0] = uppercase[random.Next(uppercase.Length)];
            password[1] = lowercase[random.Next(lowercase.Length)];
            password[2] = digits[random.Next(digits.Length)];
            password[3] = specialChars[random.Next(specialChars.Length)];

            // Điền các ký tự còn lại ngẫu nhiên
            for (int i = 4; i < length; i++)
            {
                password[i] = allChars[random.Next(allChars.Length)];
            }

            // Trộn ngẫu nhiên các ký tự trong mật khẩu
            Shuffle(password, random);

            return new string(password);
        }

        private JsonResult CreatePfx(UserModel user)
        {
            // Generate private-public key pair
            var serviceProvider = new ServiceCollection()
                  .AddCertificateManager()
                  .BuildServiceProvider();

            string passwordPublic = _configuration["Matkhau:PFX"];
            var createClientServerAuthCerts = serviceProvider.GetService<CreateCertificatesClientServerAuth>();
            var private_f = _configuration["Source:Path_Private"];
            X509Certificate2 rootCaL1 = new X509Certificate2(private_f + "\\rootca\\localhost_root.pfx", passwordPublic);
            var serverL3 = createClientServerAuthCerts.NewClientChainedCertificate(
                new DistinguishedName { CommonName = user.FullName + "<" + user.Email + ">", OrganisationUnit = "" },
                new ValidityPeriod { ValidFrom = DateTime.UtcNow, ValidTo = DateTime.UtcNow.AddYears(10) },
                "localhost", rootCaL1);
            var importExportCertificate = serviceProvider.GetService<ImportExportCertificate>();
            var serverCertL3InPfxBtyes = importExportCertificate.ExportChainedCertificatePfx(passwordPublic, serverL3, rootCaL1);
            System.IO.File.WriteAllBytes(private_f + "\\pfx\\" + user.Id + ".pfx", serverCertL3InPfxBtyes);

            user.signature = "/private/pfx/" + user.Id + ".pfx";
            _context.Update(user);
            _context.SaveChanges();
            return Json(new { success = true });

        }
        static void Shuffle(char[] array, Random random)
        {
            for (int i = array.Length - 1; i > 0; i--)
            {
                int j = random.Next(i + 1);
                char temp = array[i];
                array[i] = array[j];
                array[j] = temp;
            }
        }
    }
    public class LoginAdminResponse
    {
        [Key]
        public int id { get; set; }

        public List<UserResult>? data { get; set; }
    }
    public class UserResult
    {
        [Key]
        public string id { get; set; }
        public string email { get; set; }
        public string user { get; set; }

        //public string fullName { get; set; }
        //public string description { get; set; }
        public int? suspended_login { get; set; }
    }
    public class Valid
    {
        public bool enable { get; set; }
        public DateTime? beginTime { get; set; }
        public DateTime? endTime { get; set; }
        public string? timeType { get; set; }
    }
    public class UserInfo
    {
        public string employeeNo { get; set; }
        public string name { get; set; }
        public Valid Valid { get; set; }
    }

    public class UserInfoSearch
    {
        public string searchID { get; set; }
        public string responseStatusStrg { get; set; }
        public int numOfMatches { get; set; }
        public int totalMatches { get; set; }
        public List<UserInfo> UserInfo { get; set; }
    }

    public class UserInfoSearchResponse
    {
        public UserInfoSearch UserInfoSearch { get; set; }
    }
    class SuccesMail
    {
        public int success { get; set; }
        public Exception ex { get; set; }
    }
}
