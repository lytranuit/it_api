
using Vue.Models;
using Microsoft.AspNetCore.Mvc;
using Vue.Data;
using System.Net.Mail;
using Vue.Services;
using System.Net.Mime;
using Microsoft.EntityFrameworkCore.Infrastructure;
using System.Diagnostics;
using Microsoft.EntityFrameworkCore;

namespace Vue.Controllers
{

    public class HomeController : Controller
    {
        protected readonly ItContext _context;
        private readonly ViewRender _view;
        protected readonly HoldTimeContext _holdtimecontext;

        private readonly IConfiguration _configuration;

        public HomeController(ItContext context, HoldTimeContext holdtimecontext, IConfiguration configuration, ViewRender view)
        {
            _configuration = configuration;
            _context = context;
            _holdtimecontext = holdtimecontext;
            _view = view;
            var listener = _context.GetService<DiagnosticSource>();
            (listener as DiagnosticListener).SubscribeWithAdapter(new CommandInterceptor());
        }

        public JsonResult Index()
        {
            return Json(new { test = 1, message = DateTime.Now });

        }


        public async Task<JsonResult> cronjobHoldtimeDaily()
        {
            var timecheck = DateTime.Now;
            var timecheck1 = DateTime.Now.AddDays(1);
            var timecheck2 = DateTime.Now.AddDays(-1);
            //var tasks = _context.TaskModel.Where(d => d.deleted_at == null && (d.finished_at == null || d.progress != 100) && d.endDate != null && d.endDate < DateTime.Now.AddDays(1)).Include(d => d.assignees).ToList();
            var holdtime = _holdtimecontext.HoldTimeModel
                .Where(d => d.deleted_at == null && d.date_reality == null && d.date_theory >= timecheck && d.date_theory <= timecheck1)
                .Include(d => d.Hold).ThenInclude(d => d.stage)
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
            var query = _holdtimecontext.HoldTimeModel.Where(d => d.deleted_at == null && d.date_reality == null && d.date_theory < timecheck2).ToList();

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


    }
    class SuccesMail
    {
        public int success { get; set; }
        public Exception ex { get; set; }
    }
}
