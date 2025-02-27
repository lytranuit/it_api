using Info.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Spire.Xls;
using System;
using System.Collections;
using System.Data;
using System.Text.Json.Serialization;
using Vue.Data;
using Vue.Models;
using Vue.Services;

namespace it_template.Areas.Info.Controllers
{

    public class MeetingController : BaseController
    {
        private readonly IConfiguration _configuration;
        private UserManager<UserModel> UserManager;
        private readonly ViewRender _view;
        public MeetingController(NhansuContext context, AesOperation aes, IConfiguration configuration, UserManager<UserModel> UserMgr, ViewRender view) : base(context, aes)
        {
            _configuration = configuration;
            UserManager = UserMgr;
            _view = view;
        }
        public async Task<JsonResult> Get(int id)
        {

            var Meeting = _context.MeetingModel.Where(d => d.id == id).FirstOrDefault();

            return Json(Meeting, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                ReferenceHandler = ReferenceHandler.IgnoreCycles,
            });
        }
        [HttpPost]
        public async Task<JsonResult> Remove(int id)
        {
            var jsonData = new { success = true, message = "" };
            try
            {
                var list = _context.MeetingModel.Where(d => d.id == id).FirstOrDefault();
                list.deleted_at = DateTime.Now;
                _context.Update(list);
                _context.SaveChanges();
            }
            catch (Exception ex)
            {
                jsonData = new { success = false, message = ex.Message };
            }


            return Json(jsonData, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                ReferenceHandler = ReferenceHandler.IgnoreCycles,
            });
        }
        [HttpPost]
        public async Task<JsonResult> Save(MeetingModel MeetingModel)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user = await UserManager.GetUserAsync(currentUser);
            var properties = typeof(MeetingModel).GetProperties().Where(prop => prop.CanRead && prop.CanWrite && prop.PropertyType == typeof(DateTime?));

            foreach (var prop in properties)
            {
                DateTime? value = (DateTime?)prop.GetValue(MeetingModel, null);
                if (value != null && value.Value.Kind == DateTimeKind.Utc)
                {
                    value = value.Value.ToLocalTime();
                    prop.SetValue(MeetingModel, value, null);
                }
            }
            if (MeetingModel.id > 0)
            {
                var MeetingModel_old = _context.MeetingModel.Where(d => d.id == MeetingModel.id).FirstOrDefault();

                CopyValues<MeetingModel>(MeetingModel_old, MeetingModel);

                _context.Update(MeetingModel_old);
                _context.SaveChanges();
            }
            else
            {
                MeetingModel.created_at = DateTime.Now;
                MeetingModel.created_by = user.Id;
                _context.Add(MeetingModel);
                _context.SaveChanges();

            }
            ////Gửi email đến người nhận
            ///
            var list_notify = MeetingModel.list_notify_id;

            var user_related = _context.UserModel.Where(d => list_notify.Contains(d.Id)).Select(d => d.Email).ToList();

            //user_related = user_related.Distinct().ToList();

            var mail_string = string.Join(",", user_related);
            string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;
            var body = _view.Render("Emails/Meeting", new
            {
                link_logo = Domain + "/images/clientlogo_astahealthcare.com_f1800.png",
                link = _configuration["Application:Info:link"] + "Meeting",
                data = MeetingModel,
                user = user.FullName
            });
            var email = new EmailModel
            {
                email_to = mail_string,
                subject = "[Đăng ký phòng họp] " + MeetingModel.name,
                body = body,
                email_type = "Meeting",
                status = 1
            };
            _context.Add(email);
            _context.SaveChanges();
            return Json(new { success = true });

        }

        //[HttpPost]
        public async Task<JsonResult> Table()
        {
            var draw = Request.Form["draw"].FirstOrDefault();
            var start = Request.Form["start"].FirstOrDefault();
            var length = Request.Form["length"].FirstOrDefault();
            int pageSize = length != null ? Convert.ToInt32(length) : 0;
            var mahh = Request.Form["filters[mahh]"].FirstOrDefault();
            var tenhh = Request.Form["filters[tenhh]"].FirstOrDefault();
            var mahh_goc = Request.Form["filters[mahh_goc]"].FirstOrDefault();
            var tenhh_goc = Request.Form["filters[tenhh_goc]"].FirstOrDefault();
            var colo = Request.Form["filters[colo]"].FirstOrDefault();
            decimal? colo_d = colo != null ? Convert.ToDecimal(colo) : null;
            var mapl = Request.Form["filters[mapl]"].FirstOrDefault();
            var nvl = Request.Form["filters[nvl]"].FirstOrDefault();
            var phong_hop = Request.Form["phong_hop"].FirstOrDefault();
            int phong_hop_id = phong_hop != null ? Convert.ToInt32(phong_hop) : 0;
            var status_string = Request.Form["filters[status]"].FirstOrDefault();
            int status = status_string != null ? Convert.ToInt32(status_string) : 0;
            var date_from = Request.Form["filters[dates][0]"].FirstOrDefault();
            var date_to = Request.Form["filters[dates][1]"].FirstOrDefault();
            var date0 = DateTime.Now;
            var date1 = DateTime.Now;
            int skip = start != null ? Convert.ToInt32(start) : 0;

            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user = await UserManager.GetUserAsync(currentUser);


            var customerData = _context.MeetingModel.Where(d => d.deleted_at == null);
            int recordsTotal = customerData.Count();
            if (date_from != null)
            {

                date0 = DateTime.Parse(date_from);
                if (date0.Kind == DateTimeKind.Utc)
                {
                    date0 = date0.ToLocalTime();
                }
            }
            if (date_to != null)
            {

                date1 = DateTime.Parse(date_to);
                if (date1.Kind == DateTimeKind.Utc)
                {
                    date1 = date1.ToLocalTime();
                }
            }
            if (phong_hop_id > 0)
            {

                customerData = customerData.Where(d => d.phong_hop_id == phong_hop_id);
            }
            int recordsFiltered = customerData.Count();
            var datapost = customerData.OrderBy(d => d.id).Skip(skip).Take(pageSize).ToList();
            var jsonData = new { draw = draw, recordsFiltered = recordsFiltered, recordsTotal = recordsTotal, data = datapost };
            return Json(jsonData, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                ReferenceHandler = ReferenceHandler.IgnoreCycles,
            });
        }

        public void CopyValues<T>(T target, T source)
        {
            Type t = typeof(T);

            var properties = t.GetProperties().Where(prop => prop.CanRead && prop.CanWrite);

            foreach (var prop in properties)
            {
                var value = prop.GetValue(source, null);
                //if (value != null)
                prop.SetValue(target, value, null);
            }
        }


    }
}
