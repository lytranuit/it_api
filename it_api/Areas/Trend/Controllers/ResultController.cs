

using elFinder.NetCore.Models;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Spire.Xls;
using System;
using System.Collections;
using System.Data;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using Vue.Data;
using Vue.Models;

namespace it_template.Areas.Trend.Controllers
{

    public class ResultController : BaseController
    {
        private readonly IConfiguration _configuration;
        private UserManager<UserModel> UserManager;
        public ResultController(ItContext context, IConfiguration configuration, UserManager<UserModel> UserMgr) : base(context)
        {
            _configuration = configuration;
            UserManager = UserMgr;
        }

        [HttpPost]
        public async Task<JsonResult> Delete(int id)
        {
            var Model = _context.ResultModel.Where(d => d.id == id).FirstOrDefault();
            Model.deleted_at = DateTime.Now;
            _context.Update(Model);
            _context.SaveChanges();
            return Json(new { success = true, data = Model });
        }
        [HttpPost]
        public async Task<JsonResult> Save(ResultModel ResultModel, bool is_replace = false)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            if (ResultModel.date != null && ResultModel.date.Value.Kind == DateTimeKind.Utc)
            {
                ResultModel.date = ResultModel.date.Value.ToLocalTime();
            }
            ResultModel? ResultModel_old;
            if (is_replace)
            {
                ResultModel_old = _context.ResultModel.Where(d => d.date == ResultModel.date && d.point_id == ResultModel.point_id && d.target_id == ResultModel.target_id).FirstOrDefault();
                ResultModel_old.value = ResultModel.value;
                ResultModel_old.value_text = ResultModel.value_text;
                ResultModel_old.updated_at = DateTime.Now;
                ResultModel_old.deleted_at = null;

                _context.Update(ResultModel_old);
                _context.SaveChanges();
            }
            else if (ResultModel.id == 0)
            {
                try
                {
                    ResultModel.created_at = DateTime.Now;
                    ResultModel.created_by = user_id;


                    _context.ResultModel.Add(ResultModel);

                    _context.SaveChanges();


                    ///UPDATE Limit
                    var r = _context.ResultModel.Where(d => d.id == ResultModel.id).FirstOrDefault();
                    var limits = _context.LimitModel.Where(d => d.target_id == r.target_id).Include(d => d.points).ToList();

                    LimitModel? limit;
                    if (limits != null && limits.Count() > 0)
                    {
                        limit = limits.Where(d => d.list_point.Contains(r.point_id.Value) && d.deleted_at == null && d.date_effect <= r.date).OrderBy(d => d.date_effect).ThenBy(d => d.id).LastOrDefault();
                        r.limit_id = limit != null ? limit.id : null;
                        r.limit_action = limit != null ? limit.action_limit : null;
                        r.limit_alert = limit != null ? limit.alert_limit : null;

                        _context.Update(r);
                        _context.SaveChanges();
                    }

                }
                catch (DbUpdateException ex)
                {
                    SqlException innerException = ex.InnerException as SqlException;
                    if (innerException != null && (innerException.Number == 2627 || innerException.Number == 2601))
                    {
                        return Json(new { success = false, message = "Dữ liệu bị trùng", is_duplicate = true });
                    }
                    else
                    {
                        return Json(new { success = false, message = innerException.Message });
                    }
                }
            }
            else
            {

                ResultModel_old = _context.ResultModel.Where(d => d.id == ResultModel.id).FirstOrDefault();
                CopyValues<ResultModel>(ResultModel_old, ResultModel);
                ResultModel_old.updated_at = DateTime.Now;

                _context.Update(ResultModel_old);
                _context.SaveChanges();

                ///UPDATE Limit
                var r = _context.ResultModel.Where(d => d.id == ResultModel.id).FirstOrDefault();
                var limits = _context.LimitModel.Where(d => d.target_id == r.target_id).Include(d => d.points).ToList();

                LimitModel? limit;
                if (limits != null && limits.Count() > 0)
                {
                    limit = limits.Where(d => d.list_point.Contains(r.point_id.Value) && d.deleted_at == null && d.date_effect <= r.date).OrderBy(d => d.date_effect).ThenBy(d => d.id).LastOrDefault();
                    r.limit_id = limit != null ? limit.id : null;
                    r.limit_action = limit != null ? limit.action_limit : null;
                    r.limit_alert = limit != null ? limit.alert_limit : null;

                    _context.Update(r);
                    _context.SaveChanges();
                }

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
            var name = Request.Form["filters[name]"].FirstOrDefault();
            var point_id = Request.Form["filters[point_id]"].FirstOrDefault();
            var object_id_text = Request.Form["filters[object_id]"].FirstOrDefault();
            var target_id_text = Request.Form["filters[target_id]"].FirstOrDefault();
            var date_text = Request.Form["filters[date]"].FirstOrDefault();
            var id_text = Request.Form["filters[id]"].FirstOrDefault();

            DateTime? date = date_text != null && date_text != "" ? DateTime.Parse(date_text) : null;
            int id = id_text != null ? Convert.ToInt32(id_text) : 0;
            int object_id = object_id_text != null ? Convert.ToInt32(object_id_text) : 0;
            int target_id = target_id_text != null ? Convert.ToInt32(target_id_text) : 0;
            //var tenhh = Request.Form["filters[tenhh]"].FirstOrDefault();
            int skip = start != null ? Convert.ToInt32(start) : 0;
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;

            var user_id = UserManager.GetUserId(currentUser);
            var user_current = await UserManager.GetUserAsync(currentUser); // Get user id:
            var is_admin = await UserManager.IsInRoleAsync(user_current, "Administrator");
            var cusdata = _context.ObjectModel.Where(d => d.deleted_at == null);


            var customerData = _context.ResultModel.Where(d => d.deleted_at == null);
            if (!is_admin)
            {
                var objects = _context.UserObjectModel.Where(d => d.user_id == user_id).Select(d => d.object_id).ToList();
                customerData = customerData.Where(d => objects.Contains(d.object_id.Value));
            }

            int recordsTotal = customerData.Count();
            if (point_id != null && point_id != "")
            {
                var list_point = _context.PointModel.Where(d => d.code.Contains(point_id)).Select(d => d.id).ToList();
                customerData = customerData.Where(d => list_point.Contains(d.point_id.Value));
            }
            if (date != null)
            {
                customerData = customerData.Where(d => d.date == date);
            }
            if (id != 0)
            {
                customerData = customerData.Where(d => d.id == id);
            }
            if (object_id != 0)
            {
                customerData = customerData.Where(d => d.object_id == object_id);
            }
            if (target_id != 0)
            {
                customerData = customerData.Where(d => d.target_id == target_id);
            }
            int recordsFiltered = customerData.Count();
            var datapost = customerData.OrderByDescending(d => d.id).Skip(skip).Take(pageSize)
                .Include(d => d.target)
                .Include(d => d.point)
                .ThenInclude(d => d.obj).ToList();
            //var data = new ArrayList();
            //foreach (var record in datapost)
            //{
            //	var ngaythietke = record.ngaythietke != null ? record.ngaythietke.Value.ToString("yyyy-MM-dd") : null;
            //	var ngaysodk = record.ngaysodk != null ? record.ngaysodk.Value.ToString("yyyy-MM-dd") : null;
            //	var ngayhethanthietke = record.ngayhethanthietke != null ? record.ngayhethanthietke.Value.ToString("yyyy-MM-dd") : null;
            //	var data1 = new
            //	{
            //		mahh = record.mahh,
            //		tenhh = record.tenhh,
            //		dvt = record.dvt,
            //		mansx = record.mansx,
            //		mancc = record.mancc,
            //		tennvlgoc = record.tennvlgoc,
            //		masothietke = record.masothietke,
            //		ghichu_thietke = record.ghichu_thietke,
            //		masodk = record.masodk,
            //		ghichu_sodk = record.ghichu_sodk,
            //		nhuongquyen = record.nhuongquyen,
            //		ngaythietke = ngaythietke,
            //		ngaysodk = ngaysodk,
            //		ngayhethanthietke = ngayhethanthietke
            //	};
            //	data.Add(data1);
            //}
            var jsonData = new { draw = draw, recordsFiltered = recordsFiltered, recordsTotal = recordsTotal, data = datapost };
            return Json(jsonData);
        }

        public JsonResult Get(int id)
        {
            var data = _context.ResultModel.Where(d => d.id == id).FirstOrDefault();
            return Json(data);
        }
        private void CopyValues<T>(T target, T source)
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
