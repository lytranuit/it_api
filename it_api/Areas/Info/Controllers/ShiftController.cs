



using Holdtime.Models;
using Info.Models;
using it_template.Areas.Trend.Controllers;
using iText.Commons.Actions.Contexts;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Data;
using System.Text.Json.Serialization;
using Vue.Data;
using Vue.Models;
using Vue.Services;

namespace it_template.Areas.Info.Controllers
{

    [Authorize(Roles = "Administrator,HR")]
    public class ShiftController : BaseController
    {
        private readonly IConfiguration _configuration;
        private UserManager<UserModel> UserManager;
        public ShiftController(NhansuContext context, AesOperation aes, IConfiguration configuration, UserManager<UserModel> UserMgr) : base(context, aes)
        {
            _configuration = configuration;
            UserManager = UserMgr;
        }

        [HttpPost]
        public async Task<JsonResult> Delete(string id)
        {
            var Model = _context.ShiftModel.Where(d => d.id == id).FirstOrDefault();
            Model.deleted_at = DateTime.Now;
            _context.Update(Model);
            _context.SaveChanges();
            return Json(new { success = true, data = Model });
        }
        [HttpPost]
        public async Task<JsonResult> Save(ShiftModel ShiftModel, List<ShiftHolidayModel> list_add, List<string> list_remove, List<string> list_user)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            ShiftModel? ShiftModel_old;
            if (ShiftModel.id == null)
            {
                ShiftModel.id = Guid.NewGuid().ToString();
                ShiftModel.created_at = DateTime.Now;

                //ShiftModel.created_by = user_id;


                _context.ShiftModel.Add(ShiftModel);

                _context.SaveChanges();

                ShiftModel_old = ShiftModel;

            }
            else
            {

                ShiftModel_old = _context.ShiftModel.Where(d => d.id == ShiftModel.id).FirstOrDefault();
                CopyValues<ShiftModel>(ShiftModel_old, ShiftModel);
                ShiftModel_old.updated_at = DateTime.Now;

                _context.Update(ShiftModel_old);
                _context.SaveChanges();
            }
            /////
            if (list_remove != null && list_remove.Count() > 0)
            {
                var list = _context.ShiftHolidayModel.Where(d => list_remove.Contains(d.id)).ToList();
                _context.RemoveRange(list);
                _context.SaveChanges();
            }
            if (list_add != null && list_add.Count() > 0)
            {
                foreach (var item in list_add)
                {
                    item.id = Guid.NewGuid().ToString();
                    item.shift_id = ShiftModel_old.id;

                }
                _context.AddRange(list_add);
                _context.SaveChanges();
            }
            /////
            //////
            var list_user_old = _context.ShiftUserModel.Where(d => d.shift_id == ShiftModel_old.id).Select(d => d.person_id).ToList();
            IEnumerable<string> list_delete_user = list_user_old.Except(list_user);
            IEnumerable<string> list_add_user = list_user.Except(list_user_old);

            if (list_add_user != null)
            {
                foreach (string key in list_add_user)
                {
                    var Model = _context.PersonnelModel.Where(d => d.id == key).FirstOrDefault();
                    _context.Add(new ShiftUserModel()
                    {
                        shift_id = ShiftModel_old.id,
                        person_id = key,
                        email = Model.EMAIL
                    });
                }
                //_context.SaveChanges();
            }
            if (list_delete_user != null)
            {
                foreach (string key in list_delete_user)
                {
                    ShiftUserModel ShiftUserModel = _context.ShiftUserModel.Where(d => d.shift_id == ShiftModel_old.id && d.person_id == key).First();
                    _context.Remove(ShiftUserModel);
                }
                //_context.SaveChanges();
            }
            _context.SaveChanges();

            return Json(new { success = true, data = ShiftModel_old });
        }

        [HttpPost]
        public async Task<JsonResult> Table()
        {
            var draw = Request.Form["draw"].FirstOrDefault();
            var start = Request.Form["start"].FirstOrDefault();
            var length = Request.Form["length"].FirstOrDefault();
            int pageSize = length != null ? Convert.ToInt32(length) : 0;
            var code = Request.Form["filters[code]"].FirstOrDefault();
            var name = Request.Form["filters[name]"].FirstOrDefault();
            var id = Request.Form["filters[id]"].FirstOrDefault();
            //var tenhh = Request.Form["filters[tenhh]"].FirstOrDefault();
            int skip = start != null ? Convert.ToInt32(start) : 0;
            var customerData = _context.ShiftModel.Where(d => d.deleted_at == null);
            int recordsTotal = customerData.Count();
            if (code != null && code != "")
            {
                customerData = customerData.Where(d => d.code.Contains(code));
            }
            if (name != null && name != "")
            {
                customerData = customerData.Where(d => d.name.Contains(name));
            }
            if (id != null)
            {
                customerData = customerData.Where(d => d.id == id);
            }

            int recordsFiltered = customerData.Count();
            var datapost = customerData.OrderByDescending(d => d.created_at).Skip(skip).Take(pageSize).ToList();
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
            return Json(jsonData, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }

        public JsonResult Get(string id)
        {
            var data = _context.ShiftModel.Where(d => d.id == id).FirstOrDefault();
            return Json(data, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> GetUser(string id)
        {
            var data = _context.ShiftUserModel.Where(d => d.shift_id == id).Select(d => d.person_id).ToList();
            return Json(data, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> holidays(int year, string id)
        {
            var data = _context.ShiftHolidayModel.Where(d => d.shift_id == id && d.date.Value.Year == year).ToList();
            return Json(data, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
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
