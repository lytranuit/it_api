


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
    public class TinhController : BaseController
    {
        private readonly IConfiguration _configuration;
        private UserManager<UserModel> UserManager;
        public TinhController(NhansuContext context, AesOperation aes, IConfiguration configuration, UserManager<UserModel> UserMgr) : base(context, aes)
        {
            _configuration = configuration;
            UserManager = UserMgr;
        }

        [HttpPost]
        public async Task<JsonResult> Delete(string id)
        {
            var Model = _context.TinhModel.Where(d => d.id == id).FirstOrDefault();
            _context.Remove(Model);
            _context.SaveChanges();
            return Json(new { success = true, data = Model });
        }
        [HttpPost]
        public async Task<JsonResult> Save(TinhModel TinhModel)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            TinhModel? TinhModel_old;
            if (TinhModel.id == null)
            {
                TinhModel.id = Guid.NewGuid().ToString();
                //TinhModel.created_at = DateTime.Now;
                //TinhModel.created_by = user_id;


                _context.TinhModel.Add(TinhModel);

                _context.SaveChanges();

                TinhModel_old = TinhModel;

            }
            else
            {

                TinhModel_old = _context.TinhModel.Where(d => d.id == TinhModel.id).FirstOrDefault();
                CopyValues<TinhModel>(TinhModel_old, TinhModel);
                //TinhModel_old.updated_at = DateTime.Now;

                _context.Update(TinhModel_old);
                _context.SaveChanges();
            }



            return Json(new { success = true, data = TinhModel_old });
        }

        [HttpPost]
        public async Task<JsonResult> Table()
        {
            var draw = Request.Form["draw"].FirstOrDefault();
            var start = Request.Form["start"].FirstOrDefault();
            var length = Request.Form["length"].FirstOrDefault();
            int pageSize = length != null ? Convert.ToInt32(length) : 0;
            var MaTinh = Request.Form["filters[MaTinh]"].FirstOrDefault();
            var TenTinh = Request.Form["filters[TenTinh]"].FirstOrDefault();
            var id = Request.Form["filters[id]"].FirstOrDefault();
            //var tenhh = Request.Form["filters[tenhh]"].FirstOrDefault();
            int skip = start != null ? Convert.ToInt32(start) : 0;
            var customerData = _context.TinhModel.Where(d => 1 == 1);
            int recordsTotal = customerData.Count();
            if (MaTinh != null && MaTinh != "")
            {
                customerData = customerData.Where(d => d.MaTinh.Contains(MaTinh));
            }
            if (TenTinh != null && TenTinh != "")
            {
                customerData = customerData.Where(d => d.TenTinh.Contains(TenTinh));
            }
            if (id != null)
            {
                customerData = customerData.Where(d => d.id == id);
            }

            int recordsFiltered = customerData.Count();
            var datapost = customerData.OrderBy(d => d.MaTinh).Skip(skip).Take(pageSize).ToList();
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
            var data = _context.TinhModel.Where(d => d.id == id).FirstOrDefault();
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
