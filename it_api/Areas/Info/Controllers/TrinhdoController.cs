﻿



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
    public class TrinhdoController : BaseController
    {
        private readonly IConfiguration _configuration;
        private UserManager<UserModel> UserManager;
        public TrinhdoController(NhansuContext context, AesOperation aes, IConfiguration configuration, UserManager<UserModel> UserMgr) : base(context, aes)
        {
            _configuration = configuration;
            UserManager = UserMgr;
        }

        [HttpPost]
        public async Task<JsonResult> Delete(string id)
        {
            var Model = _context.TrinhdoModel.Where(d => d.id == id).FirstOrDefault();
            _context.Remove(Model);
            _context.SaveChanges();
            return Json(new { success = true, data = Model });
        }
        [HttpPost]
        public async Task<JsonResult> Save(TrinhdoModel TrinhdoModel)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            TrinhdoModel? TrinhdoModel_old;
            if (TrinhdoModel.id == null)
            {
                TrinhdoModel.id = Guid.NewGuid().ToString();
                //TrinhdoModel.created_at = DateTime.Now;
                //TrinhdoModel.created_by = user_id;


                _context.TrinhdoModel.Add(TrinhdoModel);

                _context.SaveChanges();

                TrinhdoModel_old = TrinhdoModel;

            }
            else
            {

                TrinhdoModel_old = _context.TrinhdoModel.Where(d => d.id == TrinhdoModel.id).FirstOrDefault();
                CopyValues<TrinhdoModel>(TrinhdoModel_old, TrinhdoModel);
                //TrinhdoModel_old.updated_at = DateTime.Now;

                _context.Update(TrinhdoModel_old);
                _context.SaveChanges();
            }



            return Json(new { success = true, data = TrinhdoModel_old });
        }

        [HttpPost]
        public async Task<JsonResult> Table()
        {
            var draw = Request.Form["draw"].FirstOrDefault();
            var start = Request.Form["start"].FirstOrDefault();
            var length = Request.Form["length"].FirstOrDefault();
            int pageSize = length != null ? Convert.ToInt32(length) : 0;
            var MATRINHDO = Request.Form["filters[MATRINHDO]"].FirstOrDefault();
            var TENTRINHDO = Request.Form["filters[TENTRINHDO]"].FirstOrDefault();
            var id = Request.Form["filters[id]"].FirstOrDefault();
            //var tenhh = Request.Form["filters[tenhh]"].FirstOrDefault();
            int skip = start != null ? Convert.ToInt32(start) : 0;
            var customerData = _context.TrinhdoModel.Where(d => 1 == 1);
            int recordsTotal = customerData.Count();
            if (MATRINHDO != null && MATRINHDO != "")
            {
                customerData = customerData.Where(d => d.MATRINHDO.Contains(MATRINHDO));
            }
            if (TENTRINHDO != null && TENTRINHDO != "")
            {
                customerData = customerData.Where(d => d.TENTRINHDO.Contains(TENTRINHDO));
            }
            if (id != null)
            {
                customerData = customerData.Where(d => d.id == id);
            }

            int recordsFiltered = customerData.Count();
            var datapost = customerData.OrderBy(d => d.MATRINHDO).Skip(skip).Take(pageSize).ToList();
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
            var data = _context.TrinhdoModel.Where(d => d.id == id).FirstOrDefault();
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
