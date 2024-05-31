

using elFinder.NetCore.Models;
using Holdtime.Models;
using iText.Commons.Actions.Contexts;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Spire.Xls;
using Spire.Xls.Core;
using System.Collections;
using System.Data;
using System.Drawing;
using System.Runtime.Intrinsics.X86;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using Vue.Data;
using Vue.Models;

namespace it_template.Areas.Holdtime.Controllers
{

    public class StageController : BaseController
    {
        private readonly IConfiguration _configuration;
        private UserManager<UserModel> UserManager;
        public StageController(HoldTimeContext context, IConfiguration configuration, UserManager<UserModel> UserMgr) : base(context)
        {
            _configuration = configuration;
            UserManager = UserMgr;
        }

        [HttpPost]
        public async Task<JsonResult> Delete(int id)
        {
            var Model = _context.StageModel.Where(d => d.id == id).FirstOrDefault();
            Model.deleted_at = DateTime.Now;
            _context.Update(Model);
            _context.SaveChanges();
            return Json(new { success = true, data = Model });
        }
        [HttpPost]
        public async Task<JsonResult> Save(StageModel StageModel, List<StageTimeModel> list_add, List<StageTimeModel>? list_update, List<StageTimeModel>? list_delete, List<int> list_target)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            StageModel? StageModel_old;
            if (StageModel.id == 0)
            {
                StageModel.created_at = DateTime.Now;


                _context.StageModel.Add(StageModel);

                _context.SaveChanges();

                StageModel_old = StageModel;

            }
            else
            {

                StageModel_old = _context.StageModel.Where(d => d.id == StageModel.id).FirstOrDefault();
                CopyValues<StageModel>(StageModel_old, StageModel);
                StageModel_old.updated_at = DateTime.Now;

                _context.Update(StageModel_old);
                _context.SaveChanges();
            }
            ////
            ///targets
            /// 
            var StageTargetModel_old = _context.StageTargetModel.Where(d => d.stage_id == StageModel_old.id).ToList();
            _context.RemoveRange(StageTargetModel_old);
            _context.SaveChanges();

            foreach (var item in list_target)
            {
                _context.Add(new StageTargetModel()
                {
                    stage_id = StageModel_old.id,
                    target_id = item
                });
            }
            _context.SaveChanges();
            ////
            ///Time
            ///
            if (list_delete != null)
                _context.RemoveRange(list_delete);
            if (list_add != null)
            {
                foreach (var item in list_add)
                {
                    item.stage_id = StageModel_old.id;
                    _context.Add(item);
                }
            }
            if (list_update != null)
            {
                foreach (var item in list_update)
                {
                    _context.Update(item);
                }
            }

            _context.SaveChanges();

            return Json(new { success = true, data = StageModel_old });
        }
        [HttpPost]
        public async Task<JsonResult> Table()
        {
            var draw = Request.Form["draw"].FirstOrDefault();
            var start = Request.Form["start"].FirstOrDefault();
            var length = Request.Form["length"].FirstOrDefault();
            int pageSize = length != null ? Convert.ToInt32(length) : 0;
            var name = Request.Form["filters[name]"].FirstOrDefault();
            var id_text = Request.Form["filters[id]"].FirstOrDefault();
            int id = id_text != null ? Convert.ToInt32(id_text) : 0;
            //var tenhh = Request.Form["filters[tenhh]"].FirstOrDefault();
            int skip = start != null ? Convert.ToInt32(start) : 0;
            var customerData = _context.StageModel.Where(d => d.deleted_at == null);
            int recordsTotal = customerData.Count();
            if (name != null && name != "")
            {
                customerData = customerData.Where(d => d.name.Contains(name));
            }
            if (id != 0)
            {
                customerData = customerData.Where(d => d.id == id);
            }
            int recordsFiltered = customerData.Count();
            var datapost = customerData.OrderByDescending(d => d.id).Skip(skip).Take(pageSize).ToList();
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
            var data = _context.StageModel.Where(d => d.id == id).FirstOrDefault();
            return Json(data);
        }
        public JsonResult GetListTarget(int id)
        {
            var data = _context.StageTargetModel.Where(d => d.stage_id == id).Select(d => d.target_id).ToList();
            return Json(data);
        }
        public JsonResult GetTime(int id)
        {
            var data = _context.StageTimeModel.Where(d => d.stage_id == id).ToList();
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
