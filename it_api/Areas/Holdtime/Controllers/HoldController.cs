

using elFinder.NetCore.Models;
using Holdtime.Models;
using it_template.Areas.Trend.Controllers;
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
using System.Linq;
using System.Runtime.Intrinsics.X86;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using Vue.Data;
using Vue.Models;

namespace it_template.Areas.Holdtime.Controllers
{

    public class HoldController : BaseController
    {
        private readonly IConfiguration _configuration;
        private UserManager<UserModel> UserManager;
        public HoldController(HoldTimeContext context, IConfiguration configuration, UserManager<UserModel> UserMgr) : base(context)
        {
            _configuration = configuration;
            UserManager = UserMgr;
        }

        [HttpPost]
        public async Task<JsonResult> Delete(int id)
        {
            var Model = _context.HoldModel.Where(d => d.id == id).FirstOrDefault();
            Model.deleted_at = DateTime.Now;
            _context.Update(Model);

            var list_time = _context.HoldTimeModel.Where(d => d.hold_id == id).ToList();
            foreach (var item in list_time)
            {
                item.deleted_at = DateTime.Now;
            }
            _context.UpdateRange(list_time);
            _context.SaveChanges();
            return Json(new { success = true, data = Model });
        }
        [HttpPost]
        public async Task<JsonResult> Save(HoldModel HoldModel, List<HoldTimeModel> list_add, List<HoldTimeModel>? list_update, List<HoldTimeModel>? list_delete, List<string> list_user)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            HoldModel? HoldModel_old;
            if (HoldModel.id == 0)
            {
                HoldModel.created_at = DateTime.Now;


                _context.HoldModel.Add(HoldModel);

                _context.SaveChanges();

                HoldModel_old = HoldModel;

            }
            else
            {

                HoldModel_old = _context.HoldModel.Where(d => d.id == HoldModel.id).FirstOrDefault();
                CopyValues<HoldModel>(HoldModel_old, HoldModel);
                HoldModel_old.updated_at = DateTime.Now;

                _context.Update(HoldModel_old);
                _context.SaveChanges();
            }
            ////
            ///User
            /// 
            var HoldAlertModel_old = _context.HoldAlertModel.Where(d => d.hold_id == HoldModel_old.id).ToList();
            _context.RemoveRange(HoldAlertModel_old);
            _context.SaveChanges();

            foreach (var item in list_user)
            {
                _context.Add(new HoldAlertModel()
                {
                    hold_id = HoldModel_old.id,
                    user_id = item
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
                    item.hold_id = HoldModel_old.id;
                    _context.Add(item);
                    ///
                    _context.SaveChanges();
                    if (item.list_target != null)
                    {
                        foreach (var target_id in item.list_target)
                        {
                            var time_target = new HoldTimeTargetModel()
                            {
                                hold_time_id = item.id,
                                target_id = target_id
                            };
                            _context.Add(time_target);
                        }
                    }
                }
            }
            if (list_update != null)
            {
                foreach (var item in list_update)
                {
                    _context.Update(item);
                    //////
                    var list_target = item.list_target != null ? item.list_target : new List<int>();
                    var time_target_old = _context.HoldTimeTargetModel.Where(d => d.hold_time_id == item.id).Select(d => d.target_id.Value).ToList();
                    IEnumerable<int> list_delete_target = time_target_old.Except(list_target);
                    IEnumerable<int> list_add_target = list_target.Except(time_target_old);

                    if (list_add_target != null)
                    {
                        foreach (int key in list_add_target)
                        {

                            _context.Add(new HoldTimeTargetModel()
                            {
                                hold_time_id = item.id,
                                target_id = key,
                            });
                        }
                        //_context.SaveChanges();
                    }
                    if (list_delete_target != null)
                    {
                        foreach (int key in list_delete_target)
                        {
                            HoldTimeTargetModel HoldTimeTargetModel = _context.HoldTimeTargetModel.Where(d => d.hold_time_id == item.id && d.target_id == key).First();
                            _context.Remove(HoldTimeTargetModel);
                        }
                        //_context.SaveChanges();
                    }

                }
            }

            _context.SaveChanges();

            return Json(new { success = true, data = HoldModel_old });
        }


        [HttpPost]
        public async Task<JsonResult> SaveHoldTimeTargetModel(List<HoldTimeTargetModel> data)
        {
            _context.UpdateRange(data);
            _context.SaveChanges();
            return Json(new { success = true });
        }
        [HttpPost]
        public async Task<JsonResult> Table()
        {
            var draw = Request.Form["draw"].FirstOrDefault();
            var start = Request.Form["start"].FirstOrDefault();
            var length = Request.Form["length"].FirstOrDefault();
            int pageSize = length != null ? Convert.ToInt32(length) : 0;
            var tensp = Request.Form["filters[tensp]"].FirstOrDefault();
            var masp = Request.Form["filters[masp]"].FirstOrDefault();
            var malo = Request.Form["filters[malo]"].FirstOrDefault();
            var manghiencuu = Request.Form["filters[manghiencuu]"].FirstOrDefault();
            var maphantich = Request.Form["filters[maphantich]"].FirstOrDefault();
            var madecuong = Request.Form["filters[madecuong]"].FirstOrDefault();
            var id_text = Request.Form["filters[id]"].FirstOrDefault();
            var id_stage_id = Request.Form["filters[stage_id]"].FirstOrDefault();
            int id = id_text != null ? Convert.ToInt32(id_text) : 0;
            int stage_id = id_stage_id != null ? Convert.ToInt32(id_stage_id) : 0;
            //var tenhh = Request.Form["filters[tenhh]"].FirstOrDefault();
            int skip = start != null ? Convert.ToInt32(start) : 0;
            var customerData = _context.HoldModel.Where(d => d.deleted_at == null);
            int recordsTotal = customerData.Count();
            if (tensp != null && tensp != "")
            {
                customerData = customerData.Where(d => d.tensp.Contains(tensp));
            }
            if (masp != null && masp != "")
            {
                customerData = customerData.Where(d => d.masp.Contains(masp));
            }
            if (malo != null && malo != "")
            {
                customerData = customerData.Where(d => d.malo.Contains(malo));
            }
            if (manghiencuu != null && manghiencuu != "")
            {
                customerData = customerData.Where(d => d.manghiencuu.Contains(manghiencuu));
            }
            if (maphantich != null && maphantich != "")
            {
                customerData = customerData.Where(d => d.maphantich.Contains(maphantich));
            }
            if (madecuong != null && madecuong != "")
            {
                customerData = customerData.Where(d => d.madecuong.Contains(madecuong));
            }
            if (id != 0)
            {
                customerData = customerData.Where(d => d.id == id);
            }
            if (stage_id != 0)
            {
                customerData = customerData.Where(d => d.stage_id == stage_id);
            }
            int recordsFiltered = customerData.Count();
            var datapost = customerData.OrderByDescending(d => d.id).Skip(skip).Take(pageSize).Include(d => d.stage).ToList();
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
            var data = _context.HoldModel.Where(d => d.id == id).FirstOrDefault();
            return Json(data);
        }
        public JsonResult GetTime(int id)
        {
            var data = _context.HoldTimeModel.Where(d => d.hold_id == id).Include(d => d.targets).ThenInclude(d => d.target).ToList();
            return Json(data);
        }
        public JsonResult GetListAlert(int id)
        {
            var data = _context.HoldAlertModel.Where(d => d.hold_id == id).Select(d => d.user_id).ToList();
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
