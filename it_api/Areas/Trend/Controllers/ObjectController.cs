

using elFinder.NetCore.Models;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Spire.Xls;
using System.Collections;
using System.Data;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using Vue.Data;
using Vue.Models;

namespace it_template.Areas.Trend.Controllers
{

    public class ObjectController : BaseController
    {
        private readonly IConfiguration _configuration;
        private UserManager<UserModel> UserManager;
        public ObjectController(ItContext context, IConfiguration configuration, UserManager<UserModel> UserMgr) : base(context)
        {
            _configuration = configuration;
            UserManager = UserMgr;
        }

        [HttpPost]
        public async Task<JsonResult> Delete(int id)
        {
            var Model = _context.ObjectModel.Where(d => d.id == id).FirstOrDefault();
            Model.deleted_at = DateTime.Now;
            _context.Update(Model);
            _context.SaveChanges();
            return Json(new { success = true, data = Model });
        }
        [HttpPost]
        public async Task<JsonResult> Save(ObjectModel ObjectModel, List<int> list_target)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            ObjectModel? ObjectModel_old;
            if (ObjectModel.id == 0)
            {
                ObjectModel.created_at = DateTime.Now;


                _context.ObjectModel.Add(ObjectModel);

                _context.SaveChanges();

                ObjectModel_old = ObjectModel;

            }
            else
            {

                ObjectModel_old = _context.ObjectModel.Where(d => d.id == ObjectModel.id).FirstOrDefault();
                CopyValues<ObjectModel>(ObjectModel_old, ObjectModel);
                ObjectModel_old.updated_at = DateTime.Now;

                _context.Update(ObjectModel_old);
                _context.SaveChanges();
            }
            if (list_target != null)
            {
                var ObjectTargetModel_old = _context.ObjectTargetModel.Where(d => d.object_id == ObjectModel_old.id).Select(d => d.target_id).ToList();
                IEnumerable<int> list_delete = ObjectTargetModel_old.Except(list_target);
                IEnumerable<int> list_add = list_target.Except(ObjectTargetModel_old);
                if (list_add != null)
                {
                    foreach (int key in list_add)
                    {

                        _context.Add(new ObjectTargetModel()
                        {
                            target_id = key,
                            object_id = ObjectModel_old.id,
                        });
                    }
                    _context.SaveChanges();
                }
                if (list_delete != null)
                {
                    foreach (int key in list_delete)
                    {
                        ObjectTargetModel ObjectTargetModel = _context.ObjectTargetModel.Where(d => d.object_id == ObjectModel_old.id && d.target_id == key).First();
                        _context.Remove(ObjectTargetModel);
                    }
                    _context.SaveChanges();
                }

                var stt = 0;
                foreach (var key in list_target)
                {
                    ObjectTargetModel ObjectTargetModel = _context.ObjectTargetModel.Where(d => d.object_id == ObjectModel_old.id && d.target_id == key).First();
                    ObjectTargetModel.order = stt++;
                    _context.Update(ObjectTargetModel);
                }
                _context.SaveChanges();
            }

            return Json(new { success = true, data = ObjectModel_old });
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
            var customerData = _context.ObjectModel.Where(d => d.deleted_at == null);
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
            var data = _context.ObjectModel.Where(d => d.id == id).Include(d => d.targets).ThenInclude(d => d.target).FirstOrDefault();
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
