
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Identity;
using Vue.Data;
using Vue.Models;
using System.Collections;
using workflow.Models;
using System.Text.Json.Serialization;

namespace it_template.Areas.Trend.Controllers
{
    public class LocationController : BaseController
    {
        private UserManager<UserModel> UserManager;
        private string _type = "Location";
        public LocationController(ItContext context, UserManager<UserModel> UserMgr) : base(context)
        {
            UserManager = UserMgr;
        }
        [HttpPost]
        public async Task<JsonResult> Save(LocationModel LocationModel, List<string> list_users_id)
        {
            var jsonData = new { success = true, message = "" };
            try
            {
                if (LocationModel.id > 0)
                {
                    _context.Update(LocationModel);

                    //var list_old = _context.UserLocationModel.Where(d => d.Location_id == LocationModel.id).ToList();
                    //_context.RemoveRange(list_old);

                    _context.SaveChanges();
                }
                else
                {
                    LocationModel.created_at = DateTime.Now;
                    LocationModel.parent = 0;
                    LocationModel.stt = 1000;
                    LocationModel.count_child = 0;
                    _context.Add(LocationModel);
                    _context.SaveChanges();
                }
                //foreach (var item in list_users_id)
                //{
                //    var UserLocationModel = new UserLocationModel()
                //    {
                //        user_id = item,
                //        Location_id = LocationModel.id
                //    };
                //    _context.Add(UserLocationModel);
                //}
                //_context.SaveChanges();
            }
            catch (Exception ex)
            {
                jsonData = new
                {
                    success = false,
                    message = ex.Message
                };
            }


            return Json(jsonData);
        }

        [HttpPost]
        public async Task<JsonResult> Remove(List<int> item)
        {
            var jsonData = new { success = true, message = "" };
            try
            {
                var list = _context.LocationModel.Where(d => item.Contains(d.id)).ToList();
                foreach (var dep in list)
                {
                    dep.deleted_at = DateTime.Now;
                    _context.Update(dep);
                }
                _context.SaveChanges();
            }
            catch (Exception ex)
            {
                jsonData = new { success = false, message = ex.Message };
            }


            return Json(jsonData);
        }
        [HttpPost]
        public async Task<JsonResult> saveorder(List<LocationOrder> data)
        {
            var index = 0;
            foreach (var item in data)
            {
                var LocationModel = _context.LocationModel.Find(item.id);
                LocationModel.parent = item.parent_id != null ? item.parent_id : 0;
                LocationModel.count_child = item.count_child;
                LocationModel.stt = index++;
                _context.Update(LocationModel);
            }
            _context.SaveChanges();
            return Json(new { success = 1 });
        }

        public async Task<JsonResult> Get()
        {
            var All = GetChild(0);
            //var jsonData = new { data = ProcessModel };
            return Json(All, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        private List<LocationModel> GetChild(int parent)
        {
            var LocationModel = _context.LocationModel.Where(d => d.deleted_at == null && d.parent == parent).OrderBy(d => d.stt).ToList();
            var list = new List<LocationModel>();
            if (LocationModel.Count() > 0)
            {
                foreach (var Location in LocationModel)
                {
                    var count_child = _context.LocationModel.Where(d => d.deleted_at == null && d.parent == Location.id).Count();
                    if (count_child > 0)
                    {
                        var child = GetChild(Location.id);
                        Location.children = child;

                    }
                    list.Add(Location);
                }
            }
            return list;
        }

    }
    public class LocationOrder
    {
        public int id { get; set; }
        public int count_child { get; set; }

        public int parent_id { get; set; }
    }
}
