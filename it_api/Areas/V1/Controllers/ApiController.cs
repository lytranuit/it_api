

using it_api.Services;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Spire.Xls;
using System.Collections;
using System.Data;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using Vue.Data;
using Vue.Models;
using Vue.Services;
using static it_template.Areas.V1.Controllers.ApiController;
using static it_template.Areas.V1.Controllers.UserController;

namespace it_template.Areas.V1.Controllers
{

    public class ApiController : BaseController
    {
        private UserManager<UserModel> UserManager;
        public ApiController(ItContext context, UserManager<UserModel> UserMgr) : base(context)
        {
            UserManager = UserMgr;
        }
        public async Task<JsonResult> objects()
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user_current = await UserManager.GetUserAsync(currentUser); // Get user id:
            var is_admin = await UserManager.IsInRoleAsync(user_current, "Administrator");
            var cusdata = _context.ObjectModel.Where(d => d.deleted_at == null);
            if (!is_admin)
            {
                var objects = _context.UserObjectModel.Where(d => d.user_id == user_id).Select(d => d.object_id).ToList();
                cusdata = cusdata.Where(d => objects.Contains(d.id));
            }
            var All = cusdata.ToList();
            //var jsonData = new { data = ProcessModel };
            return Json(All, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }

        public async Task<JsonResult> Points()
        {
            var All = _context.PointModel.Where(d => d.deleted_at == null).ToList();
            //var jsonData = new { data = ProcessModel };
            return Json(All, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> targets()
        {
            var All = _context.TargetModel.Where(d => d.deleted_at == null).ToList();
            //var jsonData = new { data = ProcessModel };
            return Json(All, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }

        public async Task<JsonResult> targetsOfObject(int object_id)
        {
            var All = _context.ObjectTargetModel.Where(d => d.object_id == object_id).Include(d => d.target).Select(d => d.target).ToList();
            //var jsonData = new { data = ProcessModel };
            return Json(All, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> locations()
        {
            var All = GetChild1(0);
            //var jsonData = new { data = ProcessModel };
            return Json(All, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }

        public async Task<JsonResult> locationswithpoint(int? filter_object, int? filter_target, int? filter_type_bc)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user_current = await UserManager.GetUserAsync(currentUser); // Get user id:
            var is_admin = await UserManager.IsInRoleAsync(user_current, "Administrator");
            var cusdata = _context.ObjectModel.Where(d => d.deleted_at == null);
            List<int>? objects;
            if (!is_admin)
            {
                objects = _context.UserObjectModel.Where(d => d.user_id == user_id).Select(d => d.object_id).ToList();
            }
            else
            {
                objects = _context.ObjectModel.Where(d => d.deleted_at == null).Select(d => d.id).ToList();
            }

            var All = await GetChild2(0, filter_object, filter_target, filter_type_bc, objects);
            //var jsonData = new { data = ProcessModel };
            return Json(All, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> departments()
        {
            var All = GetChild(0);
            //var jsonData = new { data = ProcessModel };
            return Json(All, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        private List<SelectDepartmentResponse> GetChild(int parent)
        {
            var DepartmentModel = _context.DepartmentModel.Where(d => d.deleted_at == null && d.parent == parent).OrderBy(d => d.stt).ToList();
            var list = new List<SelectDepartmentResponse>();
            if (DepartmentModel.Count() > 0)
            {
                foreach (var department in DepartmentModel)
                {
                    //if (users.Count == 0)
                    //    continue;
                    var DepartmentResponse = new SelectDepartmentResponse
                    {

                        id = department.id.ToString(),
                        label = department.name,
                        name = department.name,
                        is_department = true,
                    };
                    //var count_child = _context.DepartmentModel.Where(d => d.deleted_at == null && d.parent == department.id).Count();
                    //if (count_child > 0)
                    //{
                    var child = GetChild(department.id);
                    var users = _context.UserDepartmentModel.Where(d => d.department_id == department.id).Include(d => d.user).ToList();
                    if (users.Count() == 0 && child.Count() == 0)
                        continue;
                    foreach (var item in users)
                    {
                        var user = item.user;
                        child.Add(new SelectDepartmentResponse
                        {

                            id = user.Id.ToString(),
                            label = user.FullName + "<" + user.Email + ">",
                            name = user.FullName,
                        });
                    }
                    if (child.Count() > 0)
                        DepartmentResponse.children = child;
                    //}
                    list.Add(DepartmentResponse);



                }
            }
            return list;
        }
        private List<LocationModel> GetChild1(int parent)
        {
            var Models = _context.LocationModel.Where(d => d.deleted_at == null && d.parent == parent).OrderBy(d => d.stt).ToList();
            var list = new List<LocationModel>();
            if (Models.Count() > 0)
            {
                foreach (var model in Models)
                {
                    var count_child = _context.LocationModel.Where(d => d.deleted_at == null && d.parent == model.id).Count();
                    if (count_child > 0)
                    {
                        var child = GetChild1(model.id);
                        model.children = child;
                    }
                    list.Add(model);
                }
            }
            return list;
        }
        private async Task<List<SelectLocationResponse>> GetChild2(int parent, int? filter_object, int? filter_target, int? filter_type_bc, List<int>? objects)
        {

            var Models = _context.LocationModel.Where(d => d.deleted_at == null && d.parent == parent).OrderBy(d => d.stt).ToList();
            var list = new List<SelectLocationResponse>();
            if (Models.Count() > 0)
            {
                foreach (var model in Models)
                {
                    //if (users.Count == 0)
                    //    continue;
                    var LocationResponse = new SelectLocationResponse
                    {

                        id = model.id.ToString(),
                        label = model.code != null ? model.code + " " + model.name : model.name,
                        name = model.name,
                        is_location = true,
                    };
                    //var count_child = _context.DepartmentModel.Where(d => d.deleted_at == null && d.parent == department.id).Count();
                    //if (count_child > 0)
                    //{
                    var child = await GetChild2(model.id, filter_object, filter_target, filter_type_bc, objects);
                    var points_context = _context.PointModel.Where(d => d.location_id == model.id && objects.Contains(d.object_id.Value));

                    if (filter_object != null)
                    {
                        points_context = points_context.Where(d => d.object_id == filter_object);
                    }
                    if (filter_target != null)
                    {
                        points_context = points_context.Where(d => d.target_id == filter_target);
                    }
                    if (filter_type_bc != null)
                    {
                        var list_fre = new List<int> { };
                        if (filter_type_bc == 1)
                        {

                            list_fre = new List<int> { 1 };
                        }
                        else if (filter_type_bc == 2)
                        {

                            list_fre = new List<int> { 2 };
                        }
                        else if (filter_type_bc == 3)
                        {

                            list_fre = new List<int> { 1, 2, 3 };
                        }
                        else if (filter_type_bc == 4)
                        {

                            list_fre = new List<int> { 4, 5 };
                        }
                        points_context = points_context.Where(d => list_fre.Contains(d.frequency_id.Value));
                    }
                    var points = points_context.ToList();
                    if (points.Count() == 0 && child.Count() == 0)
                        continue;
                    foreach (var item in points)
                    {
                        child.Add(new SelectLocationResponse
                        {

                            id = item.id.ToString(),
                            label = item.code != null ? item.code + " " + item.name : item.name,
                            name = item.name,
                        });
                    }
                    if (child.Count() > 0)
                        LocationResponse.children = child;
                    //}
                    list.Add(LocationResponse);

                }
            }
            return list;
        }

    }


}
