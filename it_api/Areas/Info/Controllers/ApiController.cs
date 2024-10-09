

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
using static it_template.Areas.V1.Controllers.UserController;
namespace it_template.Areas.Info.Controllers
{

    public class ApiController : BaseController
    {
        private UserManager<UserModel> UserManager;
        public ApiController(NhansuContext context, AesOperation aes, UserManager<UserModel> UserMgr) : base(context, aes)
        {
            UserManager = UserMgr;
        }

        public async Task<JsonResult> category()
        {
            var All = _context.CategoryModel.Where(d => d.deleted_at == null).ToList();
            //var jsonData = new { data = ProcessModel };
            return Json(All, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }

        public async Task<JsonResult> Trinhdo()
        {
            var All = _context.TrinhdoModel.ToList();
            //var jsonData = new { data = ProcessModel };
            return Json(All, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> Chuyenmon()
        {
            var All = _context.ChuyenmonModel.ToList();
            //var jsonData = new { data = ProcessModel };
            return Json(All, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> Area()
        {
            var All = _context.KhoiModel.ToList();
            //var jsonData = new { data = ProcessModel };
            return Json(All, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> Department()
        {
            var All = _context.DepartmentModel.ToList();
            //var jsonData = new { data = ProcessModel };
            return Json(All, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> Loaihd()
        {
            var All = _context.LoaiHDModel.ToList();
            //var jsonData = new { data = ProcessModel };
            return Json(All, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> Chucvu()
        {
            var All = _context.PositionModel.ToList();
            //var jsonData = new { data = ProcessModel };
            return Json(All, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }

        public async Task<JsonResult> Shifts()
        {
            var All = _context.ShiftModel.Where(d => d.deleted_at == null).ToList();
            //var jsonData = new { data = ProcessModel };
            return Json(All, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> PersonDepartments(DateTime? date_from)
        {
            var All = _context.DepartmentModel.Select(d => new SelectDepartmentResponse()
            {
                id = d.id,
                label = d.TENPHONG,
                name = d.MAPHONG,
                is_department = true,
            }).ToList();
            var res = new List<SelectDepartmentResponse>(){
                new SelectDepartmentResponse
                {

                    id = "Asta",
                    label = "Asta",
                    name = "Asta",
                    is_department = true,
                    children = new List<SelectDepartmentResponse>()
                }
            };
            foreach (var item in All)
            {
                var children = _context.PersonnelModel.Where(d => d.MAPHONG == item.name && (d.NGAYNGHIVIEC == null || d.NGAYNGHIVIEC >= date_from)).Select(d => new SelectDepartmentResponse()
                {
                    is_department = false,
                    id = d.id,
                    label = d.HOVATEN,
                    name = d.HOVATEN
                }).ToList();
                if (children.Count() == 0)
                    continue;
                item.children = children;
                res[0].children.Add(item);
            }
            //var jsonData = new { data = ProcessModel };
            return Json(res, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
    }


}
