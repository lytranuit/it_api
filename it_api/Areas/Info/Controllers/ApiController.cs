

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

        public async Task<JsonResult> Phonghop()
        {
            var All = _context.PhonghopModel.ToList();
            //var jsonData = new { data = ProcessModel };
            return Json(All, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
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
        public async Task<JsonResult> Tinh()
        {
            var All = _context.TinhModel.ToList();
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
        public async Task<JsonResult> InfoUser(string user_id)
        {
            var user = _context.UserModel.SingleOrDefault(d => d.Id == user_id);

            var info = new Info()
            {
                UserId = user_id,
                quanlytructiep = new Info() { },
                truongbophan = new Info() { },
                quanlycong = new Info() { },
                BGD = new Info() { }
            };

            var person = _context.PersonnelModel.SingleOrDefault(d => d.EMAIL.ToLower() == user.Email.ToLower());
            if (person == null)
            {
                return Json(info, new System.Text.Json.JsonSerializerOptions()
                {
                    DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
                });
            }
            info.Id = person.id;
            info.Name = person.HOVATEN;
            info.Email = person.EMAIL;
            info.MANV = person.MANV;

            var quanlytructiep_id = person.MAQUANLYTRUCTIEP;
            var quanlytructiep = _context.PersonnelModel.SingleOrDefault(d => d.id == quanlytructiep_id);
            if (quanlytructiep != null)
            {
                info.quanlytructiep.Id = quanlytructiep.id;
                info.quanlytructiep.Name = quanlytructiep.HOVATEN;
                info.quanlytructiep.Email = quanlytructiep.EMAIL;
                info.quanlytructiep.MANV = quanlytructiep.MANV;

                var quanlytructiep_user = _context.UserModel.SingleOrDefault(d => d.Email.ToLower() == quanlytructiep.EMAIL.ToLower());
                if (quanlytructiep_user != null)
                {
                    info.quanlytructiep.UserId = quanlytructiep_user.Id;

                }
            }

            var bophan = _context.DepartmentModel.SingleOrDefault(d => d.MAPHONG == person.MAPHONG);
            if (bophan != null)
            {
                var truongbophan_id = bophan.truongbophan_id;
                var truongbophan = _context.PersonnelModel.SingleOrDefault(d => d.id == truongbophan_id);
                if (truongbophan != null)
                {
                    info.truongbophan.Id = truongbophan.id;
                    info.truongbophan.Name = truongbophan.HOVATEN;
                    info.truongbophan.Email = truongbophan.EMAIL;
                    info.truongbophan.MANV = truongbophan.MANV;
                    var truongbophan_user = _context.UserModel.SingleOrDefault(d => d.Email.ToLower() == truongbophan.EMAIL.ToLower());
                    if (truongbophan_user != null)
                    {
                        info.truongbophan.UserId = truongbophan_user.Id;
                    }
                }

                var quanlycong_id = bophan.quanlycong_id;
                var quanlycong = _context.PersonnelModel.SingleOrDefault(d => d.id == quanlycong_id);
                if (quanlycong != null)
                {
                    info.quanlycong.Id = quanlycong.id;
                    info.quanlycong.Name = quanlycong.HOVATEN;
                    info.quanlycong.Email = quanlycong.EMAIL;
                    info.quanlycong.MANV = quanlycong.MANV;
                    var quanlycong_user = _context.UserModel.SingleOrDefault(d => d.Email.ToLower() == quanlycong.EMAIL.ToLower());
                    if (quanlycong_user != null)
                    {
                        info.quanlycong.UserId = quanlycong_user.Id;
                    }
                }

            }
            //Mặc định
            var email_BGD = "yen.tp@astahealthcare.com";
            ///Là trưởng bộ phân thì BDG là người quản lý trực tiếp
            if (person.id == info.truongbophan.Id)
            {
                email_BGD = quanlytructiep.EMAIL;
            }
            else if (person.MAPHONG == "22.01" || person.MAPHONG == "22") /// PTTT
            {
                email_BGD = "thai.mq@astahealthcare.com";
            }
            else if (person.MAPHONG == "19") /// KINH DOANH
            {
                email_BGD = "khanh.cn@astahealthcare.com";
            }



            var BGD = _context.PersonnelModel.SingleOrDefault(d => d.EMAIL.ToLower() == email_BGD.ToLower());
            var user_BGD = _context.UserModel.SingleOrDefault(d => d.Email.ToLower() == email_BGD.ToLower());
            info.BGD.Id = BGD.id;
            info.BGD.Name = BGD.HOVATEN;
            info.BGD.Email = BGD.EMAIL;
            info.BGD.MANV = BGD.MANV;
            info.BGD.UserId = user_BGD.Id;

            //var jsonData = new { data = ProcessModel };
            return Json(info, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }

    }

    public class Info
    {
        public string Id { get; set; }
        public string Name { get; set; }

        public string Email { get; set; }

        public string UserId { get; set; }

        public string MANV { get; set; }

        public Info quanlytructiep { get; set; }

        public Info truongbophan { get; set; }

        public Info quanlycong { get; set; }

        public Info BGD { get; set; }

    }
}