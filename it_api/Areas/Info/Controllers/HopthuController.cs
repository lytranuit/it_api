


using Holdtime.Models;
using Info.Models;
using it_template.Areas.Trend.Controllers;
using iText.Commons.Actions.Contexts;
using iText.Commons.Bouncycastle.Asn1.X509;
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
using static iText.StyledXmlParser.Jsoup.Select.Evaluator;

namespace it_template.Areas.Info.Controllers
{

    public class HopthuController : BaseController
    {
        private readonly IConfiguration _configuration;
        private UserManager<UserModel> UserManager;
        private readonly ViewRender _view;
        public HopthuController(NhansuContext context, AesOperation aes, IConfiguration configuration, UserManager<UserModel> UserMgr, ViewRender view) : base(context, aes)
        {
            _configuration = configuration;
            UserManager = UserMgr;
            _view = view;
        }

        [HttpPost]
        public async Task<JsonResult> Save(HopthuModel HopthuModel)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            HopthuModel? HopthuModel_old;
            if (HopthuModel.id == null)
            {
                HopthuModel.id = Guid.NewGuid().ToString();
                HopthuModel.created_at = DateTime.Now;
                HopthuModel.tinhtrang = TinhtrangHopthu.New;
                if (HopthuModel.is_andanh != true)
                {
                    HopthuModel.created_by = user_id;
                }


                _context.HopthuModel.Add(HopthuModel);

                _context.SaveChanges();

                HopthuModel_old = HopthuModel;
            }
            else
            {

                HopthuModel_old = _context.HopthuModel.Where(d => d.id == HopthuModel.id).FirstOrDefault();
                CopyValues<HopthuModel>(HopthuModel_old, HopthuModel);

                _context.Update(HopthuModel_old);
                _context.SaveChanges();
            }

            var files = Request.Form.Files;
            var list_dinhkem = new List<string>();
            if (files != null && files.Count > 0)
            {
                foreach (var file in files)
                {
                    var timeStamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds();
                    string name = file.FileName;
                    string type = file.Name;
                    string ext = Path.GetExtension(name);
                    string mimeType = file.ContentType;
                    //var fileName = Path.GetFileName(name);

                    var newName = timeStamp + "-" + name;
                    //var muahang_id = MuahangModel_old.id;
                    newName = newName.Replace("+", "_");
                    newName = newName.Replace("%", "_");
                    var dir = _configuration["Source:Path_Private"] + "\\hopthu\\" + HopthuModel_old.id;
                    bool exists = Directory.Exists(dir);

                    if (!exists)
                        Directory.CreateDirectory(dir);


                    var filePath = dir + "\\" + newName;

                    string url = "/private/hopthu/" + HopthuModel_old.id + "/" + newName;

                    using (var fileSrteam = new FileStream(filePath, FileMode.Create))
                    {
                        await file.CopyToAsync(fileSrteam);
                        list_dinhkem.Add(url);
                    }

                }


                ///Lưu lại
                HopthuModel_old.list_dinhkem = list_dinhkem;
                _context.Update(HopthuModel_old);
                _context.SaveChanges();
            }




            /////Gửi thông tin qua mail
            var list_user = new List<string>();
            list_user.Add("90c8b28c-9d17-485a-b64e-5e75d57db638");
            list_user.Add("5e689dde-6197-4cff-a707-4556dbf31280");
            list_user.Add("342edd08-7b6d-441e-8f0f-be19484f0275");
            await thongbao(list_user, HopthuModel_old.id);

            return Json(new { success = true, data = HopthuModel_old });
        }
        [HttpPost]
        public async Task<JsonResult> thongbao(List<string> list_user, string id)
        {
            if (list_user.Count() == 0)
            {
                return Json(new { success = false, message = "Không có list user" });
            }
            string Domain = (HttpContext.Request.IsHttps ? "https://" : "http://") + HttpContext.Request.Host.Value;

            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            var HopthuModel_old = _context.HopthuModel.Where(d => d.id == id).FirstOrDefault();
            var email_type = "hopthu_add";
            var subject = "[Mới] Hộp thư góp ý";
            var body = _view.Render("Emails/Hopthu_add", new
            {
                link_logo = Domain + "/images/clientlogo_astahealthcare.com_f1800.png",
                link = _configuration["Application:Info:link"] + "hopthu",
                data = HopthuModel_old,
            });
            if (HopthuModel_old.tinhtrang == TinhtrangHopthu.Success)
            {
                email_type = "hopthu_success";
                subject = "[Đã xử lý] Hộp thư góp ý";
            }
            else if (HopthuModel_old.tinhtrang == TinhtrangHopthu.Pedding)
            {
                email_type = "hopthu_pedding";
                subject = "[Đang xử lý] Hộp thư góp ý";
            }

            /////Gửi thông tin qua mail

            //var mail_string = "tran.dl@astahealthcare.com";
            var user_related = _context.UserModel.Where(d => list_user.Contains(d.Id)).Select(d => d.Email).ToList();
            var mail_string = string.Join(",", user_related);
            var email = new EmailModel
            {
                email_to = mail_string,
                subject = subject,
                body = body,
                email_type = email_type,
                status = 1,
                data_attachments = HopthuModel_old.list_dinhkem
            };
            _context.Add(email);
            _context.SaveChanges();
            return Json(new { success = true, data = HopthuModel_old });
        }

        [HttpPost]
        public async Task<JsonResult> Table()
        {
            var draw = Request.Form["draw"].FirstOrDefault();
            var start = Request.Form["start"].FirstOrDefault();
            var length = Request.Form["length"].FirstOrDefault();
            int pageSize = length != null ? Convert.ToInt32(length) : 0;
            var id = Request.Form["filters[id]"].FirstOrDefault();
            var loai_string = Request.Form["filters[loai]"].FirstOrDefault();
            var tinhtrang_string = Request.Form["filters[tinhtrang]"].FirstOrDefault();
            var content = Request.Form["filters[content]"].FirstOrDefault();
            //var tenhh = Request.Form["filters[tenhh]"].FirstOrDefault();
            int skip = start != null ? Convert.ToInt32(start) : 0;
            int loai = loai_string != null ? Convert.ToInt32(loai_string) : 0;
            int tinhtrang = tinhtrang_string != null ? Convert.ToInt32(tinhtrang_string) : 0;
            var customerData = _context.HopthuModel.Where(d => d.deleted_at == null);
            int recordsTotal = customerData.Count();


            if (id != null)
            {
                customerData = customerData.Where(d => d.id == id);
            }
            if (content != null && content != "")
            {
                customerData = customerData.Where(d => d.content.Contains(content));
            }
            if (loai > 0)
            {
                customerData = customerData.Where(d => d.loai == (LoaiPhanHoi)loai);
            }
            if (tinhtrang > 0)
            {
                customerData = customerData.Where(d => d.tinhtrang == (TinhtrangHopthu)tinhtrang);
            }
            if (content != null && content != "")
            {
                customerData = customerData.Where(d => d.content.Contains(content));
            }

            int recordsFiltered = customerData.Count();
            var datapost = customerData.OrderByDescending(d => d.created_at).Skip(skip).Take(pageSize).Include(d => d.user_created_by).ToList();
            var jsonData = new { draw = draw, recordsFiltered = recordsFiltered, recordsTotal = recordsTotal, data = datapost };
            return Json(jsonData, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }

        public JsonResult Get(string id)
        {
            var data = _context.HopthuModel.Where(d => d.id == id).FirstOrDefault();
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
