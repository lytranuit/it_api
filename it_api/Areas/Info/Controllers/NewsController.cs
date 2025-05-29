


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
using static iText.StyledXmlParser.Jsoup.Select.Evaluator;

namespace it_template.Areas.Info.Controllers
{

    [Authorize(Roles = "Administrator,HR,Manager Website")]
    public class NewsController : BaseController
    {
        private readonly IConfiguration _configuration;
        private UserManager<UserModel> UserManager;
        public NewsController(NhansuContext context, AesOperation aes, IConfiguration configuration, UserManager<UserModel> UserMgr) : base(context, aes)
        {
            _configuration = configuration;
            UserManager = UserMgr;
        }

        [HttpPost]
        public async Task<JsonResult> Delete(string id)
        {
            var Model = _context.NewsModel.Where(d => d.id == id).FirstOrDefault();
            Model.deleted_at = DateTime.Now;
            _context.Update(Model);
            _context.SaveChanges();
            return Json(new { success = true, data = Model });
        }
        [HttpPost]
        public async Task<JsonResult> saveTinnong(string message)
        {
            var Model = _context.HotNewsModel.Where(d => d.id == 1).FirstOrDefault();
            Model.message = message;
            _context.Update(Model);
            _context.SaveChanges();
            return Json(new { success = true, data = Model });
        }
        [HttpPost]
        public async Task<JsonResult> Save(NewsModel NewsModel, List<string> list_category)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            NewsModel? NewsModel_old;
            if (NewsModel.id == null)
            {
                NewsModel.id = Guid.NewGuid().ToString();
                NewsModel.created_at = DateTime.Now;
                NewsModel.created_by = user_id;


                _context.NewsModel.Add(NewsModel);

                _context.SaveChanges();

                NewsModel_old = NewsModel;

            }
            else
            {

                NewsModel_old = _context.NewsModel.Where(d => d.id == NewsModel.id).FirstOrDefault();
                CopyValues<NewsModel>(NewsModel_old, NewsModel);
                NewsModel_old.updated_at = DateTime.Now;

                _context.Update(NewsModel_old);
                _context.SaveChanges();
            }

            var list_old = _context.NewsCategoryModel.Where(d => d.news_id == NewsModel_old.id).ToList();
            _context.RemoveRange(list_old);
            _context.SaveChanges();

            foreach (var item in list_category)
            {
                _context.Add(new NewsCategoryModel()
                {
                    category_id = item,
                    news_id = NewsModel_old.id
                });
            }
            _context.SaveChanges();

            return Json(new { success = true, data = NewsModel_old });
        }

        [HttpPost]
        public async Task<JsonResult> Table()
        {
            var draw = Request.Form["draw"].FirstOrDefault();
            var start = Request.Form["start"].FirstOrDefault();
            var length = Request.Form["length"].FirstOrDefault();
            int pageSize = length != null ? Convert.ToInt32(length) : 0;
            var title = Request.Form["filters[title]"].FirstOrDefault();
            var id = Request.Form["filters[id]"].FirstOrDefault();
            //var tenhh = Request.Form["filters[tenhh]"].FirstOrDefault();
            int skip = start != null ? Convert.ToInt32(start) : 0;
            var customerData = _context.NewsModel.Where(d => d.deleted_at == null);
            int recordsTotal = customerData.Count();
            if (title != null && title != "")
            {
                customerData = customerData.Where(d => d.title.Contains(title));
            }

            if (id != null)
            {
                customerData = customerData.Where(d => d.id == id);
            }

            int recordsFiltered = customerData.Count();
            var datapost = customerData.OrderByDescending(d => d.created_at).Skip(skip).Take(pageSize).Include(d => d.user_created_by).ToList();
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
            var data = _context.NewsModel.Where(d => d.id == id).FirstOrDefault();
            return Json(data);
        }
        public JsonResult GetCategory(string id)
        {
            var data = _context.NewsCategoryModel.Where(d => d.news_id == id).Select(d => d.category_id).ToList();
            return Json(data);
        }
        public JsonResult GetHotNews()
        {
            var data = _context.HotNewsModel.Where(d => d.id == 1).FirstOrDefault();
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
