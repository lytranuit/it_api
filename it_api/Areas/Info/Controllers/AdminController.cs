using it_template.Areas.Trend.Controllers;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Collections;
using System.Text.Json.Serialization;
using Vue.Data;
using Vue.Models;
using Vue.Services;

namespace it_template.Areas.Info.Controllers
{
    public class AdminController : BaseController
    {
        private UserManager<UserModel> UserManager;
        public AdminController(NhansuContext context, AesOperation aes, UserManager<UserModel> UserMgr) : base(context, aes)
        {
            UserManager = UserMgr;
        }
        public async Task<JsonResult> HomeBadge()
        {
            var data = _context.HotNewsModel.Where(d => d.id == 1).FirstOrDefault();
            var tin_moi = _context.NewsModel.Where(d => d.deleted_at == null && d.is_publish == true).OrderByDescending(d => d.created_at).Take(7).ToList();
            var highlight = _context.NewsModel.Where(d => d.deleted_at == null && d.is_publish == true && d.is_highlight == true).OrderByDescending(d => d.created_at).Take(4).ToList();
            var cate = _context.CategoryModel.Where(d => d.deleted_at == null).ToList();
            foreach (var item in cate)
            {
                item.list_news = _context.NewsCategoryModel.Where(d => d.category_id == item.id).Include(d => d.news).Where(d => d.news.deleted_at == null).Select(d => d.news).Take(10).ToList();
            }


            //var noibat = _context.
            return Json(new { message = data.message, tin_moi, highlight, cate }, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }

    }

}
