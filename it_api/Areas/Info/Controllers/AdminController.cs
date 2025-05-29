using it_template.Areas.Trend.Controllers;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using PH.WorkingDaysAndTimeUtility.Configuration;
using PH.WorkingDaysAndTimeUtility;
using System.Collections;
using System.Text.Json.Serialization;
using Vue.Data;
using Vue.Models;
using Vue.Services;
using Point85.ShiftSharp.Schedule;
using System.Dynamic;
using Info.Models;
using static iText.StyledXmlParser.Jsoup.Select.Evaluator;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Spire.Xls;

namespace it_template.Areas.Info.Controllers
{
    public class AdminController : BaseController
    {
        private UserManager<UserModel> UserManager;
        private readonly TinhCong _tinhcong;
        public AdminController(NhansuContext context, AesOperation aes, TinhCong tinhcong, UserManager<UserModel> UserMgr) : base(context, aes)
        {
            UserManager = UserMgr;
            _tinhcong = tinhcong;
        }
        public async Task<JsonResult> HomeBadge()
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user = await UserManager.GetUserAsync(currentUser);
            var user_id = user.Id;
            var email = user.Email;
            var date_now = DateTime.Now.Date;

            var is_admin = await UserManager.IsInRoleAsync(user, "Administrator");
            var is_manager = await UserManager.IsInRoleAsync(user, "Manager HR");
            var is_hr = await UserManager.IsInRoleAsync(user, "HR");
            var customerData = _context.PersonnelModel.Where(d => d.NGAYNGHIVIEC == null);
            if (is_manager)
            {
                var person = _context.PersonnelModel.Where(d => d.EMAIL == email).FirstOrDefault();
                if (person != null)
                {
                    var maphong = person.MAPHONG;
                    if (email == "thao.pdp@astahealthcare.com")
                    {
                        customerData = customerData.Where(d => d.MAPHONG == maphong || d.MAPHONG == "22");
                    }
                    else
                    {
                        customerData = customerData.Where(d => d.MAPHONG == maphong);
                    }
                }
                else
                {
                    customerData = customerData.Where(d => 0 == 1);
                }
            }
            else if (!is_admin && !is_hr)
            {
                customerData = customerData.Where(d => d.EMAIL == email);
            }
            var tong_nv = customerData.Count();
            var list_nv = customerData.Select(d => d.MANV).ToList();
            var list_machamcong = customerData.Select(d => d.MACC).ToList();

            var nghiphep = _context.ChamcongModel.Where(d => list_nv.Contains(d.MANV) && d.date == date_now && d.value_new != "X" && d.value_new != "").Select(d => d.MANV).Distinct().Count();
            var danglamviec = tong_nv - nghiphep;
            var dachamcong = _context.HikModel.Where(d => list_machamcong.Contains(d.id) && d.date == date_now).Select(d => d.id).Distinct().Count();
            var chuachamcong = tong_nv - dachamcong;

            /////Công
            ///
            decimal tong = 0;
            decimal phep = 0;
            decimal conglam = 0;
            decimal phepconlai = 0;
            var record = _context.PersonnelModel.Where(d => d.EMAIL == email).ToList();
            if (record.Count() > 0)
            {

                var now = DateTime.Now;

                // Nếu ngày >= 26, chuyển sang tháng sau
                if (now.Day >= 26)
                {
                    now = now.AddMonths(1);
                }

                // Tính toán tháng và năm trước đó
                int previousMonth = now.Month - 1;
                int previousYear = now.Year;

                // Nếu là tháng 1 và lùi về tháng 0, giảm năm và đặt tháng thành 12
                if (previousMonth == 0)
                {
                    previousMonth = 12;
                    previousYear -= 1;
                }

                // Tạo đối tượng DateTime
                var date_from = new DateTime(previousYear, previousMonth, 26);
                var date_to = new DateTime(now.Year, now.Month, 25);


                var list_data_cong = _tinhcong.cong(record, date_from, date_to);
                IDictionary<string, dynamic> data_cong = list_data_cong.Find(x => x["EMAIL"] == email);
                conglam = (decimal)data_cong["tong"];
                phep = (decimal)data_cong["tongphep"];
                tong = (decimal)data_cong["tongcong"];
                phepconlai = (decimal)data_cong["phepnamconlai"];
            }


            var tongthunhap = record.FirstOrDefault().tong_thunhap;
            //var noibat = _context.
            return Json(new { phep, tong, conglam, tongthunhap, phepconlai, tong_nv, nghiphep, chuachamcong, danglamviec }, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> getDashboard()
        {
            //var biendong = new { labels = data.Select(d => d.label).ToList(), datasets = new List<Chart>() { new Chart { label = "Doanh thu", data = data.Select(d => d.data).ToList() } } };
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            var person = _context.PersonnelModel.Where(d => d.EMAIL.ToLower() == user.Email.ToLower()).FirstOrDefault(); ///Check is_truongbophan


            var tong = _context.PersonnelModel.Where(d => d.NGAYNGHIVIEC == null).Count();
            var chinhthuc = _context.PersonnelModel.Where(d => d.NGAYNGHIVIEC == null && d.tinhtrang == "Chính thức" && d.LOAIHD != "DV").Count();
            var hocthuviec = _context.PersonnelModel.Where(d => d.NGAYNGHIVIEC == null && (d.tinhtrang == "Học việc" || d.tinhtrang == "Thử việc" || d.tinhtrang == "Thử việc không bảo hiểm và không phép năm") && d.LOAIHD != "DV").Count();
            var dichvu = _context.PersonnelModel.Where(d => d.NGAYNGHIVIEC == null && (d.LOAIHD == "DV" || d.tinhtrang == "Dịch vụ" || d.tinhtrang == "Dịch vụ không phép năm")).Count();

            ///Check is_truongbophan
            var phong = _context.DepartmentModel.Where(d => d.truongbophan_id == person.id).FirstOrDefault();
            if (phong != null)
            {
                tong = _context.PersonnelModel.Where(d => d.NGAYNGHIVIEC == null && d.MAPHONG == person.MAPHONG).Count();
                chinhthuc = _context.PersonnelModel.Where(d => d.NGAYNGHIVIEC == null && d.tinhtrang == "Chính thức" && d.LOAIHD != "DV" && d.MAPHONG == person.MAPHONG).Count();
                hocthuviec = _context.PersonnelModel.Where(d => d.NGAYNGHIVIEC == null && (d.tinhtrang == "Học việc" || d.tinhtrang == "Thử việc" || d.tinhtrang == "Thử việc không bảo hiểm và không phép năm") && d.LOAIHD != "DV" && d.MAPHONG == person.MAPHONG).Count();
                dichvu = _context.PersonnelModel.Where(d => d.NGAYNGHIVIEC == null && (d.LOAIHD == "DV" || d.tinhtrang == "Dịch vụ" || d.tinhtrang == "Dịch vụ không phép năm") && d.MAPHONG == person.MAPHONG).Count();

            }
            var json = new { tong, chinhthuc, hocthuviec, dichvu };
            return Json(json);
        }

        [HttpPost]
        public async Task<JsonResult> biendong(int nam, string department)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user_id = UserManager.GetUserId(currentUser);
            var user = await UserManager.GetUserAsync(currentUser);
            var person = _context.PersonnelModel.Where(d => d.EMAIL.ToLower() == user.Email.ToLower()).FirstOrDefault(); ///Check is_truongbophan

            var is_admin = await UserManager.IsInRoleAsync(user, "Administrator");
            var is_manager = await UserManager.IsInRoleAsync(user, "Manager HR");
            var is_hr = await UserManager.IsInRoleAsync(user, "HR");

            //var biendong = new { labels = data.Select(d => d.label).ToList(), datasets = new List<Chart>() { new Chart { label = "Doanh thu", data = data.Select(d => d.data).ToList() } } };
            var salaryUser = _context.SalaryUserModel.Include(d => d.salary).Where(d => d.salary.deleted_at == null && d.salary.status == "Đã khóa" && d.salary.nam == nam).ToList();
            ///Check is_truongbophan
            var phong = _context.DepartmentModel.Where(d => d.truongbophan_id == person.id).FirstOrDefault();
            var is_truongbophan = false;
            if (phong != null)
            {
                is_truongbophan = true;
                salaryUser = salaryUser.Where(d => d.MABOPHAN == person.MAPHONG).ToList();
            }
            var data = salaryUser.GroupBy(d => new { year = d.salary.nam, month = d.salary.thang }).Select(group => new ChartNhansu1
            {
                sort = $"{group.Key.year:D4}-{group.Key.month:D2}",
                label = group.Key.month + "/" + group.Key.year,
                data = (double)group.Sum(d => d.thuclanh),
                data_nv = (double)group.Count()
            }).OrderBy(d => d.sort).ToList();

            var biendong = new
            {
                labels = data.Select(d => d.label).ToList(),
                datasets = new List<ChartNhansu>() {
                    new ChartNhansu { type="bar",label = "Tổng NV", data = data.Select(d => d.data_nv).ToList(),yAxisID="A" }
                }
            };
            if (is_truongbophan)
            {
                biendong.datasets.Add(new ChartNhansu { label = "Tổng lương", data = data.Select(d => d.data).ToList(), type = "line", fill = false, yAxisID = "B" });
            }
            return Json(new
            {
                data = biendong
            }, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }

        public async Task<JsonResult> Highlight()
        {
            var highlight = _context.NewsModel.Where(d => d.deleted_at == null && d.is_publish == true && d.is_highlight == true).OrderByDescending(d => d.created_at).Take(4).Select(d => new NewsModel()
            {
                title = d.title,
                image_url = d.image_url,
                description = d.description,
                id = d.id,
            }).ToList();
            return Json(new { highlight }, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> message()
        {
            var data = _context.HotNewsModel.Where(d => d.id == 1).FirstOrDefault();
            return Json(new { message = data.message, }, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> tin_moi()
        {
            var tin_moi = _context.NewsModel.Where(d => d.deleted_at == null && d.is_publish == true).OrderByDescending(d => d.created_at).Take(7).Select(d => new NewsModel()
            {
                title = d.title,
                image_url = d.image_url,
                description = d.description,
                id = d.id,
            }).ToList();
            return Json(new { tin_moi }, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> cate()
        {
            var cate = _context.CategoryModel.Where(d => d.deleted_at == null).Take(3).ToList();
            foreach (var item in cate)
            {
                item.list_news = _context.NewsCategoryModel.Where(d => d.category_id == item.id).Include(d => d.news).Where(d => d.news.deleted_at == null).Select(d => d.news).Take(10).Select(d => new NewsModel()
                {
                    title = d.title,
                    image_url = d.image_url,
                    description = d.description,
                    id = d.id,
                }).ToList();
            }
            return Json(new { cate }, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
    }
    public class ChartNhansu
    {
        public string label { get; set; }
        public List<double> data { get; set; }
        public bool fill { get; set; }
        public string type { get; set; }
        public string yAxisID { get; set; }
    }
    public class ChartNhansu1
    {
        public string sort { get; set; }
        public string label { get; set; }
        public double data { get; set; }
        public double data_nv { get; set; }
    }
}
