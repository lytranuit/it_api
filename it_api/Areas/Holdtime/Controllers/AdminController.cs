using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Collections;
using System.Text.Json.Serialization;
using Vue.Data;
using Vue.Models;

namespace it_template.Areas.Holdtime.Controllers
{
    public class AdminController : BaseController
    {
        private UserManager<UserModel> UserManager;
        public AdminController(HoldTimeContext context, UserManager<UserModel> UserMgr) : base(context)
        {
            UserManager = UserMgr;
        }
        public async Task<JsonResult> HomeBadge()
        {
            var timecheck2 = DateTime.Now.AddDays(-1);
            var tab1 = _context.HoldTimeModel.Where(d => d.deleted_at == null).Count();
            var tab2 = _context.HoldTimeModel.Where(d => d.deleted_at == null && d.date_reality != null).Count();
            var tab3 = _context.HoldTimeModel.Where(d => d.deleted_at == null && d.date_reality != null && d.is_pass == true).Count();
            var tab4 = _context.HoldTimeModel.Where(d => d.deleted_at == null && d.date_reality == null && d.date_theory < timecheck2).Count();
            //var jsonData = new { data = ProcessModel };
            return Json(new { tab1 = tab1, tab2 = tab2, tab3 = tab3, tab4 = tab4 }, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                ReferenceHandler = ReferenceHandler.IgnoreCycles,
            });
        }
        [HttpPost]
        public async Task<JsonResult> Table1(List<DateTime> dates)
        {
            var draw = Request.Form["draw"].FirstOrDefault();
            var start = Request.Form["start"].FirstOrDefault();
            var length = Request.Form["length"].FirstOrDefault();
            int pageSize = length != null ? Convert.ToInt32(length) : 0;
            var name = Request.Form["filters[name]"].FirstOrDefault();
            var id_text = Request.Form["filters[id]"].FirstOrDefault();
            int id = id_text != null ? Convert.ToInt32(id_text) : 0;
            var type = Request.Form["type"].FirstOrDefault();
            //var tenhh = Request.Form["filters[tenhh]"].FirstOrDefault();
            int skip = start != null ? Convert.ToInt32(start) : 0;
            var customerData = _context.HoldTimeModel.Where(d => d.deleted_at == null && d.date_reality == null);

            int recordsTotal = customerData.Count();

            if (name != null && name != "")
            {
                customerData = customerData.Where(d => d.name.Contains(name));
            }
            if (id != 0)
            {
                customerData = customerData.Where(d => d.id == id);
            }
            if (dates != null)
            {
                if (dates[0] != null && dates[0].Kind == DateTimeKind.Utc)
                {
                    dates[0] = dates[0].ToLocalTime();
                }
                if (dates[1] != null && dates[1].Kind == DateTimeKind.Utc)
                {
                    dates[1] = dates[1].ToLocalTime();
                }
                customerData = customerData.Where(d => d.date_theory >= dates[0] && d.date_theory <= dates[1]);
            }
            int recordsFiltered = customerData.Count();
            var datapost = customerData.OrderByDescending(d => d.id).Skip(skip).Take(pageSize)
                .Include(d => d.Hold)
                .ThenInclude(d => d.stage)
                .OrderByDescending(d => d.date_theory)
                .ToList();
            var data = new ArrayList();
            foreach (var record in datapost)
            {
                var data1 = new
                {
                    hold_id = record.hold_id,
                    tensp = record.Hold.tensp,
                    masp = record.Hold.masp,
                    malo = record.Hold.malo,
                    stage = record.Hold.stage.name,
                    name = record.name,
                    date_theory = record.date_theory.Value.ToString("yyyy-MM-dd"),
                };
                data.Add(data1);
            }
            var jsonData = new { draw = draw, recordsFiltered = recordsFiltered, recordsTotal = recordsTotal, data = data };
            return Json(jsonData, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                ReferenceHandler = ReferenceHandler.IgnoreCycles,
            });
        }
        [HttpPost]
        public async Task<JsonResult> Table2()
        {
            var draw = Request.Form["draw"].FirstOrDefault();
            var start = Request.Form["start"].FirstOrDefault();
            var length = Request.Form["length"].FirstOrDefault();
            int pageSize = length != null ? Convert.ToInt32(length) : 0;
            var name = Request.Form["filters[name]"].FirstOrDefault();
            var id_text = Request.Form["filters[id]"].FirstOrDefault();
            int id = id_text != null ? Convert.ToInt32(id_text) : 0;
            var type = Request.Form["type"].FirstOrDefault();
            //var tenhh = Request.Form["filters[tenhh]"].FirstOrDefault();
            int skip = start != null ? Convert.ToInt32(start) : 0;
            var timecheck2 = DateTime.Now.AddDays(-1);
            var customerData = _context.HoldTimeModel.Where(d => d.deleted_at == null && d.date_reality == null && d.date_theory < timecheck2);

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
            var datapost = customerData.OrderByDescending(d => d.id).Skip(skip).Take(pageSize)
                .Include(d => d.Hold)
                .ThenInclude(d => d.stage)
                .OrderByDescending(d => d.date_theory)
                .ToList();
            var data = new ArrayList();
            foreach (var record in datapost)
            {
                var data1 = new
                {
                    hold_id = record.hold_id,
                    tensp = record.Hold.tensp,
                    masp = record.Hold.masp,
                    malo = record.Hold.malo,
                    stage = record.Hold.stage.name,
                    name = record.name,
                    date_theory = record.date_theory.Value.ToString("yyyy-MM-dd"),
                };
                data.Add(data1);
            }
            var jsonData = new { draw = draw, recordsFiltered = recordsFiltered, recordsTotal = recordsTotal, data = data };
            return Json(jsonData, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                ReferenceHandler = ReferenceHandler.IgnoreCycles,
            });
        }
    }

}
