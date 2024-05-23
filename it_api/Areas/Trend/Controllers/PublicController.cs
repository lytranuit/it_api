
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Identity;
using Vue.Data;
using Vue.Models;
using System.Collections;
using workflow.Models;
using System.Text.Json.Serialization;
using System.Data;
using Microsoft.EntityFrameworkCore.Infrastructure;
using System.Diagnostics;

namespace it_template.Areas.Trend.Controllers
{
    [Area("Trend")]
    public class PublicController : Controller
    {
        protected readonly ItContext _context;
        private UserManager<UserModel> UserManager;
        public PublicController(ItContext context, UserManager<UserModel> UserMgr) : base()
        {
            _context = context;
            var listener = _context.GetService<DiagnosticSource>();
            (listener as DiagnosticListener).SubscribeWithAdapter(new CommandInterceptor());
            UserManager = UserMgr;
        }

        public async Task<JsonResult> updateLimitForResult()
        {
            var results = _context.ResultModel.Where(d => d.deleted_at == null).ToList();
            foreach (var r in results)
            {
                var limits = _context.LimitModel.Where(d => d.target_id == r.target_id).Include(d => d.points).ToList();

                LimitModel? limit;
                if (limits != null && limits.Count() > 0)
                {
                    limit = limits.Where(d => d.list_point.Contains(r.point_id.Value) && d.deleted_at == null && d.date_effect <= r.date).OrderBy(d => d.date_effect).ThenBy(d => d.id).LastOrDefault();
                    r.limit_id = limit != null ? limit.id : null;
                    r.limit_action = limit != null ? limit.action_limit : null;
                    r.limit_alert = limit != null ? limit.alert_limit : null;
                }

            }
            _context.UpdateRange(results);
            _context.SaveChanges();
            return Json(new { success = true });
        }
    }


}
