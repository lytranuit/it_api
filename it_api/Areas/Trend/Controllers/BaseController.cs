using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authorization;

using Vue.Data;
using Microsoft.EntityFrameworkCore.Infrastructure;
using System.Diagnostics;

namespace it_template.Areas.Trend.Controllers
{
    [Area("Trend")]
    [Authorize(Roles = "Administrator,Manager Trend")]
    //[Authorize]
    public class BaseController : Controller
    {
        protected readonly ItContext _context;

        public BaseController(ItContext context)
        {
            _context = context;
            var listener = _context.GetService<DiagnosticSource>();
            (listener as DiagnosticListener).SubscribeWithAdapter(new CommandInterceptor());
        }
    }
}
