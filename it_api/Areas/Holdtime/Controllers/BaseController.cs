using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authorization;

using Vue.Data;
using Microsoft.EntityFrameworkCore.Infrastructure;
using System.Diagnostics;

namespace it_template.Areas.Holdtime.Controllers
{
    [Area("Holdtime")]
    [Authorize(Roles = "Administrator,Manager Holdtime")]
    //[Authorize]
    public class BaseController : Controller
    {
        protected readonly HoldTimeContext _context;

        public BaseController(HoldTimeContext context)
        {
            _context = context;
            var listener = _context.GetService<DiagnosticSource>();
            (listener as DiagnosticListener).SubscribeWithAdapter(new CommandInterceptor());
        }
    }
}
