using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authorization;

using Vue.Data;
using Microsoft.EntityFrameworkCore.Infrastructure;
using System.Diagnostics;
using Vue.Services;

namespace it_template.Areas.Info.Controllers
{
    [Area("Info")]
    //[Authorize(Roles = "Administrator,Manager Holdtime,User")]
    [Authorize]
    public class BaseController : Controller
    {
        private AesOperation _AesOperation;
        protected readonly NhansuContext _context;

        public BaseController(NhansuContext context, AesOperation aes)
        {
            _context = context;
            _AesOperation = aes;
            var listener = _context.GetService<DiagnosticSource>();
            (listener as DiagnosticListener).SubscribeWithAdapter(new CommandInterceptor());
        }
    }
}
