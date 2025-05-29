using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authorization;

using Vue.Data;
using Microsoft.EntityFrameworkCore.Infrastructure;
using System.Diagnostics;
using Vue.Services;

namespace it_template.Areas.Info.Controllers
{
    [Area("Info")]
    //[Authorize(Roles = "Administrator,Manager HR,HR,User")]
    //[Authorize]
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
        protected void CopyValues<T>(T target, T source)
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
