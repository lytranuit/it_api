

using it_api.Services;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Spire.Xls;
using System.Collections;
using System.Data;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using Vue.Data;
using Vue.Models;
using Vue.Services;
namespace it_template.Areas.Holdtime.Controllers
{

    public class ApiController : BaseController
    {
        private UserManager<UserModel> UserManager;
        public ApiController(HoldTimeContext context, UserManager<UserModel> UserMgr) : base(context)
        {
            UserManager = UserMgr;
        }
        public async Task<JsonResult> targets()
        {
            var All = _context.TargetModel.Where(d => d.deleted_at == null).ToList();
            //var jsonData = new { data = ProcessModel };
            return Json(All, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
        public async Task<JsonResult> stages()
        {
            var All = _context.StageModel.Where(d => d.deleted_at == null).ToList();
            //var jsonData = new { data = ProcessModel };
            return Json(All, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }
    }


}
