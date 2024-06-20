

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
namespace it_template.Areas.Info.Controllers
{

    public class ApiController : BaseController
    {
        private UserManager<UserModel> UserManager;
        public ApiController(NhansuContext context, AesOperation aes, UserManager<UserModel> UserMgr) : base(context, aes)
        {
            UserManager = UserMgr;
        }
    }


}
