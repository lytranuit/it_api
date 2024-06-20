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
    }

}
