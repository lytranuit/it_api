
using Microsoft.AspNetCore.Mvc;
using Vue.Models;
using Vue.Data;
using Microsoft.AspNetCore.Identity;
using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Authorization;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using System.Diagnostics;
using System.Text.Json.Serialization;
using it_api.Services;
using Vue.Services;
using System.IdentityModel.Tokens.Jwt;
using System.Collections;
using System.Drawing;

namespace it_template.Areas.V1.Controllers
{

    [Area("V1")]
    public class AuthController : Controller
    {
        private readonly ItContext _context;
        private readonly NhansuContext _nhansucontext;
        private readonly UserManager<UserModel> UserManager;

        private readonly IConfiguration _configuration;
        private readonly SignInManager<UserModel> _signInManager;
        private readonly LoginMailPyme _LoginMailPyme;
        private AuthManager _authManager;

        public AuthController(ItContext context, NhansuContext nhansucontext, UserManager<UserModel> UserMgr, SignInManager<UserModel> signInManager, IConfiguration configuration, LoginMailPyme LoginMailPyme, AuthManager auth)
        {
            _context = context;
            _configuration = configuration;
            UserManager = UserMgr;
            _LoginMailPyme = LoginMailPyme;
            _signInManager = signInManager;
            _authManager = auth;
            _nhansucontext = nhansucontext;

            var listener = _context.GetService<DiagnosticSource>();
            (listener as DiagnosticListener).SubscribeWithAdapter(new CommandInterceptor());
        }
        public class InputModel
        {
            /// <summary>
            ///     This API supports the ASP.NET Core Identity default UI infrastructure and is not intended to be used
            ///     directly from your code. This API may change or be removed in future releases.
            /// </summary>
            [Required]
            [DataType(DataType.Password)]
            [Display(Name = "Current password")]
            public string? oldpassword { get; set; }

            /// <summary>
            ///     This API supports the ASP.NET Core Identity default UI infrastructure and is not intended to be used
            ///     directly from your code. This API may change or be removed in future releases.
            /// </summary>
            [Required]
            [StringLength(100, ErrorMessage = "The {0} must be at least {2} and at max {1} characters long.", MinimumLength = 6)]
            [DataType(DataType.Password)]
            [Display(Name = "New password")]
            public string? newpassword { get; set; }

            /// <summary>
            ///     This API supports the ASP.NET Core Identity default UI infrastructure and is not intended to be used
            ///     directly from your code. This API may change or be removed in future releases.
            /// </summary>
            [DataType(DataType.Password)]
            [Display(Name = "Confirm new password")]
            [Compare("newpassword", ErrorMessage = "The new password and confirmation password do not match.")]
            public string? confirm { get; set; }
        }
        [HttpPost]
        public async Task<JsonResult> login(string email, string password)
        {
            var pass = password;
            var user = _context.UserModel.Where(x => x.Email.ToLower() == email.ToLower() && x.deleted_at == null).FirstOrDefault();
            LoginResponse responseJson = new LoginResponse()
            {
                authed = false
            };
            if (user == null)
            {
                /// Audittrail
                var audit = new AuditTrailsModel();
                audit.Type = AuditType.LoginFailed.ToString();
                audit.DateTime = DateTime.Now;
                audit.description = $"Tài khoản {email} đăng nhập thất bại.";
                _context.Add(audit);
                await _context.SaveChangesAsync();
                return Json(new LoginResponse { authed = false, error = "Tài khoản hoặc mật khẩu không đúng" });
            }
            var is_pyme = _LoginMailPyme.is_pyme(email);
            if (is_pyme)
            {
                responseJson = await _LoginMailPyme.login(email, password);
                if (responseJson.authed == true)
                {
                    user.last_login = DateTime.Now;
                    user.AccessFailedCount = 0;
                    _context.Update(user);
                    //await _signInManager.SignInAsync(user, true);


                    /// Audittrail
                    var audit = new AuditTrailsModel();
                    audit.UserId = user.Id;
                    audit.Type = AuditType.Login.ToString();
                    audit.DateTime = DateTime.Now;
                    audit.description = $"Tài khoản {email} đăng nhập thành công";
                    _context.Add(audit);
                    await _context.SaveChangesAsync();


                    responseJson.authed = true;
                    responseJson.token = await _authManager.CreateToken(user);
                    var roles = await UserManager.GetRolesAsync(user);
                    var person = _nhansucontext.PersonnelModel.Where(d => d.EMAIL.ToLower() == user.Email.ToLower()).FirstOrDefault();
                    string? report_for = null;
                    if (person != null)
                    {
                        var report_for_person = _nhansucontext.PersonnelModel.Where(d => d.id == person.MAQUANLYTRUCTIEP).FirstOrDefault();
                        if (report_for_person != null)
                        {
                            var report_for_user = _context.UserModel.Where(d => d.Email.ToLower() == report_for_person.EMAIL.ToLower()).FirstOrDefault();
                            report_for = report_for_user != null ? report_for_user.Id : null;
                        }
                    }
                    responseJson.user = new UserInfo()
                    {
                        roles = roles,
                        id = user.Id,
                        email = user.Email,
                        FullName = user.FullName,
                        image_url = user.image_url,
                        image_sign = user.image_sign,
                        report_for = report_for,
                        key_private = _configuration["Key_Access"]
                    };
                    ////
                    Response.Cookies.Append(
                        _configuration["JWT:NameCookieAuth"],
                        responseJson.token,
                        new CookieOptions()
                        {
                            Domain = _configuration["JWT:Domain"],
                            Expires = DateTime.Now.AddHours(Int64.Parse(_configuration["JWT:Expire"]))
                        }
                    );
                    return Json(responseJson);
                }
            }
            var result = await _signInManager.PasswordSignInAsync(email, password, true, lockoutOnFailure: true);
            if (result.Succeeded)
            {
                /// Audittrail
                user.last_login = DateTime.Now;
                user.AccessFailedCount = 0;
                _context.Update(user);

                var audit = new AuditTrailsModel();
                audit.UserId = user.Id;
                audit.Type = AuditType.Login.ToString();
                audit.DateTime = DateTime.Now;
                audit.description = $"Tài khoản {email} đăng nhập thành công";
                _context.Add(audit);
                await _context.SaveChangesAsync();


                responseJson.authed = true;
                responseJson.token = await _authManager.CreateToken(user);
                var roles = await UserManager.GetRolesAsync(user);
                var person = _nhansucontext.PersonnelModel.Where(d => d.EMAIL.ToLower() == user.Email.ToLower()).FirstOrDefault();
                string? report_for = null;
                if (person != null)
                {
                    var report_for_person = _nhansucontext.PersonnelModel.Where(d => d.id == person.id).FirstOrDefault();
                    if (report_for_person != null)
                    {
                        var report_for_user = _context.UserModel.Where(d => d.Email.ToLower() == report_for_person.EMAIL.ToLower()).FirstOrDefault();
                        report_for = report_for_user != null ? report_for_user.Id : null;
                    }
                }
                responseJson.user = new UserInfo()
                {
                    roles = roles,
                    id = user.Id,
                    email = user.Email,
                    FullName = user.FullName,
                    image_url = user.image_url,
                    image_sign = user.image_sign,
                    report_for = report_for,
                    key_private = _configuration["Key_Access"]
                };
                ////
                Response.Cookies.Append(
                    _configuration["JWT:NameCookieAuth"],
                    responseJson.token,
                    new CookieOptions()
                    {
                        Domain = _configuration["JWT:Domain"],
                        Expires = DateTime.Now.AddHours(Int64.Parse(_configuration["JWT:Expire"]))
                    }
                );
            }
            else if (result.IsLockedOut)
            {
                /// Audittrail
                var audit = new AuditTrailsModel();
                audit.Type = AuditType.Lockout.ToString();
                audit.DateTime = DateTime.Now;
                audit.description = $"Tài khoản {email} đã bị khóa trong 5 phút.";
                _context.Add(audit);
                await _context.SaveChangesAsync();
                responseJson.authed = false;
                responseJson.error = $"Tài khoản đã bị khóa trong 5 phút.";
            }
            else
            {
                /// Audittrail
                var audit = new AuditTrailsModel();
                audit.Type = AuditType.LoginFailed.ToString();
                audit.DateTime = DateTime.Now;
                audit.description = $"Tài khoản {email} đăng nhập thất bại.";
                _context.Add(audit);
                await _context.SaveChangesAsync();
                responseJson.authed = false;
                responseJson.error = $"Tài khoản hoặc mật khẩu không đúng";
            }

            return Json(responseJson);


        }
        [HttpPost]
        public async Task<JsonResult> Logout()
        {
            /// Audittrail
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            var user = await UserManager.GetUserAsync(currentUser); // Get user
            if (user != null)
            {
                var audit = new AuditTrailsModel();
                audit.UserId = user.Id;
                audit.Type = AuditType.Logout.ToString();
                audit.DateTime = DateTime.Now;
                audit.description = $"Tài khoản {user.FullName} đã đăng xuất";
                _context.Add(audit);
                await _context.SaveChangesAsync();

            }
            await _signInManager.SignOutAsync();
            ////Remove Cookie
            Response.Cookies.Delete(_configuration["JWT:NameCookieAuth"], new CookieOptions()
            {
                Domain = _configuration["JWT:Domain"]
            });
            return Json(new { success = true });
        }
        [HttpPost]
        [Authorize]
        public async Task<JsonResult> ChangePassword(InputModel input)
        {
            System.Security.Claims.ClaimsPrincipal currentUser = this.User;
            string id = UserManager.GetUserId(currentUser); // Get user id:

            var user = await UserManager.GetUserAsync(currentUser);
            if (user == null)
            {
                return Json(new { success = false, message = $"Unable to load user with ID '{UserManager.GetUserId(User)}'." });
            }

            var changePasswordResult = await UserManager.ChangePasswordAsync(user, input.oldpassword, input.newpassword);
            if (!changePasswordResult.Succeeded)
            {
                var ErrorMessage = "";
                foreach (var error in changePasswordResult.Errors)
                {
                    ErrorMessage += error.Description + "<br>";
                }
                return Json(new { success = false, message = ErrorMessage });
            }

            user.AccessFailedCount = 0;
            user.LockoutEnd = null;
            await UserManager.UpdateAsync(user);

            /// Audittrail
            var audit = new AuditTrailsModel();
            audit.UserId = user.Id;
            audit.Type = AuditType.ChangePassword.ToString();
            audit.DateTime = DateTime.Now;
            audit.description = $"Tài khoản {user.FullName} đã đổi mật khẩu";
            _context.Add(audit);
            await _context.SaveChangesAsync();

            var StatusMessage = "Mật khẩu đã được thay đổi";

            return Json(new { success = true, message = StatusMessage });
        }

        //[Authorize]
        public async Task<JsonResult> TokenInfo(string token)
        {
            JwtSecurityToken? jwt;
            if (token != null && _authManager.ValidateToken(token, out jwt))
            {
                return Json(new { success = true, });
            }
            else
            {
                return Json(new { success = false });
            }
            //var user = _context.UserModel.Where(d => d.deleted_at == null && d.Email.ToLower() == find.email.ToLower()).Include(d => d.list_users).ThenInclude(d => d.userManager).FirstOrDefault();
            //if (user != null)
            //{
            //    var roles = await UserManager.GetRolesAsync(user);
            //    var is_sign = true;
            //    if (user.image_sign == "/private/images/tick.png")
            //    {
            //        is_sign = false;
            //    }
            //    return Json(new
            //    {
            //        success = true,
            //        roles = roles,
            //        email = user.Email,
            //        FullName = user.FullName,
            //        image_url = user.image_url,
            //        is_sign = is_sign,
            //        image_sign = user.image_sign,
            //        list_users = user.list_users,
            //        id = user.Id,
            //        token = token,
            //        vaild_to = find.vaild_to.Value.ToString("yyyy-MM-dd HH:mm:ss")
            //    }); ;
            //}



        }
        public JsonResult Index()
        {
            return Json("auth");

        }

        [Authorize]
        public async Task<JsonResult> Users()
        {
            var users = _context.UserModel.Where(d => d.deleted_at == null).Select(d => new
            {
                id = d.Id,
                name = $"{d.FullName}<{d.Email}>",
                email = d.Email,
                fullName = d.FullName,
            }).ToList();
            return Json(users, new System.Text.Json.JsonSerializerOptions()
            {
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
            });
        }



    }
}
